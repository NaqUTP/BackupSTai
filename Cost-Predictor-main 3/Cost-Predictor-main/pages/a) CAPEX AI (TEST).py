# ============================
# ✅ ADD THIS: MONTE CARLO PACK
# ============================
# Paste into your script in 2 places:
# (A) HELPERS: paste after `knn_impute_numeric(...)`
# (B) UI: paste inside Predict tab, AFTER the "Run Prediction" button block
#         and BEFORE the "Batch (Excel)" uploader block (recommended)

# ============================================================
# (A) HELPERS — paste after knn_impute_numeric(...)
# ============================================================

def _fmt_money(v: float, curr: str) -> str:
    try:
        return f"{curr} {float(v):,.2f}".strip()
    except Exception:
        return str(v)

def _fmt_pct(v: float) -> str:
    try:
        return f"{float(v):.2f}%"
    except Exception:
        return str(v)

def percentile(x: np.ndarray, p: float) -> float:
    return float(np.percentile(x, p))

def sample_feature(series: pd.Series, center: float | None, mode: str, pct: float, rng: np.random.Generator):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return np.nan

    # if user did not provide center -> sample from historical distribution
    if center is None or (isinstance(center, float) and np.isnan(center)):
        return float(rng.choice(s.values))

    center = float(center)
    if mode == "Normal":
        std = abs(center) * (pct / 100.0)
        if std == 0:
            std = float(s.std(ddof=0) or 0.0)
        if std == 0:
            return center
        return float(rng.normal(loc=center, scale=std))
    else:
        half = abs(center) * (pct / 100.0)
        if half == 0:
            q1, q3 = float(s.quantile(0.25)), float(s.quantile(0.75))
            half = (q3 - q1) / 2.0
            if half == 0:
                return center
        return float(rng.uniform(center - half, center + half))

def sample_percentage(center: float, mode: str, pct: float, rng: np.random.Generator) -> float:
    center = float(center)
    if mode == "Normal":
        std = abs(center) * (pct / 100.0)
        v = float(rng.normal(center, std))
    else:
        half = abs(center) * (pct / 100.0)
        v = float(rng.uniform(center - half, center + half))
    return float(np.clip(v, 0.0, 100.0))

def run_monte_carlo(
    df_dataset: pd.DataFrame,
    feature_cols: list[str],
    pipe: Pipeline,
    base_payload: dict,
    n: int,
    seed: int,
    feat_mode: str,
    feat_pct: float,
    vary_financial: bool,
    fin_mode: str,
    fin_pct: float,
    eprr_base: dict,
    sst_pct_base: float,
    owners_pct_base: float,
    cont_pct_base: float,
    esc_pct_base: float,
):
    """
    Returns a DataFrame of trials. ✅ Includes sampled feature columns (needed for tornado + bucket summaries).
    """
    rng = np.random.default_rng(int(seed))
    df_num = df_dataset.select_dtypes(include=[np.number]).copy()

    feature_series = {c: df_num[c] for c in feature_cols if c in df_num.columns}

    rows = []
    for _ in range(int(n)):
        sampled = {}

        # sample each feature
        for c in feature_cols:
            v = base_payload.get(c, np.nan)
            if isinstance(v, str) and v.strip() == "":
                v = np.nan
            try:
                v = float(v)
            except Exception:
                v = np.nan

            sampled[c] = sample_feature(feature_series.get(c, pd.Series(dtype=float)), v, feat_mode, feat_pct, rng)

        # model prediction
        base_pred = float(pipe.predict(pd.DataFrame([sampled], columns=feature_cols))[0])

        # financials
        if vary_financial:
            sst_pct = sample_percentage(sst_pct_base, fin_mode, fin_pct, rng)
            owners_pct = sample_percentage(owners_pct_base, fin_mode, fin_pct, rng)
            cont_pct = sample_percentage(cont_pct_base, fin_mode, fin_pct, rng)
            esc_pct = sample_percentage(esc_pct_base, fin_mode, fin_pct, rng)

            eprr_trial = {k: max(0.0, sample_percentage(float(v), fin_mode, fin_pct, rng)) for k, v in eprr_base.items()}
            eprr_trial, _ = normalize_to_100(eprr_trial)
        else:
            sst_pct, owners_pct, cont_pct, esc_pct = sst_pct_base, owners_pct_base, cont_pct_base, esc_pct_base
            eprr_trial = dict(eprr_base)

        owners_cost, sst_cost, contingency_cost, escalation_cost, eprr_costs, grand_total = cost_breakdown(
            base_pred, eprr_trial, sst_pct, owners_pct, cont_pct, esc_pct
        )

        row = {
            **sampled,  # ✅ critical for tornado + bucket feature summaries
            "Base Pred": base_pred,
            "Grand Total": grand_total,
            "Owners Cost": owners_cost,
            "SST Cost": sst_cost,
            "Contingency Cost": contingency_cost,
            "Escalation Cost": escalation_cost,
            "SST %": sst_pct,
            "Owners %": owners_pct,
            "Cont %": cont_pct,
            "Esc %": esc_pct,
            **{f"EPRR_{k}_Cost": v for k, v in eprr_costs.items()},
        }
        rows.append(row)

    return pd.DataFrame(rows)

def make_scenario_buckets(mc_df: pd.DataFrame, col: str = "Grand Total", scheme: str = "3"):
    x = pd.to_numeric(mc_df[col], errors="coerce").dropna().to_numpy()
    if x.size == 0:
        out = mc_df.copy()
        out["Scenario Bucket"] = np.nan
        return out, pd.DataFrame(), {}

    if scheme == "5":
        p10 = percentile(x, 10); p30 = percentile(x, 30); p70 = percentile(x, 70); p90 = percentile(x, 90)
        bins = [-np.inf, p10, p30, p70, p90, np.inf]
        labels = ["Very Low (≤P10)", "Low (P10–P30)", "Base (P30–P70)", "High (P70–P90)", "Very High (≥P90)"]
        cuts = {"P10": p10, "P30": p30, "P70": p70, "P90": p90}
    else:
        p10 = percentile(x, 10); p90 = percentile(x, 90)
        bins = [-np.inf, p10, p90, np.inf]
        labels = ["Low (≤P10)", "Base (P10–P90)", "High (≥P90)"]
        cuts = {"P10": p10, "P90": p90}

    out = mc_df.copy()
    out["Scenario Bucket"] = pd.cut(out[col], bins=bins, labels=labels, include_lowest=True)

    g = out.groupby("Scenario Bucket", observed=True)[col]
    summary = pd.DataFrame(
        {
            "Scenario Bucket": labels,
            "Probability": (out["Scenario Bucket"].value_counts(normalize=True).reindex(labels).fillna(0.0) * 100.0).round(2),
            "Min": g.min().reindex(labels),
            "Mean": g.mean().reindex(labels),
            "Median": g.median().reindex(labels),
            "Max": g.max().reindex(labels),
        }
    )
    return out, summary, cuts

def probability_exceedance(x: np.ndarray, threshold: float) -> float:
    x = pd.to_numeric(pd.Series(x), errors="coerce").dropna().to_numpy()
    if x.size == 0:
        return float("nan")
    return float((x > threshold).mean() * 100.0)

def exceedance_by_bucket(bucketed_df: pd.DataFrame, threshold: float, col: str = "Grand Total"):
    out = []
    for b, sub in bucketed_df.groupby("Scenario Bucket", observed=True):
        arr = pd.to_numeric(sub[col], errors="coerce").dropna().to_numpy()
        out.append(
            {"Scenario Bucket": b, "P(>Budget)%": float((arr > threshold).mean() * 100.0) if arr.size else np.nan, "n": int(arr.size)}
        )
    df = pd.DataFrame(out)
    if not df.empty:
        df["P(>Budget)%"] = df["P(>Budget)%"].round(2)
    return df

def tornado_drivers_from_samples(samples_df: pd.DataFrame, feature_cols: list[str], y_col: str = "Grand Total", top_k: int = 12):
    present = [c for c in feature_cols if c in samples_df.columns]
    if not present:
        return pd.DataFrame()

    y = pd.to_numeric(samples_df[y_col], errors="coerce")
    rows = []
    for c in present:
        x = pd.to_numeric(samples_df[c], errors="coerce")
        m = x.notna() & y.notna()
        if m.sum() < 10:
            continue
        corr = float(np.corrcoef(x[m].to_numpy(), y[m].to_numpy())[0, 1])
        if np.isnan(corr):
            continue
        rows.append({"Feature": c, "Corr": corr, "AbsCorr": abs(corr)})

    return pd.DataFrame(rows).sort_values("AbsCorr", ascending=False).head(int(top_k))

def bucket_feature_summary(bucketed_df: pd.DataFrame, feature_cols: list[str], col_bucket: str = "Scenario Bucket"):
    present = [c for c in feature_cols if c in bucketed_df.columns]
    if not present or col_bucket not in bucketed_df.columns:
        return pd.DataFrame()

    grp = bucketed_df.groupby(col_bucket, observed=True)[present]
    means = grp.mean(numeric_only=True)
    stds = grp.std(numeric_only=True, ddof=0)

    out = []
    for b in means.index:
        for c in present:
            out.append(
                {
                    "Scenario Bucket": b,
                    "Feature": c,
                    "Mean": float(means.loc[b, c]) if pd.notnull(means.loc[b, c]) else np.nan,
                    "Std": float(stds.loc[b, c]) if pd.notnull(stds.loc[b, c]) else np.nan,
                }
            )
    return pd.DataFrame(out)

def auto_narrative(currency: str, p10: float, p50: float, p90: float, mean_gt: float, std_gt: float,
                   budget: float | None, prob_exceed: float | None, bucket_summary: pd.DataFrame, tornado_df: pd.DataFrame):
    lines = []
    lines.append(
        f"**Cost risk summary:** P10={_fmt_money(p10, currency)}, P50={_fmt_money(p50, currency)}, P90={_fmt_money(p90, currency)}."
    )
    lines.append(f"Mean={_fmt_money(mean_gt, currency)} with σ={_fmt_money(std_gt, currency)}.")
    if budget is not None and prob_exceed is not None and not np.isnan(prob_exceed):
        lines.append(f"**Budget check:** P(Grand Total > Budget={_fmt_money(budget, currency)}) = **{prob_exceed:.2f}%**.")
    if tornado_df is not None and not tornado_df.empty:
        top = tornado_df.head(3)["Feature"].tolist()
        lines.append(f"**Top drivers:** {', '.join(top)}.")
    if bucket_summary is not None and not bucket_summary.empty:
        idx = bucket_summary["Probability"].astype(float).idxmax()
        lines.append(
            f"**Most likely scenario bucket:** {bucket_summary.loc[idx,'Scenario Bucket']} ({float(bucket_summary.loc[idx,'Probability']):.2f}% of trials)."
        )
    return "  \n".join(lines)


# ============================================================
# (B) UI BLOCK — paste inside Predict tab
# ============================================================
# Put this right AFTER your "Run Prediction" button block,
# and BEFORE the "Batch (Excel)" uploader.
#
# It uses existing variables from your code:
# - df_pred, feat_cols, payload, pipe, currency_pred
# - eprr, sst_pct, owners_pct, cont_pct, esc_pct
# ============================================================

st.markdown("---")
st.markdown('<h4 style="margin:0;color:#000;">🎲 Monte Carlo Simulation</h4><p>Uncertainty bands • buckets • drivers • budget risk</p>', unsafe_allow_html=True)

mc_r1, mc_r2, mc_r3, mc_r4 = st.columns([1, 1, 1, 2])
with mc_r1:
    mc_n = st.number_input("Trials", min_value=200, max_value=30000, value=3000, step=200, key=f"mc_trials__{ds_name_pred}")
with mc_r2:
    mc_seed = st.number_input("Seed", min_value=0, max_value=999999, value=42, step=1, key=f"mc_seed__{ds_name_pred}")
with mc_r3:
    feat_mode = st.selectbox("Feature sampling", ["Normal", "Uniform"], index=0, key=f"mc_feat_mode__{ds_name_pred}")
with mc_r4:
    feat_pct = st.slider("Feature uncertainty ±%", 0.0, 60.0, 10.0, 1.0, key=f"mc_feat_pct__{ds_name_pred}")

mc_c1, mc_c2 = st.columns([1, 2])
with mc_c1:
    vary_financial = st.checkbox("Also vary financial % + EPRR", value=False, key=f"mc_vary_fin__{ds_name_pred}")
with mc_c2:
    fin_mode = st.selectbox("Financial sampling", ["Normal", "Uniform"], index=0, disabled=not vary_financial, key=f"mc_fin_mode__{ds_name_pred}")
    fin_pct = st.slider("Financial uncertainty ±%", 0.0, 60.0, 5.0, 0.5, disabled=not vary_financial, key=f"mc_fin_pct__{ds_name_pred}")

bkt_c1, bkt_c2 = st.columns([1, 2])
with bkt_c1:
    bucket_scheme_ui = st.selectbox("Scenario buckets", ["3 (P10/P90)", "5 (P10/P30/P70/P90)"], index=0, key=f"mc_bucket__{ds_name_pred}")
with bkt_c2:
    st.caption("Buckets are percentile-based on Monte Carlo Grand Total.")

bud_c1, bud_c2 = st.columns([1, 2])
with bud_c1:
    budget = st.number_input("Budget threshold (optional)", min_value=0.0, value=0.0, step=1000.0, key=f"mc_budget__{ds_name_pred}")
with bud_c2:
    st.caption("If > 0, we compute probability Grand Total exceeds this budget (overall and per bucket).")

run_mc = st.button("Run Monte Carlo", key=f"run_mc__{ds_name_pred}__{target_col_pred}")

if run_mc:
    with st.spinner("Running Monte Carlo..."):
        mc_df = run_monte_carlo(
            df_dataset=df_pred,
            feature_cols=feat_cols,
            pipe=pipe,
            base_payload=payload,
            n=int(mc_n),
            seed=int(mc_seed),
            feat_mode=feat_mode,
            feat_pct=float(feat_pct),
            vary_financial=bool(vary_financial),
            fin_mode=fin_mode if vary_financial else "Normal",
            fin_pct=float(fin_pct) if vary_financial else 0.0,
            eprr_base=eprr,
            sst_pct_base=sst_pct,
            owners_pct_base=owners_pct,
            cont_pct_base=cont_pct,
            esc_pct_base=esc_pct,
        )

    # --- core stats
    gt_arr = pd.to_numeric(mc_df["Grand Total"], errors="coerce").dropna().to_numpy()
    p10, p50, p90 = percentile(gt_arr, 10), percentile(gt_arr, 50), percentile(gt_arr, 90)
    mean_gt = float(np.mean(gt_arr))
    std_gt = float(np.std(gt_arr, ddof=0))

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("P10", _fmt_money(p10, currency_pred))
    s2.metric("P50", _fmt_money(p50, currency_pred))
    s3.metric("P90", _fmt_money(p90, currency_pred))
    s4.metric("Mean ± Std", f"{_fmt_money(mean_gt, currency_pred)} ± {_fmt_money(std_gt, currency_pred)}")

    # --- distribution
    st.plotly_chart(px.histogram(mc_df, x="Grand Total", nbins=50, title="Grand Total Distribution"), use_container_width=True)

    sorted_gt = np.sort(gt_arr)
    cdf = np.linspace(0, 1, len(sorted_gt))
    fig_cdf = go.Figure()
    fig_cdf.add_trace(go.Scatter(x=sorted_gt, y=cdf, mode="lines", name="CDF"))
    fig_cdf.update_layout(title="Grand Total CDF", xaxis_title="Grand Total", yaxis_title="Probability")
    st.plotly_chart(fig_cdf, use_container_width=True)

    # --- buckets
    scheme = "5" if bucket_scheme_ui.startswith("5") else "3"
    mc_bucketed, bucket_summary, cuts = make_scenario_buckets(mc_df, col="Grand Total", scheme=scheme)
    cut_txt = ", ".join([f"{k}={_fmt_money(v, currency_pred)}" for k, v in cuts.items()])
    st.caption(f"Cutpoints: {cut_txt if cut_txt else '—'}")

    if not bucket_summary.empty:
        disp = bucket_summary.copy()
        for c in ["Min", "Mean", "Median", "Max"]:
            disp[c] = disp[c].apply(lambda v: _fmt_money(v, currency_pred) if pd.notnull(v) else "—")
        disp["Probability"] = disp["Probability"].apply(_fmt_pct)
        st.dataframe(disp, use_container_width=True)

        fig_prob = px.bar(bucket_summary, x="Scenario Bucket", y="Probability", title="Scenario Probability (%)", text="Probability")
        fig_prob.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        st.plotly_chart(fig_prob, use_container_width=True)

    # --- example scenario per bucket (closest to bucket median)
    st.markdown("#### Example scenario per bucket (closest to bucket median)")
    examples = []
    if "Scenario Bucket" in mc_bucketed.columns:
        for b in mc_bucketed["Scenario Bucket"].dropna().unique():
            sub = mc_bucketed[mc_bucketed["Scenario Bucket"] == b]
            if sub.empty:
                continue
            med = float(pd.to_numeric(sub["Grand Total"], errors="coerce").median())
            sub2 = sub.assign(_dist=(pd.to_numeric(sub["Grand Total"], errors="coerce") - med).abs()).sort_values("_dist").head(1)
            examples.append(sub2.drop(columns=["_dist"]))
    if examples:
        ex_df = pd.concat(examples, ignore_index=True)
        st.dataframe(ex_df, use_container_width=True, height=240)
    else:
        st.caption("No examples available.")

    # --- budget exceedance
    budget_val = float(budget) if budget and float(budget) > 0 else None
    prob_overall = probability_exceedance(gt_arr, budget_val) if budget_val is not None else None

    st.markdown("#### Budget exceedance")
    if budget_val is not None:
        st.metric("Overall P(>Budget)", f"{prob_overall:.2f}%")
        by_bucket = exceedance_by_bucket(mc_bucketed, budget_val, col="Grand Total")
        st.dataframe(by_bucket, use_container_width=True)
        st.plotly_chart(px.bar(by_bucket, x="Scenario Bucket", y="P(>Budget)%", title="Exceedance Probability by Bucket"), use_container_width=True)
    else:
        st.caption("Set a budget > 0 to enable exceedance calculations.")

    # --- tornado drivers + bucket feature summaries
    st.markdown("#### Tornado drivers (top features driving Grand Total)")
    tornado_df = tornado_drivers_from_samples(mc_bucketed, feature_cols=feat_cols, y_col="Grand Total", top_k=12)

    if tornado_df is not None and not tornado_df.empty:
        st.dataframe(tornado_df, use_container_width=True)
        fig_tor = px.bar(tornado_df.sort_values("AbsCorr", ascending=True), x="AbsCorr", y="Feature", orientation="h",
                         title="Top drivers by |corr(feature, Grand Total)|")
        st.plotly_chart(fig_tor, use_container_width=True)
    else:
        st.caption("Tornado unavailable (not enough valid samples or features).")

    st.markdown("#### Bucket feature summaries (mean ± std)")
    feat_sum = bucket_feature_summary(mc_bucketed, feature_cols=feat_cols, col_bucket="Scenario Bucket")
    if not feat_sum.empty:
        # show top 8 by tornado (if available), else first 8 features
        show_feats = tornado_df["Feature"].tolist()[:8] if (tornado_df is not None and not tornado_df.empty) else feat_cols[:8]
        fs = feat_sum[feat_sum["Feature"].isin(show_feats)].copy()
        fs["Mean"] = fs["Mean"].apply(lambda v: f"{v:,.4g}" if pd.notnull(v) else "—")
        fs["Std"] = fs["Std"].apply(lambda v: f"{v:,.4g}" if pd.notnull(v) else "—")
        st.dataframe(fs, use_container_width=True, height=350)
    else:
        st.caption("Feature summaries unavailable.")

    # --- auto narrative
    st.markdown("#### Auto narrative")
    st.markdown(
        auto_narrative(
            currency=currency_pred,
            p10=p10,
            p50=p50,
            p90=p90,
            mean_gt=mean_gt,
            std_gt=std_gt,
            budget=budget_val,
            prob_exceed=prob_overall,
            bucket_summary=bucket_summary,
            tornado_df=tornado_df,
        )
    )

    # --- download
    st.markdown("#### Download")
    mc_xlsx = io.BytesIO()
    mc_bucketed.to_excel(mc_xlsx, index=False, engine="openpyxl")
    mc_xlsx.seek(0)
    st.download_button(
        "⬇️ Download Monte Carlo Results (Excel)",
        data=mc_xlsx.getvalue(),
        file_name=f"{ds_name_pred}_monte_carlo_{scheme}bucket.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Monte Carlo results (table)", expanded=False):
        st.dataframe(mc_bucketed, use_container_width=True, height=380)
