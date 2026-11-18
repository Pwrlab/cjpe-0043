"""Local web interface (Streamlit). Start with:  refilltool gui"""
from __future__ import annotations

import io as _io
import json
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from refilltool import __version__
from refilltool.export import excel_bytes, method_overview, save_all
from refilltool.io import (detect_columns, detect_covariates, detect_pairs, detect_rain, prepare, radiation_qc,
                           read_table)
from refilltool.methods import METHOD_LABELS
from refilltool.pipeline import DEFAULT_CONFIG, config_text, merge_config, run

st.set_page_config(page_title="Refilling Tool v2", page_icon="🌳", layout="wide")

COLORS = {"refilling": "#D55E00", "exp": "#0072B2", "line": "#009E73", "et": "#CC79A7"}
NONE = "（无 none）"


def _example_path():
    here = Path(__file__).resolve()
    cands = [here.parents[i] / "example.xlsx" for i in (3, 2, 1) if len(here.parents) > i]  # repo, bundle
    for p in [*cands, Path.cwd() / "example.xlsx"]:
        if p.exists():
            return p
    return None


@st.cache_data(show_spinner="读取文件 (reading file)…")
def _read(raw: bytes, name: str, sheet):
    return read_table(_io.BytesIO(raw), sheet=sheet, name=name)


@st.cache_data(show_spinner=False)
def _sheets(raw: bytes):
    try:
        return pd.ExcelFile(_io.BytesIO(raw)).sheet_names
    except Exception:  # noqa: BLE001
        return None


def _state(key, default):
    if key not in st.session_state:
        st.session_state[key] = default
    return st.session_state[key]


# ================================================================ sidebar: data
st.sidebar.title("🌳 Refilling Tool v2")
st.sidebar.caption("夜间液流组分划分 (nighttime sap flow partitioning) · PWRlab, BJFU")

st.sidebar.header("1 数据 (data)")
up = st.sidebar.file_uploader("上传 xlsx / csv", type=["xlsx", "xls", "csv", "txt"])
ex = _example_path()
use_example = False
if up is None and ex is not None:
    use_example = st.sidebar.checkbox("使用示例数据 example.xlsx", value=True)

if up is not None:
    raw, fname = up.getvalue(), up.name
elif use_example:
    raw, fname = ex.read_bytes(), ex.name
else:
    st.title("Refilling Tool v2")
    st.info("在左侧上传数据文件。格式：一列时间 (time)、一列太阳辐射 (solar radiation, W m⁻²)、"
            "其余每列一个探针的液流通量密度 (Fd)。可选：降雨列、VPD/温度/土壤水分等环境因子列。"
            "列名任意，程序自动识别，可手动修改。")
    st.stop()

sheet = 0
if fname.lower().endswith((".xlsx", ".xls")):
    names = _sheets(raw)
    if names and len(names) > 1:
        sheet = st.sidebar.selectbox("工作表 (sheet)", names)
df_raw = _read(raw, fname, sheet)

t0, r0, trees0 = detect_columns(df_raw)
rain0, covs0 = detect_rain(df_raw), detect_covariates(df_raw)
cols = [str(c) for c in df_raw.columns]
with st.sidebar.expander("列设置 (columns)", expanded=t0 is None or r0 is None):
    time_col = st.selectbox("时间列 (time)", cols, index=cols.index(t0) if t0 in cols else 0)
    rad_col = st.selectbox("辐射列 (solar radiation)", cols,
                           index=cols.index(r0) if r0 in cols else min(1, len(cols) - 1))
    rest = [c for c in cols if c not in (time_col, rad_col)]
    rain_sel = st.selectbox("降雨列 (rain, 可选)", [NONE, *rest],
                            index=([NONE, *rest].index(rain0) if rain0 in rest else 0))
    rain_col = None if rain_sel == NONE else rain_sel
    rest2 = [c for c in rest if c != rain_col]
    cov_cols = st.multiselect("环境因子 (covariates, 输出夜间均值)", rest2, default=[c for c in covs0 if c in rest2])
    cand = [c for c in rest2 if c not in cov_cols]
    tree_cols = st.multiselect("液流列 (sap flow)", cand, default=[c for c in trees0 if c in cand])

# ================================================================ sidebar: methods & parameters
st.sidebar.header("2 方法 (methods)")
cfg = merge_config(None)
methods = [m for m in METHOD_LABELS if st.sidebar.checkbox(METHOD_LABELS[m], value=True, key=f"use_{m}")]
cfg["methods"] = methods

st.sidebar.header("3 参数 (parameters)")
cfg["legacy"] = st.sidebar.checkbox(
    "v1 兼容模式 (legacy)", False,
    help="完全复现 v1 程序（含 v1 的 Et 算法和夜间划分），只用于与旧结果对照。")

bp = cfg["breakpoints"]
if not cfg["legacy"]:
    with st.sidebar.expander("断点 (breakpoints, 原文 1.4.2 / 1.5)"):
        mode = st.radio("拐点检测 (knee detection)", ["auto", "fixed"], horizontal=True,
                        format_func={"auto": "自动 Kneedle", "fixed": "固定时长"}.get,
                        help="自动：平滑后用 Kneedle 检测 P1（Exp 用）和 P2（Re、Et 用）；检测失败时用下面的固定值。")
        bp["mode"] = mode
        c1, c2 = st.columns(2)
        bp["p1_minutes"] = c1.number_input("P1 (min)", 10, 720, bp["p1_minutes"], 10,
                                           help="Exp 法：液流由快速衰减转为缓慢衰减的时刻（日落后）")
        bp["p2_minutes"] = c2.number_input("P2 (min)", 10, 900, bp["p2_minutes"], 10,
                                           help="Re、Et 法：液流基本停止变化、充水结束的时刻")
        bp["line_lag_minutes"] = c1.number_input("Line 滞后期 lag (min)", 10, 360, bp["line_lag_minutes"], 10,
                                                 help="原文：液流滞后于太阳辐射约 90 min")
        bp["smooth_minutes"] = c2.number_input("平滑窗口 smooth (min)", 0, 180, bp["smooth_minutes"], 10,
                                               help="拐点检测前的滑动平均窗口，消除异常值影响")
        bp["min_minutes"] = st.number_input("自动拐点最早时刻 (min)", 0, 300, bp["min_minutes"], 10,
                                            help="早于该时刻的自动拐点视为无效，改用固定值")
        cfg["et_min_strategy"] = st.checkbox("Et：后半夜最小值作为蒸腾速率（原文）", True,
                                             help="不勾选则用后半夜均值")
else:
    with st.sidebar.expander("v1 参数 (legacy parameters)"):
        for m in methods:
            d = DEFAULT_CONFIG["params"][m]
            c1, c2 = st.columns(2)
            cfg["params"][m]["default_point"] = c1.number_input(f"{m} 断点 (points)", 1, 500, d["default_point"],
                                                                key=f"bp_{m}")
            cfg["params"][m]["kneed_process"] = c2.checkbox("auto knee", d["kneed_process"], key=f"kn_{m}")

with st.sidebar.expander("夜间划分与筛选 (night & filter)"):
    nc = cfg["night"]
    nc["threshold"] = st.number_input("夜间辐射阈值 (W m⁻²)", 0.0, 100.0, 5.0, 0.5)
    nc["min_run_minutes"] = st.number_input("忽略短于此时长的昼夜切换 (min)", 0, 240, 30, 10,
                                            help="夜间辐射噪声或白天短暂阴暗不会切断夜晚。")
    nc["drop_partial"] = st.checkbox("去掉首尾不完整夜晚 (partial nights)", True)
    nc["drop_gap"] = st.checkbox("去掉有缺测的夜晚 (nights with gaps)", True)
    nc["min_day_rad_MJ"] = st.number_input("晴天筛选：当日辐射 ≥ (MJ m⁻²)", 0.0, 40.0, 0.0, 1.0,
                                           help="0 表示不筛选。方法适用于典型晴天。")
    if rain_col:
        nc["max_rain_mm"] = st.number_input("降雨筛选：当日至夜间结束降雨 ≤ (mm)", 0.0, 100.0, 0.0, 0.2,
                                            help="原文只用无降雨日的数据。")

with st.sidebar.expander("单位 (units)：Q = Σ Fd·t·As"):
    cfg["units"]["integrate_time"] = st.checkbox("乘以采样间隔 t（原文公式 1）", True, disabled=cfg["legacy"])
    st.caption("填写边材面积 As (m²) 后结果单位为 mL night⁻¹；不填则为 mL m⁻² night⁻¹。"
               "高度差法必须填写两个高度的 As。")
    area_df = pd.DataFrame({"column": tree_cols,
                            "As_m2": [_state("areas", {}).get(c) for c in tree_cols]})
    area_df = st.data_editor(area_df, hide_index=True, width="stretch", disabled=["column"],
                             column_config={"As_m2": st.column_config.NumberColumn("As (m²)", min_value=0.0,
                                                                                   format="%.5f")},
                             key="area_editor")
    areas = {r.column: float(r.As_m2) for r in area_df.itertuples() if pd.notna(r.As_m2) and r.As_m2 > 0}
    st.session_state["areas"] = areas
    cfg["units"]["sapwood_area_m2"] = areas

with st.sidebar.expander("高度差预测充水法 (height difference, 原文 1.4.3)"):
    st.caption("每棵树两个探针：下 = 胸高 1.3 m，上 = 最低活枝下。A 茎干充水 = Q下 − Q上；"
               "上部液流用所选方法分为 B 冠层充水和 C 夜间蒸腾。同时输出 4 种方法的评估（原文 1.5）。")
    auto_pairs = detect_pairs(tree_cols)
    if st.session_state.get("pairs_src") != (fname, tuple(tree_cols)):  # new file: re-detect pairs
        st.session_state["pairs_src"] = (fname, tuple(tree_cols))
        st.session_state["pairs_df"] = pd.DataFrame(auto_pairs or [], columns=["name", "lower", "upper"]).astype(str)
        st.session_state.pop("pairs_editor", None)
    if auto_pairs:
        st.caption(f"已按列名自动配对 {len(auto_pairs)} 棵树（如 T1_1.3m / T1_7.0m），可修改。")
    pairs_df = st.session_state["pairs_df"]
    pairs_df = st.data_editor(
        pairs_df, num_rows="dynamic", hide_index=True, width="stretch", key="pairs_editor",
        column_config={"name": st.column_config.TextColumn("树 tree"),
                       "lower": st.column_config.SelectboxColumn("下 lower (1.3 m)", options=tree_cols),
                       "upper": st.column_config.SelectboxColumn("上 upper", options=tree_cols)})
    pairs = [{"name": r.name if isinstance(r.name, str) and r.name else r.lower, "lower": r.lower, "upper": r.upper}
             for r in pairs_df.itertuples() if isinstance(r.lower, str) and isinstance(r.upper, str)]
    cfg["height_diff"]["pairs"] = pairs
    cfg["height_diff"]["method"] = st.selectbox("上部探针所用方法", methods or ["line"],
                                                index=(methods.index("line") if "line" in methods else 0),
                                                format_func=METHOD_LABELS.get, help="原文推荐 Line 法")

excluded = _state("excluded", set())
cfg["exclude_nights"] = sorted(excluded)

if not methods or not tree_cols:
    st.warning("请至少选择一种方法和一个液流列。")
    st.stop()

# ================================================================ compute (cached in session)
key = json.dumps([fname, len(raw), sheet, time_col, rad_col, rain_col, cov_cols, tree_cols, cfg],
                 default=str, sort_keys=True)
if st.session_state.get("key") != key:
    try:
        ds = prepare(df_raw, time_col, rad_col, tree_cols, regularize=not cfg["legacy"],
                     rain_col=rain_col, cov_cols=cov_cols)
    except Exception as e:  # noqa: BLE001
        st.error(f"数据读取失败：{e}")
        st.stop()
    bar = st.progress(0.0, text="计算中 (computing)…")
    res = run(ds, cfg, progress=lambda p: bar.progress(p, text=f"计算中 (computing)… {p:.0%}"))
    bar.empty()
    st.session_state.update(key=key, res=res)
res = st.session_state["res"]
ds = res.dataset
S = res.summary

st.title("Refilling Tool v2")
used = int((res.nights.status == "used").sum())
c = st.columns(5)
c[0].metric("液流列 (sensors)", len(ds.tree_cols))
c[1].metric("时间步长 (step)", f"{ds.step_minutes:g} min" if ds.step is not None else "?")
c[2].metric("夜晚 (nights)", f"{used} / {len(res.nights)}", help="参与计算 / 识别出的夜晚")
c[3].metric("成功划分 (ok)", int((S.status == "ok").sum()))
c[4].metric("跳过 (skipped)", int(S.status.str.startswith(("skipped", "fit failed")).sum()))
st.caption(f"单位 (unit)：{res.unit} · 模式：{'v1 兼容' if cfg['legacy'] else '原文方法 (paper)'}")
for n in res.notes:
    st.warning(n)

tabs = ["① 数据检查 (QC)", "② 逐夜检查 (night viewer)", "③ 结果 (results)"]
if len(res.evaluation) or len(res.height):
    tabs.append("④ 高度差法 (height difference)")
tabs.append("⑤ 导出 (export)")
T = dict(zip(tabs, st.tabs(tabs)))

# ---------------------------------------------------------------- QC
with T[tabs[0]]:
    for m in ds.messages:
        st.warning(m)
    qc = radiation_qc(ds, cfg["night"]["threshold"])
    if qc["night_high"] or qc["noon_low"]:
        st.warning(f"辐射可疑值 (suspicious radiation)：22–03 时高于阈值 {qc['night_high']} 个；"
                   f"10–14 时低于阈值 {qc['noon_low']} 个。请检查辐射校准与时区。")
    if not ds.messages and not (qc["night_high"] or qc["noon_low"]):
        st.success("未发现明显数据问题 (no obvious data problems)。")
    tq = st.selectbox("查看 (column)", ds.tree_cols, key="qc_tree")
    d = ds.df
    step = max(1, len(d) // 20000)
    fig = go.Figure()
    fig.add_scatter(x=d[ds.time_col][::step], y=d[tq][::step], name=tq, line=dict(width=1, color="#0072B2"))
    fig.add_scatter(x=d[ds.time_col][::step], y=d[ds.rad_col][::step], name="radiation", yaxis="y2",
                    line=dict(width=0.8, color="#E69F00"), opacity=0.6)
    for _, nr in res.nights.iterrows():
        if nr.status != "used":
            fig.add_vrect(x0=nr.night_start, x1=nr.night_end, fillcolor="red", opacity=0.12, line_width=0)
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified",
                      yaxis=dict(title="Fd"), yaxis2=dict(title="W m⁻²", overlaying="y", side="right"),
                      legend=dict(orientation="h", y=1.08))
    st.plotly_chart(fig, width="stretch")
    st.caption("红色阴影 = 未参与计算的夜晚 (excluded nights)。")
    st.dataframe(res.nights, width="stretch", height=300)

# ---------------------------------------------------------------- night viewer
with T[tabs[1]]:
    ids = res.nights.night_id.tolist()
    labels = {r.night_id: f"#{r.night_id}  {r.night_start:%Y-%m-%d %H:%M}  [{r.status}]"
              for r in res.nights.itertuples()}
    c1, c2 = st.columns([1, 2])
    tree = c1.selectbox("液流列 (sensor)", ds.tree_cols, key="nv_tree")
    i = min(_state("nv_idx", 0), len(ids) - 1)
    b1, b2, b3 = c2.columns([1, 4, 1])
    if b1.button("◀", width="stretch") and i > 0:
        i -= 1
    if b3.button("▶", width="stretch") and i < len(ids) - 1:
        i += 1
    nid = b2.selectbox("夜晚 (night)", ids, index=i, format_func=labels.get, label_visibility="collapsed")
    st.session_state["nv_idx"] = ids.index(nid)

    tt = res.times(nid)
    yy = res.observed(nid, tree)
    fig = go.Figure()
    fig.add_scatter(x=tt, y=yy, name="observed Fd", mode="lines+markers",
                    line=dict(color="black", width=1.5), marker=dict(size=4))
    rows = S[(S.night_id == nid) & (S.tree == tree)]
    tag = {"refilling": "P2", "et": "P2", "exp": "P1", "line": "lag"}
    shown = set()
    for m in methods:
        f = res.fits.get((nid, tree, m))
        if f is not None and not np.all(np.isnan(f)):
            fig.add_scatter(x=tt, y=f, name=f"{METHOD_LABELS[m]} refilling",
                            line=dict(color=COLORS[m], width=2, dash="dot"))
        r = rows[rows.method == m]
        if len(r) and pd.notna(r.breakpoint_time.iloc[0]):
            x0 = r.breakpoint_time.iloc[0]
            lab = tag[m] if not cfg["legacy"] else m
            if (lab, x0) not in shown:
                fig.add_vline(x=x0, line=dict(color=COLORS[m], width=1, dash="dash"),
                              annotation_text=lab, annotation_position="top")
                shown.add((lab, x0))
    fig.update_layout(height=430, margin=dict(l=10, r=10, t=40, b=10), hovermode="x unified",
                      yaxis_title="Fd", legend=dict(orientation="h", y=1.12))
    st.plotly_chart(fig, width="stretch")
    st.caption("实线 = 实测液流；虚线 = 各方法的充水曲线，曲线下方面积为充水量，与实测曲线之间的面积为夜间蒸腾。")
    show = rows[["method", "breakpoint", "breakpoint_time", "breakpoint_source", "night_sum", "fit_sum",
                 "remainder", "ratio", "status"]].rename(
        columns={"fit_sum": "refilling", "remainder": "transpiration", "ratio": "refilling_share"})
    st.dataframe(show.assign(method=show.method.map(METHOD_LABELS)), width="stretch", hide_index=True)

    is_ex = nid in excluded
    if st.button("✅ 恢复此夜 (include)" if is_ex else "🚫 排除此夜 (exclude this night)",
                 help="例如阴雨天、仪器异常。对所有列生效，会重新计算。"):
        (excluded.discard if is_ex else excluded.add)(nid)
        st.rerun()
    if excluded:
        st.caption(f"已手动排除 (excluded)：{sorted(excluded)}")
        if st.button("清空排除列表 (clear)"):
            excluded.clear()
            st.rerun()

# ---------------------------------------------------------------- results
with T[tabs[2]]:
    ok = S[S.status == "ok"]
    if ok.empty:
        st.warning("没有成功的结果。")
    else:
        names = {"ratio": "充水占比 refilling share", "fit_sum": "充水量 refilling",
                 "remainder": "夜间蒸腾 transpiration", "night_sum": "夜间液流 night sap flow"}
        val = st.radio("指标 (variable)", list(names), format_func=names.get, horizontal=True)
        st.subheader("各方法比较 (method comparison)")
        fig = go.Figure()
        for m in methods:
            v = ok[ok.method == m]
            fig.add_box(x=v.tree, y=v[val], name=METHOD_LABELS[m], marker_color=COLORS[m], boxpoints=False)
        fig.update_xaxes(categoryorder="array", categoryarray=ds.tree_cols)
        fig.update_layout(boxmode="group", height=420, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=val,
                          legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, width="stretch")

        st.subheader("季节动态 (seasonal dynamics)")
        c1, c2 = st.columns(2)
        mm = c1.selectbox("方法 (method)", methods, format_func=METHOD_LABELS.get)
        sel = c2.multiselect("列 (sensors)", ds.tree_cols, default=ds.tree_cols[: min(4, len(ds.tree_cols))])
        fig = go.Figure()
        for t in sel:
            v = ok[(ok.method == mm) & (ok.tree == t)]
            fig.add_scatter(x=v.night_start, y=v[val], name=t, mode="lines+markers", marker=dict(size=4))
        fig.update_layout(height=400, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=val,
                          hovermode="x unified")
        st.plotly_chart(fig, width="stretch")

        if ds.cov_cols:
            st.subheader("对环境因子的响应 (response to covariates, 原文 2.3)")
            c1, c2, c3 = st.columns(3)
            cov = c1.selectbox("环境因子", ds.cov_cols)
            mm2 = c2.selectbox("方法 ", methods, format_func=METHOD_LABELS.get, key="cov_m")
            t2 = c3.selectbox("列 ", ds.tree_cols, key="cov_t")
            v = ok[(ok.method == mm2) & (ok.tree == t2)].merge(
                res.nights[["night_id", f"night_mean_{cov}"]], on="night_id")
            fig = go.Figure(go.Scatter(x=v[f"night_mean_{cov}"], y=v[val], mode="markers",
                                       marker=dict(color=COLORS[mm2], size=6)))
            if len(v.dropna(subset=[f"night_mean_{cov}", val])) > 3:
                vv = v.dropna(subset=[f"night_mean_{cov}", val])
                r = np.corrcoef(vv[f"night_mean_{cov}"], vv[val])[0, 1]
                st.caption(f"Pearson r = {r:.3f} (n = {len(vv)})")
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=20, b=10),
                              xaxis_title=f"night mean {cov}", yaxis_title=val)
            st.plotly_chart(fig, width="stretch")

        st.subheader("汇总 (overview)")
        st.dataframe(method_overview(res).round(4), width="stretch", hide_index=True)

# ---------------------------------------------------------------- height difference
if "④ 高度差法 (height difference)" in T:
    with T["④ 高度差法 (height difference)"]:
        H, E = res.height, res.evaluation
        if len(E):
            st.subheader("方法评估 (method evaluation, 原文 1.5)")
            st.caption("C1/C2 = 用下/上探针算出的夜间蒸腾，理论上相等；A1 = (A1+B1) − B2，Ah = Q下 − Q上，"
                       "理论上相等。偏差越小、P 越大，方法越适合该树种。")
            show = E[["pair", "method", "n_nights", "C1_mean", "C2_mean", "C_bias_pct", "C1_vs_C2_P_ttest",
                      "A1_mean", "Ah_mean", "A_bias_pct", "A1_vs_Ah_P_ttest"]].copy()
            show["method"] = show.method.map(METHOD_LABELS)
            st.dataframe(show.round(3), width="stretch", hide_index=True)
            best = E[E.pair == E.pair.iloc[-1]].assign(score=lambda x: x.C_bias_pct.abs() + x.A_bias_pct.abs())
            b = best.sort_values("score").iloc[0]
            st.info(f"综合偏差最小的方法（{b.pair}）：**{METHOD_LABELS[b.method]}**（|C 偏差| + |A 偏差| 最小）。")
        if len(H):
            st.subheader(f"夜间液流组分 (components, {METHOD_LABELS[cfg['height_diff']['method']]})")
            agg = H.groupby("pair")[["A_stem_refilling", "B_canopy_refilling", "C_transpiration"]].mean()
            fig = go.Figure()
            for col, colr, lab in (("A_stem_refilling", "#8C564B", "A 茎干充水"),
                                   ("B_canopy_refilling", "#2CA02C", "B 冠层充水"),
                                   ("C_transpiration", "#1F77B4", "C 夜间蒸腾")):
                fig.add_bar(x=agg.index, y=agg[col], name=lab, marker_color=colr)
            fig.update_layout(barmode="stack", height=380, margin=dict(l=10, r=10, t=20, b=10),
                              yaxis_title=f"mean per night ({res.unit})")
            st.plotly_chart(fig, width="stretch")
            shares = H.groupby("pair")[["A_share", "B_share", "C_share"]].mean().mul(100).round(1)
            st.dataframe(shares.rename(columns=lambda c: c.replace("_share", " (%)")), width="stretch")
            st.dataframe(H, width="stretch", height=300, hide_index=True)

# ---------------------------------------------------------------- export
with T[tabs[-1]]:
    stem = Path(fname).stem + "_refilling"
    fitting = st.checkbox("包含逐点充水曲线 (include refilling curves)", False, help="文件会变大；需要自己作图时勾选。")
    c1, c2, c3 = st.columns(3)
    c1.download_button("⬇ Excel 结果", excel_bytes(res, fitting, fname), f"{stem}.xlsx", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇ CSV (长表 long format)", S.to_csv(index=False).encode("utf-8-sig"),
                       f"{stem}_summary.csv", "text/csv", width="stretch")
    c3.download_button("⬇ 参数 config.json", config_text(res.config).encode("utf-8"),
                       f"{stem}_config.json", "application/json", width="stretch")
    st.divider()
    out = st.text_input("或直接保存到本地文件夹 (save to local folder)", str(Path.home() / "refilling_results"))
    if st.button("💾 保存 (save)"):
        files = save_all(res, out, stem=stem, fitting=fitting, source_name=fname)
        st.success("已保存：\n" + "\n".join(f"- {f}" for f in files))
    st.caption("长表 (long format) 可直接在 R 中用 ggplot2 作图；config.json 记录全部参数，"
               "命令行可用 `refilltool run 数据 -c config.json` 复现。")

st.sidebar.divider()
st.sidebar.caption(f"v{__version__} · 方法见 CJPE 2023.0043 · 参数会写入导出文件")
