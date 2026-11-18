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
from refilltool.export import excel_bytes, save_all
from refilltool.io import detect_columns, prepare, radiation_qc, read_table
from refilltool.methods import METHOD_LABELS
from refilltool.pipeline import DEFAULT_CONFIG, config_text, merge_config, run

st.set_page_config(page_title="Refilling Tool v2", page_icon="🌳", layout="wide")

COLORS = {"refilling": "#D55E00", "exp": "#0072B2", "line": "#009E73", "et": "#CC79A7"}
HELP_BP = "夜间开始后的点数；前半夜/后半夜分界 (breakpoint)。"


def _example_path():
    here = Path(__file__).resolve()
    for p in [here.parents[3] / "example.xlsx", here.parents[2] / "example.xlsx", Path.cwd() / "example.xlsx"]:
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


# ---------------------------------------------------------------- sidebar: data
st.sidebar.title("🌳 Refilling Tool v2")
st.sidebar.caption("夜间液流分解 (nighttime sap flow partitioning) · PWRlab, BJFU")

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
            "其余每列一棵样树的液流 (sap flow)。列名可以任意，程序会自动识别并允许手动修改。")
    st.stop()

sheet = 0
if fname.lower().endswith((".xlsx", ".xls")):
    names = _sheets(raw)
    if names and len(names) > 1:
        sheet = st.sidebar.selectbox("工作表 (sheet)", names)
df_raw = _read(raw, fname, sheet)

t0, r0, trees0 = detect_columns(df_raw)
cols = [str(c) for c in df_raw.columns]
with st.sidebar.expander("列设置 (columns)", expanded=t0 is None or r0 is None):
    time_col = st.selectbox("时间列 (time)", cols, index=cols.index(t0) if t0 in cols else 0)
    rad_col = st.selectbox("辐射列 (solar radiation)", cols, index=cols.index(r0) if r0 in cols else min(1, len(cols) - 1))
    cand = [c for c in cols if c not in (time_col, rad_col)]
    tree_cols = st.multiselect("样树列 (sap flow)", cand, default=[c for c in trees0 if c in cand])

# ---------------------------------------------------------------- sidebar: methods
st.sidebar.header("2 方法 (methods)")
cfg = merge_config(None)
methods = [m for m in METHOD_LABELS if st.sidebar.checkbox(METHOD_LABELS[m], value=True, key=f"use_{m}")]
cfg["methods"] = methods

st.sidebar.header("3 参数 (parameters)")
with st.sidebar.expander("断点与拟合 (breakpoint & fit)"):
    for m in methods:
        d = DEFAULT_CONFIG["params"][m]
        st.markdown(f"**{METHOD_LABELS[m]}**")
        if m == "et":
            st.caption("按原文 1.4.2(4)：使用 Re 法断点 P2，后半夜液流最小值作为整夜蒸腾速率。"
                       "下面的断点仅在 v1 兼容模式下生效。")
        c1, c2 = st.columns(2)
        cfg["params"][m]["default_point"] = c1.number_input("断点 (points)", 1, 500, d["default_point"],
                                                            key=f"bp_{m}", help=HELP_BP)
        cfg["params"][m]["kneed_process"] = c2.checkbox("自动拐点 (auto knee)", d["kneed_process"],
                                                        key=f"kn_{m}",
                                                        help="Kneedle 自动检测拐点；仅适用于夜间数据质量高的情况。")
        if m == "et":
            cfg["params"][m]["min_strategy"] = st.checkbox("最小值策略 (min strategy)", d["min_strategy"],
                                                           key="et_min",
                                                           help="勾选=后半夜最小值（原文方法）；不勾选=后半夜均值。")
with st.sidebar.expander("夜间划分与筛选 (night & filter)"):
    nc = cfg["night"]
    nc["threshold"] = st.number_input("夜间辐射阈值 (W m⁻²)", 0.0, 100.0, 5.0, 0.5)
    nc["min_run_minutes"] = st.number_input("忽略短于此时长的昼夜切换 (min)", 0, 240, 30, 10,
                                            help="夜间辐射噪声或白天短暂阴暗不会切断夜晚。")
    nc["drop_partial"] = st.checkbox("去掉首尾不完整夜晚 (partial nights)", True)
    nc["drop_gap"] = st.checkbox("去掉有缺测的夜晚 (nights with gaps)", True)
    nc["min_day_rad_MJ"] = st.number_input("晴天筛选：当日辐射 ≥ (MJ m⁻²)", 0.0, 40.0, 0.0, 1.0,
                                           help="0 表示不筛选。方法适用于典型晴天。")
    cfg["legacy"] = st.checkbox("v1 兼容模式 (legacy)", False,
                                help="完全复现 v1 的夜间划分（含首尾不完整夜晚），用于与已发表结果对照。")

excluded = _state("excluded", set())
cfg["exclude_nights"] = sorted(excluded)

if not methods or not tree_cols:
    st.warning("请至少选择一种方法和一棵样树。")
    st.stop()

# ---------------------------------------------------------------- compute (cached in session)
key = json.dumps([fname, len(raw), sheet, time_col, rad_col, tree_cols, cfg], default=str, sort_keys=True)
if st.session_state.get("key") != key:
    try:
        ds = prepare(df_raw, time_col, rad_col, tree_cols, regularize=not cfg["legacy"])
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
c[0].metric("样树 (trees)", len(ds.tree_cols))
c[1].metric("时间步长 (step)", f"{ds.step_minutes:g} min" if ds.step is not None else "?")
c[2].metric("夜晚 (nights)", f"{used} / {len(res.nights)}", help="参与计算 / 识别出的夜晚")
c[3].metric("成功拟合 (fits ok)", int((S.status == "ok").sum()))
c[4].metric("跳过 (skipped)", int(S.status.str.startswith(("skipped", "fit failed")).sum()))

tab_qc, tab_night, tab_res, tab_out = st.tabs(
    ["① 数据检查 (QC)", "② 逐夜检查 (night viewer)", "③ 结果 (results)", "④ 导出 (export)"])

# ---------------------------------------------------------------- QC
with tab_qc:
    for m in ds.messages:
        st.warning(m)
    qc = radiation_qc(ds, cfg["night"]["threshold"])
    if qc["night_high"] or qc["noon_low"]:
        st.warning(f"辐射可疑值 (suspicious radiation)：22–03 时高于阈值 {qc['night_high']} 个；"
                   f"10–14 时低于阈值 {qc['noon_low']} 个。请检查辐射校准与时区。")
    if not ds.messages and not (qc["night_high"] or qc["noon_low"]):
        st.success("未发现明显数据问题 (no obvious data problems)。")
    tq = st.selectbox("查看样树 (tree)", ds.tree_cols, key="qc_tree")
    d = ds.df
    step = max(1, len(d) // 20000)  # keep the browser fast on long records
    fig = go.Figure()
    fig.add_scatter(x=d[ds.time_col][::step], y=d[tq][::step], name=tq, line=dict(width=1, color="#0072B2"))
    fig.add_scatter(x=d[ds.time_col][::step], y=d[ds.rad_col][::step], name="radiation", yaxis="y2",
                    line=dict(width=0.8, color="#E69F00"), opacity=0.6)
    for _, nr in res.nights.iterrows():
        if nr.status != "used":
            fig.add_vrect(x0=nr.night_start, x1=nr.night_end, fillcolor="red", opacity=0.12, line_width=0)
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified",
                      yaxis=dict(title="sap flow"), yaxis2=dict(title="W m⁻²", overlaying="y", side="right"),
                      legend=dict(orientation="h", y=1.08))
    st.plotly_chart(fig, use_container_width=True)
    st.caption("红色阴影 = 未参与计算的夜晚 (excluded nights)。")
    st.dataframe(res.nights, use_container_width=True, height=300)

# ---------------------------------------------------------------- night viewer
with tab_night:
    ids = res.nights.night_id.tolist()
    labels = {r.night_id: f"#{r.night_id}  {r.night_start:%Y-%m-%d %H:%M}  [{r.status}]"
              for r in res.nights.itertuples()}
    c1, c2 = st.columns([1, 2])
    tree = c1.selectbox("样树 (tree)", ds.tree_cols, key="nv_tree")
    i = _state("nv_idx", 0)
    i = min(i, len(ids) - 1)
    b1, b2, b3 = c2.columns([1, 4, 1])
    if b1.button("◀", use_container_width=True) and i > 0:
        i -= 1
    if b3.button("▶", use_container_width=True) and i < len(ids) - 1:
        i += 1
    nid = b2.selectbox("夜晚 (night)", ids, index=i, format_func=labels.get, label_visibility="collapsed")
    st.session_state["nv_idx"] = ids.index(nid)

    tt = res.times(nid)
    yy = res.observed(nid, tree)
    fig = go.Figure()
    fig.add_scatter(x=tt, y=yy, name="observed", mode="lines+markers",
                    line=dict(color="black", width=1.5), marker=dict(size=4))
    rows = S[(S.night_id == nid) & (S.tree == tree)]
    for m in methods:
        f = res.fits.get((nid, tree, m))
        if f is not None and not np.all(np.isnan(f)):
            fig.add_scatter(x=tt, y=f, name=METHOD_LABELS[m], line=dict(color=COLORS[m], width=2, dash="dot"))
        r = rows[rows.method == m]
        if len(r) and pd.notna(r.breakpoint_time.iloc[0]):
            fig.add_vline(x=r.breakpoint_time.iloc[0], line=dict(color=COLORS[m], width=1, dash="dash"))
    fig.update_layout(height=430, margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified",
                      yaxis_title="sap flow", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)
    show = rows[["method", "breakpoint", "breakpoint_time", "night_sum", "fit_sum", "remainder", "ratio", "status"]]
    st.dataframe(show.assign(method=show.method.map(METHOD_LABELS)), use_container_width=True, hide_index=True)

    is_ex = nid in excluded
    if st.button("✅ 恢复此夜 (include)" if is_ex else "🚫 排除此夜 (exclude this night)",
                 help="例如阴雨天、仪器异常。对所有样树生效，会重新计算。"):
        (excluded.discard if is_ex else excluded.add)(nid)
        st.rerun()
    if excluded:
        st.caption(f"已手动排除 (excluded)：{sorted(excluded)}")
        if st.button("清空排除列表 (clear)"):
            excluded.clear()
            st.rerun()

# ---------------------------------------------------------------- results
with tab_res:
    ok = S[S.status == "ok"]
    if ok.empty:
        st.warning("没有成功的拟合结果。")
    else:
        val = st.radio("指标 (variable)", ["ratio", "fit_sum", "remainder", "night_sum"], horizontal=True,
                       help="ratio = fit_sum / night_sum")
        st.subheader("各方法比较 (method comparison)")
        fig = go.Figure()
        for m in methods:
            v = ok[ok.method == m]
            fig.add_box(x=v.tree, y=v[val], name=METHOD_LABELS[m], marker_color=COLORS[m], boxpoints=False)
        fig.update_xaxes(categoryorder="array", categoryarray=ds.tree_cols)
        fig.update_layout(boxmode="group", height=420, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=val,
                          legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("时间动态 (seasonal dynamics)")
        c1, c2 = st.columns(2)
        mm = c1.selectbox("方法 (method)", methods, format_func=METHOD_LABELS.get)
        sel = c2.multiselect("样树 (trees)", ds.tree_cols, default=ds.tree_cols[: min(4, len(ds.tree_cols))])
        fig = go.Figure()
        for t in sel:
            v = ok[(ok.method == mm) & (ok.tree == t)]
            fig.add_scatter(x=v.night_start, y=v[val], name=t, mode="lines+markers", marker=dict(size=4))
        fig.update_layout(height=400, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=val,
                          hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("汇总统计 (summary statistics)")
        stats = ok.groupby(["method", "tree"])[val].agg(["count", "mean", "std", "median"]).reset_index()
        stats["method"] = stats.method.map(METHOD_LABELS)
        st.dataframe(stats.round(4), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------- export
with tab_out:
    stem = Path(fname).stem + "_refilling"
    fitting = st.checkbox("包含逐点拟合曲线 (include fitted curves)", False,
                          help="文件会变大；需要自己作图时勾选。")
    c1, c2, c3 = st.columns(3)
    c1.download_button("⬇ Excel 结果", excel_bytes(res, fitting, fname), f"{stem}.xlsx", use_container_width=True,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇ CSV (长表 long format)", S.to_csv(index=False).encode("utf-8-sig"),
                       f"{stem}_summary.csv", "text/csv", use_container_width=True)
    c3.download_button("⬇ 参数 config.json", config_text(res.config).encode("utf-8"),
                       f"{stem}_config.json", "application/json", use_container_width=True)
    st.divider()
    out = st.text_input("或直接保存到本地文件夹 (save to local folder)", str(Path.cwd() / "results"))
    if st.button("💾 保存 (save)"):
        files = save_all(res, out, stem=stem, fitting=fitting, source_name=fname)
        st.success("已保存：\n" + "\n".join(f"- {f}" for f in files))
    st.caption("长表 (long format) 可直接在 R 中用 ggplot2 作图；config.json 记录全部参数，便于复现 (reproducibility)。")

st.sidebar.divider()
st.sidebar.caption(f"v{__version__} · 方法见 CJPE 2023.0043 · 参数会写入导出文件")
