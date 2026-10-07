#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
لوحة مؤشرات المنشآت الاقتصادية في فلسطين (1997 – 2023)
=======================================================
تشغيل اللوحة:
    pip install dash plotly pandas openpyxl
    python dashboard.py                  # يبحث عن Book6.10.xlsx بجانب السكربت
    python dashboard.py path/to/file.xlsx

ثم افتح المتصفح على:  http://127.0.0.1:8050
"""

import sys
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from dash import Dash, dcc, html, dash_table, Input, Output, State
from dash.dash_table.Format import Format, Group, Scheme

# ════════════════════════════════════════════════════════════════════
# 1) الإعدادات العامة والألوان
# ════════════════════════════════════════════════════════════════════
FONT = "Cairo, 'Segoe UI', Tahoma, Arial, sans-serif"

PRIMARY, DARK, MUTED, GRID = "#0B6E4F", "#1F2D3D", "#6B7785", "#E6EAEE"
RED, GOLD, BLUE, PURPLE = "#B3261E", "#C8963E", "#2F6690", "#7B5EA7"
MALE, FEMALE = "#2F6690", "#C2548A"
WB_COLOR, GAZA_COLOR = PRIMARY, RED
SEQ = ["#0B6E4F", "#2F6690", "#C8963E", "#B3261E", "#7B5EA7", "#3A9D8F",
       "#E07A5F", "#6C757D", "#8AB17D", "#264653", "#D4A373", "#9A8C98"]
GREENS = ["#0B6E4F", "#2E8B6A", "#5BA98A", "#8CC7AC", "#B9DDCB", "#E2F1EA"]

CFG = {"displaylogo": False, "responsive": True,
       "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"]}

WB_GOVS = ["جنين", "طوباس والأغوار الشمالية", "طولكرم", "نابلس", "قلقيلية", "سلفيت",
           "رام الله والبيرة", "أريحا والأغوار", "القدس", "بيت لحم", "الخليل"]
GAZA_GOVS = ["شمال غزة", "غزة", "دير البلح", "خانيونس", "رفح"]
GOVS = WB_GOVS + GAZA_GOVS
MAIN_REGS = ["فلسطين", "الضفة الغربية", "قطاع غزة"]
ALL_REGS = MAIN_REGS + GOVS

# أسماء مختصرة للأنشطة الاقتصادية (للعناوين في الرسوم)
SHORT = {
    "B": "التعدين والمحاجر", "C": "الصناعات التحويلية", "D": "الكهرباء والغاز",
    "E": "المياه والنفايات", "F": "الإنشاءات", "G": "التجارة وإصلاح المركبات",
    "H": "النقل والتخزين", "I": "الإقامة والطعام", "J": "المعلومات والاتصالات",
    "K": "المالية والتأمين", "L": "الأنشطة العقارية", "M": "المهنية والعلمية",
    "N": "الإدارية والمساندة", "O": "الإدارة العامة", "P": "التعليم",
    "Q": "الصحة والعمل الاجتماعي", "R": "الفنون والترفيه", "S": "خدمات أخرى",
    "T": "الأسر المعيشية", "U": "المنظمات الدولية", "NS": "غير مبين",
}

GROUP_AR = {
    "ownership": "الملكية", "legal": "الشكل القانوني", "organization": "الشكل التنظيمي",
    "activity": "النشاط الاقتصادي", "employment": "العمالة حسب النشاط",
    "governorates": "المحافظات", "size": "حجم المنشأة (الضفة الغربية)",
    "operational_status": "الحالة التشغيلية", "history": "السلسلة التاريخية",
    "bridge": "جسر التصنيف 2007-2012", "regional_history": "التاريخ الإقليمي",
    "nonagri_derived": "المنشآت غير الزراعية",
}

# ════════════════════════════════════════════════════════════════════
# 2) تحميل البيانات وتهيئتها
# ════════════════════════════════════════════════════════════════════
COLS = ["group", "key", "cat_ar", "cat_en", "reg_ar", "reg_en",
        "year", "est", "pe", "male", "female"]


def find_data_file() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    here = Path(__file__).resolve().parent
    candidates = (
        here / "Book6.10.xlsx",
        here / "Book6_10.xlsx",
        Path.cwd() / "Book6.10.xlsx",
        Path.cwd() / "Book6_10.xlsx",
    )
    for p in candidates:
        if p.exists():
            return p
    raise SystemExit("لم يتم العثور على ملف البيانات Excel (Book6.10.xlsx أو Book6_10.xlsx).")


def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path).iloc[:, :11]
    df.columns = COLS  # الاعتماد على ترتيب الأعمدة لا أسمائها
    df["year"] = df["year"].astype(int)
    for c in ("est", "pe", "male", "female"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ("group", "key", "cat_ar", "reg_ar"):
        df[c] = df[c].astype(str).str.strip()
    return df


D = load_data(find_data_file())
TOT = D["key"].str.upper().eq("TOTAL")          # صفوف «المجموع»
GOV = D[D.group == "governorates"].copy()
EMP = D[D.group == "employment"].copy()
ACT = D[D.group == "activity"].copy()
OWN_TOT = D[(D.group == "ownership") & TOT]

OWN_YEARS = sorted(int(y) for y in OWN_TOT.year.unique())
GOV_YEARS = sorted(int(y) for y in GOV.year.unique())
ACT_YEARS = sorted(int(y) for y in ACT.year.unique())
EMP_YEARS = sorted(int(y) for y in EMP.year.unique())


# ════════════════════════════════════════════════════════════════════
# 3) دوال مساعدة
# ════════════════════════════════════════════════════════════════════
def fmt(n, d=0):
    if n is None or (isinstance(n, float) and np.isnan(n)):
        return "—"
    return f"{n:,.{d}f}"


def fem_share(m, f):
    """نسبة الإناث من (الذكور + الإناث)؛ لأن إجمالي العاملين لا يساوي مجموعهما دائماً."""
    return f / (m + f) * 100


def wrap(s, w=24):
    return "<br>".join(textwrap.wrap(str(s), w)) or str(s)


def act_label(key, full=None):
    return SHORT.get(key, full or key)


def style(fig, height=400, legend=True):
    fig.update_layout(
        height=height, margin=dict(l=10, r=10, t=30 if legend else 10, b=30),
        font=dict(family=FONT, size=13, color=DARK),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        hoverlabel=dict(font_family=FONT, font_size=13, align="right"),
        showlegend=legend, colorway=SEQ, separators=".,",
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=False, zeroline=False, linecolor=GRID)
    fig.update_yaxes(gridcolor=GRID, zeroline=False)
    return fig


def rtl_hbar(fig):
    """جعل الأعمدة الأفقية تنمو من اليمين لليسار."""
    fig.update_xaxes(autorange="reversed", showgrid=True, gridcolor=GRID)
    fig.update_yaxes(side="right", showgrid=False)
    return fig


def empty_fig(msg="لا تتوفر بيانات لهذا الاختيار", height=300):
    fig = go.Figure()
    fig.add_annotation(text=msg, showarrow=False, font=dict(size=15, color=MUTED))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return style(fig, height, legend=False)


def card(title, body, sub=None, cls=""):
    head = [html.H3(title)]
    if sub:
        head.append(html.Span(sub, className="card-sub"))
    return html.Div([html.Div(head, className="card-head"), body], className=f"card {cls}")


def graph(gid, height=400):
    return dcc.Graph(id=gid, config=CFG, style={"height": f"{height}px"})


def kpi_card(label, value, sub="", color=PRIMARY):
    return html.Div([html.Div(label, className="kpi-label"),
                     html.Div(value, className="kpi-value"),
                     html.Div(sub, className="kpi-sub")],
                    className="kpi", style={"borderTopColor": color})


def dd(cid, options, value, width=200, multi=False):
    opts = [o if isinstance(o, dict) else {"label": str(o), "value": o} for o in options]
    return dcc.Dropdown(id=cid, options=opts, value=value, clearable=False,
                        multi=multi, className="dd", style={"minWidth": f"{width}px"})


def controls(*items):
    return html.Div([html.Div([html.Label(lbl), comp], className="ctl") for lbl, comp in items],
                    className="controls")


def pct_change(new, old):
    if isinstance(old, pd.Series):
        return (new / old.replace(0, np.nan) - 1) * 100
    return (new / old - 1) * 100 if old else np.nan


# ════════════════════════════════════════════════════════════════════
# 4) بناء الرسوم — نظرة عامة
# ════════════════════════════════════════════════════════════════════
def build_kpis(region, year):
    own = OWN_TOT[OWN_TOT.reg_ar == region].set_index("year").est
    est = own.get(year, np.nan)
    prev = [y for y in own.index if y < year]
    if prev:
        py = max(prev)
        g = pct_change(est, own[py])
        growth = f"{'▲' if g >= 0 else '▼'} {abs(g):.1f}%"
        gsub, gcol = f"مقارنة بتعداد {py}", (PRIMARY if g >= 0 else RED)
    else:
        growth, gsub, gcol = "—", "لا يوجد تعداد سابق", MUTED

    pal = own_total("فلسطين", year)
    if region == "فلسطين":
        share_lbl, share_val = "حصة الضفة الغربية", own_total("الضفة الغربية", year) / pal * 100
    else:
        share_lbl, share_val = "حصة المنطقة من فلسطين", est / pal * 100

    ys = [y for y in GOV_YEARS if y <= year]
    if ys:
        yw = max(ys)
        r = GOV[(GOV.reg_ar == region) & (GOV.year == yw)].iloc[0]
        note = f"بيانات تعداد {yw}" + ("" if yw == year else " (أحدث بيانات متاحة)")
        workers, fem, avg = fmt(r.pe), f"{fem_share(r.male, r.female):.1f}%", f"{r.pe / r.est:.2f}"
    else:
        note = "غير متوفر قبل 2012"
        workers = fem = avg = "—"

    return [
        kpi_card("المنشآت العاملة", fmt(est), f"تعداد {year}", PRIMARY),
        kpi_card("التغير عن التعداد السابق", growth, gsub, gcol),
        kpi_card("إجمالي العاملين", workers, note, BLUE),
        kpi_card("نسبة الإناث بين العاملين", fem, note, FEMALE),
        kpi_card("متوسط العاملين لكل منشأة", avg, note, GOLD),
        kpi_card(share_lbl, f"{share_val:.1f}%", f"تعداد {year}", PURPLE),
    ]


def own_total(region, year):
    d = OWN_TOT[(OWN_TOT.reg_ar == region) & (OWN_TOT.year == year)].est
    return float(d.iloc[0]) if len(d) else np.nan


def fig_trend(region):
    d = OWN_TOT[OWN_TOT.reg_ar == region].sort_values("year")
    if d.empty:
        return empty_fig()
    growth = [None] + [pct_change(b, a) for a, b in zip(d.est[:-1], d.est[1:])]
    fig = go.Figure(go.Bar(
        x=d.year.astype(str), y=d.est, marker_color=PRIMARY, name="المنشآت العاملة",
        text=d.est, texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False,
        customdata=[("—" if g is None else f"{g:+.1f}%") for g in growth],
        hovertemplate="تعداد %{x}<br>%{y:,.0f} منشأة<br>التغير: %{customdata}<extra></extra>"))
    fig.update_yaxes(showticklabels=False, showgrid=False, range=[0, d.est.max() * 1.18])
    return style(fig, 380, legend=False)


def fig_ownership_snapshot(region, year):
    d = D[(D.group == "ownership") & ~TOT & (D.reg_ar == region) & (D.year == year) & (D.est > 0)]
    if d.empty:
        return empty_fig()
    d = d.sort_values("est")
    tot = own_total(region, year)
    fig = go.Figure(go.Bar(
        y=d.cat_ar, x=d.est, orientation="h", marker_color=PRIMARY,
        customdata=d.est / tot * 100, text=d.est,
        texttemplate="%{text:,.0f} (%{customdata:.1f}%)", textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>%{x:,.0f} منشأة (%{customdata:.1f}%)<extra></extra>"))
    style(fig, 380, legend=False)
    rtl_hbar(fig)
    fig.update_xaxes(range=[d.est.max() * 1.45, 0], showticklabels=False)
    return fig


def fig_west_gaza():
    d = OWN_TOT[OWN_TOT.reg_ar.isin(MAIN_REGS)].pivot(index="year", columns="reg_ar", values="est").sort_index()
    xs = d.index.astype(str)
    fig = go.Figure()
    for name, col in (("الضفة الغربية", WB_COLOR), ("قطاع غزة", GAZA_COLOR)):
        fig.add_bar(x=xs, y=d[name], name=name, marker_color=col,
                    text=d[name], texttemplate="%{text:,.0f}", textposition="inside",
                    textfont=dict(color="white", size=11),
                    hovertemplate="تعداد %{x}<br>" + name + ": %{y:,.0f}<extra></extra>")
    fig.add_scatter(x=xs, y=d["فلسطين"], mode="text", showlegend=False,
                    text=[f"{v:,.0f}" for v in d["فلسطين"]], textposition="top center",
                    textfont=dict(size=12, color=DARK), hoverinfo="skip")
    fig.update_layout(barmode="stack")
    fig.update_yaxes(showticklabels=False, showgrid=False, range=[0, d["فلسطين"].max() * 1.15])
    return style(fig, 380)


def fig_gender(region):
    d = GOV[GOV.reg_ar == region].sort_values("year")
    if d.empty:
        return empty_fig()
    xs = d.year.astype(str)
    shares = [fem_share(m, f) for m, f in zip(d.male, d.female)]
    fig = go.Figure()
    fig.add_bar(x=xs, y=d.male, name="الذكور", marker_color=MALE, text=d.male,
                texttemplate="%{text:,.0f}", textposition="inside", textfont=dict(color="white"))
    fig.add_bar(x=xs, y=d.female, name="الإناث", marker_color=FEMALE, text=d.female,
                texttemplate="%{text:,.0f}", textposition="inside", textfont=dict(color="white"),
                customdata=shares, hovertemplate="الإناث: %{y:,.0f}<br>النسبة: %{customdata:.1f}%<extra></extra>")
    fig.add_scatter(x=xs, y=d.male + d.female, mode="text", showlegend=False, hoverinfo="skip",
                    text=[f"الإناث {s:.1f}%" for s in shares], textposition="top center")
    fig.update_layout(barmode="stack")
    fig.update_yaxes(showticklabels=False, showgrid=False, range=[0, (d.male + d.female).max() * 1.15])
    return style(fig, 380)


# ════════════════════════════════════════════════════════════════════
# 5) بناء الرسوم — المحافظات
# ════════════════════════════════════════════════════════════════════
METRICS = {
    "est": ("عدد المنشآت", ",.0f"), "pe": ("إجمالي العاملين", ",.0f"),
    "fem": ("نسبة الإناث من العاملين (%)", ".1f"), "avg": ("متوسط العاملين لكل منشأة", ".2f"),
}
SCOPES = {"all": GOVS, "wb": WB_GOVS, "gaza": GAZA_GOVS}


def gov_frame(year, scope):
    g = GOV[(GOV.year == year) & GOV.reg_ar.isin(SCOPES[scope])].copy()
    g["fem"] = fem_share(g.male, g.female)
    g["avg"] = g.pe / g.est
    g["area"] = np.where(g.reg_ar.isin(WB_GOVS), "الضفة الغربية", "قطاع غزة")
    return g


def fig_gov_rank(year, metric, scope):
    g = gov_frame(year, scope)
    if g.empty:
        return empty_fig()
    label, f = METRICS[metric]
    g = g.sort_values(metric)
    fig = go.Figure()
    for area, col in (("الضفة الغربية", WB_COLOR), ("قطاع غزة", GAZA_COLOR)):
        s = g[g.area == area]
        if s.empty:
            continue
        fig.add_bar(y=s.reg_ar, x=s[metric], orientation="h", name=area, marker_color=col,
                    text=s[metric], texttemplate=f"%{{text:{f}}}", textposition="outside", cliponaxis=False,
                    hovertemplate="%{y}<br>" + label + ": %{x:" + f + "}<extra></extra>")
    h = max(380, 32 * len(g) + 80)
    style(fig, h)
    rtl_hbar(fig)
    fig.update_layout(barmode="overlay")
    fig.update_yaxes(categoryorder="array", categoryarray=list(g.reg_ar))
    fig.update_xaxes(range=[g[metric].max() * 1.25, 0], showticklabels=False)
    return fig


def fig_gov_growth(scope):
    if len(GOV_YEARS) < 2:
        return empty_fig("يلزم تعدادان على الأقل")
    y0, y1 = GOV_YEARS[0], GOV_YEARS[-1]
    a, b = gov_frame(y0, scope).set_index("reg_ar"), gov_frame(y1, scope).set_index("reg_ar")
    d = pd.DataFrame({"est": pct_change(b.est, a.est), "pe": pct_change(b.pe, a.pe)}).dropna()
    d = d.sort_values("est", ascending=False)
    fig = go.Figure()
    for col, name, color in (("est", "المنشآت", PRIMARY), ("pe", "العاملون", GOLD)):
        fig.add_bar(x=[wrap(i, 12) for i in d.index], y=d[col], name=name, marker_color=color,
                    text=d[col], texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False,
                    hovertemplate="%{x}<br>" + name + ": %{y:.1f}%<extra></extra>")
    fig.update_layout(barmode="group")
    fig.update_yaxes(ticksuffix="%")
    return style(fig, 420)


def fig_gov_heat(scope):
    d = D[(D.group == "ownership") & TOT & D.reg_ar.isin(SCOPES[scope])]
    p = d.pivot(index="reg_ar", columns="year", values="est").reindex(SCOPES[scope])
    z = p.div(p.max(axis=1), axis=0)  # تطبيع كل صف ليظهر الاتجاه لا الحجم
    fig = go.Figure(go.Heatmap(
        z=z.values, x=[str(c) for c in p.columns], y=list(p.index), customdata=p.values,
        text=[[f"{v:,.0f}" for v in row] for row in p.values], texttemplate="%{text}",
        colorscale=[[0, "#EAF4EF"], [1, PRIMARY]], showscale=False, xgap=3, ygap=3,
        hovertemplate="%{y} — %{x}<br>%{customdata:,.0f} منشأة<extra></extra>"))
    fig.update_yaxes(autorange="reversed", side="right", showgrid=False)
    fig.update_xaxes(side="top")
    return style(fig, max(380, 30 * len(p) + 90), legend=False)


def fig_gov_donut(year, scope):
    g = gov_frame(year, scope)
    if g.empty:
        return empty_fig()
    g = g.sort_values("est", ascending=False)
    fig = go.Figure(go.Pie(labels=g.reg_ar, values=g.est, hole=0.55, sort=False,
                           textinfo="percent", marker=dict(colors=SEQ * 2),
                           hovertemplate="%{label}<br>%{value:,.0f} منشأة (%{percent})<extra></extra>"))
    fig.add_annotation(text=f"<b>{g.est.sum():,.0f}</b><br>منشأة", showarrow=False, font=dict(size=15))
    return style(fig, 420)


def gov_table_data(year, scope):
    g = gov_frame(year, scope)
    prev = [y for y in GOV_YEARS if y < year]
    py = None
    if prev:
        py = max(prev)
        old = gov_frame(py, scope).set_index("reg_ar").est
        g["growth"] = [pct_change(e, old.get(r, np.nan)) for r, e in zip(g.reg_ar, g.est)]
    else:
        g["growth"] = np.nan
    g = g.sort_values("est", ascending=False)
    out = g[["reg_ar", "area", "est", "pe", "fem", "avg", "growth"]].round(2)
    return out.replace({np.nan: None}).to_dict("records")


# ════════════════════════════════════════════════════════════════════
# 6) بناء الرسوم — الأنشطة الاقتصادية
# ════════════════════════════════════════════════════════════════════
def act_frame(year, region):
    d = ACT[(ACT.year == year) & (ACT.reg_ar == region)]
    if d.empty:
        return d, np.nan
    tot = float(d[d.key == "TOTAL"].est.iloc[0])
    d = d[~d.key.isin(["TOTAL", "NS"]) & (d.est > 0)].copy()
    d["share"] = d.est / tot * 100
    d["short"] = [act_label(k, c) for k, c in zip(d.key, d.cat_ar)]
    return d, tot


def fig_act_bar(year, region, topn):
    d, tot = act_frame(year, region)
    if d.empty:
        return empty_fig()
    d = d.nlargest(topn, "est").sort_values("est")
    fig = go.Figure(go.Bar(
        y=d.short, x=d.est, orientation="h", marker_color=PRIMARY, customdata=np.c_[d.share, d.cat_ar],
        text=d.est, texttemplate="%{text:,.0f}  (%{customdata[0]:.1f}%)", textposition="outside", cliponaxis=False,
        hovertemplate="%{customdata[1]}<br>%{x:,.0f} منشأة (%{customdata[0]:.1f}%)<extra></extra>"))
    style(fig, max(360, 38 * len(d) + 70), legend=False)
    rtl_hbar(fig)
    fig.update_xaxes(range=[d.est.max() * 1.5, 0], showticklabels=False)
    return fig


def fig_act_tree(year, region):
    d, _ = act_frame(year, region)
    if d.empty:
        return empty_fig()
    fig = go.Figure(go.Treemap(
        labels=d.short, parents=[""] * len(d), values=d.est, customdata=d.cat_ar,
        texttemplate="<b>%{label}</b><br>%{percentRoot:.1%}", branchvalues="total",
        marker=dict(colors=d.est, colorscale=[[0, "#DCEFE6"], [1, PRIMARY]], line=dict(width=2, color="white")),
        hovertemplate="%{customdata}<br>%{value:,.0f} منشأة<extra></extra>"))
    return style(fig, 420, legend=False)


def fig_act_years(region, topn, year):
    d, _ = act_frame(year, region)
    if d.empty:
        return empty_fig()
    keys = list(d.nlargest(topn, "est").sort_values("est").key)
    fig = go.Figure()
    for i, y in enumerate(ACT_YEARS):
        s = ACT[(ACT.year == y) & (ACT.reg_ar == region) & ACT.key.isin(keys)].set_index("key").reindex(keys)
        fig.add_bar(y=[act_label(k) for k in keys], x=s.est, orientation="h", name=str(y),
                    marker_color=GREENS[min(i * 2, 5)] if len(ACT_YEARS) > 1 else PRIMARY,
                    hovertemplate="%{y} — " + str(y) + "<br>%{x:,.0f} منشأة<extra></extra>")
    style(fig, max(380, 22 * len(keys) * len(ACT_YEARS) + 70))
    rtl_hbar(fig)
    fig.update_layout(barmode="group")
    fig.update_xaxes(range=[None, 0])
    return fig


def fig_act_change(region):
    if len(ACT_YEARS) < 2:
        return empty_fig()
    y0, y1 = ACT_YEARS[0], ACT_YEARS[-1]
    a, _ = act_frame(y0, region)
    b, _ = act_frame(y1, region)
    a, b = a.set_index("key"), b.set_index("key")
    keys = a.index[a.est >= 50].intersection(b.index)
    if len(keys) == 0:
        return empty_fig()
    ch = pd.Series({k: pct_change(b.est[k], a.est[k]) for k in keys}).sort_values()
    fig = go.Figure(go.Bar(
        y=[act_label(k) for k in ch.index], x=ch.values, orientation="h",
        marker_color=[PRIMARY if v >= 0 else RED for v in ch.values],
        text=ch.values, texttemplate="%{text:+.0f}%", textposition="outside", cliponaxis=False,
        customdata=np.c_[[a.est[k] for k in ch.index], [b.est[k] for k in ch.index]],
        hovertemplate="%{y}<br>" + f"{y0}: " + "%{customdata[0]:,.0f}<br>" + f"{y1}: " + "%{customdata[1]:,.0f}"
                      + "<br>التغير: %{x:+.1f}%<extra></extra>"))
    style(fig, max(380, 30 * len(ch) + 70), legend=False)
    rtl_hbar(fig)
    lo, hi = min(ch.min(), 0), ch.max()
    fig.update_xaxes(range=[hi * 1.3, lo * 1.3 if lo < 0 else -hi * 0.05], ticksuffix="%")
    return fig


def act_insight(year, region):
    d, tot = act_frame(year, region)
    if d.empty:
        return ""
    d = d.sort_values("est", ascending=False)
    top = d.iloc[0]
    top3 = d.head(3).est.sum() / tot * 100
    return html.Div([
        html.B("قراءة سريعة: "),
        f"يتصدر نشاط «{top.cat_ar}» بـ {top.est:,.0f} منشأة ({top.share:.1f}% من إجمالي {tot:,.0f}) في {region} عام {year}، "
        f"وتستحوذ الأنشطة الثلاثة الأكبر معاً على {top3:.1f}% من المنشآت."], className="insight")


# ════════════════════════════════════════════════════════════════════
# 7) بناء الرسوم — العمالة والنوع الاجتماعي
# ════════════════════════════════════════════════════════════════════
def emp_frame(year, region):
    d = EMP[(EMP.year == year) & (EMP.reg_ar == region)]
    t = d[d.key == "TOTAL"]
    d = d[~d.key.isin(["TOTAL", "NS"]) & (d.pe > 0)].copy()
    d["fem"] = fem_share(d.male, d.female)
    d["avg"] = d.pe / d.est
    d["short"] = [act_label(k, c) for k, c in zip(d.key, d.cat_ar)]
    return d, (t.iloc[0] if len(t) else None)


def emp_kpis(year, region):
    d, t = emp_frame(year, region)
    if t is None:
        return []
    return [
        kpi_card("إجمالي العاملين", fmt(t.pe), f"{region} — {year}", BLUE),
        kpi_card("الذكور", fmt(t.male), f"{t.male / (t.male + t.female) * 100:.1f}% من الموزّع حسب الجنس", MALE),
        kpi_card("الإناث", fmt(t.female), f"{fem_share(t.male, t.female):.1f}% من الموزّع حسب الجنس", FEMALE),
        kpi_card("متوسط العاملين لكل منشأة", f"{t.pe / t.est:.2f}", f"{fmt(t.est)} منشأة", GOLD),
    ]


def fig_emp_stack(year, region):
    d, _ = emp_frame(year, region)
    if d.empty:
        return empty_fig()
    d = d.sort_values("pe")
    fig = go.Figure()
    fig.add_bar(y=d.short, x=d.male, orientation="h", name="الذكور", marker_color=MALE,
                hovertemplate="%{y}<br>الذكور: %{x:,.0f}<extra></extra>")
    fig.add_bar(y=d.short, x=d.female, orientation="h", name="الإناث", marker_color=FEMALE,
                hovertemplate="%{y}<br>الإناث: %{x:,.0f}<extra></extra>")
    style(fig, max(420, 28 * len(d) + 80))
    rtl_hbar(fig)
    fig.update_layout(barmode="stack")
    return fig


def fig_emp_metric(year, region, col, label, suffix, fmt_spec):
    d, t = emp_frame(year, region)
    if d.empty:
        return empty_fig()
    d = d.sort_values(col)
    overall = fem_share(t.male, t.female) if col == "fem" else t.pe / t.est
    fig = go.Figure(go.Bar(
        y=d.short, x=d[col], orientation="h", marker_color=[FEMALE if v >= overall else "#CBD3DB" for v in d[col]]
        if col == "fem" else [GOLD if v >= overall else "#CBD3DB" for v in d[col]],
        text=d[col], texttemplate=f"%{{text:{fmt_spec}}}{suffix}", textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>" + label + f": %{{x:{fmt_spec}}}{suffix}<extra></extra>"))
    style(fig, max(420, 28 * len(d) + 80), legend=False)
    rtl_hbar(fig)
    fig.add_vline(x=overall, line_dash="dash", line_color=DARK,
                  annotation_text=f"المتوسط العام {overall:{fmt_spec}}{suffix}", annotation_position="top")
    fig.update_xaxes(range=[d[col].max() * 1.3, 0], showticklabels=False)
    return fig


def fig_emp_bubble(year, region):
    d, _ = emp_frame(year, region)
    if d.empty:
        return empty_fig()
    big = set(d.nlargest(8, "pe").index)
    fig = go.Figure(go.Scatter(
        x=d.avg, y=d.fem, mode="markers+text", text=[s if i in big else "" for i, s in zip(d.index, d.short)],
        textposition="top center", textfont=dict(size=11),
        marker=dict(size=d.pe, sizemode="area", sizeref=2.0 * d.pe.max() / (60 ** 2), sizemin=6,
                    color=d.fem, colorscale=[[0, MALE], [1, FEMALE]], line=dict(width=1, color="white"), opacity=0.85),
        customdata=np.c_[d.cat_ar, d.pe],
        hovertemplate="<b>%{customdata[0]}</b><br>متوسط العاملين: %{x:.2f}<br>نسبة الإناث: %{y:.1f}%"
                      "<br>العاملون: %{customdata[1]:,.0f}<extra></extra>"))
    fig.update_xaxes(title="متوسط العاملين لكل منشأة (حجم المنشأة)", showgrid=True, gridcolor=GRID)
    fig.update_yaxes(title="نسبة الإناث %", ticksuffix="%")
    return style(fig, 480, legend=False)


# ════════════════════════════════════════════════════════════════════
# 8) بناء الرسوم — الملكية والشكل القانوني والتنظيمي
# ════════════════════════════════════════════════════════════════════
CLASS_OPTS = {"ownership": "الملكية", "legal": "الشكل القانوني", "organization": "الشكل التنظيمي"}


def class_years(group):
    return sorted(int(y) for y in D[D.group == group].year.unique())


def fig_class_bar(group, region, year, log):
    d = D[(D.group == group) & ~TOT & (D.reg_ar == region) & (D.year == year) & (D.est > 0)]
    if d.empty:
        return empty_fig()
    tot = float(D[(D.group == group) & TOT & (D.reg_ar == region) & (D.year == year)].est.iloc[0])
    d = d.sort_values("est")
    fig = go.Figure(go.Bar(
        y=[wrap(c, 26) for c in d.cat_ar], x=d.est, orientation="h", marker_color=PRIMARY,
        customdata=d.est / tot * 100, text=d.est, texttemplate="%{text:,.0f} (%{customdata:.1f}%)",
        textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>%{x:,.0f} منشأة (%{customdata:.1f}%)<extra></extra>"))
    style(fig, max(380, 36 * len(d) + 80), legend=False)
    rtl_hbar(fig)
    top = d.est.max()
    fig.update_xaxes(range=[np.log10(top * 6), 0] if log else [top * 1.5, 0],
                     type="log" if log else "linear", showticklabels=False)
    return fig


def fig_class_share(group, region):
    d = D[(D.group == group) & ~TOT & (D.reg_ar == region)]
    p = d.pivot_table(index="cat_ar", columns="year", values="est", aggfunc="sum").fillna(0)
    if p.empty:
        return empty_fig()
    sh = p / p.sum() * 100
    small = sh.index[sh.max(axis=1) < 2]
    if len(small):
        sh = sh.drop(small)
        sh.loc["فئات أخرى"] = 100 - sh.sum()
    order = sh.mean(axis=1).sort_values(ascending=False).index
    fig = go.Figure()
    for i, cat in enumerate(order):
        fig.add_bar(x=[str(c) for c in sh.columns], y=sh.loc[cat], name=cat, marker_color=SEQ[i % len(SEQ)],
                    hovertemplate=cat + " — %{x}<br>%{y:.1f}%<extra></extra>")
    fig.update_layout(barmode="stack")
    fig.update_yaxes(ticksuffix="%", range=[0, 100])
    style(fig, 440)
    fig.update_layout(legend=dict(orientation="h", y=-0.15, x=1, xanchor="right"))
    return fig


# ════════════════════════════════════════════════════════════════════
# 9) بناء الرسوم — حجم المنشآت (الضفة الغربية)
# ════════════════════════════════════════════════════════════════════
SZ = D[D.group == "size"].copy()
SZ["act"] = SZ.key.str.split(" ").str[0]
SZ["cls"] = SZ.key.str.split(" ", n=1).str[1]
SZ_ORDER = [c for c in ["1–4", "5–9", "10–19", "20–49", "50–99", "100+"] if c in set(SZ.cls)]
SZ_YEARS = sorted(int(y) for y in SZ.year.unique())
SZ_COLORS = dict(zip(SZ_ORDER, ["#E2F1EA", "#B9DDCB", "#8CC7AC", "#5BA98A", "#2E8B6A", "#0B4F38"]))


def size_dist(year, act="TOTAL"):
    d = SZ[(SZ.year == year) & (SZ.act == act) & SZ.cls.isin(SZ_ORDER)].set_index("cls").est.reindex(SZ_ORDER)
    return d


def size_kpis(year):
    s = size_dist(year)
    tot = float(SZ[(SZ.year == year) & (SZ.key == "TOTAL المجموع")].est.iloc[0])
    return [
        kpi_card("إجمالي المنشآت", fmt(tot), f"الضفة الغربية — {year}", PRIMARY),
        kpi_card("منشآت صغيرة جداً (1–4)", f"{s.iloc[0] / tot * 100:.1f}%", f"{fmt(s.iloc[0])} منشأة", GOLD),
        kpi_card("منشآت متوسطة (10–49)", f"{s[['10–19', '20–49']].sum() / tot * 100:.1f}%", "10 إلى 49 عاملاً", BLUE),
        kpi_card("منشآت كبيرة (50+)", f"{s[['50–99', '100+']].sum() / tot * 100:.1f}%",
                 f"{fmt(s[['50–99', '100+']].sum())} منشأة", RED),
    ]


def fig_size_dist():
    fig = go.Figure()
    for i, y in enumerate(SZ_YEARS):
        s = size_dist(y)
        fig.add_bar(x=SZ_ORDER, y=s / s.sum() * 100, name=str(y), marker_color=[PRIMARY, GOLD][i % 2],
                    text=s / s.sum() * 100, texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False,
                    customdata=s, hovertemplate="فئة %{x} عاملاً<br>%{y:.1f}% (%{customdata:,.0f} منشأة)<extra></extra>")
    fig.update_layout(barmode="group")
    fig.update_yaxes(showticklabels=False, showgrid=False)
    fig.update_xaxes(title="عدد العاملين في المنشأة")
    return style(fig, 380)


def fig_size_stack(year, sort):
    d = SZ[(SZ.year == year) & SZ.cls.isin(SZ_ORDER) & ~SZ.act.isin(["TOTAL", "NS"])]
    p = d.pivot_table(index="act", columns="cls", values="est", aggfunc="sum").reindex(columns=SZ_ORDER).fillna(0)
    p = p[p.sum(axis=1) >= 30]
    if p.empty:
        return empty_fig()
    sh = p.div(p.sum(axis=1), axis=0) * 100
    order = (p.sum(axis=1) if sort == "count" else sh[SZ_ORDER[0]]).sort_values().index
    fig = go.Figure()
    for c in SZ_ORDER:
        fig.add_bar(y=[act_label(a) for a in order], x=sh.loc[order, c], orientation="h", name=c,
                    marker_color=SZ_COLORS[c], customdata=p.loc[order, c],
                    text=sh.loc[order, c], texttemplate="%{text:.0f}%", textfont=dict(size=11),
                    hovertemplate="%{y} — " + c + " عاملاً<br>%{x:.1f}% (%{customdata:,.0f} منشأة)<extra></extra>")
    style(fig, max(440, 30 * len(p) + 90))
    rtl_hbar(fig)
    fig.update_layout(barmode="stack")
    fig.update_xaxes(range=[100, 0], ticksuffix="%")
    return fig


# ════════════════════════════════════════════════════════════════════
# 10) بناء الرسوم — الحالة التشغيلية
# ════════════════════════════════════════════════════════════════════
OP = D[D.group == "operational_status"]
OP_YEARS = sorted(int(y) for y in OP.year.unique())
STATUS = [("Operating", "عاملة", PRIMARY), ("Temporarily closed", "متوقفة مؤقتاً", GOLD),
          ("Under preparation", "تحت التجهيز", BLUE)]


def fig_op_stack(year):
    d = OP[OP.year == year].pivot(index="reg_ar", columns="key", values="est").reindex(MAIN_REGS)
    if d.empty or d.isna().all().all():
        return empty_fig()
    tot = d[[s[0] for s in STATUS]].sum(axis=1)
    fig = go.Figure()
    for key, ar, col in STATUS:
        fig.add_bar(y=MAIN_REGS, x=d[key] / tot * 100, orientation="h", name=ar, marker_color=col,
                    customdata=d[key], text=d[key] / tot * 100, texttemplate="%{text:.1f}%",
                    hovertemplate="%{y} — " + ar + "<br>%{customdata:,.0f} منشأة (%{x:.1f}%)<extra></extra>")
    style(fig, 300)
    rtl_hbar(fig)
    fig.update_layout(barmode="stack")
    fig.update_xaxes(range=[100, 0], ticksuffix="%")
    fig.update_yaxes(autorange="reversed")
    return fig


def fig_op_closed():
    fig = go.Figure()
    for i, r in enumerate(MAIN_REGS):
        ys, vs = [], []
        for y in OP_YEARS:
            d = OP[(OP.year == y) & (OP.reg_ar == r)].set_index("key").est
            if len(d) and "Enumerated" in d:
                ys.append(str(y))
                vs.append(d.get("Temporarily closed", 0) / d["Enumerated"] * 100)
        fig.add_bar(x=ys, y=vs, name=r, marker_color=[PRIMARY, WB_COLOR, GAZA_COLOR][i] if i else DARK,
                    text=vs, texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
    fig.update_layout(barmode="group")
    fig.update_yaxes(ticksuffix="%", showgrid=True)
    return style(fig, 300)


# ════════════════════════════════════════════════════════════════════
# 11) بناء الرسوم — التطور التاريخي
# ════════════════════════════════════════════════════════════════════
H = D[D.group == "history"].copy()
H_YEARS = sorted(int(y) for y in H.year.unique())
HM = {"est": "عدد المنشآت", "pe": "إجمالي العاملين", "fem": "نسبة الإناث %"}


def fig_hist_sector(metric):
    d = H[~TOT.loc[H.index]].copy()
    d["fem"] = fem_share(d.male, d.female)
    if d.empty:
        return empty_fig()
    last = d[d.year == H_YEARS[-1]].sort_values(metric)
    order = list(last.cat_ar)
    fig = go.Figure()
    for i, y in enumerate(H_YEARS):
        s = d[d.year == y].set_index("cat_ar").reindex(order)
        fig.add_bar(y=[wrap(o, 28) for o in order], x=s[metric], orientation="h", name=str(y),
                    marker_color=GREENS[min(i * 2, 5)],
                    hovertemplate="%{y} — " + str(y) + "<br>" + HM[metric] + ": %{x:,.1f}<extra></extra>")
    style(fig, 620)
    rtl_hbar(fig)
    fig.update_layout(barmode="group")
    fig.update_xaxes(range=[None, 0])
    return fig


def fig_hist_total():
    t = H[TOT.loc[H.index]].sort_values("year")
    fig = make_subplots(rows=1, cols=2, subplot_titles=("عدد المنشآت", "إجمالي العاملين"))
    for c, (col, color) in enumerate((("est", PRIMARY), ("pe", BLUE)), start=1):
        fig.add_bar(x=t.year.astype(str), y=t[col], marker_color=color, showlegend=False,
                    text=t[col], texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False, row=1, col=c)
        fig.update_yaxes(showticklabels=False, showgrid=False, range=[0, t[col].max() * 1.2], row=1, col=c)
    fig.update_annotations(font=dict(family=FONT, size=14))
    return style(fig, 340, legend=False)


def fig_hist_regions():
    d = D[D.group == "regional_history"].sort_values("year")
    fig = go.Figure()
    for r, col in (("الضفة الغربية", WB_COLOR), ("قطاع غزة", GAZA_COLOR)):
        s = d[d.reg_ar == r]
        fig.add_bar(x=s.year.astype(str), y=s.est, name=r, marker_color=col, text=s.est,
                    texttemplate="%{text:,.0f}", textposition="inside", textfont=dict(color="white"))
    fig.update_layout(barmode="stack")
    fig.update_yaxes(showticklabels=False, showgrid=False)
    return style(fig, 340)


# ════════════════════════════════════════════════════════════════════
# 12) واجهة المستخدم (Layout)
# ════════════════════════════════════════════════════════════════════
CSS = """
:root{--primary:#0B6E4F;--dark:#1F2D3D;--muted:#6B7785;--bg:#F3F6F8;--line:#E6EAEE}
*{box-sizing:border-box}
html,body{margin:0;background:var(--bg);color:var(--dark);direction:rtl;
  font-family:'Cairo','Segoe UI',Tahoma,Arial,sans-serif}
.header{background:linear-gradient(120deg,#0A3B2C 0%,#0B6E4F 60%,#1F8F6B 100%);color:#fff;padding:26px 34px 22px}
.header h1{margin:0 0 6px;font-size:27px;font-weight:800}
.header p{margin:0;opacity:.88;font-size:14px}
.wrap{max-width:1520px;margin:0 auto;padding:0 20px 36px}
.tabs-parent{margin-top:18px}
.custom-tab,.tab{padding:11px 18px!important;font-family:inherit!important;font-weight:600!important;font-size:14px!important;
  color:var(--muted)!important;background:transparent!important;border:none!important;
  border-bottom:3px solid transparent!important;cursor:pointer}
.custom-tab--sel,.tab--selected{color:var(--primary)!important;border-bottom:3px solid var(--primary)!important;
  background:#fff!important;border-radius:10px 10px 0 0!important}
.tabs-container,.tab-container{border-bottom:1px solid var(--line)!important;flex-wrap:wrap}
.pane{padding-top:18px}
.controls{display:flex;flex-wrap:wrap;gap:18px;background:#fff;border:1px solid var(--line);border-radius:14px;
  padding:14px 18px;margin-bottom:16px;align-items:flex-end}
.ctl{display:flex;flex-direction:column;gap:5px}
.ctl label{font-size:12.5px;font-weight:700;color:var(--muted)}
.dd{font-family:inherit;font-size:14px;direction:rtl}
.grid{display:grid;gap:16px;margin-bottom:16px}
.g2{grid-template-columns:repeat(2,minmax(0,1fr))}
.g3{grid-template-columns:repeat(3,minmax(0,1fr))}
.g21{grid-template-columns:minmax(0,2fr) minmax(0,1fr)}
.g12{grid-template-columns:minmax(0,1fr) minmax(0,2fr)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:16px}
.kpi{background:#fff;border:1px solid var(--line);border-top:4px solid var(--primary);border-radius:14px;padding:14px 16px}
.kpi-label{font-size:13px;color:var(--muted);font-weight:600}
.kpi-value{font-size:28px;font-weight:800;margin:4px 0 2px}
.kpi-sub{font-size:12px;color:var(--muted)}
.card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:16px 18px;overflow:hidden}
.card-head{display:flex;align-items:baseline;gap:10px;margin-bottom:6px;flex-wrap:wrap}
.card-head h3{margin:0;font-size:16.5px;font-weight:800}
.card-sub{font-size:12px;color:var(--muted)}
.insight{background:#EAF4EF;border-right:4px solid var(--primary);border-radius:10px;padding:12px 16px;
  margin-bottom:16px;font-size:14px;line-height:1.9}
.note{font-size:12.5px;color:var(--muted);line-height:1.9;background:#fff;border:1px dashed #cfd6dd;
  border-radius:12px;padding:14px 18px;margin-top:8px}
.note b{color:var(--dark)}
.btn{font-family:inherit;font-weight:700;background:var(--primary);color:#fff;border:none;border-radius:10px;
  padding:9px 18px;cursor:pointer;font-size:14px}
.btn:hover{filter:brightness(1.1)}
@media(max-width:1000px){.g2,.g3,.g21,.g12{grid-template-columns:1fr}.header h1{font-size:21px}}
"""

app = Dash(__name__, title="لوحة مؤشرات المنشآت الاقتصادية — فلسطين", suppress_callback_exceptions=False)
app.index_string = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
{{%metas%}}
<title>{{%title%}}</title>
{{%favicon%}}
{{%css%}}
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
{{%app_entry%}}
<footer>{{%config%}}{{%scripts%}}{{%renderer%}}</footer>
</body>
</html>"""

yr_opts = lambda ys: [{"label": f"تعداد {y}", "value": y} for y in ys]
reg_opts = lambda rs: [{"label": r, "value": r} for r in rs]


def tab(label, value, *children):
    return dcc.Tab(label=label, value=value, className="custom-tab", selected_className="custom-tab--sel",
                   children=html.Div(list(children), className="pane"))


tab_overview = tab(
    "نظرة عامة", "ov",
    controls(("المنطقة", dd("ov-region", ALL_REGS, "فلسطين", 220)),
             ("سنة التعداد", dd("ov-year", yr_opts(OWN_YEARS), OWN_YEARS[-1], 160))),
    html.Div(id="ov-kpis", className="kpis"),
    html.Div([card("تطور عدد المنشآت العاملة", graph("ov-trend", 380), "حسب بيانات الملكية"),
              card("توزيع المنشآت حسب الملكية", graph("ov-own", 380))], className="grid g2"),
    html.Div([card("الضفة الغربية مقابل قطاع غزة", graph("ov-wg", 380)),
              card("العاملون حسب الجنس", graph("ov-gender", 380), "تعدادا 2012 و2017")], className="grid g2"),
)

tab_gov = tab(
    "المحافظات", "gv",
    controls(("سنة التعداد", dd("gv-year", yr_opts(GOV_YEARS), GOV_YEARS[-1], 160)),
             ("المؤشر", dd("gv-metric", [{"label": v[0], "value": k} for k, v in METRICS.items()], "est", 260)),
             ("النطاق", dd("gv-scope", [{"label": "كل المحافظات", "value": "all"},
                                        {"label": "الضفة الغربية", "value": "wb"},
                                        {"label": "قطاع غزة", "value": "gaza"}], "all", 180))),
    html.Div([card("ترتيب المحافظات", graph("gv-rank", 520)),
              card("حصة كل محافظة من المنشآت", graph("gv-donut", 520))], className="grid g21"),
    html.Div([card("خريطة حرارية لتطور المنشآت (1997–2023)", graph("gv-heat", 520),
                   "اللون = الحجم النسبي داخل كل محافظة"),
              card("نمو المنشآت والعاملين بين التعدادين", graph("gv-growth", 520))], className="grid g2"),
    card("جدول المحافظات", dash_table.DataTable(
        id="gv-table", sort_action="native", page_size=20,
        columns=[{"name": "المحافظة", "id": "reg_ar"}, {"name": "المنطقة", "id": "area"}] + [
            {"name": n, "id": i, "type": "numeric",
             "format": Format(group=Group.yes, precision=p, scheme=Scheme.fixed)}
            for n, i, p in (("المنشآت", "est", 0), ("العاملون", "pe", 0), ("الإناث %", "fem", 1),
                            ("متوسط العاملين/منشأة", "avg", 2), ("نمو المنشآت % عن التعداد السابق", "growth", 1))],
        **{"style_table": {"overflowX": "auto", "direction": "rtl"}, "style_cell": {
            "textAlign": "right", "fontFamily": "Cairo, Tahoma, sans-serif", "padding": "9px 12px", "fontSize": 14},
           "style_header": {"backgroundColor": DARK, "color": "white", "fontWeight": "700", "border": "none"},
           "style_data_conditional": [{"if": {"row_index": "odd"}, "backgroundColor": "#F7F9FA"}]})),
)

tab_act = tab(
    "الأنشطة الاقتصادية", "ac",
    controls(("سنة التعداد", dd("ac-year", yr_opts(ACT_YEARS), ACT_YEARS[-1], 160)),
             ("المنطقة", dd("ac-region", ALL_REGS, "فلسطين", 220)),
             ("عدد الأنشطة المعروضة", dd("ac-top", [5, 8, 10, 15, 20], 10, 120))),
    html.Div(id="ac-insight"),
    html.Div([card("أكبر الأنشطة حسب عدد المنشآت", graph("ac-bar", 440)),
              card("الهيكل النسبي للأنشطة", graph("ac-tree", 440))], className="grid g2"),
    html.Div([card("مقارنة بين التعدادات", graph("ac-years", 520)),
              card("نسبة التغير بين أول وآخر تعداد", graph("ac-change", 520),
                   "للأنشطة التي لا يقل عدد منشآتها عن 50 في التعداد الأول")], className="grid g2"),
)

tab_emp = tab(
    "العمالة والنوع الاجتماعي", "em",
    controls(("سنة التعداد", dd("em-year", yr_opts(EMP_YEARS), EMP_YEARS[-1], 160)),
             ("المنطقة", dd("em-region", MAIN_REGS, "فلسطين", 200))),
    html.Div(id="em-kpis", className="kpis"),
    html.Div([card("العاملون حسب النشاط والجنس", graph("em-stack", 560)),
              card("نسبة الإناث حسب النشاط", graph("em-fem", 560), "الوردي = أعلى من المتوسط العام")],
             className="grid g2"),
    html.Div([card("متوسط العاملين لكل منشأة", graph("em-avg", 560), "الذهبي = أعلى من المتوسط العام"),
              card("خريطة الأنشطة: الحجم مقابل مشاركة المرأة", graph("em-bubble", 560),
                   "حجم الفقاعة = عدد العاملين")], className="grid g2"),
)

tab_cls = tab(
    "الملكية والشكل القانوني", "ow",
    controls(("التصنيف", dd("ow-class", [{"label": v, "value": k} for k, v in CLASS_OPTS.items()], "ownership", 220)),
             ("المنطقة", dd("ow-region", ALL_REGS, "فلسطين", 220)),
             ("سنة التعداد", dd("ow-year", yr_opts(OWN_YEARS), OWN_YEARS[-1], 160)),
             ("المقياس", dcc.RadioItems(id="ow-log", options=[{"label": " عادي", "value": 0},
                                                              {"label": " لوغاريتمي", "value": 1}],
                                        value=0, inline=True, inputStyle={"marginLeft": "5px", "marginRight": "14px"}))),
    html.Div([card("التوزيع حسب التصنيف المختار", graph("ow-bar", 460), "المقياس اللوغاريتمي يُظهر الفئات الصغيرة"),
              card("تطور الحصص النسبية عبر التعدادات", graph("ow-share", 460))], className="grid g2"),
)

tab_size = tab(
    "حجم المنشآت", "sz",
    controls(("سنة التعداد", dd("sz-year", yr_opts(SZ_YEARS), SZ_YEARS[-1], 160)),
             ("ترتيب الأنشطة حسب", dd("sz-sort", [{"label": "عدد المنشآت", "value": "count"},
                                                 {"label": "نسبة المنشآت الصغيرة جداً", "value": "micro"}],
                                     "count", 240))),
    html.Div(id="sz-kpis", className="kpis"),
    html.Div([card("توزيع المنشآت حسب فئات حجم العمالة", graph("sz-dist", 380), "الضفة الغربية"),
              card("هيكل الحجم داخل كل نشاط", graph("sz-stack", 560), "نسبة مئوية — الضفة الغربية")],
             className="grid g12"),
)

tab_op = tab(
    "الحالة التشغيلية", "op",
    controls(("سنة التعداد", dd("op-year", yr_opts(OP_YEARS), OP_YEARS[-1], 160))),
    html.Div([card("المنشآت المحصورة حسب الحالة التشغيلية", graph("op-stack", 300)),
              card("نسبة المنشآت المتوقفة مؤقتاً", graph("op-closed", 300), "من إجمالي المنشآت المحصورة")],
             className="grid g2"),
)

tab_hist = tab(
    "التطور التاريخي 1997–2007", "hi",
    controls(("المؤشر", dd("hi-metric", [{"label": v, "value": k} for k, v in HM.items()], "est", 220))),
    html.Div([card("الإجماليات عبر التعدادات", graph("hi-total", 340), "التصنيف القديم للأنشطة"),
              card("الضفة الغربية وقطاع غزة", graph("hi-reg", 340), "عدد المنشآت")], className="grid g2"),
    card("القطاعات الاقتصادية عبر التعدادات", graph("hi-sector", 620)),
)

D_DISPLAY = D.rename(columns={"group": "المجموعة", "key": "المفتاح", "cat_ar": "الفئة", "cat_en": "Category",
                              "reg_ar": "المنطقة", "reg_en": "Region", "year": "السنة", "est": "المنشآت",
                              "pe": "إجمالي العاملين", "male": "الذكور", "female": "الإناث"})
D_DISPLAY["المجموعة"] = D_DISPLAY["المجموعة"].map(GROUP_AR).fillna(D_DISPLAY["المجموعة"])

tab_data = tab(
    "البيانات", "dt",
    controls(("المجموعة", dd("dt-group", [{"label": "الكل", "value": "all"}] +
                             [{"label": v, "value": v} for v in D_DISPLAY["المجموعة"].unique()], "all", 240)),
             ("السنة", dd("dt-year", [{"label": "الكل", "value": 0}] + [{"label": str(y), "value": y}
                                                                       for y in sorted(D.year.unique())], 0, 140)),
             ("المنطقة", dd("dt-region", [{"label": "الكل", "value": "all"}] + reg_opts(ALL_REGS), "all", 220)),
             ("", html.Button("⬇ تنزيل CSV", id="dt-btn", className="btn"))),
    html.Div(id="dt-count", className="insight"),
    card("البيانات التفصيلية", dash_table.DataTable(
        id="dt-table", page_size=18, sort_action="native", filter_action="native",
        columns=[{"name": c, "id": c} for c in D_DISPLAY.columns],
        style_table={"overflowX": "auto", "direction": "rtl"},
        style_cell={"textAlign": "right", "fontFamily": "Cairo, Tahoma, sans-serif", "padding": "8px 10px",
                    "fontSize": 13, "minWidth": 70, "maxWidth": 260, "overflow": "hidden", "textOverflow": "ellipsis"},
        style_header={"backgroundColor": DARK, "color": "white", "fontWeight": "700", "border": "none"},
        style_data_conditional=[{"if": {"row_index": "odd"}, "backgroundColor": "#F7F9FA"}])),
    dcc.Download(id="dt-dl"),
)

NOTES = html.Div([
    html.B("ملاحظات منهجية: "),
    "(1) تختلف مجاميع المنشآت بين الجداول لاختلاف التغطية والتصنيف في الملف نفسه "
    "(مثلاً 2012: 144,969 في جدول الملكية والأنشطة مقابل 135,401 في جدول المحافظات والعمالة)، لذلك لا تُقارَن الأرقام عبر الجداول مباشرة. ",
    "(2) نسبة الإناث تُحسب من (الذكور + الإناث) لأن إجمالي العاملين لا يساوي مجموعهما في بعض الصفوف "
    "(كالقدس وفلسطين والضفة)، إذ لا يتوفر التوزيع حسب الجنس لجزء من المنشآت. ",
    "(3) بيانات العاملين متوفرة لتعدادي 2012 و2017 فقط (إضافة للسلسلة التاريخية 1997–2007)."
], className="note")

app.layout = html.Div([
    html.Div([html.H1("لوحة مؤشرات المنشآت الاقتصادية في فلسطين"),
              html.P(f"تحليل تفاعلي لتعدادات المنشآت {OWN_YEARS[0]} – {OWN_YEARS[-1]} · "
                     f"{len(D):,} سجلاً في {D.group.nunique()} جدولاً · الضفة الغربية وقطاع غزة و{len(GOVS)} محافظة")],
             className="header"),
    html.Div([
        dcc.Tabs(id="tabs", value="ov", parent_className="tabs-parent", className="tabs-container",
                 children=[tab_overview, tab_gov, tab_act, tab_emp, tab_cls, tab_size, tab_op, tab_hist, tab_data]),
        NOTES,
    ], className="wrap"),
])


# ════════════════════════════════════════════════════════════════════
# 13) الاستدعاءات التفاعلية (Callbacks)
# ════════════════════════════════════════════════════════════════════
@app.callback(Output("ov-kpis", "children"), Output("ov-trend", "figure"), Output("ov-own", "figure"),
              Output("ov-wg", "figure"), Output("ov-gender", "figure"),
              Input("ov-region", "value"), Input("ov-year", "value"))
def cb_overview(region, year):
    return (build_kpis(region, year), fig_trend(region), fig_ownership_snapshot(region, year),
            fig_west_gaza(), fig_gender(region))


@app.callback(Output("gv-rank", "figure"), Output("gv-donut", "figure"), Output("gv-heat", "figure"),
              Output("gv-growth", "figure"), Output("gv-table", "data"),
              Input("gv-year", "value"), Input("gv-metric", "value"), Input("gv-scope", "value"))
def cb_gov(year, metric, scope):
    return (fig_gov_rank(year, metric, scope), fig_gov_donut(year, scope), fig_gov_heat(scope),
            fig_gov_growth(scope), gov_table_data(year, scope))


@app.callback(Output("ac-insight", "children"), Output("ac-bar", "figure"), Output("ac-tree", "figure"),
              Output("ac-years", "figure"), Output("ac-change", "figure"),
              Input("ac-year", "value"), Input("ac-region", "value"), Input("ac-top", "value"))
def cb_activity(year, region, topn):
    return (act_insight(year, region), fig_act_bar(year, region, topn), fig_act_tree(year, region),
            fig_act_years(region, topn, year), fig_act_change(region))


@app.callback(Output("em-kpis", "children"), Output("em-stack", "figure"), Output("em-fem", "figure"),
              Output("em-avg", "figure"), Output("em-bubble", "figure"),
              Input("em-year", "value"), Input("em-region", "value"))
def cb_employment(year, region):
    return (emp_kpis(year, region), fig_emp_stack(year, region),
            fig_emp_metric(year, region, "fem", "نسبة الإناث", "%", ".1f"),
            fig_emp_metric(year, region, "avg", "متوسط العاملين", "", ".2f"),
            fig_emp_bubble(year, region))


@app.callback(Output("ow-year", "options"), Output("ow-year", "value"),
              Input("ow-class", "value"), State("ow-year", "value"))
def cb_class_years(group, current):
    ys = class_years(group)
    return yr_opts(ys), (current if current in ys else ys[-1])


@app.callback(Output("ow-bar", "figure"), Output("ow-share", "figure"),
              Input("ow-class", "value"), Input("ow-region", "value"),
              Input("ow-year", "value"), Input("ow-log", "value"))
def cb_class(group, region, year, log):
    if year not in class_years(group):
        year = class_years(group)[-1]
    return fig_class_bar(group, region, year, bool(log)), fig_class_share(group, region)


@app.callback(Output("sz-kpis", "children"), Output("sz-dist", "figure"), Output("sz-stack", "figure"),
              Input("sz-year", "value"), Input("sz-sort", "value"))
def cb_size(year, sort):
    return size_kpis(year), fig_size_dist(), fig_size_stack(year, sort)


@app.callback(Output("op-stack", "figure"), Output("op-closed", "figure"), Input("op-year", "value"))
def cb_op(year):
    return fig_op_stack(year), fig_op_closed()


@app.callback(Output("hi-total", "figure"), Output("hi-reg", "figure"), Output("hi-sector", "figure"),
              Input("hi-metric", "value"))
def cb_hist(metric):
    return fig_hist_total(), fig_hist_regions(), fig_hist_sector(metric)


def filtered_data(group, year, region):
    d = D_DISPLAY
    if group != "all":
        d = d[d["المجموعة"] == group]
    if year:
        d = d[d["السنة"] == year]
    if region != "all":
        d = d[d["المنطقة"] == region]
    return d


@app.callback(Output("dt-table", "data"), Output("dt-count", "children"),
              Input("dt-group", "value"), Input("dt-year", "value"), Input("dt-region", "value"))
def cb_table(group, year, region):
    d = filtered_data(group, year, region)
    return d.replace({np.nan: None}).to_dict("records"), f"عدد السجلات المعروضة: {len(d):,} من أصل {len(D_DISPLAY):,}"


@app.callback(Output("dt-dl", "data"), Input("dt-btn", "n_clicks"),
              State("dt-group", "value"), State("dt-year", "value"), State("dt-region", "value"),
              prevent_initial_call=True)
def cb_download(_, group, year, region):
    return dcc.send_data_frame(filtered_data(group, year, region).to_csv, "palestine_establishments.csv",
                               index=False, encoding="utf-8-sig")


# Flask server exposed for production WSGI servers such as Gunicorn/Render.
server = app.server

if __name__ == "__main__":
    import os
    app.run(
        debug=False,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8050)),
    )
