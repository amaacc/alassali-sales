"""
charts.py — Professional Business Intelligence Visualizations for DailySalesApp
Built with Plotly Express & Plotly Graph Objects
Designed for: Al-Assali Trading Co. (شركة العسلي للتجارة)
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from typing import List, Dict, Any, Optional

# Corporate Color Palette
NAVY_BLUE = "#1F4E79"
ROYAL_BLUE = "#2F5597"
SKY_BLUE = "#0284C7"
EMERALD_GREEN = "#059669"
CRIMSON_RED = "#DC2626"
AMBER_GOLD = "#D97706"
SLATE_GRAY = "#64748B"
MUTED_GRAY = "#94A3B8"
CARD_BG = "#FFFFFF"

COMMON_LAYOUT = dict(
    font=dict(family="'Cairo', 'Segoe UI', Tahoma, sans-serif", size=12),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=20, r=20, t=50, b=30),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="center",
        x=0.5,
        font=dict(size=11)
    )
)

PLOTLY_CONFIG = {
    'displayModeBar': True,
    'displaylogo': False,
    'responsive': True,
    'modeBarButtonsToRemove': ['lasso2d', 'select2d', 'autoScale2d']
}


# =============================================================================
# 1. PERIODIC SALES CHARTS (تقارير الفترات الزمنية)
# =============================================================================

def build_periodic_horizontal_bar(rows: List[Dict[str, Any]], top_n: int = 15, title: str = "") -> go.Figure:
    """Horizontal ranked bar chart with gradient coloring."""
    display_rows = rows[:top_n]
    df = pd.DataFrame(display_rows)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")
    # Sort ascending for horizontal bar so #1 is at the top
    df = df.iloc[::-1].reset_index(drop=True)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df["net"],
        y=df["clean_name"],
        orientation="h",
        marker=dict(
            color=df["net"],
            colorscale=[[0, "#93C5FD"], [0.5, "#2563EB"], [1, "#1E3A8A"]],
            showscale=False
        ),
        text=[f"{val:,.0f} ر.س" for val in df["net"]],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>صافي المبيعات: %{x:,.2f} ر.س<extra></extra>"
    ))

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or f"أعلى {len(display_rows)} معرضاً مبيعاً (مرتبة تنازلياً)", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title="صافي المبيعات (ريال)", gridcolor="#E2E8F0", zeroline=True, zerolinecolor="#CBD5E1"),
        yaxis=dict(title="", tickfont=dict(size=11)),
        height=max(400, len(display_rows) * 28 + 100)
    )
    fig.update_layout(layout)
    return fig


def build_periodic_donut_chart(rows: List[Dict[str, Any]], grand_total_net: float, max_slices: int = 8, title: str = "") -> go.Figure:
    """Donut chart showing market share & revenue contribution percentage."""
    if not rows:
        return go.Figure()

    top_slices = rows[:max_slices]
    other_slices = rows[max_slices:]

    labels = [r["showroom"].replace("معرض ", "") for r in top_slices]
    values = [r["net"] for r in top_slices]

    if other_slices:
        other_sum = sum(r["net"] for r in other_slices)
        labels.append(f"بقية المعارض ({len(other_slices)})")
        values.append(round(other_sum, 2))

    colors = [
        "#1F4E79", "#2F5597", "#3B82F6", "#06B6D4", "#10B981",
        "#8B5CF6", "#F59E0B", "#EC4899", "#94A3B8"
    ]

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=colors[:len(labels)]),
        textinfo="label+percent",
        textposition="outside",
        hovertemplate="<b>%{label}</b><br>الصافي: %{value:,.2f} ر.س<br>الحصة: %{percent}<extra></extra>"
    )])

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or "توزيع الحصص السوقية لصافي مبيعات المعارض", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        annotations=[
            dict(
                text=f"الإجمالي<br><b>{grand_total_net:,.0f}</b><br>ر.س",
                x=0.5, y=0.5,
                font=dict(size=12, color=NAVY_BLUE),
                showarrow=False
            )
        ],
        height=450,
        showlegend=False
    )
    fig.update_layout(layout)
    return fig


def build_periodic_treemap(rows: List[Dict[str, Any]], title: str = "") -> go.Figure:
    """Interactive Treemap hierarchy showing relative showroom weights."""
    df = pd.DataFrame(rows)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")
    df["formatted_net"] = df["net"].apply(lambda v: f"{v:,.2f} ر.س")

    fig = px.treemap(
        df,
        path=["clean_name"],
        values="net",
        color="net",
        color_continuous_scale=[[0, "#E0F2FE"], [0.5, "#0284C7"], [1, "#0369A1"]],
        title=title or "خريطة الشجرة لتوزيع الوزن النسبي لمبيعات المعارض"
    )

    fig.update_traces(
        texttemplate="<b>%{label}</b><br>%{value:,.0f} ر.س<br>(%{percentRoot:.1%})",
        hovertemplate="<b>%{label}</b><br>صافي المبيعات: %{value:,.2f} ر.س<br>الحصة من الشركة: %{percentRoot:.2%}<extra></extra>"
    )

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or "خريطة الشجرة لتوزيع الوزن النسبي لمبيعات المعارض", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        coloraxis_showscale=False,
        height=480
    )
    fig.update_layout(layout)
    return fig


def build_periodic_pareto(rows: List[Dict[str, Any]], title: str = "") -> go.Figure:
    """Pareto Chart (80/20 Rule) combining net sales bars with cumulative contribution % curve."""
    df = pd.DataFrame(rows)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")
    total_net = df["net"].sum()
    df["cum_pct"] = (df["net"].cumsum() / total_net) * 100.0

    # Display top 20 for readability
    df_disp = df.head(20)

    fig = go.Figure()

    # Bar trace (Net sales)
    fig.add_trace(go.Bar(
        x=df_disp["clean_name"],
        y=df_disp["net"],
        name="صافي المبيعات",
        marker_color="#2F5597",
        hovertemplate="<b>%{x}</b><br>الصافي: %{y:,.2f} ر.س<extra></extra>",
        yaxis="y"
    ))

    # Line trace (Cumulative %)
    fig.add_trace(go.Scatter(
        x=df_disp["clean_name"],
        y=df_disp["cum_pct"],
        name="النسبة التراكمية %",
        marker=dict(color="#DC2626", size=7),
        line=dict(color="#DC2626", width=2.5),
        hovertemplate="<b>%{x}</b><br>التراكمي: %{y:.1f}%<extra></extra>",
        yaxis="y2"
    ))

    # 80% threshold line
    fig.add_hline(
        y=80, line_dash="dash", line_color="#D97706",
        annotation_text="عتبة 80% من المبيعات (قاعدة باريتو)",
        annotation_position="top left",
        yref="y2"
    )

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or "منحنى باريتو التراكمي (تحليل قاعدة 80/20 لكبار المعارض)", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title="", tickangle=-45, tickfont=dict(size=10)),
        yaxis=dict(title="صافي المبيعات (ريال)", gridcolor="#E2E8F0"),
        yaxis2=dict(title="النسبة التراكمية (%)", overlaying="y", side="right", range=[0, 105], showgrid=False),
        height=480
    )
    fig.update_layout(layout)
    return fig


def build_periodic_area_curve(rows: List[Dict[str, Any]], title: str = "") -> go.Figure:
    """Smooth area curve visualizing revenue drop-off and concentration."""
    df = pd.DataFrame(rows)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["clean_name"],
        y=df["net"],
        mode="lines+markers",
        name="صافي المبيعات",
        fill="tozeroy",
        line=dict(color="#2563EB", width=3, shape="spline"),
        fillcolor="rgba(37, 99, 235, 0.15)",
        marker=dict(size=6, color="#1D4ED8"),
        hovertemplate="<b>%{x}</b><br>الصافي: %{y:,.2f} ر.س<extra></extra>"
    ))

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or "منحنى توزيع وانحدار مبيعات المعارض من الأعلى إلى الأقل", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title="", tickangle=-45, tickfont=dict(size=10)),
        yaxis=dict(title="صافي المبيعات (ريال)", gridcolor="#E2E8F0"),
        height=450
    )
    fig.update_layout(layout)
    return fig


# =============================================================================
# 2. COMPARATIVE SALES CHARTS (تقارير المقارنات ونسب النمو)
# =============================================================================

def build_comparative_grouped_bars(
    rows: List[Dict[str, Any]],
    current_label: str,
    previous_label: str,
    top_n: int = 15,
    title: str = ""
) -> go.Figure:
    """Side-by-side grouped bars comparing Net Sales across both periods."""
    disp = rows[:top_n]
    df = pd.DataFrame(disp)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df["clean_name"],
        y=df["previous_net"],
        name=f"صافي {previous_label}",
        marker_color="#94A3B8",
        hovertemplate="<b>%{x}</b><br>" + f"صافي {previous_label}: " + "%{y:,.2f} ر.س<extra></extra>"
    ))
    fig.add_trace(go.Bar(
        x=df["clean_name"],
        y=df["current_net"],
        name=f"صافي {current_label}",
        marker_color="#1F4E79",
        hovertemplate="<b>%{x}</b><br>" + f"صافي {current_label}: " + "%{y:,.2f} ر.س<extra></extra>"
    ))

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        barmode="group",
        title=dict(text=title or f"مقارنة مبيعات أعلى {len(disp)} معرضاً بين الفترتين", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title="", tickangle=-45, tickfont=dict(size=10)),
        yaxis=dict(title="صافي المبيعات (ريال)", gridcolor="#E2E8F0"),
        height=480
    )
    fig.update_layout(layout)
    return fig


def build_comparative_diverging_bars(
    rows: List[Dict[str, Any]],
    current_label: str,
    previous_label: str,
    top_n: int = 20,
    title: str = ""
) -> go.Figure:
    """
    Diverging Variance Bar Chart:
    Green bars extending to the right for growth, Red bars extending to the left for decline.
    """
    # Sort by diff descending
    sorted_rows = sorted(rows, key=lambda x: x["diff"], reverse=True)
    # Take top performers and bottom performers for max contrast
    if len(sorted_rows) > top_n:
        half = top_n // 2
        selected = sorted_rows[:half] + sorted_rows[-half:]
    else:
        selected = sorted_rows

    df = pd.DataFrame(selected)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")
    df = df.iloc[::-1].reset_index(drop=True)

    colors = [EMERALD_GREEN if d >= 0 else CRIMSON_RED for d in df["diff"]]
    text_labels = [f"{d:+,.0f} ر.س ({pct:+.1f}%)" for d, pct in zip(df["diff"], df["growth_pct"])]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df["diff"],
        y=df["clean_name"],
        orientation="h",
        marker_color=colors,
        text=text_labels,
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>الفارق: %{x:+,.2f} ر.س<extra></extra>"
    ))

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or f"مخطط الفوارق المالية ومقدار التغير (أعلى نمو مقابل أكبر انخفاض)", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title="الفارق المالي (ريال) — الأخضر نمو / الأحمر انخفاض", gridcolor="#E2E8F0", zeroline=True, zerolinecolor="#475569", zerolinewidth=2),
        yaxis=dict(title="", tickfont=dict(size=11)),
        height=max(450, len(selected) * 26 + 100)
    )
    fig.update_layout(layout)
    return fig


def build_comparative_growth_matrix(
    rows: List[Dict[str, Any]],
    current_label: str,
    previous_label: str,
    title: str = ""
) -> go.Figure:
    """
    Growth Matrix / Quadrant Chart:
    X-axis: Previous Net Sales (Volume)
    Y-axis: Growth Rate %
    Bubble Size: Current Net Sales
    """
    df = pd.DataFrame(rows)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")

    # Cap visual growth outliers for better chart zoom (-100% to +200%)
    df["visual_growth"] = df["growth_pct"].clip(lower=-100, upper=250)
    colors = [EMERALD_GREEN if d >= 0 else CRIMSON_RED for d in df["diff"]]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["previous_net"],
        y=df["visual_growth"],
        mode="markers+text",
        text=df["clean_name"],
        textposition="top center",
        textfont=dict(size=9),
        marker=dict(
            size=df["current_net"],
            sizemode="area",
            sizeref=2.*max(df["current_net"])/(45.**2),
            sizemin=8,
            color=colors,
            opacity=0.8,
            line=dict(width=1, color="#1E293B")
        ),
        hovertemplate="<b>%{text}</b><br>" +
                      f"صافي {previous_label}: " + "%{x:,.2f} ر.س<br>" +
                      f"صافي {current_label}: " + "%{customdata[0]:,.2f} ر.س<br>" +
                      "الفارق: %{customdata[1]:+,.2f} ر.س<br>" +
                      "نسبة النمو: %{y:+.1f}%<extra></extra>",
        customdata=df[["current_net", "diff"]].values
    ))

    # Add zero-growth horizontal line
    fig.add_hline(y=0, line_dash="solid", line_color="#475569", line_width=1.5)

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or "مصفوفة كفاءة النمو مقابل حجم المبيعات (حجم الفقاعة يمثل المبيعات الحالية)", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title=f"حجم مبيعات فترة الأساس ({previous_label}) بالريال", gridcolor="#E2E8F0"),
        yaxis=dict(title="معدل النمو (%)", gridcolor="#E2E8F0"),
        height=520
    )
    fig.update_layout(layout)
    return fig


def build_comparative_dumbbell(
    rows: List[Dict[str, Any]],
    current_label: str,
    previous_label: str,
    top_n: int = 15,
    title: str = ""
) -> go.Figure:
    """
    Dumbbell Chart:
    Connects Previous Period point to Current Period point with a directional line.
    """
    disp = rows[:top_n]
    df = pd.DataFrame(disp)
    df["clean_name"] = df["showroom"].str.replace("معرض ", "")
    df = df.iloc[::-1].reset_index(drop=True)

    fig = go.Figure()

    # Connector lines
    for _, r in df.iterrows():
        line_color = EMERALD_GREEN if r["diff"] >= 0 else CRIMSON_RED
        fig.add_trace(go.Scatter(
            x=[r["previous_net"], r["current_net"]],
            y=[r["clean_name"], r["clean_name"]],
            mode="lines",
            line=dict(color=line_color, width=2.5),
            showlegend=False,
            hoverinfo="skip"
        ))

    # Previous points
    fig.add_trace(go.Scatter(
        x=df["previous_net"],
        y=df["clean_name"],
        mode="markers",
        name=f"سابقاً ({previous_label})",
        marker=dict(color="#94A3B8", size=10, symbol="circle"),
        hovertemplate="<b>%{y}</b><br>السابق: %{x:,.2f} ر.س<extra></extra>"
    ))

    # Current points
    fig.add_trace(go.Scatter(
        x=df["current_net"],
        y=df["clean_name"],
        mode="markers",
        name=f"حالياً ({current_label})",
        marker=dict(
            color=[EMERALD_GREEN if d >= 0 else CRIMSON_RED for d in df["diff"]],
            size=12,
            symbol="diamond"
        ),
        hovertemplate="<b>%{y}</b><br>الحالي: %{x:,.2f} ر.س<extra></extra>"
    ))

    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text=title or f"مخطط قفزات المبيعات (Dumbbell Chart) لأعلى {len(disp)} معرضاً", x=0.5, font=dict(size=14, color=NAVY_BLUE)),
        xaxis=dict(title="صافي المبيعات (ريال)", gridcolor="#E2E8F0"),
        yaxis=dict(title="", tickfont=dict(size=11)),
        height=max(450, len(disp) * 28 + 100)
    )
    fig.update_layout(layout)
    return fig


# =============================================================================
# 3. DAILY SALES CHARTS (المبيعات اليومية)
# =============================================================================

def build_daily_payment_donut(gt: Dict[str, Any]) -> go.Figure:
    """Donut chart for payment method distribution (شبكة vs نقد vs تحويل)."""
    labels = ["الشبكة (مدى/بطاقات)", "النقد (كاش)", "التحويل البنكي"]
    values = [gt.get("credit", 0.0), gt.get("cash", 0.0), gt.get("trans", 0.0)]
    colors = ["#2563EB", "#059669", "#D97706"]

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=colors),
        textinfo="label+percent",
        textposition="outside",
        hovertemplate="<b>%{label}</b><br>المبلغ: %{value:,.2f} ر.س<br>النسبة: %{percent}<extra></extra>"
    )])

    tot_payments = sum(values)
    layout = dict(**COMMON_LAYOUT)
    layout.update(
        title=dict(text="توزيع طرق تحصيل المبيعات اليومية", x=0.5, font=dict(size=13, color=NAVY_BLUE)),
        annotations=[
            dict(
                text=f"إجمالي التحصيل<br><b>{tot_payments:,.0f}</b><br>ر.س",
                x=0.5, y=0.5,
                font=dict(size=11, color=NAVY_BLUE),
                showarrow=False
            )
        ],
        height=350,
        showlegend=True
    )
    fig.update_layout(layout)
    return fig
