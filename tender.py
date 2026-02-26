"""
Дашборд аналитики строительных тендеров
Запуск: streamlit run tender_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime, timedelta
import random

# ─── Конфигурация страницы ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Тендерная аналитика",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Стили ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=IBM+Plex+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}
h1, h2, h3 { font-family: 'Rajdhani', sans-serif !important; }

.main { background-color: #0d1117; }
.block-container { padding: 1.5rem 2rem; }

.metric-card {
    background: linear-gradient(135deg, #161b22 0%, #1c2333 100%);
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 0.5rem;
    box-shadow: 0 4px 24px rgba(0,0,0,0.4);
}
.metric-label {
    color: #8b949e;
    font-size: 0.78rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
.metric-value {
    color: #e6edf3;
    font-size: 2rem;
    font-family: 'Rajdhani', sans-serif;
    font-weight: 700;
    line-height: 1.2;
}
.metric-delta-pos { color: #3fb950; font-size: 0.85rem; }
.metric-delta-neg { color: #f85149; font-size: 0.85rem; }

.section-title {
    font-family: 'Rajdhani', sans-serif;
    font-size: 1.4rem;
    font-weight: 700;
    color: #e6edf3;
    border-left: 4px solid #1f6feb;
    padding-left: 0.8rem;
    margin: 1.5rem 0 1rem 0;
}
.stSelectbox label, .stMultiSelect label, .stDateInput label {
    color: #8b949e !important;
    font-size: 0.8rem !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}
</style>
""", unsafe_allow_html=True)

# ─── Генерация тестовых данных ────────────────────────────────────────────────
@st.cache_data
def generate_sample_data(n=350):
    random.seed(42)
    np.random.seed(42)

    customers = [
        "ГКУ «Мосстройинвест»", "АО «Росатом»", "ООО «СтройГрупп»",
        "ФГУП «РосСтрой»", "ПАО «Газпром»", "МКУ «УКС г. Казань»",
        "АО «Самараинвест»", "ГКУ «УКС Татарстан»", "ООО «Уралстрой»",
        "ФГКУ «Росгвардия»", "МУП «СтройСервис»", "АО «РЖД Инфрастрой»",
    ]
    work_types = [
        "Общестроительные работы", "Дорожное строительство",
        "Инженерные сети", "Капитальный ремонт", "Реконструкция",
        "Благоустройство", "Монтаж конструкций", "Земляные работы",
    ]
    regions = [
        "Москва", "Санкт-Петербург", "Татарстан", "Свердловская обл.",
        "Самарская обл.", "Краснодарский край", "Башкортостан",
        "Нижегородская обл.", "Новосибирская обл.", "Ростовская обл.",
    ]

    start = datetime(2021, 1, 1)
    dates = [start + timedelta(days=random.randint(0, 3*365)) for _ in range(n)]

    statuses = random.choices(["Выигран", "Проигран", "В процессе"], weights=[0.38, 0.45, 0.17], k=n)
    base_sums = np.random.lognormal(mean=7.5, sigma=1.2, size=n) * 100_000

    df = pd.DataFrame({
        "Дата подачи": dates,
        "Заказчик": random.choices(customers, k=n),
        "Тип работ": random.choices(work_types, k=n),
        "Регион": random.choices(regions, k=n),
        "Сумма тендера (тенге.)": base_sums.astype(int),
        "Статус": statuses,
        "Номер тендера": [f"Т-{random.randint(100000, 999999)}" for _ in range(n)],
        "Срок (дней)": [random.randint(30, 730) for _ in range(n)],
    })

    df["Дата подачи"] = pd.to_datetime(df["Дата подачи"])
    df["Год"] = df["Дата подачи"].dt.year
    df["Месяц"] = df["Дата подачи"].dt.to_period("M").dt.to_timestamp()
    df["Год-Месяц"] = df["Дата подачи"].dt.strftime("%Y-%m")
    return df


# ─── Загрузка данных ──────────────────────────────────────────────────────────
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    # Пытаемся определить нужные колонки
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if any(k in cl for k in ["дата", "date"]): col_map[col] = "Дата подачи"
        elif any(k in cl for k in ["заказчик", "customer", "client"]): col_map[col] = "Заказчик"
        elif any(k in cl for k in ["тип", "вид", "work", "type"]): col_map[col] = "Тип работ"
        elif any(k in cl for k in ["регион", "region", "oblast"]): col_map[col] = "Регион"
        elif any(k in cl for k in ["сумм", "цена", "amount", "price", "sum"]): col_map[col] = "Сумма тендера (тенге.)"
        elif any(k in cl for k in ["статус", "status", "результ"]): col_map[col] = "Статус"
    df = df.rename(columns=col_map)
    if "Дата подачи" in df.columns:
        df["Дата подачи"] = pd.to_datetime(df["Дата подачи"], errors="coerce")
    if "Год" not in df.columns and "Дата подачи" in df.columns:
        df["Год"] = df["Дата подачи"].dt.year
        df["Месяц"] = df["Дата подачи"].dt.to_period("M").dt.to_timestamp()
        df["Год-Месяц"] = df["Дата подачи"].dt.strftime("%Y-%m")
    return df


# ─── Цветовая схема Plotly ────────────────────────────────────────────────────
LAYOUT = dict(
    paper_bgcolor="#161b22",
    plot_bgcolor="#0d1117",
    font=dict(family="IBM Plex Sans", color="#c9d1d9"),
    title_font=dict(family="Rajdhani", size=18, color="#e6edf3"),
    legend=dict(bgcolor="#161b22", bordercolor="#30363d", borderwidth=1),
    margin=dict(l=40, r=20, t=50, b=40),
)
STATUS_COLORS = {
    "Выигран": "#3fb950",
    "Проигран": "#f85149",
    "В процессе": "#d29922",
}

def apply_layout(fig, **kwargs):
    fig.update_layout(**LAYOUT, **kwargs)
    fig.update_xaxes(gridcolor="#21262d", zerolinecolor="#30363d", tickfont_size=11)
    fig.update_yaxes(gridcolor="#21262d", zerolinecolor="#30363d", tickfont_size=11)
    return fig


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏗️ Тендерная аналитика")
    st.markdown("---")

    uploaded = st.file_uploader("📂 Загрузить Excel", type=["xlsx", "xls"],
                                 help="Загрузите ваш файл с тендерами")

    st.markdown("---")
    st.markdown("**Фильтры**")

    if uploaded:
        raw_df = load_excel(uploaded)
    else:
        st.info("Используются демо-данные")
        raw_df = generate_sample_data()

    years = sorted(raw_df["Год"].dropna().unique().astype(int).tolist())
    sel_years = st.multiselect("Год", years, default=years)

    statuses = raw_df["Статус"].dropna().unique().tolist()
    sel_statuses = st.multiselect("Статус", statuses, default=statuses)

    if "Регион" in raw_df.columns:
        regions = sorted(raw_df["Регион"].dropna().unique().tolist())
        sel_regions = st.multiselect("Регион", regions, default=regions)
    else:
        sel_regions = None

    if "Тип работ" in raw_df.columns:
        work_types_all = sorted(raw_df["Тип работ"].dropna().unique().tolist())
        sel_works = st.multiselect("Тип работ", work_types_all, default=work_types_all)
    else:
        sel_works = None

    st.markdown("---")
    # Кнопка скачать шаблон
    template = pd.DataFrame({
        "Дата подачи": ["2024-01-15"],
        "Заказчик": ["ООО Пример"],
        "Тип работ": ["Дорожное строительство"],
        "Регион": ["Москва"],
        "Сумма тендера (тенге.)": [5000000],
        "Статус": ["Выигран"],
        "Номер тендера": ["Т-123456"],
        "Срок (дней)": [180],
    })
    buf = io.BytesIO()
    template.to_excel(buf, index=False)
    st.download_button("⬇️ Скачать шаблон Excel", buf.getvalue(),
                       "tender_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─── Фильтрация ───────────────────────────────────────────────────────────────
df = raw_df.copy()
if sel_years:
    df = df[df["Год"].isin(sel_years)]
if sel_statuses:
    df = df[df["Статус"].isin(sel_statuses)]
if sel_regions and "Регион" in df.columns:
    df = df[df["Регион"].isin(sel_regions)]
if sel_works and "Тип работ" in df.columns:
    df = df[df["Тип работ"].isin(sel_works)]

won = df[df["Статус"] == "Выигран"]
lost = df[df["Статус"] == "Проигран"]

# ─── Заголовок ────────────────────────────────────────────────────────────────
st.markdown("# 🏗️ Дашборд строительных тендеров")
st.markdown(f"*Данные: {len(df):,} тендеров · Период: {df['Дата подачи'].min().strftime('%d.%m.%Y') if not df.empty else '—'} – {df['Дата подачи'].max().strftime('%d.%m.%Y') if not df.empty else '—'}*")

# ─── KPI карточки ─────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)

total_sum = df["Сумма тендера (тенге.)"].sum()
won_sum = won["Сумма тендера (тенге.)"].sum()
win_rate = len(won) / max(len(df[df["Статус"] != "В процессе"]), 1) * 100
avg_won = won["Сумма тендера (тенге.)"].mean() if len(won) > 0 else 0

def kpi(col, label, value, delta=None, delta_pos=True):
    delta_html = ""
    if delta is not None:
        cls = "metric-delta-pos" if delta_pos else "metric-delta-neg"
        arrow = "▲" if delta_pos else "▼"
        delta_html = f'<div class="{cls}">{arrow} {delta}</div>'
    col.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>""", unsafe_allow_html=True)

kpi(k1, "Всего тендеров", f"{len(df):,}")
kpi(k2, "Выиграно", f"{len(won):,}", f"из {len(df[df['Статус']!='В процессе'])}", True)
kpi(k3, "Конверсия", f"{win_rate:.1f}%")
kpi(k4, "Сумма выигранных", f"{won_sum/1e6:.1f} млн ₸")
kpi(k5, "Средний выигранный", f"{avg_won/1e6:.1f} млн ₸")

st.markdown("---")

# ─── График 1: Сумма выигранных по месяцам ───────────────────────────────────
st.markdown('<div class="section-title">📈 Динамика выигранных тендеров</div>', unsafe_allow_html=True)

col_g1, col_g2 = st.columns([3, 2])

with col_g1:
    granularity = st.radio("Группировка", ["Месяц", "Квартал", "Год"], horizontal=True, key="gran")

    if not won.empty:
        if granularity == "Месяц":
            monthly = won.groupby("Месяц").agg(
                Сумма=("Сумма тендера (тенге.)", "sum"),
                Количество=("Номер тендера" if "Номер тендера" in won.columns else "Статус", "count")
            ).reset_index().rename(columns={"Месяц": "Период"})
        elif granularity == "Квартал":
            won2 = won.copy()
            won2["Квартал"] = won2["Дата подачи"].dt.to_period("Q").dt.to_timestamp()
            monthly = won2.groupby("Квартал").agg(
                Сумма=("Сумма тендера (тенге.)", "sum"),
                Количество=("Статус", "count")
            ).reset_index().rename(columns={"Квартал": "Период"})
        else:
            monthly = won.groupby("Год").agg(
                Сумма=("Сумма тендера (тенге.)", "sum"),
                Количество=("Статус", "count")
            ).reset_index().rename(columns={"Год": "Период"})

        fig1 = make_subplots(specs=[[{"secondary_y": True}]])
        fig1.add_trace(go.Bar(
            x=monthly["Период"], y=monthly["Сумма"] / 1e6,
            name="Сумма (млн ₸)", marker_color="#1f6feb",
            marker_line_width=0, opacity=0.85,
        ), secondary_y=False)
        fig1.add_trace(go.Scatter(
            x=monthly["Период"], y=monthly["Количество"],
            name="Количество", line=dict(color="#3fb950", width=2),
            mode="lines+markers", marker=dict(size=5),
        ), secondary_y=True)
        fig1.update_yaxes(title_text="Сумма (млн ₸)", secondary_y=False, title_font_size=11)
        fig1.update_yaxes(title_text="Количество", secondary_y=True, title_font_size=11,
                          showgrid=False)
        apply_layout(fig1, title="Выигранные тендеры по периодам")
        st.plotly_chart(fig1, use_container_width=True)

with col_g2:
    # Donut: статусы
    status_cnt = df["Статус"].value_counts().reset_index()
    status_cnt.columns = ["Статус", "Количество"]
    colors = [STATUS_COLORS.get(s, "#8b949e") for s in status_cnt["Статус"]]

    fig_d = go.Figure(go.Pie(
        labels=status_cnt["Статус"],
        values=status_cnt["Количество"],
        hole=0.62,
        marker=dict(colors=colors, line=dict(color="#0d1117", width=3)),
        textfont=dict(size=12),
    ))
    fig_d.add_annotation(text=f"<b>{win_rate:.0f}%</b>", x=0.5, y=0.5,
                          font=dict(size=26, family="Rajdhani", color="#e6edf3"),
                          showarrow=False)
    fig_d.add_annotation(text="конверсия", x=0.5, y=0.38,
                          font=dict(size=11, color="#8b949e"), showarrow=False)
    apply_layout(fig_d, title="Распределение статусов")
    st.plotly_chart(fig_d, use_container_width=True)

# ─── График 2: По заказчикам ──────────────────────────────────────────────────
st.markdown('<div class="section-title">🏢 Аналитика по заказчикам</div>', unsafe_allow_html=True)

col_c1, col_c2 = st.columns(2)

with col_c1:
    top_n = st.slider("Топ заказчиков", 5, 20, 10, key="topn")
    by_cust = df.groupby("Заказчик").agg(
        Всего=("Сумма тендера (тенге.)", "sum"),
        Выиграно=("Статус", lambda x: (x == "Выигран").sum()),
        Проиграно=("Статус", lambda x: (x == "Проигран").sum()),
    ).nlargest(top_n, "Всего").reset_index()
    by_cust["Заказчик_short"] = by_cust["Заказчик"].str[:25]

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(y=by_cust["Заказчик_short"], x=by_cust["Выиграно"],
                           name="Выиграно", orientation="h", marker_color="#3fb950"))
    fig2.add_trace(go.Bar(y=by_cust["Заказчик_short"], x=by_cust["Проиграно"],
                           name="Проиграно", orientation="h", marker_color="#f85149"))
    apply_layout(fig2, title=f"Тендеры по заказчикам (топ {top_n})",
                  barmode="stack", height=400,
                  xaxis_title="Количество тендеров")
    st.plotly_chart(fig2, use_container_width=True)

with col_c2:
    # Scatter: сумма vs количество
    cust_scatter = df.groupby("Заказчик").agg(
        Сумма=("Сумма тендера (тенге.)", "sum"),
        Количество=("Статус", "count"),
        Конверсия=("Статус", lambda x: (x == "Выигран").sum() / max(len(x[x != "В процессе"]), 1) * 100),
    ).reset_index()

    fig3 = px.scatter(
        cust_scatter, x="Количество", y="Сумма",
        size="Конверсия", color="Конверсия",
        hover_name="Заказчик",
        color_continuous_scale=["#f85149", "#d29922", "#3fb950"],
        size_max=30,
        labels={"Сумма": "Сумма тендеров (тенге.)", "Конверсия": "Конверсия %"},
    )
    apply_layout(fig3, title="Заказчики: объём vs количество", height=400)
    fig3.update_traces(marker=dict(line=dict(width=1, color="#0d1117")))
    st.plotly_chart(fig3, use_container_width=True)

# ─── График 3: Типы работ ─────────────────────────────────────────────────────
st.markdown('<div class="section-title">🔨 Типы работ и регионы</div>', unsafe_allow_html=True)

col_w1, col_w2 = st.columns([2, 3])

with col_w1:
    work_stat = df.groupby(["Тип работ", "Статус"]).size().reset_index(name="n")
    fig4 = px.bar(work_stat, x="n", y="Тип работ", color="Статус",
                  color_discrete_map=STATUS_COLORS, orientation="h",
                  barmode="group",
                  labels={"n": "Количество", "Тип работ": ""})
    apply_layout(fig4, title="Тендеры по типу работ", height=420)
    st.plotly_chart(fig4, use_container_width=True)

with col_w2:
    if "Регион" in df.columns and "Тип работ" in df.columns:
        heatmap_df = won.groupby(["Тип работ", "Регион"])["Сумма тендера (тенге.)"].sum().reset_index()
        heatmap_pivot = heatmap_df.pivot(index="Тип работ", columns="Регион",
                                          values="Сумма тендера (тенге.)").fillna(0)
        heatmap_vals = heatmap_pivot.values / 1e6

        fig5 = go.Figure(go.Heatmap(
            z=heatmap_vals,
            x=heatmap_pivot.columns.tolist(),
            y=heatmap_pivot.index.tolist(),
            colorscale=[
                [0, "#0d1117"], [0.3, "#1f3a6b"],
                [0.6, "#1f6feb"], [1.0, "#58a6ff"],
            ],
            text=np.round(heatmap_vals, 1),
            texttemplate="%{text}",
            textfont=dict(size=10),
            hovertemplate="<b>%{y}</b><br>%{x}<br>%{z:.1f} млн ₸<extra></extra>",
            colorbar=dict(title="млн ₸", tickfont=dict(size=10)),
        ))
        apply_layout(fig5, title="Heatmap: сумма выигранных (млн ₸) — тип × регион",
                      height=420, xaxis_tickangle=-30)
        st.plotly_chart(fig5, use_container_width=True)

# ─── График 4: Воронка / Winrate по типам ─────────────────────────────────────
st.markdown('<div class="section-title">📊 Эффективность по типам работ</div>', unsafe_allow_html=True)

col_e1, col_e2 = st.columns(2)

with col_e1:
    eff = df[df["Статус"] != "В процессе"].groupby("Тип работ").apply(
        lambda x: pd.Series({
            "Win Rate %": round((x["Статус"] == "Выигран").sum() / max(len(x), 1) * 100, 1),
            "Ср. сумма выигранных (млн ₸)": round(
                x[x["Статус"] == "Выигран"]["Сумма тендера (тенге.)"].mean() / 1e6, 2)
        })
    ).reset_index()
    eff = eff.sort_values("Win Rate %", ascending=True)

    fig6 = go.Figure(go.Bar(
        x=eff["Win Rate %"], y=eff["Тип работ"],
        orientation="h",
        marker=dict(
            color=eff["Win Rate %"],
            colorscale=["#f85149", "#d29922", "#3fb950"],
            cmin=0, cmax=100,
            line=dict(width=0),
        ),
        text=[f"{v}%" for v in eff["Win Rate %"]],
        textposition="outside",
    ))
    apply_layout(fig6, title="Конверсия (Win Rate) по типу работ", height=400,
                  xaxis_range=[0, 110], xaxis_title="Win Rate %")
    st.plotly_chart(fig6, use_container_width=True)

with col_e2:
    # Treemap по заказчикам и типам (go.Treemap — без зависимости от narwhals)
    tm_df = won.copy()
    tm_df["Сумма_млн"] = (tm_df["Сумма тендера (тенге.)"] / 1e6).round(2)

    # Строим иерархию вручную: root → тип работ → заказчик
    type_totals = tm_df.groupby("Тип работ")["Сумма_млн"].sum()
    cust_totals = tm_df.groupby(["Тип работ", "Заказчик"])["Сумма_млн"].sum()

    labels, parents, values = ["Все тендеры"], [""], [tm_df["Сумма_млн"].sum()]

    for wtype, wsum in type_totals.items():
        labels.append(str(wtype))
        parents.append("Все тендеры")
        values.append(float(wsum))

    for (wtype, cust), csum in cust_totals.items():
        labels.append(f"{cust} ({wtype[:12]})")
        parents.append(str(wtype))
        values.append(float(csum))

    # Цвет — глубина узла (0=root, 1=тип, 2=заказчик)
    depth_colors = []
    for p in parents:
        if p == "":
            depth_colors.append(0)
        elif p == "Все тендеры":
            depth_colors.append(1)
        else:
            depth_colors.append(2)

    fig7 = go.Figure(go.Treemap(
        labels=labels,
        parents=parents,
        values=values,
        marker=dict(
            colors=depth_colors,
            colorscale=[[0, "#0d1117"], [0.5, "#1f6feb"], [1.0, "#58a6ff"]],
            line=dict(width=1.5, color="#0d1117"),
        ),
        textfont=dict(size=11, color="#e6edf3"),
        hovertemplate="<b>%{label}</b><br>%{value:.1f} млн ₸<extra></extra>",
        branchvalues="total",
        maxdepth=2,
    ))
    apply_layout(fig7, title="Treemap выигранных: тип работ → заказчик", height=400)
    st.plotly_chart(fig7, use_container_width=True)

# ─── Таблица ──────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📋 Таблица тендеров</div>', unsafe_allow_html=True)

show_cols = [c for c in ["Номер тендера", "Дата подачи", "Заказчик", "Тип работ",
                          "Регион", "Сумма тендера (тенге.)", "Статус", "Срок (дней)"]
             if c in df.columns]

table_df = df[show_cols].copy().sort_values("Дата подачи", ascending=False)

def color_status(val):
    colors_map = {"Выигран": "#0d2b17", "Проигран": "#2b0d0d", "В процессе": "#2b2200"}
    return f"background-color: {colors_map.get(val, '')}"

if "Статус" in table_df.columns:
    styled = table_df.style.applymap(color_status, subset=["Статус"])
    if "Сумма тендера (тенге.)" in table_df.columns:
        styled = styled.format({"Сумма тендера (тенге.)": "{:,.0f} ₸"})
    st.dataframe(styled, use_container_width=True, height=350)
else:
    st.dataframe(table_df, use_container_width=True, height=350)

# Экспорт
buf_out = io.BytesIO()
df[show_cols].to_excel(buf_out, index=False)
st.download_button("⬇️ Экспорт текущей выборки", buf_out.getvalue(),
                   "filtered_tenders.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.markdown("<center style='color:#484f58; font-size:0.75rem'>Тендерная аналитика · Построено на Streamlit + Plotly</center>",
            unsafe_allow_html=True)