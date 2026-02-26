"""
Дашборд аналитики строительных тендеров — Pro Version
Запуск: streamlit run tender.py
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
import json

# ─── Конфигурация страницы ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Тендерная аналитика Pro",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Стили ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=IBM+Plex+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
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
.metric-label { color: #8b949e; font-size: 0.78rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.08em; }
.metric-value { color: #e6edf3; font-size: 2rem; font-family: 'Rajdhani', sans-serif; font-weight: 700; line-height: 1.2; }
.metric-delta-pos { color: #3fb950; font-size: 0.85rem; }
.metric-delta-neg { color: #f85149; font-size: 0.85rem; }
.metric-delta-neu { color: #d29922; font-size: 0.85rem; }

.section-title {
    font-family: 'Rajdhani', sans-serif;
    font-size: 1.4rem; font-weight: 700; color: #e6edf3;
    border-left: 4px solid #1f6feb;
    padding-left: 0.8rem; margin: 1.5rem 0 1rem 0;
}

.score-pill-green {
    display:inline-block; background:#0d2b17; color:#3fb950;
    border:1px solid #3fb950; border-radius:20px;
    padding:2px 12px; font-size:0.78rem; font-weight:700;
}
.score-pill-yellow {
    display:inline-block; background:#2b1e00; color:#d29922;
    border:1px solid #d29922; border-radius:20px;
    padding:2px 12px; font-size:0.78rem; font-weight:700;
}
.score-pill-red {
    display:inline-block; background:#2b0d0d; color:#f85149;
    border:1px solid #f85149; border-radius:20px;
    padding:2px 12px; font-size:0.78rem; font-weight:700;
}

.insight-box {
    background: linear-gradient(135deg, #0d2040 0%, #0d1117 100%);
    border: 1px solid #1f6feb;
    border-radius: 10px; padding: 1rem 1.2rem; margin-bottom: 0.6rem;
}
.insight-title { color: #58a6ff; font-size: 0.85rem; font-weight: 600; }
.insight-body { color: #c9d1d9; font-size: 0.82rem; margin-top: 0.3rem; }

.whatif-card {
    background: linear-gradient(135deg, #1a1f2e 0%, #161b22 100%);
    border: 1px solid #30363d; border-radius: 10px;
    padding: 1rem 1.2rem; margin-top: 0.4rem;
}

.currency-badge {
    display:inline-block; background:#1f3a6b;
    color:#58a6ff; border-radius:6px;
    padding:2px 10px; font-size:0.8rem; font-weight:600; margin-left:6px;
}

.deadline-card {
    background: linear-gradient(135deg, #2b0d0d 0%, #1a0505 100%);
    border: 1px solid #f85149; border-radius: 8px; padding: 0.8rem 1rem; margin-bottom: 0.4rem;
}
.deadline-card-warn {
    background: linear-gradient(135deg, #2b1e00 0%, #1a1200 100%);
    border: 1px solid #d29922; border-radius: 8px; padding: 0.8rem 1rem; margin-bottom: 0.4rem;
}
.deadline-title { color: #e6edf3; font-size: 0.85rem; font-weight: 500; }
.deadline-meta { color: #8b949e; font-size: 0.75rem; margin-top: 0.2rem; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# СЛОВАРЬ КООРДИНАТ КАЗАХСТАНА
# Ключи — названия регионов/городов в вашей БД (Регион)
# Значения — (широта, долгота)
# Fallback (если регион не найден): центр Казахстана (48.02, 66.92)
# ═══════════════════════════════════════════════════════════
KZ_COORDS = {
    # ── Города-миллионники и областные центры ──────────────
    "Астана":                           (51.1801, 71.4460),
    "Нур-Султан":                       (51.1801, 71.4460),   # старое название
    "Алматы":                           (43.2220, 76.8512),
    "Шымкент":                          (42.3000, 69.6000),
    "Актобе":                           (50.2839, 57.1670),
    "Атырау":                           (47.1167, 51.8833),
    "Усть-Каменогорск":                 (49.9489, 82.6285),
    "Павлодар":                         (52.2873, 76.9674),
    "Семей":                            (50.4114, 80.2275),
    "Тараз":                            (42.9000, 71.3667),
    "Костанай":                         (53.2144, 63.6244),
    "Кызылорда":                        (44.8479, 65.5093),
    "Уральск":                          (51.2333, 51.3833),
    "Петропавловск":                    (54.8647, 69.1386),
    "Актау":                            (43.6511, 51.1680),
    "Кокшетау":                         (53.2833, 69.3833),
    "Талдыкорган":                      (45.0000, 78.3833),
    "Туркестан":                        (43.2975, 68.2728),
    "Экибастуз":                        (51.7167, 75.3167),
    "Темиртау":                         (50.0578, 72.9619),
    "Рудный":                           (52.9667, 63.1167),
    "Жезказган":                        (47.7833, 67.7667),
    "Балхаш":                           (46.8480, 74.9950),
    "Сатпаев":                          (47.9000, 67.5333),
    "Жанаозен":                         (43.3347, 52.8585),
    "Кентау":                           (43.5167, 68.5000),
    "Степногорск":                      (52.3500, 71.8833),
    "Риддер":                           (50.3483, 83.5131),
    "Лисаковск":                        (52.6500, 62.5000),
    "Аркалык":                          (50.2500, 66.9000),
    "Байконур":                         (45.9647, 63.3052),

    # ── Области (по центру области) ─────────────────────────
    "Акмолинская область":              (51.5000, 69.5000),
    "Актюбинская область":              (50.2839, 57.1670),
    "Алматинская область":              (43.5000, 77.0000),
    "Атырауская область":               (47.1167, 51.8833),
    "Восточно-Казахстанская область":   (49.9489, 82.6285),
    "ВКО":                              (49.9489, 82.6285),
    "Жамбылская область":               (42.9000, 71.3667),
    "Западно-Казахстанская область":    (51.2333, 51.3833),
    "ЗКО":                              (51.2333, 51.3833),
    "Карагандинская область":           (49.8046, 73.1094),
    "Костанайская область":             (53.2144, 63.6244),
    "Кызылординская область":           (44.8479, 65.5093),
    "Мангистауская область":            (43.6511, 51.1680),
    "Павлодарская область":             (52.2873, 76.9674),
    "Северо-Казахстанская область":     (54.8647, 69.1386),
    "СКО":                              (54.8647, 69.1386),
    "Туркестанская область":            (43.2975, 68.2728),
    "Абайская область":                 (49.9489, 80.0000),
    "Жетысуская область":               (44.5000, 79.0000),
    "Улытауская область":               (48.6262, 67.7270),

    # ── Районы и малые города (частые в тендерах) ──────────
    "Алматинская обл.":                 (43.5000, 77.0000),
    "Карагандинская обл.":              (49.8046, 73.1094),
    "Павлодарская обл.":                (52.2873, 76.9674),
    "Костанайская обл.":                (53.2144, 63.6244),
    "Актюбинская обл.":                 (50.2839, 57.1670),
    "Акмолинская обл.":                 (51.5000, 69.5000),
    "Жамбылская обл.":                  (42.9000, 71.3667),
    "Кызылординская обл.":              (44.8479, 65.5093),
    "Мангистауская обл.":               (43.6511, 51.1680),
    "Туркестанская обл.":               (43.2975, 68.2728),
    "Восточно-Казахстанская обл.":      (49.9489, 82.6285),
    "Западно-Казахстанская обл.":       (51.2333, 51.3833),
    "Северо-Казахстанская обл.":        (54.8647, 69.1386),
}

# Центр Казахстана — fallback если регион не найден
KZ_CENTER = (48.0196, 66.9237)


def get_coords(region_name: str) -> tuple:
    """Возвращает (lat, lon) для региона. Нечёткий поиск по подстроке."""
    if not isinstance(region_name, str):
        return KZ_CENTER
    # Точное совпадение
    if region_name in KZ_COORDS:
        return KZ_COORDS[region_name]
    # Поиск по подстроке (например "г. Алматы" → "Алматы")
    for key, coords in KZ_COORDS.items():
        if key.lower() in region_name.lower() or region_name.lower() in key.lower():
            return coords
    return KZ_CENTER

# ═══════════════════════════════════════════════════════════
# ГЕНЕРАЦИЯ ДЕМО-ДАННЫХ
# ═══════════════════════════════════════════════════════════
@st.cache_data
def generate_sample_data(n=350):
    random.seed(42); np.random.seed(42)

    customers = [
        "ГКУ «Мосстройинвест»","АО «Росатом»","ООО «СтройГрупп»",
        "ФГУП «РосСтрой»","ПАО «Газпром»","МКУ «УКС г. Казань»",
        "АО «Самараинвест»","ГКУ «УКС Татарстан»","ООО «Уралстрой»",
        "ФГКУ «Росгвардия»","МУП «СтройСервис»","АО «РЖД Инфрастрой»",
    ]
    work_types = [
        "Общестроительные работы","Дорожное строительство","Инженерные сети",
        "Капитальный ремонт","Реконструкция","Благоустройство",
        "Монтаж конструкций","Земляные работы",
    ]
    regions = [
        "Астана","Шымкент","Алматы","Северо-Казахстанская обл.","Карагандинская обл.",
        "Жамбылская обл.","Туркестанская обл.","Алматинская обл.","Актюбинская обл.","Акмолинская обл.",
        "Кызылординская обл.","Восточно-Казахстанская обл.","Мангистауская обл.","Павлодарская обл.","Западно-Казахстанская обл.","Туркестанская обл.",
    ]
    competitors = [
        "ООО «АльфаСтрой»","АО «МегаСтрой»","ЗАО «СтройМастер»",
        "ООО «ГрандСтрой»","АО «ТехноСтрой»","ООО «КапСтрой»", None, None,
    ]
    region_coords = {
        "Москва":(55.75,37.62),"Санкт-Петербург":(59.95,30.32),
        "Татарстан":(55.80,49.18),"Свердловская обл.":(56.84,60.60),
        "Самарская обл.":(53.20,50.15),"Краснодарский край":(45.04,38.98),
        "Башкортостан":(54.74,55.97),"Нижегородская обл.":(56.33,44.00),
        "Новосибирская обл.":(54.99,82.90),"Ростовская обл.":(47.23,39.72),
    }
    work_margin = {
        "Общестроительные работы":0.18,"Дорожное строительство":0.22,
        "Инженерные сети":0.28,"Капитальный ремонт":0.15,"Реконструкция":0.20,
        "Благоустройство":0.12,"Монтаж конструкций":0.25,"Земляные работы":0.17,
    }

    start = datetime(2021, 1, 1)
    dates = [start + timedelta(days=random.randint(0, 3*365)) for _ in range(n)]
    statuses = random.choices(["Выигран","Проигран","В процессе"], weights=[0.38,0.45,0.17], k=n)
    base_sums = np.random.lognormal(mean=7.5, sigma=1.2, size=n) * 100_000

    # НМЦК (начальная максимальная цена) — базовая цена до снижения
    nmck_factor = np.random.uniform(1.05, 1.30, n)
    nmck = base_sums * nmck_factor
    price_drop_pct = ((nmck - base_sums) / nmck * 100).round(1)

    work_types_list = random.choices(work_types, k=n)
    regions_list = random.choices(regions, k=n)
    margins = np.array([work_margin[wt] + np.random.uniform(-0.05, 0.05) for wt in work_types_list])
    costs = base_sums * (1 - margins)
    currencies = random.choices(["KZT","RUB"], weights=[0.65,0.35], k=n)
    deadline_days = [random.randint(-2,30) if s=="В процессе" else None for s in statuses]

    df = pd.DataFrame({
        "Дата подачи": dates,
        "Заказчик": random.choices(customers, k=n),
        "Тип работ": work_types_list,
        "Регион": regions_list,
        "Сумма тендера": base_sums.astype(int),
        "НМЦК": nmck.astype(int),
        "Снижение %": price_drop_pct,
        "Себестоимость": costs.astype(int),
        "Статус": statuses,
        "Номер тендера": [f"Т-{random.randint(100000,999999)}" for _ in range(n)],
        "Срок (дней)": [random.randint(30,730) for _ in range(n)],
        "Конкурент": random.choices(competitors, k=n),
        "Валюта": currencies,
        "Дней до дедлайна": deadline_days,
    })
    df["Дата подачи"] = pd.to_datetime(df["Дата подачи"])
    df["Год"] = df["Дата подачи"].dt.year
    df["Месяц"] = df["Дата подачи"].dt.to_period("M").dt.to_timestamp()
    df["Год-Месяц"] = df["Дата подачи"].dt.strftime("%Y-%m")
    df["Чистая прибыль"] = df["Сумма тендера"] - df["Себестоимость"]
    df["Маржа %"] = ((df["Чистая прибыль"] / df["Сумма тендера"]) * 100).round(1)
    # ── Координаты через словарь ──────────────────────────
    df["Лат"] = df["Регион"].apply(lambda r: get_coords(r)[0])
    df["Лон"] = df["Регион"].apply(lambda r: get_coords(r)[1])
    # Небольшой разброс, чтобы точки одного региона не сливались
    df["Лат"] += np.random.uniform(-0.4, 0.4, n)
    df["Лон"] += np.random.uniform(-0.4, 0.4, n)
    return df


def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if any(k in cl for k in ["дата","date"]): col_map[col]="Дата подачи"
        elif any(k in cl for k in ["заказчик","customer"]): col_map[col]="Заказчик"
        elif any(k in cl for k in ["тип","вид","work","type"]): col_map[col]="Тип работ"
        elif any(k in cl for k in ["регион","region"]): col_map[col]="Регион"
        elif any(k in cl for k in ["нмцк","nmck","начальн"]): col_map[col]="НМЦК"
        elif any(k in cl for k in ["себест","cost"]): col_map[col]="Себестоимость"
        elif any(k in cl for k in ["сумм","цена","amount","price","sum"]): col_map[col]="Сумма тендера"
        elif any(k in cl for k in ["статус","status"]): col_map[col]="Статус"
        elif any(k in cl for k in ["конкур","competitor"]): col_map[col]="Конкурент"
        elif any(k in cl for k in ["валют","currency"]): col_map[col]="Валюта"
    df = df.rename(columns=col_map)
    if "Дата подачи" in df.columns:
        df["Дата подачи"] = pd.to_datetime(df["Дата подачи"], errors="coerce")
    if "Год" not in df.columns and "Дата подачи" in df.columns:
        df["Год"] = df["Дата подачи"].dt.year
        df["Месяц"] = df["Дата подачи"].dt.to_period("M").dt.to_timestamp()
        df["Год-Месяц"] = df["Дата подачи"].dt.strftime("%Y-%m")
    if "Себестоимость" not in df.columns and "Сумма тендера" in df.columns:
        df["Себестоимость"] = df["Сумма тендера"] * 0.82
    if "Чистая прибыль" not in df.columns:
        df["Чистая прибыль"] = df["Сумма тендера"] - df["Себестоимость"]
    if "Маржа %" not in df.columns:
        df["Маржа %"] = ((df["Чистая прибыль"] / df["Сумма тендера"]) * 100).round(1)
    if "НМЦК" not in df.columns and "Сумма тендера" in df.columns:
        df["НМЦК"] = (df["Сумма тендера"] * 1.15).astype(int)
    if "Снижение %" not in df.columns:
        df["Снижение %"] = ((df["НМЦК"] - df["Сумма тендера"]) / df["НМЦК"] * 100).round(1)
    if "Валюта" not in df.columns:
        df["Валюта"] = "KZT"
    
    # ── Координаты через словарь (работает для любого Excel) ──
    if "Регион" in df.columns:
        df["Лат"] = df["Регион"].apply(lambda r: get_coords(r)[0])
        df["Лон"] = df["Регион"].apply(lambda r: get_coords(r)[1])
        # Разброс только если в регионе несколько записей
        df["Лат"] += np.random.uniform(-0.3, 0.3, len(df))
        df["Лон"] += np.random.uniform(-0.3, 0.3, len(df))
    return df


# ═══════════════════════════════════════════════════════════
# СКОРИНГ: алгоритм перспективности
# ═══════════════════════════════════════════════════════════
def compute_scoring(df: pd.DataFrame) -> pd.DataFrame:
    """Вычисляет балл перспективности (0–100) для каждого тендера."""
    won = df[df["Статус"] == "Выигран"]

    # 1. Конверсия по заказчику
    cust_wr = df[df["Статус"] != "В процессе"].groupby("Заказчик").apply(
        lambda x: (x["Статус"] == "Выигран").sum() / max(len(x), 1)
    ).to_dict()

    # 2. Средняя маржа по типу работ (выигранные)
    type_margin = won.groupby("Тип работ")["Маржа %"].mean().to_dict() if "Маржа %" in won.columns else {}

    # 3. Топ-3 региона по win rate
    reg_wr = df[df["Статус"] != "В процессе"].groupby("Регион").apply(
        lambda x: (x["Статус"] == "Выигран").sum() / max(len(x), 1)
    ).to_dict()
    top_regions = sorted(reg_wr, key=lambda r: reg_wr[r], reverse=True)[:3]

    scores = []
    for _, row in df.iterrows():
        score = 40  # база

        # +20 за высокую конверсию у этого заказчика
        wr = cust_wr.get(row["Заказчик"], 0.3)
        score += wr * 20

        # +15 за высокую маржу в этом типе работ
        margin = type_margin.get(row.get("Тип работ",""), 18)
        if margin > 22:   score += 15
        elif margin > 17: score += 8
        else:             score += 0

        # +10 за «свой» регион
        if row.get("Регион","") in top_regions:
            score += 10

        # +10 за большую сумму (свидетельствует о серьёзном заказчике)
        if row.get("Сумма тендера", 0) > df["Сумма тендера"].quantile(0.75):
            score += 10

        # -10 за просроченный дедлайн
        ddl = row.get("Дней до дедлайна", None)
        if ddl is not None and not pd.isna(ddl) and ddl < 0:
            score -= 10

        scores.append(min(max(round(score), 0), 100))

    df = df.copy()
    df["Балл"] = scores
    df["Рекомендация"] = df["Балл"].apply(
        lambda s: "✅ Рекомендуем" if s >= 65 else ("⚠️ Средний риск" if s >= 45 else "❌ Высокий риск")
    )
    return df


# ═══════════════════════════════════════════════════════════
# PDF ОТЧЁТ (с полной поддержкой кириллицы через TTF-шрифты)
# ═══════════════════════════════════════════════════════════
def generate_pdf_report(df, won, currency_symbol, display_rate, win_rate, avg_margin):
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, KeepTogether)
    from reportlab.graphics.shapes import Drawing, Rect, String, Line
    from reportlab.graphics import renderPDF

    # ── Шрифты с поддержкой кириллицы — автопоиск TTF ──
    import os as _os, glob as _glob

    def _find_ttf(names):
        dirs = [
            "/usr/share/fonts", "/usr/local/share/fonts",
            _os.path.expanduser("~/.fonts"),
            _os.path.expanduser("~/Library/Fonts"),
            "/Library/Fonts", "C:/Windows/Fonts",
            "/System/Library/Fonts",
        ]
        for d in dirs:
            for nm in names:
                hits = _glob.glob(_os.path.join(d, "**", nm), recursive=True)
                if hits:
                    return hits[0]
        return None

    try:
        _fn = _find_ttf(["DejaVuSans.ttf","Carlito-Regular.ttf","FreeSans.ttf","LiberationSans-Regular.ttf","Arial.ttf","arial.ttf"])
        _fb = _find_ttf(["DejaVuSans-Bold.ttf","Carlito-Bold.ttf","FreeSansBold.ttf","LiberationSans-Bold.ttf","Arialbd.ttf","arialbd.ttf"])
        if not _fn or not _fb:
            raise FileNotFoundError("No Cyrillic TTF found")
        pdfmetrics.registerFont(TTFont("DJ",  _fn))
        pdfmetrics.registerFont(TTFont("DJB", _fb))
        pdfmetrics.registerFontFamily("DJ", normal="DJ", bold="DJB")
        F, FB = "DJ", "DJB"
    except Exception:
        F, FB = "Helvetica", "Helvetica-Bold"

    # ── Цвета ──
    C_BG     = colors.HexColor('#0d1117')
    C_CARD   = colors.HexColor('#161b22')
    C_CARD2  = colors.HexColor('#1c2333')
    C_BORDER = colors.HexColor('#30363d')
    C_BLUE   = colors.HexColor('#1f6feb')
    C_LBLUE  = colors.HexColor('#58a6ff')
    C_GREEN  = colors.HexColor('#3fb950')
    C_RED    = colors.HexColor('#f85149')
    C_YELLOW = colors.HexColor('#d29922')
    C_TEXT   = colors.HexColor('#e6edf3')
    C_MUTED  = colors.HexColor('#8b949e')

    def sty(name, font=None, size=9, color=C_TEXT, align=TA_LEFT, sb=0, sa=4, leading=None):
        return ParagraphStyle(name, fontName=font or F, fontSize=size, textColor=color,
                               alignment=align, spaceBefore=sb, spaceAfter=sa,
                               leading=leading or size * 1.35)

    S_H2   = sty('h2', FB, 12, C_LBLUE, sb=12, sa=6)
    S_FOOT = sty('ft',  F,  8, C_MUTED, TA_CENTER)

    def dark_table(data, widths, hdr_bg=C_BLUE):
        t = Table(data, colWidths=widths)
        ts = [
            ('BACKGROUND',    (0,0),  (-1,0),  hdr_bg),
            ('TEXTCOLOR',     (0,0),  (-1,0),  colors.white),
            ('FONTNAME',      (0,0),  (-1,0),  FB),
            ('FONTSIZE',      (0,0),  (-1,-1), 9),
            ('FONTNAME',      (0,1),  (-1,-1), F),
            ('TEXTCOLOR',     (0,1),  (-1,-1), C_TEXT),
            ('TOPPADDING',    (0,0),  (-1,-1), 5),
            ('BOTTOMPADDING', (0,0),  (-1,-1), 5),
            ('LEFTPADDING',   (0,0),  (-1,-1), 8),
            ('RIGHTPADDING',  (0,0),  (-1,-1), 8),
            ('GRID',          (0,0),  (-1,-1), 0.4, C_BORDER),
            ('VALIGN',        (0,0),  (-1,-1), 'MIDDLE'),
        ]
        for i in range(1, len(data)):
            ts.append(('BACKGROUND', (0,i), (-1,i), C_CARD if i%2==1 else C_CARD2))
        t.setStyle(TableStyle(ts))
        return t

    def mini_bar(value, max_val=35, width=70, height=10,
                  fill=C_BLUE, bg=C_CARD2):
        d = Drawing(width, height)
        d.add(Rect(0, 1, width, height-2, fillColor=bg, strokeColor=None))
        bw = max(2, width * min(float(value), max_val) / max_val)
        d.add(Rect(0, 1, bw, height-2, fillColor=fill, strokeColor=None))
        return d

    # ── Документ ──
    buf = io.BytesIO()
    W, H = A4
    LM = RM = 1.8*cm
    CW = W - LM - RM

    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=LM, rightMargin=RM,
                             topMargin=1.5*cm, bottomMargin=1.5*cm,
                             title="Тендерная аналитика — Отчёт для руководителя")
    story = []

    # ══ ШАПКА ══
    hd = Drawing(CW, 3.6*cm)
    hd.add(Rect(0, 0, CW, 3.6*cm, fillColor=C_CARD, strokeColor=C_BLUE, strokeWidth=1.5))
    hd.add(Rect(0, 3.4*cm, CW, 0.2*cm, fillColor=C_BLUE, strokeColor=None))
    hd.add(String(CW/2, 2.3*cm, "ТЕНДЕРНАЯ АНАЛИТИКА",
                   fontName=FB, fontSize=22, fillColor=C_TEXT, textAnchor='middle'))
    hd.add(String(CW/2, 1.55*cm, "Отчёт для руководителя",
                   fontName=F, fontSize=10, fillColor=C_MUTED, textAnchor='middle'))
    n_total = len(df)
    hd.add(String(CW/2, 0.7*cm,
                   f"Сгенерировано: {datetime.now().strftime('%d.%m.%Y  %H:%M')}   |   "
                   f"{n_total} тендеров   |   Валюта: {currency_symbol}",
                   fontName=F, fontSize=8.5, fillColor=C_MUTED, textAnchor='middle'))
    story.append(hd)
    story.append(Spacer(1, 0.45*cm))

    # ══ KPI ПЛАШКИ ══
    story.append(Paragraph("Ключевые показатели", S_H2))
    total_won    = won["Сумма тендера"].sum() * display_rate / 1e6
    total_profit = won["Чистая прибыль"].sum() * display_rate / 1e6 if "Чистая прибыль" in won.columns else 0
    in_prog_cnt  = len(df[df["Статус"] == "В процессе"])

    kpi_items = [
        ("Тендеров",    f"{n_total:,}",          "в выборке",    C_LBLUE),
        ("Выиграно",    f"{len(won):,}",          f"из {len(df[df['Статус']!='В процессе'])}",C_GREEN),
        ("Конверсия",   f"{win_rate:.1f}%",       "Win Rate",     C_YELLOW if win_rate<40 else C_GREEN),
        ("Выручка",     f"{total_won:.0f} млн",   currency_symbol, C_BLUE),
        ("Прибыль",     f"{total_profit:.0f} млн",currency_symbol, C_GREEN),
        ("Маржа",       f"{avg_margin:.1f}%",     "средняя",      C_YELLOW if avg_margin<20 else C_GREEN),
        ("В процессе",  f"{in_prog_cnt}",         "тендеров",     C_YELLOW),
    ]
    cw = CW / len(kpi_items)
    kpi_d = Drawing(CW, 2.6*cm)
    for i, (lbl, val, sub, col) in enumerate(kpi_items):
        x = i * cw
        kpi_d.add(Rect(x+1, 1, cw-2, 2.4*cm, fillColor=C_CARD, strokeColor=col, strokeWidth=0.8))
        kpi_d.add(Rect(x+1, 2.3*cm, cw-2, 0.1*cm, fillColor=col, strokeColor=None))
        kpi_d.add(String(x+cw/2, 1.45*cm, val,  fontName=FB, fontSize=12, fillColor=C_TEXT,  textAnchor='middle'))
        kpi_d.add(String(x+cw/2, 0.7*cm,  lbl,  fontName=FB, fontSize=7,  fillColor=col,     textAnchor='middle'))
        kpi_d.add(String(x+cw/2, 0.2*cm,  sub,  fontName=F,  fontSize=6.5,fillColor=C_MUTED, textAnchor='middle'))
    story.append(kpi_d)
    story.append(Spacer(1, 0.5*cm))

    # ══ МАРЖА ПО ТИПАМ РАБОТ ══
    if "Тип работ" in won.columns and "Маржа %" in won.columns:
        story.append(Paragraph("Маржинальность по типам работ", S_H2))
        mbt = won.groupby("Тип работ").agg(
            Кол=("Статус","count"),
            Выручка=("Сумма тендера","sum"),
            Прибыль=("Чистая прибыль","sum"),
            Маржа=("Маржа %","mean"),
        ).reset_index().sort_values("Маржа", ascending=False)

        tbl = [["Тип работ", "Кол", "Выручка (млн)", "Прибыль (млн)", "Маржа %", "Уровень"]]
        for _, r in mbt.iterrows():
            fill = C_GREEN if r["Маржа"]>=20 else (C_YELLOW if r["Маржа"]>=15 else C_RED)
            tbl.append([
                r["Тип работ"],
                str(int(r["Кол"])),
                f"{r['Выручка']*display_rate/1e6:.1f}",
                f"{r['Прибыль']*display_rate/1e6:.1f}",
                f"{r['Маржа']:.1f}%",
                mini_bar(r["Маржа"], max_val=35, width=70, height=11, fill=fill),
            ])
        story.append(dark_table(tbl, [5.8*cm, 1.4*cm, 3.0*cm, 3.0*cm, 2.0*cm, 3.2*cm]))
        story.append(Spacer(1, 0.4*cm))

    # ══ ДЕДЛАЙНЫ ══
    if "Дней до дедлайна" in df.columns:
        urgent = df[(df["Статус"]=="В процессе") &
                    df["Дней до дедлайна"].notna() &
                    (df["Дней до дедлайна"] <= 7)].sort_values("Дней до дедлайна")
        if len(urgent) > 0:
            story.append(Paragraph("Ближайшие дедлайны (следующие 7 дней)", S_H2))
            dl = [["Номер тендера", "Заказчик", "Сумма (млн)", "До дедлайна"]]
            for _, r in urgent.head(10).iterrows():
                d = int(r["Дней до дедлайна"])
                dl.append([
                    r.get("Номер тендера","—"),
                    str(r.get("Заказчик","—"))[:32],
                    f"{r['Сумма тендера']*display_rate/1e6:.1f}",
                    "ПРОСРОЧЕН" if d<0 else ("СЕГОДНЯ" if d==0 else f"через {d} дн."),
                ])
            dl_t = Table(dl, colWidths=[3.0*cm, 8.0*cm, 3.0*cm, 3.0*cm])
            dl_styles = [
                ('BACKGROUND',    (0,0),  (-1,0), C_YELLOW),
                ('TEXTCOLOR',     (0,0),  (-1,0), C_BG),
                ('FONTNAME',      (0,0),  (-1,0), FB),
                ('FONTSIZE',      (0,0),  (-1,-1), 9),
                ('FONTNAME',      (0,1),  (-1,-1), F),
                ('TOPPADDING',    (0,0),  (-1,-1), 5),
                ('BOTTOMPADDING', (0,0),  (-1,-1), 5),
                ('LEFTPADDING',   (0,0),  (-1,-1), 8),
                ('RIGHTPADDING',  (0,0),  (-1,-1), 8),
                ('GRID',          (0,0),  (-1,-1), 0.4, C_BORDER),
                ('VALIGN',        (0,0),  (-1,-1), 'MIDDLE'),
            ]
            for i, (_, r) in enumerate(urgent.head(10).iterrows(), start=1):
                dv = r["Дней до дедлайна"]
                if dv <= 1:
                    dl_styles += [('BACKGROUND',(0,i),(-1,i),colors.HexColor('#2b0d0d')),
                                   ('TEXTCOLOR',(0,i),(-1,i),C_RED)]
                elif dv <= 3:
                    dl_styles += [('BACKGROUND',(0,i),(-1,i),colors.HexColor('#2b1e00')),
                                   ('TEXTCOLOR',(0,i),(-1,i),C_YELLOW)]
                else:
                    dl_styles.append(('BACKGROUND',(0,i),(-1,i), C_CARD if i%2==1 else C_CARD2))
                    dl_styles.append(('TEXTCOLOR',(0,i),(-1,i), C_TEXT))
            dl_t.setStyle(TableStyle(dl_styles))
            story.append(dl_t)
            story.append(Spacer(1, 0.4*cm))

    # ══ СКОРИНГ ══
    if "Балл" in df.columns:
        scored_ip = df[df["Статус"]=="В процессе"].sort_values("Балл", ascending=False).head(10)
        if len(scored_ip) > 0:
            story.append(Paragraph("Скоринг: рейтинг тендеров «В процессе»", S_H2))
            sc = [["Номер", "Заказчик", "Тип работ", "Сумма (млн)", "Балл", "Рекомендация"]]
            for _, r in scored_ip.iterrows():
                sc.append([
                    r.get("Номер тендера","—"),
                    str(r.get("Заказчик","—"))[:20],
                    str(r.get("Тип работ","—"))[:18],
                    f"{r['Сумма тендера']*display_rate/1e6:.1f}",
                    str(int(r["Балл"])),
                    r.get("Рекомендация","—"),
                ])
            sc_t = Table(sc, colWidths=[2.8*cm, 4.5*cm, 4.0*cm, 2.6*cm, 1.5*cm, 3.5*cm])
            sc_styles = [
                ('BACKGROUND',    (0,0),  (-1,0), C_BLUE),
                ('TEXTCOLOR',     (0,0),  (-1,0), colors.white),
                ('FONTNAME',      (0,0),  (-1,0), FB),
                ('FONTSIZE',      (0,0),  (-1,-1), 9),
                ('FONTNAME',      (0,1),  (-1,-1), F),
                ('TOPPADDING',    (0,0),  (-1,-1), 5),
                ('BOTTOMPADDING', (0,0),  (-1,-1), 5),
                ('LEFTPADDING',   (0,0),  (-1,-1), 8),
                ('RIGHTPADDING',  (0,0),  (-1,-1), 8),
                ('GRID',          (0,0),  (-1,-1), 0.4, C_BORDER),
                ('VALIGN',        (0,0),  (-1,-1), 'MIDDLE'),
                ('ALIGN',         (3,0),  (4,-1),  'CENTER'),
            ]
            for i, (_, r) in enumerate(scored_ip.iterrows(), start=1):
                sc_styles.append(('BACKGROUND',(0,i),(-1,i), C_CARD if i%2==1 else C_CARD2))
                ball = r["Балл"]
                col = C_GREEN if ball>=65 else (C_YELLOW if ball>=45 else C_RED)
                sc_styles += [('TEXTCOLOR',(4,i),(5,i), col),
                               ('FONTNAME', (4,i),(5,i), FB)]
            sc_t.setStyle(TableStyle(sc_styles))
            story.append(sc_t)
            story.append(Spacer(1, 0.4*cm))

    # ══ ТОП ЗАКАЗЧИКОВ ══
    story.append(Paragraph("Топ заказчиков по выигранным тендерам", S_H2))
    top_c = won.groupby("Заказчик").agg(
        Кол=("Статус","count"),
        Выручка=("Сумма тендера","sum"),
        Прибыль=("Чистая прибыль","sum"),
    ).reset_index().nlargest(8,"Выручка")
    total_w = won["Сумма тендера"].sum()

    cust_tbl = [["Заказчик","Кол","Выручка (млн)","Прибыль (млн)","Доля","  "]]
    for _, r in top_c.iterrows():
        share = r["Выручка"]/max(total_w,1)*100
        cust_tbl.append([
            str(r["Заказчик"])[:30],
            str(int(r["Кол"])),
            f"{r['Выручка']*display_rate/1e6:.1f}",
            f"{r['Прибыль']*display_rate/1e6:.1f}",
            f"{share:.1f}%",
            mini_bar(share, max_val=30, width=65, height=11, fill=C_BLUE),
        ])
    story.append(dark_table(cust_tbl, [6.2*cm, 1.4*cm, 3.0*cm, 3.0*cm, 1.8*cm, 2.5*cm]))
    story.append(Spacer(1, 0.5*cm))

    # ══ ИНСАЙТЫ ══
    if "Тип работ" in won.columns and "Маржа %" in won.columns:
        mbt2 = won.groupby("Тип работ")["Маржа %"].mean()
        best = mbt2.idxmax(); worst = mbt2.idxmin()
        ins_items = [
            f"Лучшая маржа: «{best}» — {mbt2[best]:.1f}%. Приоритизируйте этот сегмент для роста прибыли.",
            f"Зона риска: «{worst}» — маржа {mbt2[worst]:.1f}%. Пересмотрите ценообразование или снизьте себестоимость.",
            f"Конверсия {win_rate:.1f}% {'выше' if win_rate>=40 else 'ниже'} целевых 40%."
            + (" Отличный результат!" if win_rate>=40 else " Проанализируйте причины проигрышей."),
        ]
        ins_d = Drawing(CW, 2.9*cm)
        ins_d.add(Rect(0, 0, CW, 2.9*cm,
                        fillColor=colors.HexColor('#0d2040'),
                        strokeColor=C_BLUE, strokeWidth=1))
        ins_d.add(String(10, 2.55*cm, "АВТОМАТИЧЕСКИЕ ИНСАЙТЫ",
                          fontName=FB, fontSize=8.5, fillColor=C_LBLUE))
        ins_d.add(Line(0, 2.35*cm, CW, 2.35*cm, strokeColor=C_BORDER, strokeWidth=0.5))
        for j, txt in enumerate(ins_items):
            ins_d.add(String(12, (1.85 - j*0.6)*cm, f"•  {txt}",
                              fontName=F, fontSize=8.5, fillColor=C_TEXT))
        story.append(ins_d)

    # ══ ФУТЕР ══
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.8, color=C_BORDER, spaceAfter=5))
    story.append(Paragraph(
        f"Сгенерировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}   |   "
        f"Тендерная аналитика Pro   |   Конфиденциально",
        S_FOOT))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════
# PLOTLY LAYOUT
# ═══════════════════════════════════════════════════════════
LAYOUT = dict(
    paper_bgcolor="#161b22", plot_bgcolor="#0d1117",
    font=dict(family="IBM Plex Sans", color="#c9d1d9"),
    title_font=dict(family="Rajdhani", size=18, color="#e6edf3"),
    legend=dict(bgcolor="#161b22", bordercolor="#30363d", borderwidth=1),
    margin=dict(l=40, r=20, t=50, b=40),
)
STATUS_COLORS = {"Выигран":"#3fb950","Проигран":"#f85149","В процессе":"#d29922"}

def apply_layout(fig, **kwargs):
    fig.update_layout(**LAYOUT, **kwargs)
    fig.update_xaxes(gridcolor="#21262d", zerolinecolor="#30363d", tickfont_size=11)
    fig.update_yaxes(gridcolor="#21262d", zerolinecolor="#30363d", tickfont_size=11)
    return fig


# ═══════════════════════════════════════════════════════════
# ПОЛУЧЕНИЕ КУРСОВ ВАЛЮТ
# ═══════════════════════════════════════════════════════════
@st.cache_data(ttl=3600)
def fetch_exchange_rates():
    """Пытается получить курсы с открытого API. Fallback на дефолтные значения."""
    try:
        import urllib.request
        url = "https://open.er-api.com/v6/latest/KZT"
        with urllib.request.urlopen(url, timeout=3) as r:
            data = json.loads(r.read().decode())
        rates = data.get("rates", {})
        return {
            "RUB": rates.get("RUB", 0.21),
            "USD": rates.get("USD", 0.0022),
            "EUR": rates.get("EUR", 0.0020),
            "source": "live",
            "updated": data.get("time_last_update_utc","—"),
        }
    except Exception:
        return {"RUB": 0.21, "USD": 0.0022, "EUR": 0.0020,
                "source": "fallback", "updated": "офлайн"}


# ═══════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🏗️ Тендерная аналитика Pro")
    st.markdown("---")

    uploaded = st.file_uploader("📂 Загрузить Excel", type=["xlsx","xls"])

    # Курсы валют (авто)
    st.markdown("---")
    st.markdown("**💱 Валюта отображения**")

    rates_data = fetch_exchange_rates()
    if rates_data["source"] == "live":
        st.success(f"✅ Курсы загружены (live)", icon=None)
    else:
        st.warning("⚠️ Офлайн-режим (дефолтные курсы)")

    display_currency = st.selectbox("Валюта", ["KZT (тенге)","RUB (рубли)","USD (доллары)","EUR (евро)"], index=0)
    currency_code = display_currency.split(" ")[0]

    auto_rate_map = {"KZT": 1.0, "RUB": rates_data["RUB"],
                     "USD": rates_data["USD"], "EUR": rates_data["EUR"]}
    auto_rate = auto_rate_map.get(currency_code, 1.0)

    symbol_map = {"KZT":"₸","RUB":"₽","USD":"$","EUR":"€"}
    currency_symbol = symbol_map.get(currency_code, "₸")

    if currency_code != "KZT":
        use_manual = st.checkbox("Задать курс вручную")
        if use_manual:
            display_rate = st.number_input(f"Курс KZT → {currency_code}",
                                            value=float(f"{auto_rate:.5f}"), format="%.5f", step=0.0001)
        else:
            display_rate = auto_rate
            st.caption(f"1 KZT = {auto_rate:.4f} {currency_code}")
    else:
        display_rate = 1.0

    # Фильтры
    st.markdown("---")
    st.markdown("**Фильтры**")

    if uploaded:
        raw_df = load_excel(uploaded)
    else:
        st.info("Используются демо-данные")
        raw_df = generate_sample_data()

    years = sorted(raw_df["Год"].dropna().unique().astype(int).tolist())
    sel_years = st.multiselect("Год", years, default=years)
    statuses_list = raw_df["Статус"].dropna().unique().tolist()
    sel_statuses = st.multiselect("Статус", statuses_list, default=statuses_list)
    regions_all = sorted(raw_df["Регион"].dropna().unique().tolist()) if "Регион" in raw_df.columns else []
    sel_regions = st.multiselect("Регион", regions_all, default=regions_all) if regions_all else None
    work_types_all = sorted(raw_df["Тип работ"].dropna().unique().tolist()) if "Тип работ" in raw_df.columns else []
    sel_works = st.multiselect("Тип работ", work_types_all, default=work_types_all) if work_types_all else None

    # ── Диагностика: регионы без координат ─────────────────
    if "Регион" in raw_df.columns:
        unknown = [r for r in raw_df["Регион"].dropna().unique()
                   if get_coords(r) == KZ_CENTER and r not in KZ_COORDS]
        if unknown:
            with st.expander(f"⚠️ Регионы без точных координат ({len(unknown)})"):
                for u in unknown:
                    st.caption(f"• {u} → центр КЗ")

    # What-if симуляция
    st.markdown("---")
    st.markdown("**🔮 Симуляция «Что если?»**")

    wif_margin_delta = st.slider("Изменение маржи (%)", -10, +15, 0, step=1,
                                  help="Смоделировать рост/снижение маржи на X%")
    wif_win_rate_delta = st.slider("Изменение конверсии (%)", -20, +20, 0, step=1,
                                    help="Смоделировать изменение Win Rate")
    wif_lose_all = st.checkbox("Сценарий: проиграть все «В процессе»")
    wif_win_all  = st.checkbox("Сценарий: выиграть все «В процессе»")

    st.markdown("---")
    template = pd.DataFrame({
        "Дата подачи":["2024-01-15"],"Заказчик":["ООО Пример"],
        "Тип работ":["Дорожное строительство"],"Регион":["Москва"],
        "Сумма тендера":[5000000],"НМЦК":[6000000],"Себестоимость":[4000000],
        "Статус":["Выигран"],"Номер тендера":["Т-123456"],"Срок (дней)":[180],
        "Конкурент":["ООО «АльфаСтрой»"],"Валюта":["KZT"],
    })
    buf_t = io.BytesIO(); template.to_excel(buf_t, index=False)
    st.download_button("⬇️ Шаблон Excel", buf_t.getvalue(), "tender_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════════
# ФИЛЬТРАЦИЯ + СКОРИНГ
# ═══════════════════════════════════════════════════════════
df_all = raw_df.copy()  # для скоринга нужны все данные

df = raw_df.copy()
if sel_years:    df = df[df["Год"].isin(sel_years)]
if sel_statuses: df = df[df["Статус"].isin(sel_statuses)]
if sel_regions and "Регион" in df.columns:   df = df[df["Регион"].isin(sel_regions)]
if sel_works and "Тип работ" in df.columns:  df = df[df["Тип работ"].isin(sel_works)]

# Считаем скоринг на ВСЁМ датасете, потом применяем к отфильтрованному
df_all_scored = compute_scoring(df_all)
score_map = df_all_scored.set_index(df_all_scored.index)[["Балл","Рекомендация"]]
df = df.copy()
df["Балл"] = score_map.reindex(df.index)["Балл"].values
df["Рекомендация"] = score_map.reindex(df.index)["Рекомендация"].values

won  = df[df["Статус"] == "Выигран"]
lost = df[df["Статус"] == "Проигран"]
in_progress = df[df["Статус"] == "В процессе"]

def c(val): return val * display_rate


# ═══════════════════════════════════════════════════════════
# ЗАГОЛОВОК
# ═══════════════════════════════════════════════════════════
st.markdown("# 🏗️ Дашборд строительных тендеров Pro")
date_range = ""
if not df.empty:
    dmin, dmax = df["Дата подачи"].min(), df["Дата подачи"].max()
    if pd.notna(dmin) and pd.notna(dmax):
        date_range = f"{dmin.strftime('%d.%m.%Y')} – {dmax.strftime('%d.%m.%Y')}"
rate_badge = f"<span class='currency-badge'>1 KZT = {display_rate:.4f} {currency_code}</span>" if currency_code != "KZT" else f"<span class='currency-badge'>{currency_code}</span>"
st.markdown(f"*Данные: {len(df):,} тендеров · {date_range}* {rate_badge}", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
# KPI
# ═══════════════════════════════════════════════════════════
total_sum  = df["Сумма тендера"].sum()
won_sum    = won["Сумма тендера"].sum()
won_profit = won["Чистая прибыль"].sum() if "Чистая прибыль" in won.columns else 0
win_rate   = len(won) / max(len(df[df["Статус"] != "В процессе"]), 1) * 100
avg_won    = won["Сумма тендера"].mean() if len(won) > 0 else 0
avg_margin = won["Маржа %"].mean() if len(won) > 0 and "Маржа %" in won.columns else 0

def kpi(col, label, value, delta=None, delta_pos=True, delta_class="pos"):
    delta_html = ""
    if delta is not None:
        cls_map = {"pos":"metric-delta-pos","neg":"metric-delta-neg","neu":"metric-delta-neu"}
        cls = cls_map.get(delta_class,"metric-delta-pos")
        arrow = "▲" if delta_pos else "▼"
        delta_html = f'<div class="{cls}">{arrow} {delta}</div>'
    col.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>""", unsafe_allow_html=True)

k1,k2,k3,k4,k5,k6 = st.columns(6)
def fmt(val): return f"{c(val)/1e6:.1f} млн {currency_symbol}"

kpi(k1,"Всего тендеров",f"{len(df):,}")
kpi(k2,"Выиграно",f"{len(won):,}",f"из {len(df[df['Статус']!='В процессе'])}")
kpi(k3,"Конверсия",f"{win_rate:.1f}%")
kpi(k4,f"Выиграно",fmt(won_sum))
kpi(k5,f"Чистая прибыль",fmt(won_profit))
kpi(k6,"Средняя маржа",f"{avg_margin:.1f}%",
    f"{'выше' if avg_margin>=20 else 'ниже'} 20%",
    avg_margin>=20, "pos" if avg_margin>=20 else "neg")

# PDF кнопка рядом с KPI
st.markdown("")
if st.button("📄 Сформировать отчёт для директора (PDF)", type="primary"):
    with st.spinner("Генерация PDF..."):
        pdf_bytes = generate_pdf_report(df, won, currency_symbol, display_rate, win_rate, avg_margin)
    st.download_button(
        label="⬇️ Скачать PDF-отчёт",
        data=pdf_bytes,
        file_name=f"tender_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
    )

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# ДЕДЛАЙНЫ
# ═══════════════════════════════════════════════════════════
if "Дней до дедлайна" in df.columns and len(in_progress) > 0:
    st.markdown('<div class="section-title">⏰ Ближайшие дедлайны</div>', unsafe_allow_html=True)
    deadlines_df = in_progress[in_progress["Дней до дедлайна"].notna()].copy()
    deadlines_df = deadlines_df[deadlines_df["Дней до дедлайна"] <= 7].sort_values("Дней до дедлайна")
    if len(deadlines_df) > 0:
        d_cols = st.columns(min(len(deadlines_df), 4))
        for i, (_, row) in enumerate(deadlines_df.head(4).iterrows()):
            days = int(row["Дней до дедлайна"])
            card_class = "deadline-card" if days <= 2 else "deadline-card-warn"
            icon = "🔴" if days <= 2 else "🟡"
            day_str = "СЕГОДНЯ" if days==0 else (f"через {days} д." if days>0 else f"просрочен {abs(days)} д.")
            d_cols[i%4].markdown(f"""
            <div class="{card_class}">
                <div class="deadline-title">{icon} {row.get('Номер тендера','—')}</div>
                <div class="deadline-meta">{str(row.get('Заказчик','—'))[:28]}</div>
                <div class="deadline-meta">{fmt(row['Сумма тендера'])} · <b>{day_str}</b></div>
            </div>""", unsafe_allow_html=True)
    st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК 1: СКОРИНГ И РЕКОМЕНДАЦИИ
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🎯 Скоринг и Рекомендации (AI Light)</div>', unsafe_allow_html=True)

sc_c1, sc_c2 = st.columns([2, 3])

with sc_c1:
    st.markdown("**Тендеры «В процессе» — рейтинг перспективности**")
    if len(in_progress) > 0:
        scored_ip = in_progress[["Номер тендера","Заказчик","Тип работ","Регион",
                                   "Сумма тендера","Балл","Рекомендация"]].copy()
        scored_ip = scored_ip.sort_values("Балл", ascending=False)
        scored_ip[f"Сумма (млн {currency_symbol})"] = (scored_ip["Сумма тендера"] * display_rate / 1e6).round(1)

        def color_score(val):
            if val >= 65: return "background-color:#0d2b17; color:#3fb950"
            if val >= 45: return "background-color:#2b1e00; color:#d29922"
            return "background-color:#2b0d0d; color:#f85149"

        def color_rec(val):
            if "Рекомендуем" in val: return "color:#3fb950"
            if "Средний" in val: return "color:#d29922"
            return "color:#f85149"

        disp = scored_ip[["Номер тендера","Заказчик",f"Сумма (млн {currency_symbol})","Балл","Рекомендация"]].copy()
        disp["Заказчик"] = disp["Заказчик"].str[:22]
        styled_sc = disp.style.applymap(color_score, subset=["Балл"]).applymap(color_rec, subset=["Рекомендация"])
        st.dataframe(styled_sc, use_container_width=True, height=380)
    else:
        st.info("Нет тендеров «В процессе»")

with sc_c2:
    # Распределение баллов
    if "Балл" in df.columns:
        fig_score = go.Figure()
        for status, color in STATUS_COLORS.items():
            sub = df[df["Статус"] == status]
            if len(sub) > 0:
                fig_score.add_trace(go.Histogram(
                    x=sub["Балл"], name=status,
                    marker_color=color, opacity=0.75,
                    xbins=dict(start=0, end=100, size=5),
                ))
        fig_score.add_vline(x=65, line_dash="dash", line_color="#3fb950",
                             annotation_text="Рекомендуем", annotation_font_color="#3fb950",
                             annotation_position="top right")
        fig_score.add_vline(x=45, line_dash="dash", line_color="#d29922",
                             annotation_text="Средний риск", annotation_font_color="#d29922",
                             annotation_position="top right")
        apply_layout(fig_score, title="Распределение баллов скоринга по статусам",
                      barmode="overlay", xaxis_title="Балл", yaxis_title="Количество", height=200)
        st.plotly_chart(fig_score, use_container_width=True)

    # Scatter: балл vs сумма для тендеров «В процессе»
    if len(in_progress) > 0 and "Балл" in in_progress.columns:
        ip_plot = in_progress.copy()
        ip_plot["Сумма_conv"] = ip_plot["Сумма тендера"] * display_rate / 1e6
        fig_sc2 = px.scatter(
            ip_plot, x="Балл", y="Сумма_conv",
            color="Балл",
            color_continuous_scale=["#f85149","#d29922","#3fb950"],
            range_color=[0,100],
            hover_name="Заказчик",
            hover_data={"Тип работ":True,"Регион":True,"Рекомендация":True,
                        "Балл":True,"Сумма_conv":":.1f"},
            size_max=20,
            labels={"Сумма_conv":f"Сумма (млн {currency_symbol})","Балл":"Балл скоринга"},
        )
        fig_sc2.add_vrect(x0=65, x1=100, fillcolor="#3fb950", opacity=0.06, line_width=0)
        fig_sc2.add_vrect(x0=0, x1=45, fillcolor="#f85149", opacity=0.06, line_width=0)
        apply_layout(fig_sc2, title="Тендеры «В процессе»: балл vs сумма", height=180)
        st.plotly_chart(fig_sc2, use_container_width=True)

    # Топ-инсайты
    if "Тип работ" in won.columns and "Маржа %" in won.columns:
        top_type = won.groupby("Тип работ")["Маржа %"].mean().idxmax()
        top_type_margin = won.groupby("Тип работ")["Маржа %"].mean().max()
        top_cust = won.groupby("Заказчик").size().idxmax()
        top_cust_wr = (df[df["Заказчик"]==top_cust]["Статус"]=="Выигран").sum() / max(len(df[df["Заказчик"]==top_cust]), 1) * 100
        st.markdown(f"""
        <div class="insight-box">
            <div class="insight-title">💡 Автоинсайт: лучший сегмент</div>
            <div class="insight-body">
            Наивысшая маржа у <b>«{top_type}»</b> — {top_type_margin:.1f}%.
            Приоритизируйте этот тип работ для максимальной прибыльности.
            </div>
        </div>
        <div class="insight-box">
            <div class="insight-title">💡 Автоинсайт: ключевой заказчик</div>
            <div class="insight-body">
            Лучший Win Rate у <b>«{top_cust}»</b> — {top_cust_wr:.0f}%.
            Тендеры этого заказчика получают максимальный балл скоринга.
            </div>
        </div>""", unsafe_allow_html=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК 2: СИМУЛЯЦИЯ «ЧТО ЕСЛИ?»
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🔮 Симуляция «Что если?» (What-if Analysis)</div>', unsafe_allow_html=True)

wif_c1, wif_c2 = st.columns([1, 3])

with wif_c1:
    base_won_sum    = won["Сумма тендера"].sum() * display_rate / 1e6
    base_profit     = won["Чистая прибыль"].sum() * display_rate / 1e6 if "Чистая прибыль" in won.columns else 0
    base_margin_avg = avg_margin

    in_prog_sum  = in_progress["Сумма тендера"].sum() * display_rate / 1e6
    sim_win_rate = (win_rate + wif_win_rate_delta) / 100

    # Сценарий: тендеры «В процессе»
    if wif_win_all:
        extra_won_sum = in_prog_sum
        extra_profit  = in_prog_sum * (base_margin_avg + wif_margin_delta) / 100
    elif wif_lose_all:
        extra_won_sum = 0
        extra_profit  = 0
    else:
        extra_won_sum = in_prog_sum * sim_win_rate
        extra_profit  = extra_won_sum * (base_margin_avg + wif_margin_delta) / 100

    sim_total_sum    = base_won_sum + extra_won_sum
    sim_total_profit = base_profit * (1 + wif_margin_delta/100) + extra_profit
    delta_sum    = sim_total_sum - base_won_sum
    delta_profit = sim_total_profit - base_profit

    def sim_kpi(label, base_v, sim_v, delta_v):
        is_pos = delta_v >= 0
        color = "#3fb950" if is_pos else "#f85149"
        arrow = "▲" if is_pos else "▼"
        st.markdown(f"""
        <div class="whatif-card">
            <div class="metric-label">{label}</div>
            <div style="display:flex;align-items:baseline;gap:8px;">
                <span style="font-family:Rajdhani;font-size:1.6rem;color:#e6edf3;font-weight:700">{sim_v:.1f}</span>
                <span style="color:{color};font-size:0.85rem">{arrow} {abs(delta_v):.1f} ({abs(delta_v/max(base_v,0.01)*100):.0f}%)</span>
            </div>
            <div style="color:#8b949e;font-size:0.75rem">База: {base_v:.1f} млн {currency_symbol}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown(f"**Параметры:** маржа {'+' if wif_margin_delta>=0 else ''}{wif_margin_delta}%, конверсия {'+' if wif_win_rate_delta>=0 else ''}{wif_win_rate_delta}%")
    sim_kpi(f"Прогноз выручки (млн {currency_symbol})", base_won_sum, sim_total_sum, delta_sum)
    sim_kpi(f"Прогноз прибыли (млн {currency_symbol})", base_profit, sim_total_profit, delta_profit)

    scenario_label = "Выиграть все" if wif_win_all else ("Проиграть все" if wif_lose_all else "Базовый + изменения")
    st.caption(f"Сценарий «В процессе»: **{scenario_label}**")

with wif_c2:
    # График с пунктирным прогнозом
    if not won.empty:
        monthly_won = won.groupby("Месяц").agg(
            Сумма=("Сумма тендера","sum"),
            Прибыль=("Чистая прибыль","sum"),
        ).reset_index()
        monthly_won["Сумма_conv"]   = monthly_won["Сумма"]   * display_rate / 1e6
        monthly_won["Прибыль_conv"] = monthly_won["Прибыль"] * display_rate / 1e6

        # Генерируем прогноз: последние 6 месяцев + 6 вперёд
        last_date = monthly_won["Месяц"].max()
        future_months = pd.date_range(last_date + pd.offsets.MonthBegin(1), periods=6, freq="MS")
        avg_monthly = monthly_won["Сумма_conv"].tail(6).mean()
        avg_monthly_p = monthly_won["Прибыль_conv"].tail(6).mean()

        sim_factor_sum    = (1 + wif_win_rate_delta/100) * (1 + wif_margin_delta/200)
        sim_factor_profit = (1 + wif_margin_delta/100) * (1 + wif_win_rate_delta/100)

        if wif_win_all:   sim_factor_sum = 1.3; sim_factor_profit = 1.3
        elif wif_lose_all: sim_factor_sum = 0.3; sim_factor_profit = 0.3

        future_sum    = [avg_monthly    * sim_factor_sum    * (1 + i*0.02) for i in range(6)]
        future_profit = [avg_monthly_p  * sim_factor_profit * (1 + i*0.02) for i in range(6)]

        fig_wif = go.Figure()
        fig_wif.add_trace(go.Bar(
            x=monthly_won["Месяц"], y=monthly_won["Сумма_conv"],
            name=f"Факт выручка (млн {currency_symbol})", marker_color="#1f6feb", opacity=0.8,
        ))
        fig_wif.add_trace(go.Scatter(
            x=monthly_won["Месяц"], y=monthly_won["Прибыль_conv"],
            name=f"Факт прибыль", line=dict(color="#3fb950", width=2),
            mode="lines+markers",
        ))
        # Соединительная точка
        bridge_x = [monthly_won["Месяц"].iloc[-1], future_months[0]]
        bridge_sum = [monthly_won["Сумма_conv"].iloc[-1], future_sum[0]]
        bridge_prf = [monthly_won["Прибыль_conv"].iloc[-1], future_profit[0]]

        fig_wif.add_trace(go.Scatter(
            x=list(future_months), y=future_sum,
            name=f"Прогноз выручка", line=dict(color="#58a6ff", width=2, dash="dot"),
            mode="lines+markers", marker=dict(symbol="diamond", size=7),
        ))
        fig_wif.add_trace(go.Scatter(
            x=list(future_months), y=future_profit,
            name=f"Прогноз прибыль", line=dict(color="#7ee787", width=2, dash="dot"),
            mode="lines+markers", marker=dict(symbol="diamond", size=7),
        ))
        # Зона прогноза
        fig_wif.add_vrect(
            x0=str(future_months[0])[:7], x1=str(future_months[-1])[:7],
            fillcolor="#1f6feb", opacity=0.05, line_width=0,
            annotation_text="Прогноз", annotation_position="top left",
            annotation_font_color="#58a6ff",
        )
        apply_layout(fig_wif, title=f"Факт + прогноз с симуляцией (маржа {'+' if wif_margin_delta>=0 else ''}{wif_margin_delta}%, конверсия {'+' if wif_win_rate_delta>=0 else ''}{wif_win_rate_delta}%)",
                      height=360)
        st.plotly_chart(fig_wif, use_container_width=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК 3: PRICE INTELLIGENCE
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">💹 Price Intelligence — Анализ цен и НМЦК</div>', unsafe_allow_html=True)

pi_c1, pi_c2, pi_c3 = st.columns([2, 2, 2])

with pi_c1:
    if "Снижение %" in df.columns:
        fig_pi1 = go.Figure()
        fig_pi1.add_trace(go.Histogram(
            x=df[df["Статус"]=="Выигран"]["Снижение %"],
            name="Выигранные", marker_color="#3fb950", opacity=0.8,
            xbins=dict(start=0, end=40, size=1),
        ))
        fig_pi1.add_trace(go.Histogram(
            x=df[df["Статус"]=="Проигран"]["Снижение %"],
            name="Проигранные", marker_color="#f85149", opacity=0.6,
            xbins=dict(start=0, end=40, size=1),
        ))
        median_drop_won = df[df["Статус"]=="Выигран"]["Снижение %"].median()
        fig_pi1.add_vline(x=median_drop_won, line_dash="dash", line_color="#3fb950",
                           annotation_text=f"Медиана выигранных: {median_drop_won:.1f}%",
                           annotation_font_color="#3fb950")
        apply_layout(fig_pi1, title="Распределение снижения от НМЦК",
                      barmode="overlay", xaxis_title="Снижение %", yaxis_title="Кол-во", height=300)
        st.plotly_chart(fig_pi1, use_container_width=True)

with pi_c2:
    if "Снижение %" in df.columns and "Заказчик" in df.columns:
        price_by_cust = df[df["Статус"]=="Выигран"].groupby("Заказчик").agg(
            Медиана_снижения=("Снижение %","median"),
            Кол=("Статус","count"),
            Ср_НМЦК=("НМЦК","mean"),
        ).reset_index().sort_values("Медиана_снижения").head(10)

        fig_pi2 = go.Figure(go.Bar(
            y=price_by_cust["Заказчик"].str[:22],
            x=price_by_cust["Медиана_снижения"],
            orientation="h",
            marker=dict(color=price_by_cust["Медиана_снижения"],
                        colorscale=["#3fb950","#d29922","#f85149"],
                        line=dict(width=0)),
            text=[f"{v:.1f}%" for v in price_by_cust["Медиана_снижения"]],
            textposition="outside",
        ))
        apply_layout(fig_pi2, title="Медиана снижения по заказчикам",
                      height=300, xaxis_title="Снижение от НМЦК %", xaxis_range=[0,35])
        st.plotly_chart(fig_pi2, use_container_width=True)

with pi_c3:
    # Инсайт-рекомендации по ценам
    if "Снижение %" in df.columns and "Заказчик" in df.columns:
        st.markdown("**💡 Ценовые инсайты**")
        top5_cust_price = df[df["Статус"]=="Выигран"].groupby("Заказчик").agg(
            Медиана=("Снижение %","median"),
            Ср_НМЦК=("НМЦК","mean"),
        ).reset_index().nlargest(5, "Ср_НМЦК")

        for _, r in top5_cust_price.iterrows():
            rec_price = r["Ср_НМЦК"] * (1 - r["Медиана"]/100) * display_rate / 1e6
            st.markdown(f"""
            <div class="insight-box">
                <div class="insight-title">📌 {r['Заказчик'][:28]}</div>
                <div class="insight-body">
                Победители обычно снижают на <b>{r['Медиана']:.1f}%</b>.<br>
                Рекомендуемая цена: <b>≤ {rec_price:.2f} млн {currency_symbol}</b>
                </div>
            </div>""", unsafe_allow_html=True)

        # Scatter НМЦК vs итоговая цена
        pi_scatter = df[df["Статус"]=="Выигран"].copy()
        pi_scatter["НМЦК_conv"] = pi_scatter["НМЦК"] * display_rate / 1e6
        pi_scatter["Цена_conv"] = pi_scatter["Сумма тендера"] * display_rate / 1e6

st.markdown("---")

# Детальный график: НМЦК vs итоговая цена
if "НМЦК" in df.columns:
    pi_full = df.copy()
    pi_full["НМЦК_conv"] = pi_full["НМЦК"] * display_rate / 1e6
    pi_full["Цена_conv"] = pi_full["Сумма тендера"] * display_rate / 1e6

    fig_pi3 = px.scatter(
        pi_full.sample(min(200, len(pi_full)), random_state=42),
        x="НМЦК_conv", y="Цена_conv", color="Статус",
        color_discrete_map=STATUS_COLORS,
        hover_name="Заказчик",
        hover_data={"Снижение %":":.1f","Тип работ":True,"НМЦК_conv":":.2f","Цена_conv":":.2f"},
        opacity=0.7,
        labels={"НМЦК_conv":f"НМЦК (млн {currency_symbol})", "Цена_conv":f"Итоговая цена (млн {currency_symbol})"},
        trendline="ols",
    )
    # Линия 1:1
    max_val = max(pi_full["НМЦК_conv"].max(), pi_full["Цена_conv"].max())
    fig_pi3.add_trace(go.Scatter(
        x=[0, max_val], y=[0, max_val],
        mode="lines", name="НМЦК = цена",
        line=dict(color="#8b949e", dash="dash", width=1),
    ))
    apply_layout(fig_pi3, title="НМЦК vs итоговая цена контракта (ниже диагонали = снижение)", height=380)
    st.plotly_chart(fig_pi3, use_container_width=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК: ДИНАМИКА
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📈 Динамика выигранных тендеров</div>', unsafe_allow_html=True)
col_g1, col_g2 = st.columns([3,2])

with col_g1:
    granularity = st.radio("Группировка", ["Месяц","Квартал","Год"], horizontal=True)
    if not won.empty:
        if granularity == "Месяц":
            monthly = won.groupby("Месяц").agg(Сумма=("Сумма тендера","sum"),Прибыль=("Чистая прибыль","sum"),Количество=("Статус","count")).reset_index().rename(columns={"Месяц":"Период"})
        elif granularity == "Квартал":
            w2 = won.copy(); w2["Q"] = w2["Дата подачи"].dt.to_period("Q").dt.to_timestamp()
            monthly = w2.groupby("Q").agg(Сумма=("Сумма тендера","sum"),Прибыль=("Чистая прибыль","sum"),Количество=("Статус","count")).reset_index().rename(columns={"Q":"Период"})
        else:
            monthly = won.groupby("Год").agg(Сумма=("Сумма тендера","sum"),Прибыль=("Чистая прибыль","sum"),Количество=("Статус","count")).reset_index().rename(columns={"Год":"Период"})

        monthly["S"] = monthly["Сумма"]*display_rate/1e6
        monthly["P"] = monthly["Прибыль"]*display_rate/1e6
        fig1 = make_subplots(specs=[[{"secondary_y":True}]])
        fig1.add_trace(go.Bar(x=monthly["Период"],y=monthly["S"],name=f"Выручка (млн {currency_symbol})",marker_color="#1f6feb",opacity=0.85),secondary_y=False)
        fig1.add_trace(go.Bar(x=monthly["Период"],y=monthly["P"],name=f"Прибыль (млн {currency_symbol})",marker_color="#3fb950",opacity=0.7),secondary_y=False)
        fig1.add_trace(go.Scatter(x=monthly["Период"],y=monthly["Количество"],name="Кол-во",line=dict(color="#d29922",width=2),mode="lines+markers"),secondary_y=True)
        fig1.update_yaxes(title_text=f"млн {currency_symbol}",secondary_y=False,title_font_size=11)
        fig1.update_yaxes(title_text="Количество",secondary_y=True,title_font_size=11,showgrid=False)
        apply_layout(fig1,title="Выручка и прибыль по периодам",barmode="overlay")
        st.plotly_chart(fig1,use_container_width=True)

with col_g2:
    sc = df["Статус"].value_counts().reset_index(); sc.columns=["Статус","Количество"]
    colors_pie = [STATUS_COLORS.get(s,"#8b949e") for s in sc["Статус"]]
    fig_d = go.Figure(go.Pie(labels=sc["Статус"],values=sc["Количество"],hole=0.62,
                              marker=dict(colors=colors_pie,line=dict(color="#0d1117",width=3)),textfont=dict(size=12)))
    fig_d.add_annotation(text=f"<b>{win_rate:.0f}%</b>",x=0.5,y=0.5,font=dict(size=26,family="Rajdhani",color="#e6edf3"),showarrow=False)
    fig_d.add_annotation(text="конверсия",x=0.5,y=0.38,font=dict(size=11,color="#8b949e"),showarrow=False)
    apply_layout(fig_d,title="Распределение статусов")
    st.plotly_chart(fig_d,use_container_width=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК: LFL
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📅 Сравнение периодов (Like-for-Like)</div>', unsafe_allow_html=True)
lfl_c1, lfl_c2 = st.columns([1,3])
avail_years = sorted(raw_df["Год"].dropna().unique().astype(int).tolist())
with lfl_c1:
    if len(avail_years) >= 2:
        year_a = st.selectbox("Базовый год", avail_years[:-1], index=len(avail_years)-2)
        year_b = st.selectbox("Сравниваемый год", avail_years[1:], index=len(avail_years)-2)
    else:
        year_a = year_b = avail_years[0] if avail_years else 2021
with lfl_c2:
    today_doy = datetime.now().timetuple().tm_yday
    df_a = raw_df[(raw_df["Год"]==year_a)&(raw_df["Статус"]=="Выигран")]
    df_b = raw_df[(raw_df["Год"]==year_b)&(raw_df["Статус"]=="Выигран")]
    df_a_lfl = df_a[df_a["Дата подачи"].dt.dayofyear <= today_doy]
    df_b_lfl = df_b[df_b["Дата подачи"].dt.dayofyear <= today_doy]
    metrics_lfl = {
        "Кол-во тендеров":(len(df_a_lfl),len(df_b_lfl)),
        f"Выручка (млн {currency_symbol})":(round(c(df_a_lfl["Сумма тендера"].sum())/1e6,1),round(c(df_b_lfl["Сумма тендера"].sum())/1e6,1)),
        f"Прибыль (млн {currency_symbol})":(round(c(df_a_lfl["Чистая прибыль"].sum())/1e6,1) if "Чистая прибыль" in df_a_lfl.columns else 0,
                                             round(c(df_b_lfl["Чистая прибыль"].sum())/1e6,1) if "Чистая прибыль" in df_b_lfl.columns else 0),
    }
    lfl_cols = st.columns(len(metrics_lfl))
    for i,(metric,(va,vb)) in enumerate(metrics_lfl.items()):
        dv = vb-va; dp = (dv/max(abs(va),1))*100
        col = "#3fb950" if dv>=0 else "#f85149"; arrow = "▲" if dv>=0 else "▼"
        lfl_cols[i].markdown(f"""
        <div class="metric-card">
            <div class="metric-label">{metric}</div>
            <div class="metric-value">{vb}</div>
            <div style="color:{col};font-size:0.85rem">{arrow} {abs(dp):.1f}% vs {year_a} ({va})</div>
        </div>""",unsafe_allow_html=True)

if year_a != year_b:
    ma_ = df_a.copy(); ma_["M"] = ma_["Дата подачи"].dt.month
    mb_ = df_b.copy(); mb_["M"] = mb_["Дата подачи"].dt.month
    ma_g = ma_.groupby("M")["Сумма тендера"].sum().reset_index()
    mb_g = mb_.groupby("M")["Сумма тендера"].sum().reset_index()
    fig_lfl = go.Figure()
    fig_lfl.add_trace(go.Scatter(x=ma_g["M"],y=ma_g["Сумма тендера"]*display_rate/1e6,name=str(year_a),line=dict(color="#8b949e",width=2,dash="dash")))
    fig_lfl.add_trace(go.Scatter(x=mb_g["M"],y=mb_g["Сумма тендера"]*display_rate/1e6,name=str(year_b),line=dict(color="#1f6feb",width=2.5)))
    fig_lfl.update_xaxes(tickvals=list(range(1,13)),ticktext=["Янв","Фев","Мар","Апр","Май","Июн","Июл","Авг","Сен","Окт","Ноя","Дек"])
    apply_layout(fig_lfl,title=f"Выручка по месяцам: {year_a} vs {year_b} (млн {currency_symbol})")
    st.plotly_chart(fig_lfl,use_container_width=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК: ФИНАНСЫ
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">💰 Финансовый блок — Маржинальность</div>', unsafe_allow_html=True)
fc1,fc2 = st.columns(2)
with fc1:
    if "Тип работ" in won.columns:
        mt = won.groupby("Тип работ").agg(Выручка=("Сумма тендера","sum"),Себест=("Себестоимость","sum"),Прибыль=("Чистая прибыль","sum"),Маржа=("Маржа %","mean")).reset_index().sort_values("Маржа",ascending=True)
        mt["S"]=mt["Выручка"]*display_rate/1e6; mt["C"]=mt["Себест"]*display_rate/1e6; mt["P"]=mt["Прибыль"]*display_rate/1e6
        fig_f1 = go.Figure()
        fig_f1.add_trace(go.Bar(y=mt["Тип работ"],x=mt["C"],name=f"Себестоимость",orientation="h",marker_color="#f85149",opacity=0.8))
        fig_f1.add_trace(go.Bar(y=mt["Тип работ"],x=mt["P"],name=f"Прибыль",orientation="h",marker_color="#3fb950",opacity=0.9))
        apply_layout(fig_f1,title="Структура: Себестоимость + Прибыль",barmode="stack",height=380,xaxis_title=f"млн {currency_symbol}")
        st.plotly_chart(fig_f1,use_container_width=True)
with fc2:
    if "Тип работ" in won.columns:
        colors_m = ["#f85149" if m<15 else "#d29922" if m<22 else "#3fb950" for m in mt["Маржа"]]
        fig_f2 = go.Figure()
        fig_f2.add_trace(go.Bar(y=mt["Тип работ"],x=mt["Маржа"],orientation="h",
                                 marker=dict(color=colors_m,line=dict(width=0)),
                                 text=[f"{v:.1f}%" for v in mt["Маржа"]],textposition="outside"))
        fig_f2.add_shape(type="line",x0=20,x1=20,y0=-0.5,y1=len(mt)-0.5,line=dict(color="#d29922",width=1.5,dash="dash"))
        apply_layout(fig_f2,title="Маржинальность по типу работ",height=380,xaxis_range=[0,40],xaxis_title="Маржа %")
        st.plotly_chart(fig_f2,use_container_width=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК: КАРТА
# ═══════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🗺️ Географическая карта тендеров</div>', unsafe_allow_html=True)

if "Лат" in df.columns and "Лон" in df.columns:
    map_filter = st.radio("Показать на карте", ["Все", "Выигран", "Проигран", "В процессе"], horizontal=True)
    map_df = df if map_filter == "Все" else df[df["Статус"] == map_filter]
    map_df = map_df.copy()
    map_df["Сумма_conv"] = map_df["Сумма тендера"] * display_rate / 1e6
    map_df["Размер"]     = np.sqrt(map_df["Сумма тендера"] / map_df["Сумма тендера"].max()) * 30 + 4

    fig_map = px.scatter_geo(
        map_df,
        lat="Лат", lon="Лон",
        color="Статус",
        size="Размер",
        hover_name="Заказчик",
        hover_data={
            "Регион":    True,
            "Сумма_conv": ":.1f",
            "Статус":    True,
            "Лат":       False,
            "Лон":       False,
            "Размер":    False,
        },
        color_discrete_map=STATUS_COLORS,
        projection="natural earth",
        labels={"Сумма_conv": f"Сумма (млн {currency_symbol})"},
    )
    fig_map.update_geos(
        # ── Настроено на Казахстан ──────────────────────────
        scope="asia",
        center={"lat": 48.0, "lon": 68.0},
        projection_scale=4.8,
        # ── Стили ───────────────────────────────────────────
        bgcolor="#0d1117",
        landcolor="#1c2333",
        oceancolor="#0d1117",
        showocean=True,
        lakecolor="#161b22",
        showrivers=True,    rivercolor="#21262d",
        showcountries=True, countrycolor="#30363d",
        showcoastlines=True, coastlinecolor="#30363d",
        showframe=False,
    )
    apply_layout(fig_map, title="Карта тендеров по регионам Казахстана (размер = сумма)",
                  height=520, geo=dict(bgcolor="#0d1117"))
    st.plotly_chart(fig_map, use_container_width=True)

    # ── Сводная таблица по регионам под картой ─────────────
    if "Регион" in df.columns:
        region_summary = df.groupby("Регион").agg(
            Всего=("Сумма тендера", "count"),
            Выиграно=("Статус", lambda x: (x == "Выигран").sum()),
            Сумма=("Сумма тендера", "sum"),
        ).reset_index()
        region_summary["Конверсия %"] = (
            region_summary["Выиграно"] / region_summary["Всего"] * 100
        ).round(1)
        region_summary[f"Сумма (млн {currency_symbol})"] = (
            region_summary["Сумма"] * display_rate / 1e6
        ).round(1)
        region_summary = region_summary.drop(columns=["Сумма"]).sort_values("Всего", ascending=False)

        with st.expander("📊 Сводка по регионам"):
            st.dataframe(region_summary, use_container_width=True, height=300)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# БЛОК: ЗАКАЗЧИКИ + КОНКУРЕНТЫ
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🏢 Заказчики</div>', unsafe_allow_html=True)
cc1,cc2 = st.columns(2)
with cc1:
    topn = st.slider("Топ заказчиков",5,20,10)
    bc = df.groupby("Заказчик").agg(Всего=("Сумма тендера","sum"),Выиграно=("Статус",lambda x:(x=="Выигран").sum()),Проиграно=("Статус",lambda x:(x=="Проигран").sum())).nlargest(topn,"Всего").reset_index()
    bc["Ж"]=bc["Заказчик"].str[:25]
    fig_c1=go.Figure()
    fig_c1.add_trace(go.Bar(y=bc["Ж"],x=bc["Выиграно"],name="Выиграно",orientation="h",marker_color="#3fb950"))
    fig_c1.add_trace(go.Bar(y=bc["Ж"],x=bc["Проиграно"],name="Проиграно",orientation="h",marker_color="#f85149"))
    apply_layout(fig_c1,title=f"Топ-{topn} заказчиков",barmode="stack",height=400)
    st.plotly_chart(fig_c1,use_container_width=True)
with cc2:
    cs = df.groupby("Заказчик").agg(Сумма=("Сумма тендера","sum"),Кол=("Статус","count"),Конверсия=("Статус",lambda x:(x=="Выигран").sum()/max(len(x[x!="В процессе"]),1)*100)).reset_index()
    cs["Сумма_c"]=cs["Сумма"]*display_rate/1e6
    fig_c2=px.scatter(cs,x="Кол",y="Сумма_c",size="Конверсия",color="Конверсия",hover_name="Заказчик",
                       color_continuous_scale=["#f85149","#d29922","#3fb950"],size_max=30,
                       labels={"Сумма_c":f"Сумма (млн {currency_symbol})","Конверсия":"Конверсия %"})
    apply_layout(fig_c2,title="Объём vs количество",height=400)
    st.plotly_chart(fig_c2,use_container_width=True)

if "Конкурент" in df.columns:
    st.markdown('<div class="section-title">⚔️ Конкуренты</div>', unsafe_allow_html=True)
    cmp_df = df[df["Конкурент"].notna()].copy()
    cm1,cm2 = st.columns(2)
    with cm1:
        cs2 = cmp_df.groupby("Конкурент").agg(Всего=("Статус","count"),Win=("Статус",lambda x:(x=="Выигран").sum()),Loss=("Статус",lambda x:(x=="Проигран").sum())).reset_index()
        cs2["WR"]=(cs2["Win"]/cs2["Всего"]*100).round(1)
        fig_cm=go.Figure()
        fig_cm.add_trace(go.Bar(y=cs2["Конкурент"],x=cs2["Win"],name="Выиграли",orientation="h",marker_color="#3fb950"))
        fig_cm.add_trace(go.Bar(y=cs2["Конкурент"],x=cs2["Loss"],name="Проиграли",orientation="h",marker_color="#f85149"))
        apply_layout(fig_cm,title="Win/Loss vs конкурентов",barmode="stack",height=300)
        st.plotly_chart(fig_cm,use_container_width=True)
    with cm2:
        fig_cm2=go.Figure(go.Bar(x=cs2["Конкурент"],y=cs2["WR"],
                                  marker=dict(color=cs2["WR"],colorscale=["#f85149","#d29922","#3fb950"],cmin=0,cmax=100,line=dict(width=0)),
                                  text=[f"{v}%" for v in cs2["WR"]],textposition="outside"))
        apply_layout(fig_cm2,title="Наш Win Rate по конкурентам",height=300,yaxis_range=[0,110],xaxis_tickangle=-20)
        st.plotly_chart(fig_cm2,use_container_width=True)

st.markdown("---")


# ═══════════════════════════════════════════════════════════
# ТАБЛИЦА С УМНЫМ ПОИСКОМ
# ═══════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📋 Таблица тендеров</div>', unsafe_allow_html=True)

sr1,sr2 = st.columns([3,1])
with sr1:
    sq = st.text_input("🔍 Поиск по заказчику, номеру, типу работ, региону",placeholder="Ключевое слово...")
with sr2:
    tsort = st.selectbox("Сортировка",["Дата ↓","Дата ↑","Сумма ↓","Сумма ↑","Балл ↓"])

show_cols = [c for c in ["Номер тендера","Дата подачи","Заказчик","Тип работ","Регион",
                          "Сумма тендера","Себестоимость","Чистая прибыль","Маржа %",
                          "Снижение %","Валюта","Статус","Конкурент","Срок (дней)",
                          "Дней до дедлайна","Балл","Рекомендация"] if c in df.columns]
tdf = df[show_cols].copy()

if sq:
    mask = pd.Series([False]*len(tdf), index=tdf.index)
    for col in ["Заказчик","Номер тендера","Тип работ","Регион","Конкурент","Рекомендация"]:
        if col in tdf.columns:
            mask |= tdf[col].astype(str).str.contains(sq, case=False, na=False)
    tdf = tdf[mask]

smap = {"Дата ↓":("Дата подачи",False),"Дата ↑":("Дата подачи",True),
        "Сумма ↓":("Сумма тендера",False),"Сумма ↑":("Сумма тендера",True),
        "Балл ↓":("Балл",False)}
sc_col, sc_asc = smap.get(tsort, ("Дата подачи",False))
if sc_col in tdf.columns: tdf = tdf.sort_values(sc_col, ascending=sc_asc)

st.markdown(f"*Найдено: {len(tdf):,} тендеров*")

def color_status(val):
    return {"Выигран":"background-color:#0d2b17","Проигран":"background-color:#2b0d0d","В процессе":"background-color:#2b2200"}.get(val,"")
def color_score(val):
    try:
        v=float(val)
        if v>=65: return "background-color:#0d2b17;color:#3fb950;font-weight:700"
        if v>=45: return "background-color:#2b1e00;color:#d29922;font-weight:700"
        return "background-color:#2b0d0d;color:#f85149;font-weight:700"
    except: return ""
def color_rec(val):
    if "Рекомендуем" in str(val): return "color:#3fb950"
    if "Средний" in str(val): return "color:#d29922"
    return "color:#f85149"
def color_ddl(val):
    try:
        v=float(val)
        if v<=2: return "background-color:#2b0d0d;color:#f85149"
        if v<=7: return "background-color:#2b1e00;color:#d29922"
    except: pass
    return ""

styled = tdf.style
if "Статус" in tdf.columns: styled = styled.applymap(color_status, subset=["Статус"])
if "Балл" in tdf.columns: styled = styled.applymap(color_score, subset=["Балл"])
if "Рекомендация" in tdf.columns: styled = styled.applymap(color_rec, subset=["Рекомендация"])
if "Дней до дедлайна" in tdf.columns: styled = styled.applymap(color_ddl, subset=["Дней до дедлайна"])
if "Маржа %" in tdf.columns: styled = styled.format({"Маржа %":"{:.1f}%"})
if "Снижение %" in tdf.columns: styled = styled.format({"Снижение %":"{:.1f}%"})
st.dataframe(styled, use_container_width=True, height=420)

buf_out = io.BytesIO(); tdf.to_excel(buf_out, index=False)
st.download_button("⬇️ Экспорт выборки Excel", buf_out.getvalue(), "filtered_tenders.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.markdown(f"<center style='color:#484f58;font-size:0.75rem'>Тендерная аналитика Pro · Streamlit + Plotly · {currency_code}</center>",
            unsafe_allow_html=True)