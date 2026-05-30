import pandas as pd
import re
from datetime import datetime
from typing import Tuple
from dataclasses import dataclass
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class Config:
    INPUT_FILE: str = "dashboard_data.xlsx"
    OUTPUT_FILE: str = "mf_ranked_screener.xlsx"
    TREND_BONUS: float = 5.0

CONFIG = Config()

QUALITY_FILTERS = {
    "cagr_2y_min": 0.10,
    "cagr_3y_min": 0.12,
    "min_1y_return": -30.0,
}

ENGINE1_WEIGHTS = {
    "return_6m": 0.30,
    "return_3m": 0.20,
    "return_1y": 0.25,
    "return_1m": 0.25,
}

ENGINE2_WEIGHTS = {
    "return_1y": 0.25,
    "return_2y": 0.30,
    "return_3y": 0.45,
}

COMPOSITE_BLEND = {
    "engine1_momentum": 0.55,
    "engine2_quality": 0.45,
}

FILTERS = {
    "cat_level_1": "Open Ended Schemes",
    "cat_level_2": "Other Scheme",
    "plan_type": "Regular",
    "option_type": "Growth",
}

COLUMN_MAP = {
    "scheme_name": "scheme_name",
    "category": "cat_level_3",
    "amc": "amc_name",
    "return_1m": "return_30d",
    "return_3m": "return_90d",
    "return_6m": "return_180d",
    "return_1y": "return_365d",
    "return_2y": "return_730d",
    "return_3y": "return_1095d",
    "nav": "latest_nav",
}

# ══════════════════════════════════════════════════════════════════════════════
# ASSET TAGGING RULES - Priority-Ordered
# ══════════════════════════════════════════════════════════════════════════════
ASSET_TAG_RULES = [
    {"tag": "Silver", "contains": ["silver"], "priority": 1},
    {"tag": "Gold", "contains": ["gold"], "priority": 2},
    {"tag": "G-SEC", "contains": ["g-sec", "gsec"], "priority": 3},
    {"tag": "GILT", "contains": ["gilt"], "priority": 4},
    {"tag": "NASDAQ", "contains": ["nasdaq", "nq100"], "priority": 5},
    {"tag": "S&P 500", "contains": ["s&p 500", "s&p500", "s&p 50"], "priority": 6},
    {"tag": "International", "contains": ["overseas", "global", "us equity", "emerging market", "china", "hang seng"], "priority": 7},
    {"tag": "Pharma/Healthcare", "contains": ["pharma", "healthcare", "health"], "priority": 8},
    {"tag": "Technology", "contains": ["tech", "artificial intelligence", "ai ", "fang"], "priority": 9},
    {"tag": "Commodities", "contains": ["commodit", "metal", "energy", "mining"], "priority": 10},
    {"tag": "Index Fund", "contains": ["index", "nifty", "sensex", "bse", "midcap", "smallcap"], "priority": 11},
]

# ══════════════════════════════════════════════════════════════════════════════
# COLORS
# ══════════════════════════════════════════════════════════════════════════════
class Colors:
    TITLE_BG = "0D1117"
    TITLE_FG = "FFFFFF"
    INFO_BG = "F0F4F8"
    INFO_FG = "445566"
    BORDER = "D0D8E4"
    COL_HDR_BG = "1A3A5C"
    COL_HDR_FG = "FFFFFF"
    MOMENTUM_BG = "E8560A"
    MOMENTUM_FG = "FFFFFF"
    LONGTERM_BG = "1E6B8C"
    LONGTERM_FG = "FFFFFF"
    ENGINE_BG = "2D5016"
    ENGINE_FG = "FFFFFF"
    RANK1_BG = "FFD700"
    RANK2_BG = "E8E8E8"
    RANK3_BG = "D4956A"
    ALT_ROW = "F5F8FC"
    WHITE = "FFFFFF"
    POSITIVE = "1E7A4B"
    NEGATIVE = "C0392B"
    ENGINE1_TINT = "FFF3E0"
    ENGINE2_TINT = "E3F2FD"
    COMP_TINT = "F0FFF4"
    MISSING = "888888"
    ASMP_TITLE = "0D1117"
    ASMP_BLEND = "2D5016"
    ASMP_ROW_BL = "EAFAF1"

C = Colors()

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for col_key, value in FILTERS.items():
        actual = cols_lower.get(col_key.lower().strip())
        if actual:
            df = df[df[actual].astype(str).str.strip().str.lower() == str(value).strip().lower()]
    return df

def load_data() -> pd.DataFrame:
    sheets = pd.read_excel(CONFIG.INPUT_FILE, sheet_name=None)
    frames = [df for df in sheets.values() if "scheme_name" in df.columns]
    if not frames:
        raise ValueError("No valid sheets found with 'scheme_name' column")
    return apply_filters(pd.concat(frames, ignore_index=True))

# ══════════════════════════════════════════════════════════════════════════════
# SCORING ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def to_num(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace('%', '', regex=False).str.replace(',', '', regex=False).str.strip()
    return pd.to_numeric(s, errors='coerce')

def pct_rank(series: pd.Series) -> pd.Series:
    valid = series.dropna()
    if valid.empty:
        return series.fillna(0.0)
    ranks = series.rank(method='min', na_option='bottom')
    return (ranks - 1) / max(len(series) - 1, 1) * 100

def cagr_2y(r: float) -> float:
    if pd.isna(r) or r <= -100:
        return float('nan')
    return ((1 + float(r) / 100) ** 0.5 - 1) * 100

def assign_asset_tag(scheme_name: str) -> str:
    name_lower = scheme_name.lower()
    for rule in sorted(ASSET_TAG_RULES, key=lambda x: x.get("priority", 99)):
        for keyword in rule.get("contains", []):
            if keyword in name_lower:
                return rule["tag"]
    return "Standard Equity/Debt"

def calculate_trend_strength(row: pd.Series) -> Tuple[float, str]:
    r1m, r3m, r6m = row["_r1m"], row["_r3m"], row["_r6m"]
    if pd.isna(r1m) or pd.isna(r3m) or pd.isna(r6m):
        return 0, ""
    
    if r6m > r3m > r1m:
        return CONFIG.TREND_BONUS, "📈 Uptrend"
    elif r6m > r3m or r3m > r1m:
        return CONFIG.TREND_BONUS / 2, "↗️ Moderate"
    elif r6m < r3m < r1m:
        return -CONFIG.TREND_BONUS / 2, "📉 Downtrend"
    return 0, ""

def score_funds(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    
    # Numeric conversions
    df["_r1m"] = to_num(df[COLUMN_MAP["return_1m"]])
    df["_r3m"] = to_num(df[COLUMN_MAP["return_3m"]])
    df["_r6m"] = to_num(df[COLUMN_MAP["return_6m"]])
    df["_r1y"] = to_num(df[COLUMN_MAP["return_1y"]])
    df["_r2y_raw"] = to_num(df[COLUMN_MAP["return_2y"]])
    df["_r2y_cagr"] = df["_r2y_raw"].apply(cagr_2y)
    df["_r3y"] = to_num(df[COLUMN_MAP["return_3y"]])
    df["_cat"] = df[COLUMN_MAP["category"]].astype(str).str.strip().str.title()
    
    # Asset classification
    df["_asset_class"] = df[COLUMN_MAP["scheme_name"]].apply(assign_asset_tag)
    
    # Missing data detection
    essential_returns = ["_r1m", "_r3m", "_r6m", "_r1y", "_r2y_cagr", "_r3y"]
    df["_has_missing_data"] = df[essential_returns].isna().any(axis=1)
    
    # Quality filters with drawdown protection
    mask_2y = df["_r2y_cagr"] > (QUALITY_FILTERS["cagr_2y_min"] * 100)
    mask_3y = df["_r3y"] > (QUALITY_FILTERS["cagr_3y_min"] * 100)
    mask_drawdown = df["_r1y"] > QUALITY_FILTERS["min_1y_return"]
    df["_qualifies"] = mask_2y & mask_3y & mask_drawdown & (~df["_has_missing_data"])
    
    # Trend calculation
    trend_results = df.apply(calculate_trend_strength, axis=1)
    df["_trend_bonus"] = trend_results.apply(lambda x: x[0])
    df["_trend"] = trend_results.apply(lambda x: x[1])
    
    # Initialize engine scores
    df["_e1"] = 0.0
    df["_e2"] = 0.0
    
    # Category-wise scoring
    for cat in df["_cat"].unique():
        cm = df["_cat"] == cat
        cq = cm & df["_qualifies"]
        valid_m_mask = cm & (~df["_has_missing_data"])
        
        # Engine 1: Momentum
        if valid_m_mask.sum() > 0:
            e1 = (pct_rank(df.loc[valid_m_mask, "_r6m"]) * ENGINE1_WEIGHTS["return_6m"] +
                  pct_rank(df.loc[valid_m_mask, "_r3m"]) * ENGINE1_WEIGHTS["return_3m"] +
                  pct_rank(df.loc[valid_m_mask, "_r1y"]) * ENGINE1_WEIGHTS["return_1y"] +
                  pct_rank(df.loc[valid_m_mask, "_r1m"]) * ENGINE1_WEIGHTS["return_1m"])
            e1 = e1 + df.loc[valid_m_mask, "_trend_bonus"]
            df.loc[valid_m_mask, "_e1"] = e1.clip(0, 100)
        
        # Engine 2: Quality
        if cq.sum() > 0:
            e2 = (pct_rank(df.loc[cq, "_r1y"]) * ENGINE2_WEIGHTS["return_1y"] +
                  pct_rank(df.loc[cq, "_r2y_cagr"]) * ENGINE2_WEIGHTS["return_2y"] +
                  pct_rank(df.loc[cq, "_r3y"]) * ENGINE2_WEIGHTS["return_3y"])
            df.loc[cq, "_e2"] = e2
    
    # Composite score
    df["_comp"] = (df["_e1"] * COMPOSITE_BLEND["engine1_momentum"] +
                   df["_e2"] * COMPOSITE_BLEND["engine2_quality"])
    
    df.loc[df["_has_missing_data"], "_comp"] = -1.0
    
    # Ranking
    df["_rank"] = (df.groupby("_cat")["_comp"]
                   .rank(method='min', ascending=False)
                   .astype(int))
    
    return df

# ══════════════════════════════════════════════════════════════════════════════
# FORMATTING HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def border() -> Border:
    s = Side(style='thin', color=C.BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def fmt_pct(val) -> str:
    if pd.isna(val) or val is None:
        return "—"
    try:
        return f"{float(val):+.2f}%"
    except:
        return "—"

def score_col(val) -> str:
    try:
        v = float(val)
        if v < 0:
            return C.MISSING
    except:
        return C.MISSING
    if v >= 75:
        return C.POSITIVE
    if v >= 50:
        return "E67E00"
    return C.NEGATIVE

def signal(comp: float, trend: str = "", e1: float = 0, e2: float = 0) -> str:
    if comp < 0:
        return "❌ Missing Data"
    if comp >= 85 and "Uptrend" in str(trend):
        return "🚀 Strong Conviction"
    if comp >= 75:
        return "⭐ Strong Buy"
    if comp >= 60:
        if e1 > e2:
            return "📈 Momentum Play"
        return "🏛️ Quality Hold"
    if comp >= 55:
        return "✅ Buy"
    if comp >= 40:
        return "⚠️ Watch"
    return "🔴 Avoid"

def clean_name(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]]', '', str(name))[:31]

def hfont(size: int = 9, color: str = "FFFFFF", bold: bool = True) -> Font:
    return Font(name="Arial", bold=bold, size=size, color=color)

def dfont(size: int = 9, bold: bool = False, color: str = "000000") -> Font:
    return Font(name="Arial", bold=bold, size=size, color=color)

# Column layout
MOMENTUM_COLS = (5, 7)
LONGTERM_COLS = (8, 10)
ENGINE_COLS = (11, 13)

COL_HEADERS = [
    "Rank", "Scheme Name", "AMC", "Asset Class",
    "1M\nReturn", "3M\nReturn", "6M\nReturn",
    "1Y\nReturn", "2Y\nCAGR", "3Y\nCAGR",
    "Engine 1\n(Momentum)", "Engine 2\n(Quality)", "Composite\nScore",
]

COL_WIDTHS = [6, 50, 20, 18, 9, 9, 9, 9, 9, 9, 14, 14, 12]
RETURN_COLS_IDX = {5, 6, 7, 8, 9, 10}

# ══════════════════════════════════════════════════════════════════════════════
# BUILD CATEGORY SHEET
# ══════════════════════════════════════════════════════════════════════════════
def build_category_sheet(wb, cat, cat_df):
    ws = wb.create_sheet(clean_name(cat))
    bd = border()
    total = len(cat_df)
    qualified = int(cat_df["_qualifies"].sum())
    ncols = len(COL_HEADERS)

    # Title row
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]
    c.value = f"✦  {cat.upper()}  ✦"
    c.font = Font(name="Arial", bold=True, size=14, color=C.TITLE_FG)
    c.fill = fill(C.TITLE_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Info row
    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    c = ws["A2"]
    e1_pct = int(COMPOSITE_BLEND["engine1_momentum"] * 100)
    e2_pct = int(COMPOSITE_BLEND["engine2_quality"] * 100)
    c.value = f"Dual Engine Model: {e1_pct}% Momentum (ST) + {e2_pct}% Quality (LT) | Passed Gates: {qualified}/{total}"
    c.font = Font(name="Arial", italic=True, size=8.5, color=C.INFO_FG)
    c.fill = fill(C.INFO_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 16

    # Group headers row 3
    for ci in range(1, 5):
        ws.cell(row=3, column=ci).fill = fill(C.COL_HDR_BG)
        ws.cell(row=3, column=ci).border = bd

    for g_start, g_end, label, bg, fg in [
        (MOMENTUM_COLS[0], MOMENTUM_COLS[1], "◄  Momentum  ►", C.MOMENTUM_BG, C.MOMENTUM_FG),
        (LONGTERM_COLS[0], LONGTERM_COLS[1], "◄  Long-Term  ►", C.LONGTERM_BG, C.LONGTERM_FG),
        (ENGINE_COLS[0], ENGINE_COLS[1], "◄  Engine Scores  ►", C.ENGINE_BG, C.ENGINE_FG)
    ]:
        ws.merge_cells(f"{get_column_letter(g_start)}3:{get_column_letter(g_end)}3")
        cell = ws.cell(row=3, column=g_start, value=label)
        cell.font = hfont(color=fg)
        cell.fill = fill(bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for ci in range(g_start, g_end + 1):
            ws.cell(row=3, column=ci).fill = fill(bg)
            ws.cell(row=3, column=ci).border = bd
    ws.row_dimensions[3].height = 18

    # Column headers row 4
    for ci, hdr in enumerate(COL_HEADERS, 1):
        c = ws.cell(row=4, column=ci, value=hdr)
        c.font = hfont()
        c.fill = fill(C.COL_HDR_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bd
    ws.row_dimensions[4].height = 28

    # Data rows
    for i, (_, row) in enumerate(cat_df.iterrows(), 5):
        rank = row["_rank"]
        is_missing = row["_has_missing_data"]
        
        if is_missing:
            rbg = C.WHITE
        else:
            rbg = C.RANK1_BG if rank == 1 else (C.RANK2_BG if rank == 2 else (C.RANK3_BG if rank == 3 else (C.ALT_ROW if i % 2 == 0 else C.WHITE)))
        
        vals = [
            "—" if is_missing else rank,
            row.get(COLUMN_MAP["scheme_name"], "—"),
            row.get(COLUMN_MAP["amc"], "—"),
            row["_asset_class"],
            fmt_pct(row["_r1m"]), fmt_pct(row["_r3m"]), fmt_pct(row["_r6m"]),
            fmt_pct(row["_r1y"]), fmt_pct(row["_r2y_cagr"]), fmt_pct(row["_r3y"]),
            "—" if is_missing else round(row["_e1"], 1),
            "—" if is_missing else round(row["_e2"], 1),
            "—" if is_missing else round(row["_comp"], 1)
        ]

        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=ci, value=val)
            c.border = bd
            c.fill = fill(rbg)
            c.alignment = Alignment(horizontal="left" if ci in {2, 3, 4} else "center", vertical="center")
            
            if ci in RETURN_COLS_IDX and isinstance(val, str) and val != "—":
                num = float(val.replace('%', ''))
                c.font = Font(name="Arial", bold=(rank <= 3 and not is_missing), size=9, color=C.POSITIVE if num >= 0 else C.NEGATIVE)
            elif ci == 11 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val))
                c.fill = fill(C.ENGINE1_TINT)
            elif ci == 12 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val))
                c.fill = fill(C.ENGINE2_TINT)
            elif ci == 13 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val))
                c.fill = fill(C.COMP_TINT)
            else:
                c.font = dfont(bold=(rank <= 3 and not is_missing))
                if is_missing:
                    c.font = Font(name="Arial", italic=True, size=9, color=C.MISSING)
        ws.row_dimensions[i].height = 16

    # Column widths
    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    
    # Auto-filter and freeze
    ws.auto_filter.ref = f"A4:{get_column_letter(ncols)}{4 + len(cat_df)}"
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════════════════════
# BUILD SUMMARY SHEET
# ══════════════════════════════════════════════════════════════════════════════
def build_summary(wb, df_scored, categories):
    ws = wb.create_sheet("🏆 SUMMARY", 0)
    bd = border()
    ncols = len(COL_HEADERS) + 1

    # Title
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]
    c.value = "MF INTELLIGENCE — UNIFIED DUAL ENGINE RANKED SCREENER"
    c.font = Font(name="Arial", bold=True, size=15, color=C.TITLE_FG)
    c.fill = fill(C.TITLE_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    # Info row
    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    c = ws["A2"]
    e1_pct = int(COMPOSITE_BLEND["engine1_momentum"] * 100)
    e2_pct = int(COMPOSITE_BLEND["engine2_quality"] * 100)
    c.value = f"Composite Master Framework: {e1_pct}% Engine 1 (Momentum) + {e2_pct}% Engine 2 (Quality Setup) | Filter via 'Asset Class' column"
    c.font = Font(name="Arial", italic=True, size=8.5, color=C.INFO_FG)
    c.fill = fill(C.INFO_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 16

    # Group headers
    for ci in range(1, 5):
        ws.cell(row=3, column=ci).fill = fill(C.COL_HDR_BG)
        ws.cell(row=3, column=ci).border = bd

    for g_start, g_end, label, bg, fg in [
        (MOMENTUM_COLS[0], MOMENTUM_COLS[1], "◄  Momentum  ►", C.MOMENTUM_BG, C.MOMENTUM_FG),
        (LONGTERM_COLS[0], LONGTERM_COLS[1], "◄  Long-Term  ►", C.LONGTERM_BG, C.LONGTERM_FG),
        (ENGINE_COLS[0], ENGINE_COLS[1], "◄  Engine Scores  ►", C.ENGINE_BG, C.ENGINE_FG)
    ]:
        ws.merge_cells(f"{get_column_letter(g_start)}3:{get_column_letter(g_end)}3")
        cell = ws.cell(row=3, column=g_start, value=label)
        cell.font = hfont(color=fg)
        cell.fill = fill(bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for ci in range(g_start, g_end + 1):
            ws.cell(row=3, column=ci).fill = fill(bg)
            ws.cell(row=3, column=ci).border = bd
    ws.row_dimensions[3].height = 18

    # Column headers
    for ci, hdr in enumerate(COL_HEADERS, 1):
        c = ws.cell(row=4, column=ci, value=hdr)
        c.font = hfont()
        c.fill = fill(C.COL_HDR_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bd

    sig_col = len(COL_HEADERS) + 1
    ws.cell(row=4, column=sig_col, value="Signal").font = hfont()
    ws.cell(row=4, column=sig_col).fill = fill(C.COL_HDR_BG)
    ws.cell(row=4, column=sig_col).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=4, column=sig_col).border = bd
    ws.column_dimensions[get_column_letter(sig_col)].width = 16

    # Data rows (top fund per category)
    for i, cat in enumerate(categories, 5):
        cat_df = df_scored[df_scored["_cat"] == cat].sort_values("_rank")
        if cat_df.empty:
            continue
        top = cat_df.iloc[0]
        rbg = C.ALT_ROW if i % 2 == 0 else C.WHITE
        is_missing = top["_has_missing_data"]

        vals = [
            cat, top.get(COLUMN_MAP["scheme_name"], "—"), top.get(COLUMN_MAP["amc"], "—"),
            top["_asset_class"],
            fmt_pct(top["_r1m"]), fmt_pct(top["_r3m"]), fmt_pct(top["_r6m"]),
            fmt_pct(top["_r1y"]), fmt_pct(top["_r2y_cagr"]), fmt_pct(top["_r3y"]),
            "—" if is_missing else round(top["_e1"], 1),
            "—" if is_missing else round(top["_e2"], 1),
            "—" if is_missing else round(top["_comp"], 1)
        ]

        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=ci, value=val)
            c.border = bd
            c.fill = fill(rbg)
            c.alignment = Alignment(horizontal="left" if ci in {1, 2, 3, 4} else "center", vertical="center")
            
            if ci in RETURN_COLS_IDX and isinstance(val, str) and val != "—":
                num = float(val.replace('%', ''))
                c.font = Font(name="Arial", size=9, color=C.POSITIVE if num >= 0 else C.NEGATIVE)
            elif ci == 11 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val))
                c.fill = fill(C.ENGINE1_TINT)
            elif ci == 12 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val))
                c.fill = fill(C.ENGINE2_TINT)
            elif ci == 13 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val))
                c.fill = fill(C.COMP_TINT)
            else:
                c.font = dfont()

        # Signal column
        sig_val = signal(top["_comp"], top.get("_trend", ""), top["_e1"], top["_e2"])
        c_sig = ws.cell(row=i, column=sig_col, value=sig_val)
        c_sig.font = Font(name="Arial", bold=True, size=9)
        c_sig.fill = fill(rbg)
        c_sig.border = bd
        c_sig.alignment = Alignment(horizontal="center", vertical="center")
        if is_missing:
            c_sig.font = Font(name="Arial", italic=True, size=9, color=C.MISSING)
        ws.row_dimensions[i].height = 18

    # Column widths
    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    
    ws.auto_filter.ref = f"A4:{get_column_letter(ncols)}{4 + len(categories)}"
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════════════════════
# BUILD ASSUMPTIONS SHEET
# ══════════════════════════════════════════════════════════════════════════════
def build_assumptions(wb, df_scored):
    ws = wb.create_sheet("📋 ASSUMPTIONS", 1)
    bd = border()

    def section(row, title, bg, fg="FFFFFF", ncols=3):
        ws.merge_cells(f"A{row}:{get_column_letter(ncols)}{row}")
        c = ws[f"A{row}"]
        c.value = title
        c.font = Font(name="Arial", bold=True, size=11, color=fg)
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 22
        return row + 1

    def kv_row(row, key, val, desc="", row_bg="FFFFFF"):
        for ci, text in [(1, key), (2, val), (3, desc)]:
            c = ws.cell(row=row, column=ci, value=text)
            c.font = Font(name="Arial", size=9)
            c.fill = fill(row_bg)
            c.border = bd
            c.alignment = Alignment(horizontal="left" if ci in {1, 3} else "center", vertical="center")
        ws.row_dimensions[row].height = 15
        return row + 1

    def col_headers(row, headers, bg):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
            c.fill = fill(bg)
            c.border = bd
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 16
        return row + 1

    # Title
    ws.merge_cells("A1:C1")
    ws["A1"] = "DUAL ENGINE MODEL — ASSUMPTIONS & METHODOLOGY"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = fill(C.ASMP_TITLE)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Generated timestamp
    ws.merge_cells("A2:C2")
    ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Enhanced Asset Classification"
    ws["A2"].font = Font(name="Arial", italic=True, size=8.5, color=C.INFO_FG)
    ws["A2"].fill = fill(C.INFO_BG)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 14

    r = 4
    r = section(r, "⚖  COMPOSITE BLEND WEIGHTS", C.ASMP_BLEND)
    r = col_headers(r, ["Parameter", "Value", "Description"], C.ASMP_BLEND)
    r = kv_row(r, "Engine 1 Weight (Momentum)", f"{int(COMPOSITE_BLEND['engine1_momentum']*100)}%", "Short-term momentum configuration weight", C.ASMP_ROW_BL)
    r = kv_row(r, "Engine 2 Weight (Quality)", f"{int(COMPOSITE_BLEND['engine2_quality']*100)}%", "Long-term core health score weight", C.ASMP_ROW_BL)
    
    r += 1
    r = section(r, "📊  QUALITY GATES", C.ASMP_BLEND)
    r = col_headers(r, ["Parameter", "Value", "Description"], C.ASMP_BLEND)
    r = kv_row(r, "Min 2Y CAGR", f"{int(QUALITY_FILTERS['cagr_2y_min']*100)}%", "Minimum 2-year CAGR to qualify", C.ASMP_ROW_BL)
    r = kv_row(r, "Min 3Y CAGR", f"{int(QUALITY_FILTERS['cagr_3y_min']*100)}%", "Minimum 3-year CAGR to qualify", C.ASMP_ROW_BL)
    r = kv_row(r, "Drawdown Tolerance", f"{int(QUALITY_FILTERS['min_1y_return'])}%", "Max 1Y loss allowed", C.ASMP_ROW_BL)
    
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 44

# ══════════════════════════════════════════════════════════════════════════════
# MAIN EXECUTOR
# ══════════════════════════════════════════════════════════════════════════════
def main():
    print("🚀 Loading data matrix...")
    df = load_data()
    
    print("⚙️ Executing Scoring Engine...")
    df_scored = score_funds(df)
    
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    categories = sorted(df_scored["_cat"].unique())
    
    print("📊 Building sheets...")
    build_summary(wb, df_scored, categories)
    build_assumptions(wb, df_scored)
    
    for cat in categories:
        cat_df = df_scored[df_scored["_cat"] == cat].sort_values("_rank")
        if not cat_df.empty:
            build_category_sheet(wb, cat, cat_df)
    
    wb.save(CONFIG.OUTPUT_FILE)
    print(f"\n✅ Done! Output: '{CONFIG.OUTPUT_FILE}'")

if __name__ == "__main__":
    main()
