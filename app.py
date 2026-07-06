from __future__ import annotations

import io
import os
import re
import base64
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import TypedDict

os.environ.setdefault("MPLCONFIGDIR", "/private/tmp/cpc_matplotlib")

try:
    import fitz
except ImportError:
    fitz = None

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as RLImage
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

matplotlib.use("Agg")

APP_TITLE = "CPC 광고 운영보고서 자동 생성"
PRIMARY = colors.HexColor("#111111")
ACCENT = colors.HexColor("#FEE500")
ACCENT_SOFT = colors.HexColor("#FFF8CC")
TEXT = colors.HexColor("#1F1F1F")
MUTED = colors.HexColor("#737373")
LIGHT_BG = colors.HexColor("#FAFAF7")
LINE = colors.HexColor("#E8E4D8")
SPEND_COL = "소진비용(위안)"
CPC_COL = "클릭당 평균비용(위안)"
BASE_DIR = Path(__file__).resolve().parent
FONT_DIR = BASE_DIR / "assets" / "fonts"
NOTO_FONT = BASE_DIR / "assets" / "fonts" / "NotoSansKR-Static-Regular.ttf"
NOTO_VAR_FONT = BASE_DIR / "assets" / "fonts" / "NotoSansKR-Regular.ttf"

CANONICAL_COLUMNS = {
    "date": "날짜",
    "impressions": "노출수",
    "clicks": "클릭수",
    "spend": SPEND_COL,
    "cpc": CPC_COL,
    "voucher": "상품권 조회수",
    "save": "저장하기",
    "share": "공유하기",
}

COLUMN_PATTERNS = {
    "date": ["日期", "时间", "日", "date", "day", "날짜", "일자"],
    "impressions": ["曝光（次）", "曝光", "展示", "展现", "impression", "노출"],
    "clicks": ["点击（次）", "点击", "click", "클릭"],
    "spend": ["花费（元）", "花费", "消耗", "消费", "金额", "费用", "cost", "spend", "소진", "비용"],
    "cpc": ["平均点击价格（元）", "平均点击价格", "平均点击", "点击价格", "CPC", "平均价格", "클릭당", "평균비용"],
    "voucher": ["查看团购", "团购", "券", "商品券", "voucher", "product view", "상품권", "조회"],
    "save": ["收藏", "收藏数", "save", "saved", "저장", "저장하기"],
    "share": ["分享", "分享数", "share", "shared", "공유", "공유하기"],
}

REQUIRED_COLUMNS = ["날짜", SPEND_COL, "노출수", "클릭수", CPC_COL, "상품권 조회수", "저장하기", "공유하기"]
APP_NAME = "내부 업무 자동화"


@dataclass(frozen=True)
class NavItem:
    key: str
    label: str
    group: str
    description: str = ""


NAV_ITEMS = [
    NavItem("cpc_report", "CPC 보고서 생성", "업무 자동화", "중국어 CPC 데이터로 점주 전달용 운영 보고서를 생성합니다."),
    NavItem("contract_generator", "계약서 자동 생성", "업무 자동화", "계약서 자동 작성 기능을 준비 중입니다."),
    NavItem("manual_cpc_registration", "CPC 등록 가이드", "업무 매뉴얼"),
    NavItem("manual_permission_transfer", "CPT/CPC 권한 이전 가이드", "업무 매뉴얼"),
    NavItem("manual_backoffice_extract", "백오피스 데이터 추출 가이드", "업무 매뉴얼"),
    NavItem("manual_backoffice_cpc", "백오피스 CPC 등록 방법", "업무 매뉴얼"),
    NavItem("manual_review_delete", "리뷰 삭제 요청 가이드", "업무 매뉴얼"),
    NavItem("manual_faq", "자주 묻는 질문(FAQ)", "업무 매뉴얼"),
    NavItem("manual_search", "검색 기능", "업무 매뉴얼"),
]

NAV_BY_KEY = {item.key: item for item in NAV_ITEMS}
DEFAULT_NAV_KEY = "cpc_report"


@dataclass
class StoreInputs:
    store_name: str
    store_type: str
    period_start: date
    period_end: date
    pause_date: date | None
    memo: str


class ReportFonts(TypedDict):
    medium: str
    bold: str
    extra_bold: str
    chart_path: str | None


def setup_page() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="📊", layout="wide", initial_sidebar_state="expanded")
    st.markdown(
        f"""
        <style>
        {report_font_face_css()}
        .stApp {{
            background: #f6f7f9;
            font-family: "Pretendard", -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Noto Sans KR", "Malgun Gothic", sans-serif;
            font-variant-numeric: tabular-nums;
            font-feature-settings: "tnum";
        }}
        [data-testid="stSidebar"] {{
            background: #111827;
        }}
        [data-testid="stSidebar"] * {{
            color: #f9fafb;
        }}
        [data-testid="stSidebar"] .stRadio label,
        [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {{
            color: #f9fafb;
        }}
        .sidebar-title {{
            padding: 10px 4px 18px;
            font-size: 18px;
            font-weight: 800;
            letter-spacing: 0;
        }}
        .sidebar-caption {{
            margin-top: -14px;
            margin-bottom: 18px;
            color: #9ca3af;
            font-size: 13px;
            line-height: 1.5;
        }}
        .app-header {{
            padding: 24px 28px;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            background: #ffffff;
            margin-bottom: 22px;
        }}
        .app-header .eyebrow {{
            margin: 0 0 8px;
            color: #6b7280;
            font-size: 13px;
            font-weight: 700;
        }}
        .app-header h1 {{
            margin: 0;
            color: #111827;
            font-size: 30px;
            font-weight: 800;
            letter-spacing: 0;
        }}
        .app-header p {{
            margin: 8px 0 0;
            color: #4b5563;
            line-height: 1.55;
        }}
        .coming-soon {{
            padding: 52px 32px;
            border: 1px dashed #cbd5e1;
            border-radius: 12px;
            background: #ffffff;
            text-align: center;
        }}
        .coming-soon h2 {{
            margin: 0;
            color: #111827;
            font-size: 24px;
            font-weight: 800;
            letter-spacing: 0;
        }}
        .coming-soon p {{
            margin: 10px auto 0;
            max-width: 520px;
            color: #6b7280;
            line-height: 1.6;
        }}
        div[data-testid="stMetric"] {{
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            padding: 12px;
            min-width: 120px;
        }}
        div[data-testid="stMetricValue"] {{ white-space: normal; overflow-wrap: anywhere; font-weight: 800; }}
        .hint {{ color: #4b5563; font-size: 14px; line-height: 1.6; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_page_header(title: str, description: str, eyebrow: str = APP_NAME) -> None:
    st.markdown(
        f"""
        <div class="app-header">
            <p class="eyebrow">{eyebrow}</p>
            <h1>{title}</h1>
            <p>{description}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> str:
    st.sidebar.markdown(
        """
        <div class="sidebar-title">내부 업무 자동화</div>
        <div class="sidebar-caption">반복 업무 도구와 매뉴얼을 한 곳에서 관리합니다.</div>
        """,
        unsafe_allow_html=True,
    )
    selected_key = st.session_state.get("active_nav", DEFAULT_NAV_KEY)
    items_by_group: dict[str, list[NavItem]] = {}
    for item in NAV_ITEMS:
        items_by_group.setdefault(item.group, []).append(item)

    for group, items in items_by_group.items():
        st.sidebar.markdown(f"**{group}**")
        for item in items:
            button_type = "primary" if item.key == selected_key else "secondary"
            if st.sidebar.button(item.label, key=f"nav_{item.key}", type=button_type, use_container_width=True):
                selected_key = item.key

    st.session_state["active_nav"] = selected_key
    return selected_key


def normalize_column_name(value: object) -> str:
    text = str(value).strip().lower()
    return re.sub(r"[\s_\-()（）【】\[\]:：/\\]+", "", text)


def score_column(column: str, patterns: list[str]) -> float:
    normalized = normalize_column_name(column)
    best = 0.0
    for pattern in patterns:
        target = normalize_column_name(pattern)
        if target and (target in normalized or normalized in target):
            best = max(best, 1.0)
        best = max(best, SequenceMatcher(None, normalized, target).ratio())
    return best


def build_column_mapping(columns: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    used: set[str] = set()
    for canonical, patterns in COLUMN_PATTERNS.items():
        scored = sorted(((score_column(col, patterns), col) for col in columns if col not in used), reverse=True)
        if not scored:
            continue
        score, source = scored[0]
        threshold = 0.48 if canonical in {"date", "impressions", "clicks", "spend"} else 0.55
        if score >= threshold:
            mapping[source] = CANONICAL_COLUMNS[canonical]
            used.add(source)
    return mapping


def parse_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text in {"-", "—", "无", "없음"}:
        return 0.0
    text = text.replace(",", "")
    text = re.sub(r"[元￥₩%\s]", "", text)
    match = re.search(r"-?\d+(\.\d+)?", text)
    return float(match.group()) if match else 0.0


def read_best_sheet(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    xls = pd.ExcelFile(uploaded_file)
    best_sheet = xls.sheet_names[0]
    best_score = -1
    for sheet in xls.sheet_names:
        sample = pd.read_excel(xls, sheet_name=sheet, nrows=20)
        mapping = build_column_mapping([str(col) for col in sample.columns])
        score = len(mapping) + len(sample)
        if score > best_score:
            best_score = score
            best_sheet = sheet
    return pd.read_excel(xls, sheet_name=best_sheet)


def preprocess_dataframe(raw_df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    df = raw_df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    df = df.dropna(how="all")
    mapping = build_column_mapping(list(df.columns))
    df = df.rename(columns=mapping)
    messages: list[str] = []

    if "날짜" not in df.columns:
        possible_date = None
        for col in df.columns:
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().sum() >= max(2, len(df) * 0.4):
                possible_date = col
                break
        if possible_date:
            df = df.rename(columns={possible_date: "날짜"})
            messages.append(f"'{possible_date}' 컬럼을 날짜로 자동 인식했습니다.")
        else:
            df["날짜"] = pd.date_range(datetime.today().date(), periods=len(df), freq="D")
            messages.append("날짜 컬럼을 찾지 못해 행 순서 기준 임시 날짜를 사용했습니다.")

    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    df = df[df["날짜"].notna()].copy()
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = 0.0
            if col != "날짜":
                messages.append(f"'{col}' 컬럼이 없어 0으로 계산했습니다.")

    for col in [SPEND_COL, "노출수", "클릭수", CPC_COL, "상품권 조회수", "저장하기", "공유하기"]:
        df[col] = df[col].apply(parse_number)

    missing_cpc = (df[CPC_COL] <= 0) & (df["클릭수"] > 0)
    df.loc[missing_cpc, CPC_COL] = df.loc[missing_cpc, SPEND_COL] / df.loc[missing_cpc, "클릭수"]

    agg_rules = {SPEND_COL: "sum", "노출수": "sum", "클릭수": "sum", CPC_COL: "mean", "상품권 조회수": "sum", "저장하기": "sum", "공유하기": "sum"}
    df = df[REQUIRED_COLUMNS].sort_values("날짜").groupby("날짜", as_index=False).agg(agg_rules)
    df[CPC_COL] = df.apply(lambda row: row[SPEND_COL] / row["클릭수"] if row["클릭수"] > 0 else row[CPC_COL], axis=1)
    return df, mapping, messages


def split_latest_months(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame | None, date, date, list[str]]:
    if df.empty:
        today = date.today()
        return df, None, today, today, ["업로드 데이터에서 월별 데이터를 확인하지 못했습니다."]

    working = df.copy()
    working["월"] = working["날짜"].dt.to_period("M")
    months = sorted(working["월"].dropna().unique())
    latest_month = months[-1]
    previous_month = latest_month - 1

    current_df = working[working["월"] == latest_month].drop(columns=["월"]).copy()
    previous_df = working[working["월"] == previous_month].drop(columns=["월"]).copy()
    previous_result = previous_df if not previous_df.empty else None

    period_start = current_df["날짜"].min().date()
    period_end = current_df["날짜"].max().date()
    messages = [f"가장 최근 월({latest_month.strftime('%Y.%m')})을 대상월 데이터로 자동 인식했습니다."]
    if previous_result is not None:
        messages.append(f"직전 월({previous_month.strftime('%Y.%m')})을 전월 비교 데이터로 자동 인식했습니다.")
    else:
        messages.append("직전 월 데이터가 없어 전월 대비 비교 없이 현재월 기준으로 보고서를 생성했습니다.")

    return current_df, previous_result, period_start, period_end, messages


def calculate_metrics(df: pd.DataFrame) -> dict[str, float]:
    total_spend = float(df[SPEND_COL].sum()) if not df.empty else 0.0
    total_impressions = float(df["노출수"].sum()) if not df.empty else 0.0
    total_clicks = float(df["클릭수"].sum()) if not df.empty else 0.0
    total_voucher = float(df["상품권 조회수"].sum()) if not df.empty else 0.0
    avg_cpc = total_spend / total_clicks if total_clicks else 0.0
    click_rate = total_clicks / total_impressions if total_impressions else 0.0
    return {
        "total_spend": total_spend,
        "total_impressions": total_impressions,
        "total_clicks": total_clicks,
        "total_voucher": total_voucher,
        "avg_cpc": avg_cpc,
        "click_rate": click_rate,
    }


def format_currency(value: float) -> str:
    return f"{value:,.0f}위안"


def format_cpc_currency(value: float) -> str:
    if abs(value) >= 10:
        return format_currency(value)
    text = f"{value:,.2f}".rstrip("0").rstrip(".")
    return f"{text}위안"


def format_daily_cpc(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{value:,.2f}위안"


def format_rate(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{value * 100:.2f}%"


def format_number(value: float) -> str:
    return f"{value:,.0f}"


def format_short_date(value: date | datetime | pd.Timestamp | None) -> str:
    if value is None or pd.isna(value):
        return "-"
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        value = value.date()
    return value.strftime("%m.%d")


def format_period(start: date | datetime | pd.Timestamp, end: date | datetime | pd.Timestamp) -> str:
    return f"{format_short_date(start)} ~ {format_short_date(end)}"


def format_month_full(value: date | datetime | pd.Timestamp | None) -> str:
    if value is None or pd.isna(value):
        return "-"
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        value = value.date()
    return f"{value.month}월"


def format_report_title_month(value: date | datetime | pd.Timestamp | None) -> str:
    if value is None or pd.isna(value):
        return "월간"
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        value = value.date()
    return f"{value.month}월"


def comparison_value(current: float, previous: float) -> float | None:
    if previous == 0:
        return 0.0 if current == 0 else None
    return ((current - previous) / previous) * 100


def format_change_percent(value: float | None) -> str:
    if value is None:
        return "-"
    if abs(value) < 0.0001:
        return "0%"
    return f"{'▲' if value > 0 else '▼'}{abs(value):.0f}%"


def build_month_comparison(
    current: dict[str, float],
    previous: dict[str, float] | None,
    current_days: int = 1,
    previous_days: int = 1,
    include_change: bool = True,
) -> list[dict[str, object]]:
    fields = [
        ("노출수", "total_impressions", "count"),
        ("클릭수", "total_clicks", "count"),
        ("광고비", "total_spend", "money"),
        ("클릭당 평균비용(CPC)", "avg_cpc", "cpc"),
    ]
    rows = []
    for label, key, kind in fields:
        current_value = float(current[key])
        previous_value = float(previous[key]) if previous else 0.0
        if previous and include_change and kind in {"count", "money"}:
            current_basis = current_value / max(current_days, 1)
            previous_basis = previous_value / max(previous_days, 1)
            change = comparison_value(current_basis, previous_basis)
        else:
            change = comparison_value(current_value, previous_value) if previous and include_change else None
        if kind == "money":
            current_text = format_currency(current_value)
            previous_text = format_currency(previous_value) if previous else "-"
        elif kind == "cpc":
            current_text = format_cpc_currency(current_value)
            previous_text = format_cpc_currency(previous_value) if previous else "-"
        else:
            current_text = f"{format_number(current_value)}회"
            previous_text = f"{format_number(previous_value)}회" if previous else "-"
        rows.append(
            {
                "label": label,
                "current": current_value,
                "previous": previous_value,
                "current_text": current_text,
                "previous_text": previous_text,
                "change": change,
                "change_text": format_change_percent(change) if include_change else "",
            }
        )
    return rows


def build_operation_comment(
    df: pd.DataFrame,
    previous_df: pd.DataFrame | None,
    comparison_excluded: bool,
) -> str | None:
    active_days = int(((df.get("노출수", 0) > 0) | (df.get("클릭수", 0) > 0)).sum()) if not df.empty else 0
    if active_days < 7 or comparison_excluded:
        return "집행 기간이 짧아 데이터 참고성이 제한적입니다."

    current_metrics = calculate_metrics(df)
    if previous_df is None or previous_df.empty:
        return "광고가 정상적으로 운영되며 노출 및 클릭 데이터가 발생하였습니다."

    previous_metrics = calculate_metrics(previous_df)
    current_impressions = current_metrics["total_impressions"]
    previous_impressions = previous_metrics["total_impressions"]
    current_clicks = current_metrics["total_clicks"]
    previous_clicks = previous_metrics["total_clicks"]
    current_ctr = current_metrics["click_rate"]
    previous_ctr = previous_metrics["click_rate"]

    if current_impressions > previous_impressions or current_clicks > previous_clicks:
        return "전월 대비 광고 반응이 개선된 것으로 확인됩니다."
    if current_impressions < previous_impressions or current_clicks < previous_clicks:
        return "전월 대비 광고 반응이 감소하여 지속적인 모니터링이 필요합니다."
    if current_ctr > previous_ctr:
        return "광고 효율이 전월 대비 개선된 것으로 확인됩니다."
    if current_ctr < previous_ctr:
        return "광고 효율 변화 추이를 지속적으로 관찰할 필요가 있습니다."
    return None


def market_trend_note(month: int) -> str | None:
    notes = {
        1: "1월은 중국 설 연휴 전후로 해외여행 수요 변화를 참고하기 좋은 시기입니다.",
        2: "2월은 중국 설 연휴 이후 여행 수요 흐름을 함께 살펴볼 수 있는 시기입니다.",
        5: "5월은 중국 노동절 연휴가 포함된 시기로 관광 관련 수요 변화를 참고할 수 있습니다.",
        6: "6월은 중국 주요 시험 종료 이후 여름방학 수요로 넘어가는 시기입니다.",
        7: "7월은 중국 여름방학 시즌이 시작되는 시기로 중국인 관광객 증가가 예상됩니다.",
        8: "중국 여름 휴가 시즌이 지속되는 기간으로 관광 관련 수요 변화를 모니터링할 예정입니다.",
        10: "10월은 중국 국경절 연휴가 포함된 시기로 방한 관광객 증가가 예상됩니다.",
    }
    return notes.get(month)


def find_pretendard_font(weight: str) -> Path | None:
    patterns = [
        f"Pretendard-{weight}.ttf",
        f"Pretendard-{weight}.otf",
        f"*Pretendard*{weight}*.ttf",
        f"*Pretendard*{weight}*.otf",
    ]
    for pattern in patterns:
        matches = sorted(FONT_DIR.glob(pattern))
        if matches:
            return matches[0]
    return None


def find_dongle_font(weight: str) -> Path | None:
    patterns = [
        f"Dongle-{weight}.ttf",
        f"Dongle-{weight}.otf",
        f"*Dongle*{weight}*.ttf",
        f"*Dongle*{weight}*.otf",
    ]
    for pattern in patterns:
        matches = sorted(FONT_DIR.glob(pattern))
        if matches:
            return matches[0]
    return None


def font_data_uri(path: Path) -> str:
    mime = "font/ttf" if path.suffix.lower() == ".ttf" else "font/otf"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def report_font_face_css() -> str:
    font_files = {
        "Dongle": {
            "Light": 300,
            "Regular": 400,
            "Bold": 700,
        },
        "Pretendard": {
            "Regular": 400,
            "Medium": 500,
            "SemiBold": 600,
            "Bold": 700,
            "ExtraBold": 800,
        },
    }
    rules = []
    for family, weights in font_files.items():
        for label, weight in weights.items():
            path = find_dongle_font(label) if family == "Dongle" else find_pretendard_font(label)
            if not path:
                continue
            font_format = "truetype" if path.suffix.lower() == ".ttf" else "opentype"
            rules.append(
                f"""
                @font-face {{
                    font-family: '{family}';
                    font-style: normal;
                    font-weight: {weight};
                    src: url('{font_data_uri(path)}') format('{font_format}');
                    font-display: swap;
                }}
                """
            )
    return "\n".join(rules)


def register_pdf_fonts() -> ReportFonts:
    pretendard_regular = find_pretendard_font("Regular")
    pretendard_medium = find_pretendard_font("Medium") or pretendard_regular
    pretendard_semi_bold = find_pretendard_font("SemiBold")
    pretendard_bold = find_pretendard_font("Bold") or pretendard_semi_bold or pretendard_medium
    pretendard_extra_bold = find_pretendard_font("ExtraBold") or pretendard_bold
    if pretendard_medium and pretendard_bold and pretendard_extra_bold:
        pdfmetrics.registerFont(TTFont("Pretendard-Medium", str(pretendard_medium)))
        pdfmetrics.registerFont(TTFont("Pretendard-Bold", str(pretendard_bold)))
        pdfmetrics.registerFont(TTFont("Pretendard-ExtraBold", str(pretendard_extra_bold)))
        pdfmetrics.registerFontFamily(
            "Pretendard",
            normal="Pretendard-Medium",
            bold="Pretendard-Bold",
            italic="Pretendard-Medium",
            boldItalic="Pretendard-Bold",
        )
        return {
            "medium": "Pretendard-Medium",
            "bold": "Pretendard-Bold",
            "extra_bold": "Pretendard-ExtraBold",
            "chart_path": str(pretendard_medium),
        }

    dongle_regular = find_dongle_font("Regular")
    dongle_bold = find_dongle_font("Bold") or dongle_regular
    if dongle_regular and dongle_bold:
        pdfmetrics.registerFont(TTFont("Dongle-Regular", str(dongle_regular)))
        pdfmetrics.registerFont(TTFont("Dongle-Bold", str(dongle_bold)))
        pdfmetrics.registerFontFamily(
            "Dongle",
            normal="Dongle-Regular",
            bold="Dongle-Bold",
            italic="Dongle-Regular",
            boldItalic="Dongle-Bold",
        )
        return {
            "medium": "Dongle-Regular",
            "bold": "Dongle-Bold",
            "extra_bold": "Dongle-Bold",
            "chart_path": str(dongle_regular),
        }

    if NOTO_FONT.exists():
        pdfmetrics.registerFont(TTFont("NotoSansKR", str(NOTO_FONT)))
        return {"medium": "NotoSansKR", "bold": "NotoSansKR", "extra_bold": "NotoSansKR", "chart_path": str(NOTO_FONT)}
    if NOTO_VAR_FONT.exists():
        pdfmetrics.registerFont(TTFont("NotoSansKR", str(NOTO_VAR_FONT)))
        return {"medium": "NotoSansKR", "bold": "NotoSansKR", "extra_bold": "NotoSansKR", "chart_path": str(NOTO_VAR_FONT)}
    pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
    return {"medium": "HYGothic-Medium", "bold": "HYGothic-Medium", "extra_bold": "HYGothic-Medium", "chart_path": None}


def describe_active_report_font() -> str:
    fonts = register_pdf_fonts()
    if fonts["medium"].startswith("Dongle"):
        return "Dongle 둥근 폰트 적용됨: 본문 Regular, 제목/KPI Bold"
    if fonts["medium"].startswith("Pretendard"):
        return "Pretendard 적용됨: 본문 Medium, 제목 Bold, KPI 숫자 ExtraBold"
    if fonts["medium"] == "NotoSansKR":
        return "Noto Sans KR 적용 중: assets/fonts에 Pretendard 파일이 없으면 이 폰트로 대체됩니다."
    return f"{fonts['medium']} 적용 중"


def find_korean_ttf_font() -> str | None:
    pretendard_medium = find_pretendard_font("Medium") or find_pretendard_font("Regular")
    if pretendard_medium:
        return str(pretendard_medium)
    dongle_regular = find_dongle_font("Regular")
    if dongle_regular:
        return str(dongle_regular)
    for path in [
        NOTO_FONT,
        NOTO_VAR_FONT,
        Path("/System/Library/Fonts/Supplemental/AppleGothic.ttf"),
        Path("/Library/Fonts/AppleGothic.ttf"),
    ]:
        if path.exists():
            return str(path)
    return None


def make_styles(fonts: ReportFonts) -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("TitleKR", parent=base["Title"], fontName=fonts["extra_bold"], fontSize=20, leading=25, textColor=TEXT, alignment=TA_LEFT),
        "subtitle": ParagraphStyle("SubtitleKR", parent=base["Normal"], fontName=fonts["medium"], fontSize=10.5, leading=15, textColor=TEXT, alignment=TA_LEFT),
        "h2": ParagraphStyle("H2KR", parent=base["Heading2"], fontName=fonts["bold"], fontSize=10.5, leading=13, textColor=TEXT, spaceAfter=0),
        "body": ParagraphStyle("BodyKR", parent=base["BodyText"], fontName=fonts["medium"], fontSize=8.5, leading=12.5, textColor=TEXT, alignment=TA_LEFT),
        "small": ParagraphStyle("SmallKR", parent=base["BodyText"], fontName=fonts["medium"], fontSize=6.5, leading=8.2, textColor=MUTED, alignment=TA_CENTER),
        "card_label": ParagraphStyle("CardLabel", parent=base["BodyText"], fontName=fonts["medium"], fontSize=7.2, leading=9, textColor=MUTED, alignment=TA_CENTER),
        "card_value": ParagraphStyle("CardValue", parent=base["BodyText"], fontName=fonts["extra_bold"], fontSize=10.8, leading=13, textColor=TEXT, alignment=TA_CENTER),
        "chart_label": ParagraphStyle("ChartLabel", parent=base["BodyText"], fontName=fonts["bold"], fontSize=7.8, leading=8.8, textColor=TEXT, alignment=TA_CENTER),
        "notice": ParagraphStyle("NoticeKR", parent=base["BodyText"], fontName=fonts["medium"], fontSize=7.2, leading=9, textColor=MUTED, alignment=TA_LEFT),
    }


def draw_page_background(canvas, doc, inputs: StoreInputs, fonts: ReportFonts) -> None:
    canvas.saveState()
    width, height = A4
    canvas.setFillColor(LIGHT_BG)
    canvas.rect(0, 0, width, height, fill=1, stroke=0)
    if doc.page == 1:
        header_h = 34 * mm
        canvas.setFillColor(ACCENT)
        canvas.rect(0, height - header_h, width, header_h, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor("#111111"))
        canvas.roundRect(width - 46 * mm, height - 24 * mm, 34 * mm, 8 * mm, 3 * mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(fonts["bold"], 6.5)
        canvas.drawCentredString(width - 29 * mm, height - 21.3 * mm, "CPC Report")
        canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.22))
        canvas.circle(width - 23 * mm, height - 14 * mm, 18 * mm, fill=1, stroke=0)
    canvas.restoreState()


def paragraph(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(str(text).replace("\n", "<br/>"), style)


def section_heading(title: str, styles: dict[str, ParagraphStyle]) -> Table:
    title_row = Table([[paragraph(title, styles["h2"])]], colWidths=[176 * mm], hAlign="LEFT")
    title_row.setStyle(
        TableStyle(
            [
                ("LINEBEFORE", (0, 0), (0, 0), 3, ACCENT),
                ("LEFTPADDING", (0, 0), (0, 0), 6),
                ("BOTTOMPADDING", (0, 0), (0, 0), 2.5),
            ]
        )
    )
    line = Table([[""]], colWidths=[176 * mm], rowHeights=[1])
    line.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LINE), ("FONTNAME", (0, 0), (-1, -1), styles["body"].fontName)]))
    return Table([[title_row], [line]], colWidths=[176 * mm], hAlign="LEFT", spaceBefore=3, spaceAfter=2.2)


def metric_cards(items: list[tuple[str, str]], styles: dict[str, ParagraphStyle]) -> Table:
    card_width = (176 * mm - 12 * mm) / 5
    row = []
    for idx, (label, value) in enumerate(items):
        if idx:
            row.append("")
        card = Table(
            [[paragraph(label, styles["card_label"])], [paragraph(value, styles["card_value"])]],
            colWidths=[card_width],
        )
        card.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                    ("BOX", (0, 0), (-1, -1), 0.35, LINE),
                    ("LINEABOVE", (0, 0), (-1, 0), 2, ACCENT),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 5.5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5.5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        row.append(card)
    table = Table([row], colWidths=[card_width, 3 * mm, card_width, 3 * mm, card_width, 3 * mm, card_width, 3 * mm, card_width], hAlign="LEFT")
    table.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, -1), styles["body"].fontName), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    return table


def simple_table(data: list[list[object]], fonts: ReportFonts, widths: list[float], font_size: float = 7.6, row_padding: float = 3.2) -> Table:
    table = Table(data, colWidths=widths, hAlign="LEFT", repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), fonts["medium"]),
                ("FONTNAME", (0, 0), (-1, 0), fonts["bold"]),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("LEADING", (0, 0), (-1, -1), font_size + 2),
                ("BACKGROUND", (0, 0), (-1, 0), ACCENT_SOFT),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#2B2B2B")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.22, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), row_padding),
                ("BOTTOMPADDING", (0, 0), (-1, -1), row_padding),
            ]
        )
    )
    return table


def comparison_info_table(current_df: pd.DataFrame, previous_df: pd.DataFrame, fonts: ReportFonts) -> Table:
    data = [
        ["비교 기간 :", f"{format_month_full(previous_df['날짜'].min())} ↔ {format_month_full(current_df['날짜'].min())}"],
        ["광고 운영일수 :", f"전월 {len(previous_df)}일 | 대상월 {len(current_df)}일"],
    ]
    table = Table(data, colWidths=[36 * mm, 140 * mm], hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), fonts["medium"]),
                ("FONTNAME", (0, 0), (0, -1), fonts["bold"]),
                ("FONTSIZE", (0, 0), (-1, -1), 7.3),
                ("LEADING", (0, 0), (-1, -1), 9.3),
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#555555")),
                ("BOX", (0, 0), (-1, -1), 0.25, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2.0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.0),
            ]
        )
    )
    return table


def comparison_table(rows: list[dict[str, object]], fonts: ReportFonts, include_change: bool) -> Table:
    data = [["항목", "대상월", "전월"] + (["전월 대비"] if include_change else [])]
    for row in rows:
        line = [row["label"], row["current_text"], row["previous_text"]]
        if include_change:
            line.append(row["change_text"])
        data.append(line)
    col_widths = [46 * mm, 62 * mm, 62 * mm] if not include_change else [42 * mm, 44 * mm, 44 * mm, 46 * mm]
    table = simple_table(data, fonts, col_widths, 7.1, 2.15)
    commands = []
    for idx, row in enumerate(rows, start=1):
        if include_change:
            current_value = float(row["current"])
            previous_value = float(row["previous"])
            color = MUTED if current_value == previous_value else colors.HexColor("#DC2626") if current_value > previous_value else colors.HexColor("#2563EB")
            commands.append(("TEXTCOLOR", (3, idx), (3, idx), color))
        commands.append(("ALIGN", (1, idx), (-1, idx), "RIGHT"))
    table.setStyle(TableStyle(commands))
    return table


def configure_matplotlib_font() -> None:
    font_path = find_korean_ttf_font()
    if font_path:
        from matplotlib import font_manager

        font_manager.fontManager.addfont(font_path)
        plt.rcParams["font.family"] = font_manager.FontProperties(fname=font_path).get_name()
    plt.rcParams["axes.unicode_minus"] = False


def chart_image(df: pd.DataFrame, column: str, color: str) -> io.BytesIO:
    configure_matplotlib_font()
    plot_df = df.copy()
    plot_df["표기"] = plot_df["날짜"].dt.strftime("%m.%d")
    plot_df["x"] = range(len(plot_df))
    tick_positions = list(plot_df["x"])
    tick_labels = [plot_df["표기"].iloc[idx] for idx in tick_positions]

    fig, ax = plt.subplots(figsize=(11.6, 3.75), dpi=220)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.plot(plot_df["x"], plot_df[column], color=color, linewidth=2.1, marker="o", markersize=2.8)
    ax.fill_between(plot_df["x"], plot_df[column], color=color, alpha=0.12)
    ax.grid(True, color="#E8E4D8", alpha=0.55, linewidth=0.55)
    for spine in ax.spines.values():
        spine.set_color("#D8D3C7")
        spine.set_linewidth(0.7)
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, fontsize=6.2, rotation=45, ha="right")
    ax.tick_params(axis="y", labelsize=8.5)
    ax.tick_params(colors="#333333")
    ax.margins(x=0.015)
    fig.subplots_adjust(left=0.07, right=0.99, top=0.96, bottom=0.27)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.06)
    plt.close(fig)
    buf.seek(0)
    return buf


def daily_chart_table(df: pd.DataFrame, styles: dict[str, ParagraphStyle]) -> Table:
    chart_width = 176 * mm
    chart_height = 57 * mm
    rows = []
    for idx, (title, column, color) in enumerate([("일별 노출수 추이", "노출수", "#111111"), ("일별 클릭수 추이", "클릭수", "#F2B705")]):
        if idx:
            rows.append([Spacer(1, 1 * mm)])
        rows.append([paragraph(title, styles["chart_label"])])
        rows.append([Spacer(1, 0.4 * mm)])
        rows.append([RLImage(chart_image(df, column, color), width=chart_width, height=chart_height)])
    table = Table(rows, colWidths=[chart_width], hAlign="LEFT")
    table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    return table


def build_pdf(df: pd.DataFrame, inputs: StoreInputs, mapping: dict[str, str], previous_df: pd.DataFrame | None = None) -> bytes:
    fonts = register_pdf_fonts()
    styles = make_styles(fonts)
    metrics = calculate_metrics(df)
    previous_metrics = calculate_metrics(previous_df) if previous_df is not None and not previous_df.empty else None
    current_days = len(df)
    previous_days = len(previous_df) if previous_df is not None and not previous_df.empty else 0
    has_previous = previous_metrics is not None
    comparison_excluded = has_previous and abs(current_days - previous_days) >= 6
    include_change = has_previous and not comparison_excluded
    comparison_rows = build_month_comparison(
        metrics,
        previous_metrics,
        current_days=current_days,
        previous_days=previous_days or 1,
        include_change=include_change,
    )
    show_day_gap_notice = comparison_excluded and inputs.store_type != "신규 매장"
    operation_comment = build_operation_comment(df, previous_df, comparison_excluded)
    report_month = int(df["날짜"].dt.month.mode().iloc[0]) if not df.empty else 0
    market_note = market_trend_note(report_month) if report_month else None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=13 * mm, leftMargin=13 * mm, topMargin=9 * mm, bottomMargin=9 * mm)
    story = []

    # Page 1: store, period, key metrics, comparison, charts.
    story.append(Spacer(1, 1 * mm))
    story.append(paragraph(inputs.store_name, styles["title"]))
    story.append(paragraph(f"{format_report_title_month(inputs.period_start)} CPC 운영 리포트", styles["subtitle"]))
    story.append(paragraph(f"운영기간: {format_period(inputs.period_start, inputs.period_end)}", styles["subtitle"]))
    story.append(Spacer(1, 4 * mm))

    story.append(section_heading("핵심 성과 지표", styles))
    story.append(
        metric_cards(
            [
                ("노출수", f"{format_number(metrics['total_impressions'])}회"),
                ("클릭수", f"{format_number(metrics['total_clicks'])}회"),
                ("광고비", format_currency(metrics["total_spend"])),
                ("클릭당 평균비용", format_cpc_currency(metrics["avg_cpc"])),
                ("상품권", f"{format_number(metrics['total_voucher'])}회"),
            ],
            styles,
        )
    )
    if has_previous:
        story.append(section_heading("전월 대비 성과 비교", styles))
        story.append(comparison_info_table(df, previous_df, fonts))
        story.append(Spacer(1, 1 * mm))
        story.append(comparison_table(comparison_rows, fonts, include_change))
        if show_day_gap_notice:
            story.append(Spacer(1, 0.5 * mm))
            story.append(paragraph("※ 광고 운영일수 차이로 인해 증감률 표시는 제외되었습니다.", styles["notice"]))

    story.append(section_heading("일별 성과 추이", styles))
    if not df.empty:
        story.append(daily_chart_table(df, styles))
    else:
        story.append(paragraph("그래프를 생성할 수 있는 일별 데이터가 충분하지 않습니다.", styles["body"]))

    story.append(PageBreak())

    # Page 2: daily table and owner-facing monthly summary.
    story.append(section_heading("일별 상세 데이터", styles))
    show_voucher_column = bool(df["상품권 조회수"].sum() > 0) if "상품권 조회수" in df.columns else False
    daily_headers = ["날짜", "광고비", "노출수", "클릭수", "클릭당 평균비용"]
    if show_voucher_column:
        daily_headers.append("상품권")
    daily_rows = [daily_headers]
    for _, row in df.head(31).iterrows():
        clicks = float(row["클릭수"])
        spend = float(row[SPEND_COL])
        daily_cpc = spend / clicks if clicks > 0 else None
        daily_line = [
            format_short_date(row["날짜"]),
            format_currency(spend),
            format_number(row["노출수"]),
            format_number(clicks),
            format_daily_cpc(daily_cpc),
        ]
        if show_voucher_column:
            daily_line.append(f"{format_number(row['상품권 조회수'])}건")
        daily_rows.append(daily_line)
    daily_widths = [24 * mm, 34 * mm, 31 * mm, 28 * mm, 41 * mm, 18 * mm] if show_voucher_column else [27 * mm, 37 * mm, 34 * mm, 31 * mm, 47 * mm]
    story.append(simple_table(daily_rows, fonts, daily_widths, 7.4, 2.55))
    if len(df) > 31:
        story.append(paragraph("일별 데이터가 많아 PDF 표에는 앞의 31개 행을 표시했습니다.", styles["body"]))

    if operation_comment:
        story.append(section_heading("운영요약", styles))
        story.append(paragraph(operation_comment, styles["body"]))

    if market_note:
        story.append(section_heading("시장동향 및 참고사항", styles))
        story.append(paragraph(f"{market_note}<br/>해당 내용은 외부 시장 일정에 따른 참고사항이며, 매장별 광고 성과 평가에는 반영하지 않습니다.", styles["body"]))

    def on_page(canvas, doc_obj):
        draw_page_background(canvas, doc_obj, inputs, fonts)
        canvas.saveState()
        canvas.setFont(fonts["medium"], 8)
        canvas.setFillColor(MUTED)
        canvas.drawRightString(A4[0] - 16 * mm, 8 * mm, f"{doc_obj.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buffer.getvalue()


def make_file_name(inputs: StoreInputs) -> str:
    return f"{inputs.store_name}_{inputs.period_start.year}년{inputs.period_start.month:02d}월_CPC운영보고서.pdf"


def make_image_file_name(inputs: StoreInputs) -> str:
    return f"{inputs.store_name}_{inputs.period_start.year}년{inputs.period_start.month:02d}월_CPC운영보고서.png"


def pdf_to_png_bytes(pdf_bytes: bytes, zoom: float = 2.0) -> bytes:
    if fitz is None:
        raise RuntimeError("이미지 생성을 위해 PyMuPDF 패키지가 필요합니다. requirements.txt 설치 후 다시 실행해주세요.")
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    images: list[PILImage.Image] = []
    matrix = fitz.Matrix(zoom, zoom)
    for page in document:
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        images.append(PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples))
    gap = int(20 * zoom)
    width = max(image.width for image in images)
    height = sum(image.height for image in images) + gap * (len(images) - 1)
    combined = PILImage.new("RGB", (width, height), "white")
    y = 0
    for image in images:
        combined.paste(image, ((width - image.width) // 2, y))
        y += image.height + gap
    output = io.BytesIO()
    combined.save(output, format="PNG", optimize=True)
    return output.getvalue()


def preview_metrics(df: pd.DataFrame) -> None:
    metrics = calculate_metrics(df)
    items = [
        ("노출수", f"{format_number(metrics['total_impressions'])}회"),
        ("클릭수", f"{format_number(metrics['total_clicks'])}회"),
        ("광고비", format_currency(metrics["total_spend"])),
        ("클릭당 평균비용(CPC)", format_cpc_currency(metrics["avg_cpc"])),
        ("상품권 조회수", f"{format_number(metrics['total_voucher'])}회"),
    ]
    cols = st.columns(len(items))
    for col, (label, value) in zip(cols, items):
        col.metric(label, value)


def render_cpc_report_page() -> None:
    render_page_header(
        APP_TITLE,
        "중국어 CPC 데이터를 업로드하면 점주 전달용 2페이지 PDF/PNG 리포트를 생성합니다.",
        "업무 자동화",
    )
    left, right = st.columns([1, 1.1], gap="large")
    with left:
        with st.container(border=True):
            uploaded_file = st.file_uploader("CPC 데이터 파일 업로드", type=["xlsx", "xls", "csv"])
            store_name = st.text_input("매장명", placeholder="예: 홍대점")
            store_type = "선택 안 함"
            requires_store_type = False
            if uploaded_file:
                try:
                    preview_raw_df = read_best_sheet(uploaded_file)
                    preview_all_df, _, _ = preprocess_dataframe(preview_raw_df)
                    preview_df, preview_previous_df, _, _, _ = split_latest_months(preview_all_df)
                    requires_store_type = preview_previous_df is not None and abs(len(preview_df) - len(preview_previous_df)) >= 6
                except Exception:
                    requires_store_type = False

            if requires_store_type:
                st.info("전월과 대상월의 광고 운영일수 차이가 큽니다. 매장 유형은 필요할 때만 선택해주세요.")
                store_type = st.radio("매장 유형 선택", ["선택 안 함", "기존 매장", "신규 매장"], horizontal=True)
            use_pause = st.checkbox("광고 중단일 입력")
            pause_date = st.date_input("광고 중단일", value=date.today()) if use_pause else None
            memo = st.text_area("비고", placeholder="선택 입력")
            generate = st.button("보고서 생성", type="primary", use_container_width=True)

    with right:
        with st.container(border=True):
            st.subheader("생성 안내")
            st.markdown(
                """
                <div class="hint">
                여러 월이 포함된 CPC 파일 1개를 업로드하면 가장 최근 월을 대상월 데이터로, 직전 월을 전월 비교 데이터로 자동 인식합니다. 전월 데이터는 비교표에만 사용되며 별도 전월 상세표나 페이지는 생성하지 않습니다.
                </div>
                """,
                unsafe_allow_html=True,
            )

    if not generate:
        return
    if not uploaded_file:
        st.error("엑셀 또는 CSV 파일을 먼저 업로드해주세요.")
        return
    if not store_name.strip():
        st.error("매장명을 입력해주세요.")
        return
    try:
        with st.spinner("데이터를 정리하고 2페이지 보고서를 만드는 중입니다..."):
            raw_df = read_best_sheet(uploaded_file)
            all_df, mapping, messages = preprocess_dataframe(raw_df)
            if all_df.empty:
                st.warning("날짜가 포함된 데이터 행을 찾지 못했습니다. 업로드 파일의 날짜 컬럼을 확인해주세요.")
                return
            df, previous_df, period_start, period_end, month_messages = split_latest_months(all_df)
            inputs = StoreInputs(store_name.strip(), store_type, period_start, period_end, pause_date, memo)
            pdf_bytes = build_pdf(df, inputs, mapping, previous_df)
            png_bytes = pdf_to_png_bytes(pdf_bytes)

        st.success("PDF와 이미지 보고서가 생성되었습니다.")
        st.caption(f"현재 보고서 폰트: {describe_active_report_font()}")
        preview_metrics(df)
        if messages or month_messages:
            with st.expander("자동 처리 안내"):
                for message in messages + month_messages:
                    st.info(message)
        cols = st.columns(2)
        cols[0].download_button("PDF 다운로드", data=pdf_bytes, file_name=make_file_name(inputs), mime="application/pdf", use_container_width=True)
        cols[1].download_button("이미지(PNG) 다운로드", data=png_bytes, file_name=make_image_file_name(inputs), mime="image/png", use_container_width=True)
        with st.expander("일별 데이터 미리보기"):
            preview = df.copy()
            preview["날짜"] = preview["날짜"].dt.strftime("%m.%d")
            preview_columns = ["날짜", SPEND_COL, "노출수", "클릭수", "상품권 조회수", "저장하기", "공유하기"]
            st.dataframe(preview[[col for col in preview_columns if col in preview.columns]], use_container_width=True)
    except Exception as exc:
        st.error("보고서를 생성하는 중 문제가 발생했습니다.")
        st.info(f"확인할 내용: 파일 형식, 날짜 컬럼, 숫자 컬럼 값을 확인해주세요. 오류 내용: {exc}")


def render_coming_soon_page(title: str, description: str, group: str) -> None:
    render_page_header(title, description, group)
    st.markdown(
        """
        <div class="coming-soon">
            <h2>준비 중</h2>
            <p>이 메뉴는 내부 업무 자동화 사이트 구조에 먼저 추가해두었습니다. 이후 입력 폼, 문서 템플릿, 승인 흐름 등을 이 화면에 연결할 수 있습니다.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_manual_page(item: NavItem) -> None:
    if item.key == "manual_search":
        render_page_header(item.label, "업무 매뉴얼 콘텐츠를 준비 중입니다.", item.group)
        search_term = st.text_input("매뉴얼 검색", placeholder="검색어를 입력하세요")
        if search_term:
            st.info("아직 등록된 매뉴얼 본문이 없어 검색 결과가 없습니다.")
        else:
            st.caption("매뉴얼 문서가 등록되면 제목과 본문 기준으로 검색 결과가 표시됩니다.")
        return
    render_coming_soon_page(item.label, "업무 매뉴얼 콘텐츠를 준비 중입니다.", item.group)


def render_page(selected_key: str) -> None:
    item = NAV_BY_KEY.get(selected_key, NAV_BY_KEY[DEFAULT_NAV_KEY])
    if item.key == "cpc_report":
        render_cpc_report_page()
        return
    if item.key == "contract_generator":
        render_coming_soon_page(item.label, item.description, item.group)
        return
    render_manual_page(item)


def main() -> None:
    setup_page()
    selected_key = render_sidebar()
    render_page(selected_key)


if __name__ == "__main__":
    main()
