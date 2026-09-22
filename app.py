# -*- coding: utf-8 -*-
"""
TTA 인증서 발급용 데이터 추출기 (Streamlit)

- 시험결과요약서(PDF, 여러 개) + (선택) 회의록(hwp)을 업로드하면
- 표지 정보를 추출해서 인증서 제작에 필요한 필드를 표로 뽑아주고
- 표는 화면에서 직접 수정 가능하며, 엑셀로 다운로드할 수 있습니다.

"""
import io
import os
import re
import subprocess
import tempfile
from datetime import date, timedelta

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="TTA 인증서 데이터 추출기", layout="wide")
st.title("📄 TTA 인증서 발급용 데이터 추출기")

st.write("") 

st.markdown(
    """
    1. **업체명 영문명 리스트 엑셀**을 업로드 (업체명_국문 / 업체명_영문 두 컬럼짜리 엑셀)
    2. **시험결과요약서 PDF**를 업로드 (여러 개 가능, 표지에서 자동으로 필드를 추출)
    3. 인터넷전화(MMoIP) 인증 건은 **인증심의위원회 회의록(.hwp)** 도 업로드 (인증범위 및 제조국가를 회의록에서 가져옴)
    4. 인증 정보 추출 결과 표를 화면에서 확인 및 수정한 뒤 엑셀로 다운로드
    """
)

st.write("")

# ──────────────────────────────────────────────────────────────
# 업로드
# ──────────────────────────────────────────────────────────────
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

top_left, top_right = st.columns([5, 1])
with top_left:
    st.subheader("📎 파일 업로드")
with top_right:
    st.write("")  # 세로 정렬용 여백
    if st.button("🔄 업로드 초기화", use_container_width=True):
        st.session_state.uploader_key += 1
        st.rerun()

col1, col2, col3 = st.columns(3)
with col1:
    company_map_file = st.file_uploader(
        "① 업체명 영문명 리스트 (.xlsx) :red[*]",
        type=["xlsx"],
        key=f"companymap_{st.session_state.uploader_key}",
    )
with col2:
    summary_files = st.file_uploader(
        "② 시험결과요약서 (.PDF, 여러 개) :red[*]",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"summary_{st.session_state.uploader_key}",
    )
with col3:
    minutes_file = st.file_uploader(
        "③ 회의록 (.hwp, 선택)",
        type=["hwp"],
        key=f"minutes_{st.session_state.uploader_key}",
    )

st.caption(":red[*] 표시된 항목은 필수 업로드입니다. (회의록은 MMoIP 인증범위 보완용 선택 항목입니다)")


# ──────────────────────────────────────────────────────────────
# 시험결과요약서 표지 필드 추출
# ──────────────────────────────────────────────────────────────
LABEL_PATTERNS = {
    "업체명": [r"업체명"],
    "제조자": [r"제조자", r"제조사"],
    "제조국가": [r"제조국가"],
    "제품명": [r"제품명"],
    "모델명": [r"모델명"],
    "인증연월일": [r"인증\s*연월일"],
    "인증번호": [r"인증\s*번호"],
    "인증기준": [r"인증\s*기준\s*번호", r"인증기준\s*번호"],
    "인증범위": [r"인증\s*범위"],
}

# 라벨 뒤에 값이 바로 붙어 나오는 한 줄짜리 패턴들
# 예) "인증 번호 TTA-IT-C-26-007" / "제조국가: 대한민국" / "제조자/제조국가: 에이엠(주)/대한민국"
def _search_value(text, label_variants, combined_with=None):
    for lab in label_variants:
        # "라벨1/라벨2: 값1/값2" 형태 우선 시도
        if combined_with:
            pat = rf"{lab}\s*(?:및|/)\s*{combined_with}\s*[:\-]?\s*([^\n]+)"
            m = re.search(pat, text)
            if m:
                return m.group(1).strip(), True  # combined=True
        pat = rf"{lab}\s*[:\-]?\s*([^\n]+)"
        m = re.search(pat, text)
        if m:
            return m.group(1).strip(), False
    return None, False


def extract_test_highlights_from_page(page, x_split=195, y_tol=3.0):
    """표지가 2단 레이아웃(좌: 목차/ABOUT TTA, 우: Test Highlights 박스)인 문서에서
    x좌표 기준으로 우측 칼럼 단어만 모아 줄 단위로 재구성한 뒤,
    'Test Highlights' 다음부터 '개요' 헤딩 전까지를 하이라이트로 추출한다."""
    words = [w for w in page.extract_words() if w["x0"] >= x_split]
    if not words:
        return ""
    words.sort(key=lambda w: (round(w["top"] / y_tol), w["x0"]))

    lines = []
    cur_top, cur_words = None, []
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
            cur_words.append(w)
            cur_top = w["top"] if cur_top is None else cur_top
        else:
            lines.append((cur_top, cur_words))
            cur_words = [w]
            cur_top = w["top"]
    if cur_words:
        lines.append((cur_top, cur_words))

    GAP_LIMIT = 25  # 이 값보다 줄 간격이 크게 벌어지면 하이라이트 박스를 벗어난 것으로 간주
    started = False
    out_lines = []
    prev_top = None
    for top, lw in lines:
        lw_sorted = sorted(lw, key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in lw_sorted).strip()
        compact = re.sub(r"\s+", "", text)
        if not started:
            if "TestHighlights" in compact.replace(":", ""):
                started = True
                prev_top = top
                # 제목과 같은 줄에 내용이 더 있는 경우가 있는데, 그게 실제 불릿 항목일 때만 포함
                # (제목 뒤에 모델명만 덧붙는 경우 — 예: "Test Highlights: CUVIA" — 는 제외)
                m = re.search(r"Highlights\s*:?\s*(.*)", text)
                if m:
                    trailing = m.group(1).strip()
                    if trailing and re.match(r"^[•ŸlL▪∎⚫\-]", trailing):
                        out_lines.append(trailing)
            continue
        # '개요'/'시험환경' 등 다음 섹션 헤딩이 다른 단어와 같은 줄로 묶인 경우까지 감지
        if re.match(r"^(개\s*요|시험\s*목적|시험\s*환경|시험\s*방법|시험\s*절차|시험\s*결과)\b", text):
            break
        # 이미 한 줄 이상 담은 뒤, 줄 간격이 갑자기 크게 벌어지면(=하이라이트 박스를 벗어나
        # 옆 칸의 다른 문단으로 넘어간 것) 그 줄부터는 버린다. 박스 제목→첫 줄 사이의
        # 첫 간격은 원래도 넓은 경우가 많아 이 검사에서 제외한다.
        if out_lines and (top - prev_top) > GAP_LIMIT:
            break
        out_lines.append(text)
        prev_top = top

    return "\n".join(l for l in out_lines if l)


def _find_matching_paren(s, start_idx):
    depth = 0
    for i in range(start_idx, len(s)):
        if s[i] == "(":
            depth += 1
        elif s[i] == ")":
            depth -= 1
            if depth == 0:
                return i
    return -1


def extract_product_model_from_title(text, company_hint=None):
    """표지 상단 제목 영역(예: '유한회사 아라콤, 5.8 GHz DSRC OBE(모델명:AHRAPASS)')에서
    제품명/모델명/파생모델명을 뽑아낸다. 모델명 쪽 표기가 문서마다 달라
    (모델명:X) 형태와 '제품명, 모델명' 형태를 모두 처리하고, 모델명 자체에 괄호가
    포함된 경우('NHN-1041(N)SS')도 괄호 깊이를 세어 안전하게 잘라낸다."""
    m = re.search(r"TS00[^\n]*\n(.*?)Summary", text, re.DOTALL)
    if not m:
        return None, None, None
    raw = re.sub(r"\bTest\b", " ", m.group(1))

    model, derived = None, None
    idx = raw.find("(모델명")
    if idx == -1:
        idx = raw.find("(모델")
    if idx != -1:
        end = _find_matching_paren(raw, idx)
        if end != -1:
            inner = raw[idx + 1 : end].strip()
            mm = re.match(r"모델\s*명?\s*[:：]?\s*(.+?)(?:,\s*파생\s*모델(?:명)?\s*[:：]?\s*(.+))?$", inner)
            if mm:
                model = mm.group(1).strip()
                derived = mm.group(2).strip() if mm.group(2) else None
            raw = raw[:idx] + raw[end + 1 :]

    product = re.sub(r"\s+", " ", raw).strip(" ,")
    if company_hint and product.startswith(company_hint.strip()):
        product = product[len(company_hint.strip()) :].strip(" ,")

    if model is None and "," in product:
        # 라벨 없이 '제품명, 모델명' 형태로만 구분된 경우 (예: '인터넷전화 서버, IPEX100-ECPS')
        left, _, right = product.rpartition(",")
        if right.strip():
            product, model = left.strip(), right.strip()

    return product or None, model, derived


@st.cache_data(show_spinner=False)
def extract_summary_fields(file_bytes, filename):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    with pdfplumber.open(tmp_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        highlights = extract_test_highlights_from_page(page)
        # 제조자/제조국가 등은 표지가 아니라 2페이지 'IUT 세부사양' 표에만 있는 경우가 있어
        # (예: 스마트시티 통합플랫폼/데이터허브 양식) 라벨 검색 대상 텍스트에 2페이지도 포함시킨다.
        search_text = text
        if len(pdf.pages) > 1:
            search_text += "\n" + (pdf.pages[1].extract_text() or "")

    result = {"__파일명": filename}

    # 제조자/제조국가가 한 줄로 붙어 있는 경우 우선 탐지
    combined_val, is_combined = _search_value(search_text, LABEL_PATTERNS["제조자"], combined_with="제조국가")
    manu, country = None, None
    if is_combined and combined_val:
        # "에이엠(주)/대한민국" 또는 "에이엠(주) / 대한민국" 둘 다 처리
        parts = re.split(r"\s*/\s*", combined_val)
        if len(parts) >= 2:
            manu, country = parts[0].strip(), parts[1].strip()

    if manu is None:
        manu, _ = _search_value(search_text, LABEL_PATTERNS["제조자"])
    if country is None:
        country, _ = _search_value(search_text, LABEL_PATTERNS["제조국가"])

    company, _ = _search_value(text, LABEL_PATTERNS["업체명"])
    if manu is None:
        manu = company  # 제조자 표기가 없으면 업체명으로 대체

    cert_no, _ = _search_value(text, LABEL_PATTERNS["인증번호"])
    cert_date, _ = _search_value(text, LABEL_PATTERNS["인증연월일"])
    criteria, _ = _search_value(text, LABEL_PATTERNS["인증기준"])
    scope, _ = _search_value(search_text, LABEL_PATTERNS["인증범위"])
    product_name, model_name, derived_model = extract_product_model_from_title(text, company_hint=company)

    # 시험번호 (문서 상단 "No. TTA-26-XXXXX-TS00")
    m = re.search(r"No\.\s*(TTA-\d{2}-\d+)", text)
    test_no = m.group(1) if m else None

    result.update(
        {
            "업체명(추출)": company,
            "제조자(추출)": manu,
            "제조국가(추출)": country,
            "인증번호": cert_no,
            "시험번호": test_no,
            "인증연월일(추출)": cert_date,
            "인증기준(추출)": criteria,
            "인증범위(추출, 있는경우)": scope,
            "제품명(추출)": product_name,
            "모델명(추출)": model_name,
            "파생모델명(추출)": derived_model,
            "Test_Highlights": highlights,
        }
    )
    return result


# ──────────────────────────────────────────────────────────────
# 회의록(.hwp) 표 파싱 → 인증범위 매칭용
# ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def extract_minutes_table(file_bytes):
    """hwp5html로 변환 후 표를 대략적으로 파싱. 실패하면 빈 DataFrame 반환."""
    try:
        with tempfile.NamedTemporaryFile(suffix=".hwp", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        out_dir = tempfile.mkdtemp()
        subprocess.run(
            ["hwp5html", "--output=" + out_dir, tmp_path],
            check=True, capture_output=True, timeout=60,
        )
        with open(out_dir + "/index.xhtml", encoding="utf-8") as f:
            html = f.read()

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "lxml")
        table = soup.find("table")
        if table is None:
            return pd.DataFrame()
        rows = []
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cells:
                rows.append(cells)
        if len(rows) < 2:
            return pd.DataFrame()
        header, body = rows[0], rows[1:]
        # 컬럼 수가 안 맞으면 잘라서 맞춤
        ncol = len(header)
        body = [r[:ncol] + [""] * (ncol - len(r)) for r in body]
        return pd.DataFrame(body, columns=header)
    except Exception as e:
        st.warning(f"회의록(.hwp) 파싱 실패 — 수동으로 입력해 주세요. ({e})")
        return pd.DataFrame()


COUNTRY_EN_MAP = {
    # 정식 영문 국가명 (Official Name) 기준
    "대한민국": "Republic of Korea",
    "일본": "Japan",
    "중국": "China",
    "대만": "Taiwan",
    "홍콩": "Hong Kong",
    "미국": "United States of America",
    "캐나다": "Canada",
    "멕시코": "Mexico",
    "브라질": "Brazil",
    "아르헨티나": "Argentina",
    "독일": "Germany",
    "영국": "United Kingdom",
    "프랑스": "France",
    "이탈리아": "Italy",
    "스페인": "Kingdom of Spain",
    "포르투갈": "Portugal",
    "네덜란드": "The Netherlands",
    "벨기에": "Belgium",
    "스웨덴": "Sweden",
    "노르웨이": "Norway",
    "덴마크": "Denmark",
    "핀란드": "Finland",
    "스위스": "Switzerland",
    "오스트리아": "Austria",
    "그리스": "Greece",
    "폴란드": "Republic of Poland",
    "체코": "Czechina",
    "헝가리": "Hungary",
    "아일랜드": "Ireland",
    "러시아": "Russian Federation",
    "호주": "Australia",
    "뉴질랜드": "New Zealand",
    "인도": "India",
    "싱가포르": "Singapore",
    "말레이시아": "Malaysia",
    "태국": "Kingdom of Thailand",
    "베트남": "Socialist Republic of Vietnam",
    "인도네시아": "Republic of Indonesia",
    "필리핀": "Republic of the Philippines",
    "이스라엘": "State of Israel",
    "터키": "Türkiye",
    "튀르키예": "Republic of Türkiye",
    "사우디아라비아": "Kingdom of Saudi Arabia",
    "아랍에미리트": "United Arab Emirates",
    "남아프리카공화국": "Republic of South Africa",
    "이집트": "Egypt",
}


def to_country_en(country_kr):
    if not country_kr:
        return ""
    country_kr = str(country_kr).strip()
    return COUNTRY_EN_MAP.get(country_kr, country_kr)


def calc_expiry(date_str):
    """'2026. 8. 21.' 형태 문자열 → (시작일 문자열, 만료일 문자열)"""
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", date_str or "")
    if not m:
        return date_str, ""
    y, mo, d = map(int, m.groups())
    start = date(y, mo, d)
    end = start.replace(year=start.year + 5) - timedelta(days=1)
    fmt = lambda dt: f"{dt.year}. {dt.month}. {dt.day}."
    return fmt(start), fmt(end)


# ──────────────────────────────────────────────────────────────
# 인증범위: 요약서/회의록에 문구가 없어도, 해당 인증서 "양식"이 정해져 있으면
# 그 양식에서 항상 쓰는 고정 문구를 기본값으로 채운다. (인증기준 코드로 양식 판별)
# ──────────────────────────────────────────────────────────────
SCOPE_FIXED_BY_PREFIX = {
    "TCN-0025": "5.8GHz DSRC 표준적합성 TTA Certified 인증",              # DSRC
    "TCC-0007": "스마트시티 데이터허브\n(인터페이스, 데이터 모델, 기능 적합성)",  # 데이터허브
    "TCC-0005": "스마트시티 통합플랫폼\n(기본기능, 상호연동기능, 통합기능)",     # 통합플랫폼
    "TCC-0009": "지능형 홈네트워크 세대단말기 TTA Verified 인증\n일반, 기능, 규격, 신뢰성 및 상호운용성",  # 홈네트워크 세대단말기
}
MMOIP_TITLE_BY_PREFIX = {
    "TCN-0048": "행정기관 인터넷전화 서버 보안 성능품질 Ver.4",   # 인터넷전화 서버
    "TCN-0049": "행정기관 인터넷전화 단말 보안 성능품질 Ver.4",   # 인터넷전화 단말
}


def criteria_prefix(criteria):
    if not criteria:
        return None
    m = re.match(r"\s*(T[A-Z]{2}-\d+)", str(criteria))
    return m.group(1) if m else None


def resolve_scope(criteria, extracted_scope, minutes_scope):
    """
    양식별로 인증범위를 채운다.
    - DSRC/데이터허브/통합플랫폼/홈네트워크: 항상 고정 문구 (제품마다 안 바뀜)
    - IPv6: 요약서의 기기분류(예: IPv6 Router core) + 고정 뒷문구
    - MMoIP(인터넷전화 서버/단말): 고정 제목 + 회의록(우선)/요약서에서 뽑은 상세 문구
    - 자급단말기 등 그 외: 요약서에서 뽑은 값을 그대로 (제품마다 다름)
    """
    prefix = criteria_prefix(criteria)

    if prefix in SCOPE_FIXED_BY_PREFIX:
        return SCOPE_FIXED_BY_PREFIX[prefix]

    if prefix == "TCN-0024":  # IPv6
        base = (extracted_scope or "IPv6 Router core").strip()
        return f"{base}\n표준적합성 및 상호운용성"

    if prefix in MMOIP_TITLE_BY_PREFIX:
        title = MMOIP_TITLE_BY_PREFIX[prefix]
        detail = (minutes_scope or extracted_scope or "").strip()
        return f"{title}\n({detail})" if detail else title

    # 그 외(자급단말기 등)는 제품마다 실제 시험범위가 달라서 고정 문구가 없음 → 추출값 사용
    return extracted_scope or minutes_scope or ""


def _display_width(s):
    """한글/한자 등 CJK 문자는 폭이 라틴 문자의 약 2배이므로 가중치를 둬서 길이를 계산."""
    w = 0
    for ch in s:
        code = ord(ch)
        if (
            0x1100 <= code <= 0x11FF   # 한글 자모
            or 0x3130 <= code <= 0x318F  # 한글 호환 자모
            or 0xAC00 <= code <= 0xD7A3  # 한글 음절
            or 0x4E00 <= code <= 0x9FFF  # 한자
            or 0x3000 <= code <= 0x303F  # CJK 기호
            or 0xFF00 <= code <= 0xFFEF  # 전각
        ):
            w += 2
        else:
            w += 1
    return w


def autosize_worksheet(ws, df, max_width=60, min_width=8, char_px=0.95, row_height=20, header_height=18):
    """열 너비는 각 컬럼에서 가장 긴 줄(개행 기준, 한글 가중치 반영) 길이에 맞추고,
    헤더 행은 header_height, 데이터 행은 row_height로 고정 높이를 설정한다."""
    from openpyxl.utils import get_column_letter

    for idx, col in enumerate(df.columns, start=1):
        col_letter = get_column_letter(idx)
        lines = [str(col)]
        for val in df[col].astype(str):
            lines.extend(val.split("\n"))
        longest = max((_display_width(l) for l in lines), default=10)
        width = min(max_width, max(min_width, longest * char_px + 2))
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = header_height
    for row_idx in range(2, len(df) + 2):  # 2행부터 (1행은 헤더)
        ws.row_dimensions[row_idx].height = row_height


# ──────────────────────────────────────────────────────────────
# 메인 처리
# ──────────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────────
# 업체명 영문명 매핑 저장소
# - ③번에서 매핑 엑셀(.xlsx)을 업로드하면 그걸 사용
# - 업로드 안 하면 앱에 기본으로 들어있는 company_english_names.xlsx를 사용
# - 컬럼: 업체명_국문 / 업체명_영문 (2개 컬럼짜리 단순한 표)
# ──────────────────────────────────────────────────────────────
COMPANY_MAP_PATH = os.path.join(os.path.dirname(__file__), "company_english_names.xlsx")
CORP_DESIGNATOR_RE = re.compile(r"(주식회사|유한회사|유한책임회사|㈜|\(주\)|\(유\))")


def normalize_corp_name(name):
    """'㈜아이디엠테크놀러지', '주식회사 엑스게이트', '에이엠(주)' 등을 비교 가능한 형태로 정규화."""
    if not name:
        return ""
    n = CORP_DESIGNATOR_RE.sub("", str(name))
    return re.sub(r"\s+", "", n).strip()


def _company_map_from_df(df):
    m = {}
    if "업체명_국문" in df.columns and "업체명_영문" in df.columns:
        for _, r in df.iterrows():
            kr, en = r.get("업체명_국문"), r.get("업체명_영문")
            if pd.notna(kr) and pd.notna(en) and str(en).strip():
                m[str(kr).strip()] = str(en).strip()
    return m


def load_company_map(uploaded_file=None):
    if uploaded_file is not None:
        try:
            return _company_map_from_df(pd.read_excel(uploaded_file))
        except Exception as e:
            st.warning(f"업로드한 매핑 엑셀을 읽지 못했어요. 컬럼명이 '업체명_국문' / '업체명_영문'인지 확인해 주세요. ({e})")
    if os.path.exists(COMPANY_MAP_PATH):
        try:
            return _company_map_from_df(pd.read_excel(COMPANY_MAP_PATH))
        except Exception:
            return {}
    return {}


def get_english_name_from_map(company_map, company_kr):
    key = normalize_corp_name(company_kr)
    for raw_name, eng in company_map.items():
        if normalize_corp_name(raw_name) == key:
            return eng
    return None


# ──────────────────────────────────────────────────────────────
# 메인 처리
# 시험결과요약서에서 추출한 값으로 채우고, 업체명 영문명은 업로드한 리스트(.xlsx)에
# 있는 것만 그대로 사용합니다 (리스트에 없으면 빈 칸).
# ──────────────────────────────────────────────────────────────
if summary_files and company_map_file:
    company_map = load_company_map(company_map_file)

    with st.expander("회의록 표 미리보기 (MMoIP 인증범위 참고용)"):
        minutes_df = pd.DataFrame()
        if minutes_file:
            minutes_df = extract_minutes_table(minutes_file.getvalue())
            st.dataframe(minutes_df, use_container_width=True)
        else:
            st.caption("회의록을 업로드하지 않았습니다.")

    rows = []
    for f in summary_files:
        extracted = extract_summary_fields(f.getvalue(), f.name)
        cert_no = extracted.get("인증번호")

        업체명_국문 = extracted.get("업체명(추출)")
        업체명_영문 = get_english_name_from_map(company_map, 업체명_국문) or ""
        제품명 = extracted.get("제품명(추출)") or ""
        모델명 = extracted.get("모델명(추출)") or ""
        파생모델명 = extracted.get("파생모델명(추출)") or ""
        if 파생모델명:
            모델명_전체 = f"{모델명}\n(파생모델명: {파생모델명})"
        else:
            모델명_전체 = 모델명
        인증기준 = extracted.get("인증기준(추출)")

        # 회의록에서 이 건에 해당하는 행 미리 찾아두기 (인증범위 + 제조국 둘 다 참고)
        minutes_scope, minutes_country = None, None
        if minutes_file:
            minutes_df = extract_minutes_table(minutes_file.getvalue())
            if not minutes_df.empty:
                proj_col = next((c for c in minutes_df.columns if "프로젝트" in c or "시험번호" in c), None)
                scope_col = next((c for c in minutes_df.columns if "인증범위" in c), None)
                country_col = next((c for c in minutes_df.columns if "제조국" in c), None)
                if proj_col:
                    m = minutes_df[minutes_df[proj_col].astype(str).str.contains(extracted.get("시험번호") or "!!!", na=False)]
                    if not m.empty:
                        if scope_col:
                            minutes_scope = m.iloc[0][scope_col]
                        if country_col:
                            minutes_country = m.iloc[0][country_col]

        manu = extracted.get("제조자(추출)") or 업체명_국문
        # 제조국: 요약서 → 회의록 → 그래도 없으면 대한민국 기본값
        country = extracted.get("제조국가(추출)") or minutes_country or "대한민국"
        제조자국가_국문 = f"{manu} / {country}".strip(" /")

        # 제조자 영문명: 요약서에 "업체명"과 다른 별도 제조자가 명시된 경우(OEM/ODM 등)에는
        # 관리목록의 업체명 영문명으로 덮어쓰면 안 되고, 추출된 제조자명을 그대로 사용해야 함.
        # (관리목록 업체명_영문은 "발주/의뢰사" 기준이라 실제 제조자와 다를 수 있음)
        if manu == 업체명_국문:
            manu_en = 업체명_영문 or manu
        else:
            manu_en = manu  # 요약서에 적힌 제조자명을 그대로 사용 (이미 영문 표기인 경우가 많음)
        country_en = to_country_en(country)
        제조자국가_영문 = f"{manu_en} / {country_en}".strip(" /")

        # 인증범위: 요약서 추출값 + 회의록 값을 양식(인증기준) 기준 고정 문구 로직(resolve_scope)에 넣어 최종 결정
        scope = resolve_scope(인증기준, extracted.get("인증범위(추출, 있는경우)"), minutes_scope)

        시작일, 만료일 = calc_expiry(extracted.get("인증연월일(추출)"))

        rows.append(
            {
                "인증번호": cert_no,
                "시험번호": extracted.get("시험번호"),
                "업체명_국문": 업체명_국문,
                "업체명_영문": 업체명_영문,
                "제품명": 제품명,
                "모델명": 모델명_전체,
                "제조자및제조국가_국문": 제조자국가_국문,
                "제조자및제조국가_영문": 제조자국가_영문,
                "인증연월일": extracted.get("인증연월일(추출)"),
                "유효기간_시작일": 시작일,
                "유효기간_만료일": 만료일,
                "인증기준": 인증기준,
                "인증범위": scope,
                "Test_Highlights": extracted.get("Test_Highlights"),
            }
        )

    result_df = pd.DataFrame(rows)
    result_df = result_df.sort_values("인증번호", na_position="last").reset_index(drop=True)
    result_df.insert(0, "순번", range(1, len(result_df) + 1))

    st.write("") 
    st.subheader("인증 정보 추출 결과 (직접 수정 가능)")
    edited_df = st.data_editor(
        result_df,
        use_container_width=True,
        num_rows="dynamic",
        height=500,
    )

    # 엑셀 다운로드 (헤더 색상 + 열 너비/행 높이 자동 맞춤 + 줄바꿈 서식)
    export_df = edited_df
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="인증서데이터")
        ws = writer.sheets["인증서데이터"]

        from openpyxl.styles import Alignment, Font, PatternFill

        wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="4472C4")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = wrap
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align

        autosize_worksheet(ws, export_df)
        ws.freeze_panes = "A2"

    st.download_button(
        "📥 엑셀로 다운로드",
        data=buf.getvalue(),
        file_name="인증서_발급용_데이터.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.write("")
    st.info("① 업체명 영문명 리스트와 ② 시험결과요약서를 모두 업로드하면 인증정보가 표시됩니다.")
