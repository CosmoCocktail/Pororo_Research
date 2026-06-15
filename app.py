import streamlit as st
import pandas as pd
import re
import os
import random
import base64
import gspread
from google.oauth2.service_account import Credentials

# ======================================================
# 페이지 설정
# ======================================================
st.set_page_config(
    page_title="팀플 빌런즈 : 눈 속 마을 빌런 테스트",
    page_icon="🐧",
    layout="centered"
)

# ======================================================
# 변경 가능 상수
# ======================================================
QUESTION_FILE   = "questions.xlsx"
CHARACTERS      = ["에디", "크롱", "뽀로로", "루피", "포비"]
SHEET_NAME      = "result"          # Google Sheets 파일명
BG_IMAGE_FILE   = "background.png"  # 배경 이미지 (없으면 무시)
OG_IMAGE_URL    = "https://raw.githubusercontent.com/your-username/your-repo/main/단체컷.png"
OG_URL          = "https://pororoteamprojecttest.streamlit.app"

# 폰트 설정
FONT_IMPORT_URL = "https://fonts.googleapis.com/css2?family=Gugi&family=Black+Han+Sans&display=swap"
FONT_BODY       = "'Gugi', sans-serif"        # 전체 기본 폰트
FONT_QUESTION   = "'Black Han Sans', sans-serif"  # 질문 텍스트 전용 폰트

# 컬러 설정
COLOR_PRIMARY   = "#1A7FD4"   # 메인 색상
COLOR_HOVER     = "#155FA0"   # 버튼 호버 색상

# ======================================================
# 동점 추가 질문
# ======================================================
TIEBREAK_QUESTION = {
    "질문":  "주제를 정하는데 의견이 갈린다면?",
    "에디":   '"잠깐, 지금 의견 말고 나 완전 좋은 생각 났어."',
    "크롱":   '"다 괜찮은 것 같은데…"',
    "뽀로로": '"제일 재밌어 보이는 거 하면 안 돼?"',
    "루피":   '"일단 기준부터 정하고 제일 괜찮은 안으로 가자."',
    "포비":   '"다들 말해봐. 내가 의견을 정리해볼게."',
}

# ======================================================
# 데이터 로드
# ======================================================
@st.cache_data
def load_questions():
    df = pd.read_excel(QUESTION_FILE)
    df = df.dropna(subset=["질문", "답변A", "답변B"])
    df = df[df["질문"].astype(str).str.strip() != ""]
    df = df[df["답변A"].astype(str).str.strip().str.lower() != "nan"]
    df = df[df["답변B"].astype(str).str.strip().str.lower() != "nan"]
    return df.reset_index(drop=True)

questions_df = load_questions()

# ======================================================
# 유틸 함수
# ======================================================
def parse_types(type_str: str) -> list:
    return [c for c in re.findall(r'[\w가-힣]+', str(type_str)) if c in CHARACTERS]

def safe_score(val) -> int:
    return int(float(val)) if str(val) not in ("", "nan") else 2

def show_image(path: str, columns=(1, 3, 1)):
    if os.path.exists(path):
        try:
            _, col_c, _ = st.columns(columns)
            with col_c:
                st.image(path, use_container_width=True)
        except Exception:
            pass

# ======================================================
# Google Sheets 연동
# ------------------------------------------------------
# [최초 설정 방법]
# 1. Google Cloud Console → 서비스 계정 생성 → JSON 키 발급
# 2. Google Sheets 파일명 "result" 생성
#    1행: 에디 | 크롱 | 뽀로로 | 루피 | 포비
#    2행:  0  |  0  |   0   |  0  |  0
# 3. 서비스 계정 이메일에 Google Sheets 편집 권한 공유
# 4. Streamlit Cloud → Manage App → Secrets 에 JSON 내용 등록
#    [gcp_service_account]
#    type = "service_account"
#    project_id = "..."
#    private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
#    client_email = "...@....iam.gserviceaccount.com"
#    ...
# ======================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_sheet():
    creds  = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME).sheet1

def save_result(character: str):
    try:
        sheet   = get_sheet()
        header  = sheet.row_values(1)
        if character in header:
            col_idx = header.index(character) + 1
            cur_val = sheet.cell(2, col_idx).value
            sheet.update_cell(2, col_idx, (int(cur_val) if cur_val else 0) + 1)
    except Exception as e:
        st.warning("결과 저장 오류: " + str(e))

def load_result_counts():
    try:
        sheet  = get_sheet()
        header = sheet.row_values(1)
        values = sheet.row_values(2)
        counts = {
            c: (int(values[header.index(c)]) if c in header and header.index(c) < len(values) and values[header.index(c)] else 0)
            for c in CHARACTERS
        }
        return counts, sum(counts.values())
    except Exception:
        return {c: 0 for c in CHARACTERS}, 0

# ======================================================
# 라우팅
# ======================================================
def evaluate_and_route():
    scores    = st.session_state.scores
    max_score = max(scores.values())
    top       = [c for c, s in scores.items() if s == max_score]
    if len(top) == 1:
        finalize(top[0])
    else:
        st.session_state.tied_chars = top
        st.session_state.tb_order   = None
        st.session_state.page       = "tiebreak"
    st.rerun()

def finalize(character: str):
    st.session_state.result_character = character
    if not st.session_state.result_saved:
        save_result(character)
        st.session_state.result_saved = True
    st.session_state.page = "result"

# ======================================================
# 세션 관리
# ======================================================
DEFAULTS = {
    "page":             "home",
    "question_idx":     0,
    "scores":           {c: 0 for c in CHARACTERS},
    "result_character": None,
    "result_saved":     False,
    "tied_chars":       [],
    "tb_order":         None,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_session():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v if k != "scores" else {c: 0 for c in CHARACTERS}
    for k in [k for k in st.session_state if k.startswith("q_order_")]:
        del st.session_state[k]

# ======================================================
# 스타일 — 배경 이미지
# ======================================================
if os.path.exists(BG_IMAGE_FILE):
    bg_b64 = base64.b64encode(open(BG_IMAGE_FILE, "rb").read()).decode()
    st.markdown(
        "<style>.stApp {"
        + "background-image:url('data:image/png;base64," + bg_b64 + "');"
        + "background-size:cover;background-position:center;"
        + "background-repeat:no-repeat;background-attachment:fixed;"
        + "}</style>",
        unsafe_allow_html=True
    )

# ======================================================
# 스타일 — 폰트 / 컴포넌트
# ======================================================
_fi = ("@import url('" + FONT_IMPORT_URL + "');") if FONT_IMPORT_URL else ""
st.markdown("\n".join([
    "<style>",
    _fi,
    # 전체 기본 폰트
    "html,body,[class*='css'],.stMarkdown{font-family:" + FONT_BODY + ";}",
    # 버튼 폰트 (다층 셀렉터)
    "button,.stButton button,.stButton>button,"
    "div[data-testid='stBaseButton-secondary'],"
    "div[data-testid='stBaseButton-primary'],"
    "button[kind='secondary'],button[kind='primary']"
    "{font-family:" + FONT_BODY + " !important;}",
    # 공통 클래스
    ".big-title{text-align:center;font-size:2rem;font-weight:800;margin-bottom:0.5rem;}",
    ".sub-title{text-align:center;font-size:1.1rem;color:#666;margin-bottom:1.5rem;}",
    ".res-char{text-align:center;font-size:2.4rem;font-weight:900;margin:1rem 0;}",
    ".stat-box{background:#f0f4ff;border-radius:12px;padding:1.2rem;margin-top:1rem;text-align:center;}",
    ".stat-pct{font-size:1.2rem;margin-top:0.5rem;}",
    # 질문 텍스트 — 별도 폰트
    ".q-text{text-align:center;font-size:1.35rem;font-weight:700;margin-bottom:1.2rem;"
    "line-height:1.5;color:" + COLOR_PRIMARY + ";font-family:" + FONT_QUESTION + ";}",
    # 진행 상황
    ".prog-text{text-align:center;font-size:0.9rem;color:" + COLOR_PRIMARY + ";margin-bottom:0.4rem;}",
    ".stProgress>div>div>div>div{background-color:" + COLOR_PRIMARY + ";}",
    # 답변 버튼
    "div[data-testid='stHorizontalBlock'] .stButton button"
    "{background-color:" + COLOR_PRIMARY + ";color:#fff;border:none;border-radius:10px;font-weight:700;}",
    "div[data-testid='stHorizontalBlock'] .stButton button:hover"
    "{background-color:" + COLOR_HOVER + ";color:#fff;}",
    # 동점 배지
    ".tie-badge{display:inline-block;background:" + COLOR_PRIMARY + ";color:#fff;"
    "border-radius:20px;padding:4px 18px;font-size:0.9rem;margin-bottom:1rem;}",
    "</style>",
]), unsafe_allow_html=True)

# ======================================================
# OG 메타 태그 (URL 공유 시 대표 이미지)
# OG_IMAGE_URL, OG_URL 을 실제 값으로 변경 후 사용
# ======================================================
st.markdown(
    "<meta property='og:title' content='팀플 빌런즈 : 눈 속 마을 빌런 테스트'/>"
    + "<meta property='og:description' content='당신은 어떤 팀플 빌런 캐릭터일까요?'/>"
    + "<meta property='og:image' content='" + OG_IMAGE_URL + "'/>"
    + "<meta property='og:url' content='" + OG_URL + "'/>"
    + "<meta property='og:type' content='website'/>"
    + "<meta name='twitter:card' content='summary_large_image'/>"
    + "<meta name='twitter:image' content='" + OG_IMAGE_URL + "'/>",
    unsafe_allow_html=True
)

# ══════════════════════════════════════════
# 홈 페이지
# ══════════════════════════════════════════
if st.session_state.page == "home":

    st.markdown("<div class='big-title'>팀플 빌런즈</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>눈 속 마을 빌런 테스트</div>", unsafe_allow_html=True)
    show_image("단체컷.png")
    st.write("")

    if st.button("🐧 시작하기", use_container_width=True):
        reset_session()
        st.session_state.page = "question"
        st.rerun()


# ══════════════════════════════════════════
# 질문 페이지
# ══════════════════════════════════════════
elif st.session_state.page == "question":

    total = len(questions_df)
    idx   = st.session_state.question_idx

    if idx >= total:
        evaluate_and_route()
        st.stop()

    row = questions_df.iloc[idx]

    st.markdown(f"<div class='prog-text'>질문 {idx+1} / {total}</div>", unsafe_allow_html=True)
    st.progress(idx / total)
    st.write("")

    st.markdown(f"<div class='q-text'>{row['질문']}</div>", unsafe_allow_html=True)
    show_image(f"질문 이미지/question.{idx+1}.png")
    st.write("")

    order_key = f"q_order_{idx}"
    if order_key not in st.session_state:
        opts = [
            ("A", str(row["답변A"]), str(row["답변A_유형"]), safe_score(row["답변A_점수"])),
            ("B", str(row["답변B"]), str(row["답변B_유형"]), safe_score(row["답변B_점수"])),
        ]
        random.shuffle(opts)
        st.session_state[order_key] = opts

    col1, col2 = st.columns(2)
    for col, (label, text, types_str, score) in zip([col1, col2], st.session_state[order_key]):
        with col:
            if not text or text.strip() in ("", "nan"):
                continue
            if st.button(text, use_container_width=True, key=f"q{idx}_{label}"):
                for char in parse_types(types_str):
                    st.session_state.scores[char] += score
                st.session_state.question_idx += 1
                st.rerun()


# ══════════════════════════════════════════
# 동점 추가 질문 페이지
# ══════════════════════════════════════════
elif st.session_state.page == "tiebreak":

    tied       = st.session_state.tied_chars
    valid_tied = [c for c in tied if c in TIEBREAK_QUESTION]

    if len(valid_tied) <= 1:
        finalize(valid_tied[0] if valid_tied else tied[0])
        st.rerun()

    st.markdown(
        "<div style='text-align:center'>"
        "<span class='tie-badge'>⚖️ 두구두구! 마지막 질문</span>"
        "</div>",
        unsafe_allow_html=True
    )
    st.write("")
    st.markdown(f"<div class='q-text'>{TIEBREAK_QUESTION['질문']}</div>", unsafe_allow_html=True)
    st.write("")

    if st.session_state.tb_order is None or len(st.session_state.tb_order) != len(valid_tied):
        opts = [(c, TIEBREAK_QUESTION[c]) for c in valid_tied]
        random.shuffle(opts)
        st.session_state.tb_order = opts

    cols = st.columns(len(valid_tied))
    for col, (char, text) in zip(cols, st.session_state.tb_order):
        with col:
            if st.button(text, use_container_width=True, key=f"tb_{char}"):
                finalize(char)
                st.rerun()


# ══════════════════════════════════════════
# 결과 페이지
# ══════════════════════════════════════════
elif st.session_state.page == "result":

    if not st.session_state.result_character:
        st.session_state.page = "home"
        st.rerun()

    character     = st.session_state.result_character
    counts, total = load_result_counts()

    st.markdown("<div class='big-title'>🎉 결과 발표!</div>", unsafe_allow_html=True)
    st.write("")

    img_path = "result_page/" + character + ".png"
    if os.path.exists(img_path):
        show_image(img_path)
    else:
        st.info("📁 이미지 파일 미등록: result_page/" + character + ".png")

    st.markdown(
        "<div class='res-char'>당신은 <span style='color:" + COLOR_PRIMARY + "'>"
        + character + "</span>입니다!</div>",
        unsafe_allow_html=True
    )
    st.write("")

    count   = counts.get(character, 0)
    pct     = round(count / total * 100, 1) if total > 0 else 0.0
    st.markdown(
        "<div class='stat-box'>"
        + "<div style='font-size:1rem;color:#444;margin-bottom:0.4rem'>지금까지 이 테스트에 참여한 인원</div>"
        + "<div class='stat-pct'><b>" + str(pct) + "%</b> 의 인원이 <b>" + character + "</b> 를 선택했습니다.</div>"
        + "</div>",
        unsafe_allow_html=True
    )
    st.write("")

    if st.button("🔄 다시하기", use_container_width=True):
        reset_session()
        st.rerun()
