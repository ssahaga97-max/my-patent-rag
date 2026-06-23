import streamlit as st
import pandas as pd
import os
import io
import re
import chromadb
from chromadb.api.client import SharedSystemClient
from chromadb.utils import embedding_functions
from langchain_groq import ChatGroq
import openpyxl
import json
import base64
import hashlib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from urllib.request import Request, urlopen
from urllib.error import HTTPError

# --- 1. 클라우드 서버 전용 절대 경로 고정 및 초기화 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_EXCEL_PATH    = os.path.join(BASE_DIR, "my_patent_folder", "master_patents.xlsx")
USER_REGISTRY_PATH   = os.path.join(BASE_DIR, "my_patent_folder", "user_registry.json")
os.makedirs(os.path.join(BASE_DIR, "my_patent_folder"), exist_ok=True)

st.set_page_config(page_title="AI 경쟁사 특허 조사 분석", layout="wide", page_icon="🔬")


# ==========================================
# 0. 전역 스코프 세션 상태 격리 및 초기화
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
# 프로세스 재시작 감지용 플래그 — @st.cache_resource가 초기화되면 항상 True
if "github_synced" not in st.session_state:
    st.session_state.github_synced = False
if "upload_feedback" not in st.session_state:
    st.session_state.upload_feedback = None
if "sync_feedback" not in st.session_state:
    st.session_state.sync_feedback = None
if "auto_reindex_attempted" not in st.session_state:
    st.session_state.auto_reindex_attempted = False
if "_last_chroma_count" not in st.session_state:
    st.session_state._last_chroma_count = 0
if "chroma_gap_sync_done" not in st.session_state:
    st.session_state.chroma_gap_sync_done = False


# ==========================================
# 공통 유틸리티
# ==========================================
def _clean_ascii(value: str) -> str:
    """비가시적 유니코드 문자(BOM, Zero-Width Space 등) 제거."""
    return value.encode("ascii", errors="ignore").decode("ascii").strip()


def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _normalize_patent_id(value) -> str:
    """
    출원번호 정규화(canonical ID) — 특허DB·Excel마다 다른 표기를 하나의 키로 통일.

    처리 순서:
      1) Excel float / 지수표기(1.02E+12) → 정수 문자열
      2) 구분자(하이픈·공백·슬래시·점) 제거
      3) KR/kr 접두사만 선두에서 제거(국제출원 고유 문자열 보존)
      4) 영숫자만 유지(대문자) — PCT/US 등 국제출원번호 대응

    동일 특허의 서로 다른 표기(10-2020-0012345 / 1020200012345)는 같은 ID가 됨.
    정규화 후에도 다른 ID면 별도 건으로 유지(누락 방지 우선).
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            if float(value) == int(float(value)):
                return str(int(float(value)))
        except (ValueError, OverflowError):
            pass

    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return ""

    if re.fullmatch(r"-?\d+\.?\d*[eE][+\-]?\d+", s):
        try:
            return str(int(float(s)))
        except (ValueError, OverflowError):
            pass

    if s.endswith(".0") and s[:-2].replace(".", "", 1).isdigit():
        s = s[:-2]

    s = re.sub(r"[\s\-_./\\]+", "", s)
    if s.lower().startswith("kr"):
        s = s[2:]
    s = re.sub(r"[^A-Za-z0-9]", "", s).upper()
    return s


# ── 통계 집계: 복수값 분리 · 경쟁사 대표명화 ─────────────────────────────
_MULTI_VALUE_SPLIT_RE = re.compile(r"\s*[|;/]\s*")

# lookup key(대문자·기호 제거) → 화면용 대표명 (경쟁사 20사 내외 수동 매핑)
_APPLICANT_CANONICAL_KEYS: dict[str, str] = {
    "DIEBOLD":                         "DIEBOLD NIXDORF",
    "DIEBOLDNIXDORF":                  "DIEBOLD NIXDORF",
    "DIEBOLDNIXDORFINCORPORATED":      "DIEBOLD NIXDORF",
    "DIEBOLDNIXDORFSYSTEMSGMBH":       "DIEBOLD NIXDORF",
    "WINCORNIXDORF":                   "DIEBOLD NIXDORF",
    "WINCORNIXDORFINTERNATIONALGMBH":"DIEBOLD NIXDORF",
    "NCR":                             "NCR",
    "NCRCORPORATION":                  "NCR",
    "NAUTILUSHYOSUNG":                 "HYOSUNG NAUTILUS",
    "HYOSUNG":                         "HYOSUNG NAUTILUS",
    "HYOSUNGNAUTILUS":                 "HYOSUNG NAUTILUS",
    "HOTS":                            "HOTS",
    "GLORY":                           "GLORY",
    "GLORYLTD":                        "GLORY",
    "OKIELECTRIC":                     "OKI ELECTRIC",
    "OKIELECTRICINDUSTRY":             "OKI ELECTRIC",
    "OKIELECTRICINDUSTRYCOLTD":        "OKI ELECTRIC",
    "FTEC":                            "FTEC",
    "GRGBANKING":                      "GRG BANKING",
    "GRGBANKINGEQUIPMENT":             "GRG BANKING",
    "GRGBANKINGEQUIPMENTCOLTD":        "GRG BANKING",
    "SHENZHENYIHUA":                   "SHENZHEN YIHUA",
    "SHENZHENYIHUACOMPCOLTD":          "SHENZHEN YIHUA",
    "SHENZHENYIHUATIMETECHNOLOGY":     "SHENZHEN YIHUA",
    "SHENZHENYIHUAFINANCIALINTELLIGENTRESINST": "SHENZHEN YIHUA",
    "CASHWAY":                         "CASHWAY",
    "CASHWAYTECHNOLOGY":               "CASHWAY",
    "GUARDIAN":                        "GUARDIAN",
    "GUARDIANANALYTICS":               "GUARDIAN",
    "FUJITSU":                         "FUJITSU",
    "HITACHI":                         "HITACHI",
    "TOSHIBA":                         "TOSHIBA",
    "RICOH":                           "RICOH",
    "CANON":                           "CANON",
    "PANASONIC":                       "PANASONIC",
    "CUMMINSALLISON":                  "CUMMINS ALLISON",
    "DE LA RUE":                       "DE LA RUE",
    "DELARUE":                         "DE LA RUE",
    "GIESSECKE":                       "GIECKE+DEVRIENT",
    "GIESECKE":                        "GIECKE+DEVRIENT",
    "GIECKE":                          "GIECKE+DEVRIENT",
}

# lookup key 앞부분 일치 시 대표명 (매핑표에 없는 변형 포착)
_APPLICANT_PREFIX_RULES: list[tuple[str, str]] = [
    ("DIEBOLD",          "DIEBOLD NIXDORF"),
    ("WINCOR",           "DIEBOLD NIXDORF"),
    ("NCR",              "NCR"),
    ("NAUTILUS",         "HYOSUNG NAUTILUS"),
    ("HYOSUNG",          "HYOSUNG NAUTILUS"),
    ("GLORY",            "GLORY"),
    ("OKIELECTRIC",      "OKI ELECTRIC"),
    ("GRGBANKING",       "GRG BANKING"),
    ("SHENZHENYIHUA",    "SHENZHEN YIHUA"),
    ("CASHWAY",          "CASHWAY"),
    ("GUARDIAN",         "GUARDIAN"),
    ("CUMMINSALLISON",   "CUMMINS ALLISON"),
    ("GIESSECKE",        "GIECKE+DEVRIENT"),
    ("GIESECKE",         "GIECKE+DEVRIENT"),
    ("GIECKE",           "GIECKE+DEVRIENT"),
    ("DELARUE",          "DE LA RUE"),
]

_CORP_SUFFIX_PATTERNS = (
    r"incorporated", r"inc\.?", r"corp\.?", r"corporation", r"ltd\.?", r"limited",
    r"gmbh", r"co\.?", r"company", r"llc", r"plc", r"pte\.?", r"ag", r"sa", r"bv",
    r"주식회사", r"\(주\)", r"㈜", r"유한회사", r"\(유\)", r"coltd", r"col\.?",
    r"holdings?", r"group", r"international", r"systems?", r"equipment",
    r"technology", r"technologies", r"industry", r"industries", r"financial",
    r"intelligent", r"research", r"inst(?:itute)?", r"res", r"inst",
)


def _applicant_lookup_key(name: str) -> str:
    """출원인 문자열을 alias 조회용 키로 변환."""
    s = str(name).strip()
    for _ in range(3):
        prev = s
        s = s.replace(",", " ")
        for pat in _CORP_SUFFIX_PATTERNS:
            s = re.sub(rf"\b{pat}\b", "", s, flags=re.IGNORECASE)
        s = re.sub(r"[\s\-_./\\()（）\[\]]+", "", s)
        if s == prev:
            break
    return s.upper()


def _canonical_applicant_name(name: str) -> str:
    """경쟁사 출원인 표기를 대표명으로 통합 (집계 전용, 메타데이터 원본은 유지)."""
    raw = str(name).strip()
    if not raw or raw.lower() in ("nan", "none", "없음", "정보없음", "미기재"):
        return ""

    key = _applicant_lookup_key(raw)
    if not key:
        return raw

    if key in _APPLICANT_CANONICAL_KEYS:
        return _APPLICANT_CANONICAL_KEYS[key]

    for prefix, canonical in _APPLICANT_PREFIX_RULES:
        if key.startswith(prefix):
            return canonical

    # 한글 포함 시 공백·접미사만 정리한 표시명 반환
    if re.search(r"[가-힣]", raw):
        cleaned = re.sub(r"\s*(\(주\)|주식회사|㈜|유한회사|\(유\))\s*", "", raw).strip()
        return cleaned if cleaned else raw

    return raw


def _explode_multi_values(value, *, allow_comma: bool = False) -> list[str]:
    """한 셀에 묶인 복수 값(출원인·발명자·IPC)을 개별 항목 리스트로 분리."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "없음", "정보없음", "미기재"):
        return []

    parts = [p.strip() for p in _MULTI_VALUE_SPLIT_RE.split(s) if p.strip()]
    if allow_comma and len(parts) <= 1:
        parts = [p.strip() for p in re.split(r"\s*,\s*", s) if p.strip()]

    return [
        p for p in parts
        if p and p.lower() not in ("nan", "none", "없음", "정보없음", "미기재")
    ]


def _flatten_for_counts(
    series: pd.Series,
    *,
    normalizer=None,
    explode: bool = False,
    allow_comma: bool = False,
) -> pd.Series:
    """Series를 value_counts 가능한 flat Series로 변환."""
    items: list[str] = []
    for val in series:
        if explode:
            parts = _explode_multi_values(val, allow_comma=allow_comma)
        else:
            parts = [str(val).strip()] if pd.notna(val) else []
        for part in parts:
            item = normalizer(part) if normalizer else part.strip()
            if item and item.lower() not in ("nan", "none", "없음"):
                items.append(item)
    return pd.Series(items, dtype=str)


def _top_counts_table(
    series: pd.Series,
    n: int = 10,
    *,
    label: str = "항목",
    normalizer=None,
    explode: bool = False,
    allow_comma: bool = False,
) -> tuple[pd.DataFrame, str]:
    """상위 N건 집계표(DataFrame)와 LLM용 마크다운 표 문자열 반환."""
    flat = _flatten_for_counts(
        series, normalizer=normalizer, explode=explode, allow_comma=allow_comma
    )
    if flat.empty:
        empty = pd.DataFrame(columns=[label, "건수"])
        return empty, f"| {label} | 건수 |\n| --- | ---: |\n| (데이터 없음) | 0 |"

    counts = flat.value_counts().head(n)
    df = pd.DataFrame({label: counts.index, "건수": counts.values})

    lines = [f"| {label} | 건수 |", "| --- | ---: |"]
    for _, row in df.iterrows():
        lines.append(f"| {row[label]} | {int(row['건수'])} |")
    return df, "\n".join(lines)


def _year_counts_table(series: pd.Series) -> tuple[pd.DataFrame, str]:
    """출원 연도별 건수 집계표."""
    years = (
        series.astype(str).str[:4]
        .replace("", pd.NA)
        .replace("없음", pd.NA)
        .dropna()
    )
    if years.empty:
        empty = pd.DataFrame(columns=["연도", "건수"])
        return empty, "| 연도 | 건수 |\n| --- | ---: |\n| (데이터 없음) | 0 |"

    counts = years.value_counts().sort_index()
    df = pd.DataFrame({"연도": counts.index, "건수": counts.values})

    lines = ["| 연도 | 건수 |", "| --- | ---: |"]
    for _, row in df.iterrows():
        lines.append(f"| {row['연도']} | {int(row['건수'])} |")
    return df, "\n".join(lines)


def _extract_year_filter_from_query(query: str) -> tuple[int | None, int | None]:
    """질의문에서 출원연도 필터 추출. (min_year, max_year) — max_year None이면 상한 없음."""
    min_y, max_y = None, None

    m = re.search(r"(\d{4})\s*년?\s*[~\-–]\s*(\d{4})\s*년?", query)
    if m:
        return int(m.group(1)), int(m.group(2))

    for pat in (
        r"(\d{4})\s*년?\s*(이후|부터|이상)",
        r"(\d{4})\s*년?\s*~\s*(현재|지금|now)",
        r"after\s*(\d{4})",
        r"since\s*(\d{4})",
    ):
        m = re.search(pat, query, re.IGNORECASE)
        if m:
            min_y = int(m.group(1))
            break

    m = re.search(r"(\d{4})\s*년?\s*(이전|까지|미만)", query)
    if m:
        max_y = int(m.group(1))

    return min_y, max_y


def _wants_applicant_cohort(query: str) -> bool:
    """출원인별·경쟁사별 코호트 분석 의도 감지."""
    q = query.lower()
    keywords = (
        "출원인별", "출원인 별", "권리자별", "회사별", "경쟁사별", "업체별",
        "applicant", "by applicant", "출원인마다", "출원인 마다", "각 출원인",
    )
    if any(k in q for k in keywords):
        return True
    return "각각" in query and "출원인" in query


def _filter_metadata_df_by_year(
    df: pd.DataFrame, min_year: int | None, max_year: int | None = None
) -> pd.DataFrame:
    """메타데이터 DataFrame을 출원연도 범위로 필터."""
    if min_year is None and max_year is None:
        return df
    if "출원일" not in df.columns:
        return df
    years = pd.to_numeric(df["출원일"].astype(str).str[:4], errors="coerce")
    mask = years.notna()
    if min_year is not None:
        mask &= years >= min_year
    if max_year is not None:
        mask &= years <= max_year
    return df[mask].copy()


def _metadata_df_with_applicant_rows(df: pd.DataFrame) -> pd.DataFrame:
    """복수 출원인(구분자) 행을 출원인 단위로 펼침."""
    records: list[dict] = []
    for _, row in df.iterrows():
        apps = _explode_multi_values(row.get("출원인", ""), allow_comma=False)
        if not apps:
            canon = _canonical_applicant_name(row.get("출원인", ""))
            apps = [canon] if canon else []
        else:
            apps = [c for a in apps if (c := _canonical_applicant_name(a))]
        if not apps:
            apps = ["미기재"]
        for app in apps:
            rec = row.to_dict()
            rec["_대표출원인"] = app
            records.append(rec)
    return pd.DataFrame(records) if records else pd.DataFrame()


def _build_applicant_cohort_context(
    df: pd.DataFrame,
    max_applicants: int = 25,
    titles_per: int = 15,
    ipc_per: int = 8,
) -> tuple[str, pd.DataFrame]:
    """
    필터된 메타데이터에서 출원인별 IPC·특허명 코호트 요약 생성.
    RAG 10건 제한 없이 전체(필터 범위) 집계 데이터를 LLM에 전달.
    """
    exploded = _metadata_df_with_applicant_rows(df)
    if exploded.empty:
        return "(조건에 맞는 데이터 없음)", pd.DataFrame(columns=["출원인(대표명)", "특허건수"])

    parts: list[str] = []
    summary_rows: list[dict] = []
    app_counts = exploded["_대표출원인"].value_counts()

    for app, cnt in app_counts.head(max_applicants).items():
        sub = exploded[exploded["_대표출원인"] == app]
        _, ipc_md = _top_counts_table(
            sub.get("IPC", pd.Series(dtype=str)),
            n=ipc_per,
            label="IPC",
            explode=True,
            allow_comma=True,
        )
        titles = [
            str(t).strip()
            for t in sub.get("명칭", pd.Series(dtype=str))
            if str(t).strip() not in ("", "없음", "정보없음", "nan", "none")
        ][:titles_per]
        title_block = "\n".join(f"  - {t}" for t in titles) if titles else "  - (명칭 없음)"

        parts.append(
            f"#### {app} — {int(cnt)}건\n"
            f"주요 IPC:\n{ipc_md}\n"
            f"대표 특허명 ({len(titles)}건):\n{title_block}\n"
        )
        summary_rows.append({"출원인(대표명)": app, "특허건수": int(cnt)})

    return "\n".join(parts), pd.DataFrame(summary_rows)


# ==========================================
# GitHub API
# ==========================================
def _get_github_secrets():
    """
    GITHUB_TOKEN / GITHUB_REPO_URL을 Streamlit Secrets에서 로드.
    urllib은 HTTP 헤더를 latin-1로 인코딩하므로, Secrets에서 복사·붙여넣기 시
    섞여 들어온 비가시적 유니코드 문자(BOM, Zero-Width Space 등)를 ASCII 필터로 제거.
    """
    try:
        token    = st.secrets["GITHUB_TOKEN"]
        repo_url = st.secrets["GITHUB_REPO_URL"]
        if not token or token.startswith("ghp_본인의"):
            return None, None

        token    = _clean_ascii(token)
        repo_url = _clean_ascii(repo_url)

        if not token:
            return None, None
        if not repo_url.startswith("https://"):
            repo_url = "https://" + repo_url.lstrip("http://")
        return token, repo_url
    except Exception:
        return None, None


def _get_github_repo_path() -> tuple[str | None, str | None, str | None]:
    """(token, repo_url, repo_path) 반환. 실패 시 (None, None, None)."""
    token, repo_url = _get_github_secrets()
    if not token:
        return None, None, None
    repo_path = repo_url.replace(".git", "").split("github.com/")[-1]
    return token, repo_url, repo_path


def diagnose_github() -> dict:
    """
    GitHub 연결 상태를 단계별로 진단하여 dict로 반환.
    keys: secret_ok, token_prefix, repo_url, api_reachable, repo_accessible, error
    """
    result = {
        "secret_ok": False, "token_prefix": "", "repo_url": "",
        "api_reachable": False, "repo_accessible": False, "error": ""
    }
    token, repo_url, repo_path = _get_github_repo_path()
    if not token:
        result["error"] = (
            "GITHUB_TOKEN 또는 GITHUB_REPO_URL을 Streamlit Secrets에서 찾을 수 없습니다.\n"
            "TOML 구조에서 두 키가 [USER_CREDENTIALS] 섹션 헤더보다 위에 있는지 확인하세요."
        )
        return result
    result["secret_ok"]    = True
    result["token_prefix"] = token[:12] + "..."
    result["repo_url"]     = repo_url

    # 토큰 원본에 비가시적 유니코드 문자가 있었는지 체크 (진단 정보용)
    raw_token = str(st.secrets.get("GITHUB_TOKEN", ""))
    if raw_token != _clean_ascii(raw_token):
        result["error"] = (
            "⚠️ GITHUB_TOKEN에 비가시적 유니코드 문자(복사·붙여넣기 오염)가 감지되었습니다. "
            "Streamlit Secrets 편집기에서 토큰 값을 지우고 직접 다시 입력하세요."
        )
        return result

    # 2단계: GitHub API 서버 도달 여부
    try:
        ping = Request("https://api.github.com", headers={"Accept": "application/vnd.github.v3+json"})
        with urlopen(ping, timeout=5):
            result["api_reachable"] = True
    except Exception as e:
        result["error"] = f"GitHub API 서버에 접근할 수 없습니다: {e}"
        return result

    # 3단계: 레포지토리 접근 권한 (토큰 유효성)
    try:
        req = Request(
            f"https://api.github.com/repos/{repo_path}",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github.v3+json"}
        )
        with urlopen(req, timeout=8) as resp:
            repo_info = json.loads(resp.read().decode())
            result["repo_accessible"] = True
            result["repo_name"]       = repo_info.get("full_name", "")
    except HTTPError as e:
        if e.code == 401:
            result["error"] = "토큰 인증 실패(401) — 토큰이 만료되었거나 잘못되었습니다. 새 토큰을 발급하세요."
        elif e.code == 404:
            result["error"] = f"레포지토리를 찾을 수 없습니다(404) — GITHUB_REPO_URL을 확인하세요: {repo_url}"
        else:
            result["error"] = f"GitHub API 오류 ({e.code}): {e.reason}"
    except Exception as e:
        result["error"] = f"레포지토리 접근 중 오류: {e}"

    return result


def upload_file_to_github_api(local_file_path, github_target_path):
    token, _, repo_path = _get_github_repo_path()
    if not token or not os.path.exists(local_file_path):
        return False

    try:
        with open(local_file_path, "rb") as f:
            content = base64.b64encode(f.read()).decode("utf-8")

        api_url = f"https://api.github.com/repos/{repo_path}/contents/{github_target_path}"
        
        sha = None
        req_get = Request(api_url, headers={
            "Authorization": f"Bearer {token}", 
            "Accept": "application/vnd.github.v3+json"
        })
        try:
            with urlopen(req_get, timeout=15) as response:
                res_data = json.loads(response.read().decode())
                sha = res_data.get("sha")
        except HTTPError as e:
            if e.code != 404:
                print(f"[API Warning] SHA 획득 실패: HTTP {e.code}")

        payload = {
            "message": f"[Automated Sync] {github_target_path}",
            "content": content,
            "branch": "main"
        }
        if sha:
            payload["sha"] = sha
            
        req_put = Request(
            api_url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
                "Accept": "application/vnd.github.v3+json"
            },
            method="PUT"
        )
        
        try:
            with urlopen(req_put, timeout=120) as response:
                if response.status in [200, 201]:
                    return True
                print(f"GitHub PUT 예상 외 응답: HTTP {response.status}")
                return False
        except HTTPError as e:
            print(f"GitHub PUT 실패: HTTP {e.code} {e.reason} — {e.read().decode(errors='ignore')[:200]}")
            return False
    except Exception as e:
        print(f"GitHub API 통신 중 예외 제어: {e}")
    return False

def commit_and_push_data() -> bool:
    """마스터 엑셀을 GitHub에 업로드하여 컨테이너 재시작 후에도 데이터가 복원될 수 있도록 보존.
    반환값: True=GitHub 업로드 성공, False=실패.
    """
    excel_status = upload_file_to_github_api(MASTER_EXCEL_PATH, "my_patent_folder/master_patents.xlsx")
    if excel_status:
        st.toast("💾 GitHub 백업 완료 — 재시작 후에도 데이터가 자동 복원됩니다!")
    else:
        st.error(
            "❌ **GitHub 백업 실패!**  \n"
            "특허 데이터가 현재 세션 ChromaDB에는 적재되었지만 GitHub에 저장되지 않았습니다.  \n"
            "**컨테이너 재시작 시 데이터가 소실됩니다.**  \n\n"
            "👉 좌측 사이드바 **'🔍 GitHub 연결 진단'** 버튼으로 원인 파악 후,  \n"
            "**'💾 GitHub 마스터 백업 재시도'** 버튼으로 재업로드하세요."
        )
    return excel_status


def delete_file_from_github_api(github_target_path: str) -> bool:
    """GitHub Contents API로 파일 삭제. 성공 또는 404(이미 없음) 시 True."""
    token, _, repo_path = _get_github_repo_path()
    if not token:
        return False

    api_url = f"https://api.github.com/repos/{repo_path}/contents/{github_target_path}"
    try:
        req_get = Request(api_url, headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json"
        })
        with urlopen(req_get, timeout=15) as r:
            sha = json.loads(r.read().decode()).get("sha", "")

        if not sha:
            return False

        del_payload = json.dumps({
            "message": f"[Format] Delete {github_target_path}",
            "sha": sha,
            "branch": "main"
        }).encode("utf-8")
        req_del = Request(api_url, data=del_payload, headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/vnd.github.v3+json"
        }, method="DELETE")
        with urlopen(req_del, timeout=30):
            return True
    except HTTPError as e:
        if e.code == 404:
            return True
        print(f"GitHub 삭제 실패: HTTP {e.code} {e.reason}")
        return False
    except Exception as e:
        print(f"GitHub 삭제 중 오류: {e}")
        return False


def _download_file_from_github(github_path: str, local_path: str):
    """
    GitHub에서 단일 파일을 내려받아 local_path에 저장.
    반환값: True=성공, None=404(파일 없음), False=오류
    """
    token, _, repo_path = _get_github_repo_path()
    if not token:
        return False
    try:
        api_url = f"https://api.github.com/repos/{repo_path}/contents/{github_path}"
        req = Request(api_url, headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json"
        })
        with urlopen(req) as resp:
            data = json.loads(resp.read().decode())

        if data.get("content"):
            raw = base64.b64decode(data["content"])
        elif data.get("download_url"):
            with urlopen(Request(
                data["download_url"],
                headers={"Authorization": f"Bearer {token}"}
            )) as r:
                raw = r.read()
        else:
            return False

        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(raw)
        return True
    except HTTPError as e:
        if e.code == 404:
            return None
        print(f"GitHub 파일 다운로드 실패 ({github_path}): HTTP {e.code} {e.reason}")
        return False
    except Exception as e:
        print(f"GitHub 파일 다운로드 실패 ({github_path}): {e}")
        return False


# ==========================================
# [회원 관리] 사용자 레지스트리
# ==========================================

def load_user_registry() -> dict:
    if os.path.exists(USER_REGISTRY_PATH):
        try:
            with open(USER_REGISTRY_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_user_registry(registry: dict):
    os.makedirs(os.path.dirname(USER_REGISTRY_PATH), exist_ok=True)
    with open(USER_REGISTRY_PATH, "w", encoding="utf-8") as f:
        json.dump(registry, f, ensure_ascii=False, indent=2)


def download_user_registry_from_github():
    return _download_file_from_github(
        "my_patent_folder/user_registry.json", USER_REGISTRY_PATH
    )


def upload_user_registry_to_github() -> bool:
    return upload_file_to_github_api(USER_REGISTRY_PATH, "my_patent_folder/user_registry.json")


def sync_all_from_github():
    """
    세션 최초 진입 시(프로세스 재시작 감지) GitHub에서 모든 영구 데이터를 무조건 동기화.
    반환값: (excel_result, registry_result)
      각 값: True=다운로드 성공, None=GitHub에 파일 없음(정상), False=실제 오류
    """
    excel_ok    = download_master_excel_from_github()
    registry_ok = download_user_registry_from_github()
    return excel_ok, registry_ok


# ==========================================
# [관리자 알림] 가입 이메일 발송 엔진
# ==========================================
def send_admin_signup_email(user_info: dict) -> bool:
    """
    신규 회원 가입 시 관리자 이메일로 알림 발송.
    Streamlit Secrets에 SMTP_EMAIL, SMTP_PASSWORD, ADMIN_EMAIL 설정 필요.
    Gmail 사용 시 앱 비밀번호(App Password) 사용 권장.
    """
    try:
        smtp_email    = st.secrets.get("SMTP_EMAIL", "")
        smtp_password = st.secrets.get("SMTP_PASSWORD", "")
        admin_email   = st.secrets.get("ADMIN_EMAIL", "")
        if not smtp_email or not smtp_password or not admin_email:
            return False

        msg = MIMEMultipart("alternative")
        msg["From"]    = smtp_email
        msg["To"]      = admin_email
        msg["Subject"] = f"[PatentRAG] 신규 연구원 가입 알림 — {user_info['username']}"

        html_body = f"""
<html><body style="font-family:sans-serif;">
<h2>🔔 PatentRAG 포털 신규 회원 가입 알림</h2>
<table border="1" cellpadding="10" style="border-collapse:collapse;min-width:400px;">
  <tr style="background:#f0f4ff;"><td><b>사용자 ID</b></td><td>{user_info['username']}</td></tr>
  <tr><td><b>이름</b></td><td>{user_info.get('name','')}</td></tr>
  <tr style="background:#f0f4ff;"><td><b>이메일</b></td><td>{user_info.get('email','')}</td></tr>
  <tr><td><b>부서</b></td><td>{user_info.get('department','미입력')}</td></tr>
  <tr style="background:#f0f4ff;"><td><b>가입 일시</b></td><td>{user_info.get('registered_at','')}</td></tr>
</table>
<p style="color:#888;font-size:12px;">본 메일은 PatentRAG 시스템이 자동 발송한 알림입니다.</p>
</body></html>"""

        msg.attach(MIMEText(html_body, "html", "utf-8"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10) as server:
            server.login(smtp_email, smtp_password)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"관리자 이메일 발송 실패: {e}")
        return False


def download_logo_from_github():
    """컨테이너 재시작 시 로고 파일이 없으면 GitHub에서 복원."""
    logo_path = os.path.join(BASE_DIR, "atec_logo.png")
    if os.path.exists(logo_path):
        return True
    return _download_file_from_github("atec_logo.png", logo_path) is True


def download_master_excel_from_github():
    return _download_file_from_github(
        "my_patent_folder/master_patents.xlsx", MASTER_EXCEL_PATH
    )


# ==========================================
# 1. 인증 시스템 — 관리자/사용자 이원화
# ==========================================
def _get_admin_credentials() -> dict:
    """Streamlit Secrets의 [USER_CREDENTIALS] 에서 관리자 자격증명 반환."""
    if "USER_CREDENTIALS" in st.secrets:
        return dict(st.secrets["USER_CREDENTIALS"])
    return {}


def check_authentication():
    if st.session_state.logged_in:
        return True

    # 앱 시작 시 사용자 레지스트리가 없으면 GitHub에서 복원
    if not os.path.exists(USER_REGISTRY_PATH):
        download_user_registry_from_github()

    logo_path = os.path.join(BASE_DIR, "atec_logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=160)
    st.title("AI 경쟁사 특허 조사 분석")

    # 로그인 화면 진입 시에도 레지스트리가 없으면 GitHub에서 복원
    if not os.path.exists(USER_REGISTRY_PATH):
        download_user_registry_from_github()

    tab_login, tab_register = st.tabs(["🔑 로그인", "📝 신규 회원 가입"])

    # ── 로그인 탭 ──
    with tab_login:
        st.subheader("사내 연구원 로그인")
        with st.form("login_form"):
            username = st.text_input("계정 ID", key="login_id")
            password = st.text_input("비밀번호", type="password", key="login_pw")
            submitted = st.form_submit_button("접속")

        if submitted:
            admin_creds = _get_admin_credentials()
            registry    = load_user_registry()

            # 관리자 계정 확인 (Secrets 평문 또는 SHA-256 해시)
            if username in admin_creds:
                stored = admin_creds[username]
                is_hex_hash = len(stored) == 64 and all(c in "0123456789abcdef" for c in stored.lower())
                stored_hash = stored if is_hex_hash else _hash_pw(stored)
                if stored_hash == _hash_pw(password):
                    st.session_state.logged_in = True
                    st.session_state.user_id   = username
                    st.session_state.is_admin  = True
                    st.rerun()
                else:
                    st.error("❌ 비밀번호가 올바르지 않습니다.")
            # 일반 사용자 확인 (레지스트리 해시 비교)
            elif username in registry:
                user_rec = registry[username]
                if not user_rec.get("active", True):
                    st.error("⛔ 비활성화된 계정입니다. 관리자에게 문의하세요.")
                elif user_rec.get("password_hash") == _hash_pw(password):
                    st.session_state.logged_in = True
                    st.session_state.user_id   = username
                    st.session_state.is_admin  = False
                    st.rerun()
                else:
                    st.error("❌ 비밀번호가 올바르지 않습니다.")
            else:
                st.error("❌ 등록되지 않은 계정입니다. '신규 회원 가입' 탭을 이용해 주세요.")

    # ── 회원가입 탭 ──
    with tab_register:
        st.subheader("신규 연구원 계정 등록")
        st.caption("가입 완료 시 관리자에게 이메일로 자동 통보됩니다.")
        with st.form("register_form"):
            r_id   = st.text_input("사용자 ID (영문·숫자, 4자 이상)")
            r_name = st.text_input("이름 *")
            r_email= st.text_input("이메일 *")
            r_dept = st.text_input("부서 (선택)")
            r_pw   = st.text_input("비밀번호 (6자 이상)", type="password")
            r_pw2  = st.text_input("비밀번호 확인", type="password")
            reg_submitted = st.form_submit_button("가입 신청")

        if reg_submitted:
            admin_creds = _get_admin_credentials()
            registry    = load_user_registry()
            errors = []

            if len(r_id) < 4:
                errors.append("ID는 4자 이상이어야 합니다.")
            if r_id in admin_creds or r_id in registry:
                errors.append("이미 사용 중인 ID입니다.")
            if not r_name or not r_email:
                errors.append("이름과 이메일은 필수 항목입니다.")
            if len(r_pw) < 6:
                errors.append("비밀번호는 6자 이상이어야 합니다.")
            if r_pw != r_pw2:
                errors.append("비밀번호가 일치하지 않습니다.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                user_info = {
                    "username":      r_id,
                    "name":          r_name,
                    "email":         r_email,
                    "department":    r_dept,
                    "password_hash": _hash_pw(r_pw),
                    "registered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "active":        True,
                }
                registry[r_id] = user_info
                save_user_registry(registry)

                with st.spinner("💾 계정 정보를 저장하는 중..."):
                    upload_user_registry_to_github()

                email_ok = send_admin_signup_email(user_info)

                st.success(f"✅ '{r_id}' 계정 가입이 완료되었습니다! 로그인 탭에서 접속해 주세요.")
                if email_ok:
                    st.info("📧 관리자에게 가입 알림 이메일이 발송되었습니다.")
                else:
                    st.caption("(이메일 발송 미설정 — Secrets에 SMTP_EMAIL / SMTP_PASSWORD / ADMIN_EMAIL 추가 시 활성화)")

    return False


# ==========================================
# 2. [프로세스 레벨 싱글톤] @st.cache_resource 기반 인프라 팩토리
# ==========================================

# ── ChromaDB / LLM 인프라 (@st.cache_resource) ───────────────────────────────
# Streamlit은 스크립트를 매 리런마다 처음부터 재실행하므로 모듈 전역 변수는 매번 None으로
# 초기화된다. chromadb SharedSystemClient 레지스트리는 프로세스 수준에서 유지되므로,
# 모듈 전역 싱글톤 + EphemeralClient() 재호출 조합은 "An instance already exists"를 유발한다.
_EMBED_MODEL      = "sentence-transformers/paraphrase-multilingual-mpnet-base-v2"
_COLLECTION_NAME  = "competitor_patents"

# Groq on_demand 무료 티어: 단일 요청 = 입력 토큰 + max_tokens(출력 예약) ≤ 6,000 TPM
GROQ_TPM_LIMIT = 6000
GROQ_MAX_OUTPUT_TOKENS = 1024
GROQ_REQUEST_MARGIN = 700  # 추정 오차·API 오버헤드 여유


@st.cache_resource
def _get_chroma_client():
    """
    EphemeralClient를 Streamlit cache_resource로 프로세스당 1회만 생성.
    chromadb 레지스트리 잔존 시 SharedSystemClient.clear_system_cache() 후 재시도.
    """
    for attempt in range(2):
        try:
            return chromadb.EphemeralClient()
        except ValueError:
            SharedSystemClient.clear_system_cache()
            if attempt == 1:
                raise


@st.cache_resource
def _get_embedding_fn():
    return embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name=_EMBED_MODEL
    )


@st.cache_resource
def _get_llm():
    groq_api_key = _clean_ascii(st.secrets.get("GROQ_API_KEY", ""))
    if not groq_api_key:
        raise ValueError("Streamlit Secrets에 GROQ_API_KEY가 설정되어 있지 않습니다.")
    return ChatGroq(
        model="llama-3.3-70b-versatile",
        groq_api_key=groq_api_key,
        temperature=0.1,
        max_tokens=GROQ_MAX_OUTPUT_TOKENS,
    )


def _get_collection():
    """캐시된 EphemeralClient에서 컬렉션 참조. 포맷 후 reset_collection()이 재생성."""
    client = _get_chroma_client()
    return client.get_or_create_collection(
        name=_COLLECTION_NAME,
        embedding_function=_get_embedding_fn()
    )


def reset_collection():
    """
    포맷 버튼 전용: 기존 컬렉션을 삭제하고 빈 컬렉션을 재생성.
    EphemeralClient는 @st.cache_resource로 유지 — 재생성하지 않음.
    """
    client = _get_chroma_client()
    try:
        client.delete_collection(_COLLECTION_NAME)
    except Exception:
        pass
    return client.get_or_create_collection(
        name=_COLLECTION_NAME,
        embedding_function=_get_embedding_fn()
    )


def load_permanent_infra_singleton():
    """캐시된 chroma_client, collection, llm을 반환."""
    if "infra_initialized" not in st.session_state:
        with st.spinner("📦 가상 특허 가동 커널 및 AI 전문 임베딩 엔진 초기화 중..."):
            _get_chroma_client()
            _get_llm()
        st.session_state.infra_initialized = True
    return _get_chroma_client(), _get_collection(), _get_llm()


def safe_count(collection) -> int:
    """
    collection.count()를 안전하게 호출.
    일시 오류 시 0 대신 마지막 정상 값을 반환 — 0이면 자동 재인덱싱이 오작동함.
    """
    try:
        n = collection.count()
        st.session_state._last_chroma_count = n
        return n
    except Exception as e:
        print(f"[ChromaDB safe_count 오류] {e}")
        return st.session_state.get("_last_chroma_count", 0)


def _estimate_tokens(text: str) -> int:
    """한국어 특허 텍스트 보수적 토큰 추정 (과소 추정 방지)."""
    if not text:
        return 0
    # 한국어 혼합: 약 1.2~1.7자/토큰 → 1.25자/토큰 가정
    return max(1, int(len(text) / 1.25) + 5)


def _chars_for_token_budget(tokens: int) -> int:
    """토큰 예산에 대응하는 최대 문자 수."""
    return max(100, int(tokens * 1.25))


def _max_prompt_input_tokens(system_prompt: str, user_query: str) -> int:
    """입력+출력 합계가 Groq TPM 한도 내가 되도록 참고 데이터에 쓸 수 있는 토큰."""
    instruction = "답변 시 참고한 특허 번호·명칭은 [번호](URL) 마크다운 링크 형식을 그대로 유지하세요."
    fixed = (
        f"[SYSTEM] {system_prompt}\n{instruction}\n"
        f"[참고 데이터]\n\n[사용자 요청]\n{user_query}\n\n"
        "보고서는 마크다운 양식으로 한국어로 작성하세요."
    )
    overhead = _estimate_tokens(fixed) + 30
    return max(
        1200,
        GROQ_TPM_LIMIT - GROQ_MAX_OUTPUT_TOKENS - GROQ_REQUEST_MARGIN - overhead,
    )


def _truncate_text(text: str, max_chars: int, suffix: str = "\n…(분량 제한으로 일부 생략)") -> str:
    text = text.strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rstrip() + suffix


def _format_patent_meta_line(i: int, m: dict) -> tuple[str, str, str]:
    app_num    = m.get("출원번호", "번호없음")
    p_name     = m.get("명칭", "제목없음")
    applicant  = m.get("출원인", "미기재")
    inventor   = m.get("발명자", "미기재")
    ipc        = m.get("IPC", "없음")
    cpc        = m.get("CPC", "없음")
    app_date   = m.get("출원일", "없음")
    patent_url = m.get("URL", "")

    if patent_url and patent_url.startswith("http"):
        display_num  = f"[{app_num}]({patent_url})"
        display_name = f"[{p_name}]({patent_url})"
    else:
        display_num  = app_num
        display_name = p_name

    header = (
        f"[특허 {i+1}] 번호: {display_num} | 명칭: {display_name} | "
        f"출원인: {applicant} | 발명자: {inventor} | "
        f"IPC: {ipc} | CPC: {cpc} | 출원일: {app_date}\n"
    )
    return header, display_num, display_name


def _parse_patent_doc(doc: str) -> dict:
    """ChromaDB 문서 문자열을 명칭·요약·청구항으로 분리."""
    doc = doc.strip()
    title, abstract, claims = "", "", ""

    if "특허요약:" in doc:
        before, rest = doc.split("특허요약:", 1)
        title = before.strip()
        if "특허청구항:" in rest:
            abstract, claims = rest.split("특허청구항:", 1)
            abstract, claims = abstract.strip(), claims.strip()
        else:
            abstract = rest.strip()
    else:
        abstract = doc

    return {"title": title, "abstract": abstract, "claims": claims}


def _compress_patent_doc_for_llm(doc: str, max_chars: int, *, claims_first: bool = False) -> tuple[str, bool]:
    """
    토큰 예산 내로 특허 본문 축약.
    claims_first=False: 요약 우선 — 청구항을 먼저 생략·축소, 요약을 최대한 유지.
    claims_first=True : 침해 분석용 — 청구항 우선, 요약을 먼저 축소.
    """
    parsed = _parse_patent_doc(doc)
    title = parsed["title"] or "특허명칭: (미기재)"
    abstract = parsed["abstract"]
    claims = parsed["claims"]

    def _join(title_line: str, abs_text: str, claims_text: str | None) -> str:
        parts = [title_line, f"특허요약: {abs_text}"]
        if claims_text is not None:
            parts.append(f"특허청구항: {claims_text}")
        return "\n".join(parts)

    omitted_claims = "(토큰 예산 초과로 생략)"
    omitted_abstract = "(토큰 예산 초과로 생략)"

    full = _join(title, abstract, claims if claims else None)
    if len(full) <= max_chars:
        return full, False

    truncated = True

    if not claims_first:
        # 1) 요약 전체 + 청구항 생략
        abstract_only = _join(title, abstract, omitted_claims if claims else None)
        if len(abstract_only) <= max_chars:
            return abstract_only, truncated

        # 2) 요약 일부 + 청구항 생략
        prefix = f"{title}\n특허요약: "
        suffix = f"\n특허청구항: {omitted_claims}" if claims else ""
        abs_budget = max_chars - len(prefix) - len(suffix)
        if abs_budget >= 80:
            compressed = prefix + _truncate_text(abstract, abs_budget, suffix="…") + suffix
            if len(compressed) <= max_chars:
                return compressed, truncated
    else:
        # 1) 청구항 전체 + 요약 생략
        if claims:
            claims_only = _join(title, omitted_abstract, claims)
            if len(claims_only) <= max_chars:
                return claims_only, truncated

            # 2) 청구항 일부 + 요약 생략
            prefix = f"{title}\n특허요약: {omitted_abstract}\n특허청구항: "
            claims_budget = max_chars - len(prefix)
            if claims_budget >= 80:
                compressed = prefix + _truncate_text(claims, claims_budget, suffix="…")
                if len(compressed) <= max_chars:
                    return compressed, truncated

        # 청구항 없으면 요약 우선으로 폴백
        abstract_only = _join(title, abstract, None)
        if len(abstract_only) <= max_chars:
            return abstract_only, truncated

    return _truncate_text(full, max_chars), truncated


def _build_rag_context_for_llm(
    docs: list, metas: list, token_budget: int, *, claims_first: bool = False
) -> tuple[str, bool]:
    """
    검색된 특허 문서를 Groq 입력 토큰 예산 내로 축소.
    건별로 요약/청구항 우선순위에 따라 축약한 뒤, 여전히 초과하면 건당 한도를 낮춘다.
    """
    n = len(docs)
    if n == 0:
        return "", False

    per_doc_chars = max(300, min(1200, _chars_for_token_budget(token_budget) // n))
    truncated = False

    while True:
        parts = []
        for i, (doc, m) in enumerate(zip(docs, metas)):
            header, _, _ = _format_patent_meta_line(i, m)
            body, doc_truncated = _compress_patent_doc_for_llm(
                doc, per_doc_chars, claims_first=claims_first
            )
            if doc_truncated:
                truncated = True
            parts.append(f"{header}{body}\n\n")

        combined = "".join(parts)
        if _estimate_tokens(combined) <= token_budget or per_doc_chars <= 250:
            if _estimate_tokens(combined) > token_budget:
                combined = _truncate_text(combined, _chars_for_token_budget(token_budget))
                truncated = True
            return combined, truncated

        per_doc_chars = int(per_doc_chars * 0.75)


def _assemble_llm_prompt(system_prompt: str, context_text: str, user_query: str) -> tuple[str, bool, str]:
    instruction = "답변 시 참고한 특허 번호·명칭은 [번호](URL) 마크다운 링크 형식을 그대로 유지하세요."
    context_budget = _max_prompt_input_tokens(system_prompt, user_query)
    truncated = False

    if _estimate_tokens(context_text) > context_budget:
        context_text = _truncate_text(context_text, _chars_for_token_budget(context_budget))
        truncated = True

    max_input = GROQ_TPM_LIMIT - GROQ_MAX_OUTPUT_TOKENS - GROQ_REQUEST_MARGIN

    for _ in range(8):
        prompt = (
            f"[SYSTEM] {system_prompt}\n"
            f"{instruction}\n\n"
            f"[참고 데이터]\n{context_text}\n\n"
            f"[사용자 요청]\n{user_query}\n\n"
            "보고서는 마크다운 양식으로 한국어로 작성하세요."
        )
        if _estimate_tokens(prompt) <= max_input:
            return prompt, truncated, context_text

        truncated = True
        context_text = _truncate_text(context_text, max(200, int(len(context_text) * 0.82)))

    prompt = (
        f"[SYSTEM] {system_prompt}\n"
        f"{instruction}\n\n"
        f"[참고 데이터]\n{context_text}\n\n"
        f"[사용자 요청]\n{user_query}\n\n"
        "보고서는 마크다운 양식으로 한국어로 작성하세요."
    )
    return prompt, truncated, context_text


# --- 3. 엑셀 파싱 및 무결성 메타데이터 적재 로직 ---
def _find_column(col_map: dict, *keywords: str) -> str | None:
    """col_map(대문자 키)에서 keywords 순서대로 첫 매칭 컬럼명 반환."""
    for kw in keywords:
        for k, v in col_map.items():
            if kw in k:
                return v
    return None


def _detect_columns(df: pd.DataFrame) -> dict:
    """DataFrame 컬럼명을 분석해 각 필드에 해당하는 실제 컬럼명 반환."""
    col_map = {str(c).strip().replace(" ", "").upper(): c for c in df.columns}

    # 출원번호: '공개번호'/'등록번호' 등 '번호'만 포함된 열을 먼저 잡지 않도록 우선순위 지정
    id_col = (
        _find_column(col_map, "출원번호", "APPLICATIONNO", "APPNO", "APPLNO")
        or _find_column(col_map, "특허번호", "PATENTNO")
    )
    if not id_col:
        for k, v in col_map.items():
            if "번호" in k and not any(x in k for x in ("공개", "등록", "연번", "일련", "SEQ")):
                id_col = v
                break
    if not id_col:
        id_col = df.columns[0]

    return {
        "id":        id_col,
        "title":     _find_column(col_map, "명칭", "제목", "특허명", "INVENTIONTITLE") or None,
        "abstract":  _find_column(col_map, "요약", "초록", "ABSTRACT") or None,
        "claims":    _find_column(col_map, "청구항", "청구", "범위", "CLAIM") or None,
        "app_date":  _find_column(col_map, "출원일", "출원일자", "APPDATE", "FILINGDATE") or None,
        "reg_date":  _find_column(col_map, "등록일", "등록일자", "REGDATE") or None,
        "ipc":       _find_column(col_map, "IPC") or None,
        "cpc":       _find_column(col_map, "CPC") or None,
        "inventor":  _find_column(col_map, "발명자", "발명인", "INVENTOR") or None,
        "applicant": _find_column(col_map, "출원인", "권리자", "APPLICANT", "ASSIGNEE") or None,
    }


def _get_cell(row, col, default: str = "없음") -> str:
    """컬럼이 존재하고 값이 있으면 str 반환."""
    if col and pd.notna(row.get(col)):
        val = str(row[col]).strip()
        return val if val else default
    return default


def _get_info_cell(row, col) -> str:
    """임베딩용 — 빈 값은 '정보없음'."""
    return _get_cell(row, col, default="정보없음")


def _build_metadata(row, cols: dict, patent_url: str = "") -> dict:
    """ChromaDB 메타데이터 딕셔너리 생성."""
    title = _get_info_cell(row, cols["title"]) if cols["title"] else "정보없음"
    return {
        "출원번호": str(row[cols["id"]]),
        "명칭":     title,
        "출원일":   _get_cell(row, cols["app_date"]),
        "등록일":   _get_cell(row, cols["reg_date"]),
        "IPC":      _get_cell(row, cols["ipc"]),
        "CPC":      _get_cell(row, cols["cpc"]),
        "발명자":   _get_cell(row, cols["inventor"]),
        "출원인":   _get_cell(row, cols["applicant"]),
        "URL":      patent_url,
    }


def _build_document(row, cols: dict) -> str:
    """ChromaDB 임베딩용 문서 텍스트 생성."""
    title    = _get_info_cell(row, cols["title"])
    abstract = _get_info_cell(row, cols["abstract"])
    claims   = _get_info_cell(row, cols["claims"])
    return f"특허명칭: {title}\n특허요약: {abstract}\n특허청구항: {claims}"


def _count_master_excel_patents() -> int | None:
    """마스터 엑셀의 고유 출원번호 건수. 파일 없으면 None."""
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return None
    try:
        df   = pd.read_excel(MASTER_EXCEL_PATH)
        cols = _detect_columns(df)
        ids  = {
            _normalize_patent_id(v)
            for v in df[cols["id"]]
            if _normalize_patent_id(v)
        }
        return len(ids)
    except Exception as e:
        print(f"[마스터 엑셀 건수 조회 오류] {e}")
        return None


def extract_excel_hyperlinks(uploaded_file):
    link_dict = {}
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        sheet = wb.active
        for row in sheet.iter_rows():
            for cell in row:
                if cell.hyperlink and cell.hyperlink.target:
                    cell_text = str(cell.value).strip().replace("-", "")
                    link_dict[cell_text] = cell.hyperlink.target
    except Exception as e:
        print(f"링크 파싱 스킵: {e}")
    return link_dict


def _get_chroma_ids(collection) -> set:
    """ChromaDB에 이미 적재된 id 집합."""
    try:
        if collection.count() == 0:
            return set()
        return set(collection.get(include=[])["ids"])
    except Exception as e:
        print(f"[ChromaDB id 조회 오류] {e}")
        return set()


def _count_chroma_unique_ids(collection) -> int:
    """ChromaDB 고유 document id 수 (중복 ID 문서 제외한 실질 특허 수)."""
    return len(_get_chroma_ids(collection))


def _render_feedback_box(feedback: dict | None) -> None:
    """session_state 피드백 메시지 렌더."""
    if not feedback:
        return
    level = feedback.get("level", "info")
    message = feedback.get("message", "")
    if level == "success":
        st.success(message)
    elif level == "warning":
        st.warning(message)
    elif level == "error":
        st.error(message)
    else:
        st.info(message)


def _build_db_count_status(chroma_n: int, master_n: int | None, unique_chroma_n: int) -> str:
    """Chroma·마스터 건수 비교 상태 문구."""
    count_md = f"📊 **누적 적재 데이터 (ChromaDB):** `{chroma_n}` 건"
    if master_n is None:
        return count_md

    count_md += f"  \n📄 **마스터 엑셀 (고유 출원번호):** `{master_n}` 건"
    if unique_chroma_n and unique_chroma_n != chroma_n:
        count_md += f"  \n🔑 **ChromaDB 고유 ID:** `{unique_chroma_n}` 건"

    if master_n == chroma_n:
        return count_md

    if chroma_n > master_n:
        dup_docs = chroma_n - unique_chroma_n
        if master_n >= unique_chroma_n:
            count_md += (
                f"  \n✅ 마스터 엑셀이 Chroma **고유 특허 전체({master_n}건)** 와 동기화되었습니다."
            )
            if dup_docs > 0:
                count_md += (
                    f"  \nℹ️ ChromaDB 문서 수에는 동일 특허 **중복 ID {dup_docs}건**이 "
                    f"포함되어 있습니다 (분석·검색에는 영향 없음)."
                )
        else:
            gap = unique_chroma_n - master_n
            count_md += (
                f"  \n⚠️ 마스터가 Chroma 고유 특허보다 **{gap}건** 부족합니다. "
                f"**ChromaDB → 마스터 엑셀 역동기화** 후 GitHub 백업하세요."
            )
    else:
        gap = master_n - chroma_n
        count_md += (
            f"  \n⚠️ 마스터가 ChromaDB보다 **{gap}건** 많습니다. "
            f"**ChromaDB 누락분 복구** 버튼을 실행하세요."
        )
    return count_md


def _upsert_patent_batches(collection, ids: list, docs: list, metas: list, batch_size: int = 100) -> None:
    """배치 upsert. 실패 시 예외 전파."""
    for i in range(0, len(ids), batch_size):
        collection.upsert(
            ids=ids[i:i + batch_size],
            documents=docs[i:i + batch_size],
            metadatas=metas[i:i + batch_size],
        )


def _compact_master_excel() -> int:
    """canonical 출원번호 기준 마스터 엑셀 중복 행 제거(마지막 행 유지). 제거된 행 수 반환."""
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return 0
    try:
        df     = pd.read_excel(MASTER_EXCEL_PATH)
        before = len(df)
        cols   = _detect_columns(df)
        deduped = _dedupe_dataframe_by_patent_id(df, cols)
        after  = len(deduped)
        if after < before:
            deduped.to_excel(MASTER_EXCEL_PATH, index=False)
            return before - after
    except Exception as e:
        print(f"[마스터 compact 오류] {e}")
    return 0


def sync_chroma_missing_from_master(collection, hyperlink_map: dict | None = None) -> int:
    """
    마스터 엑셀에는 있으나 ChromaDB에 없는 출원번호만 upsert.
    엑셀 저장 후 Chroma 타임아웃·부분 실패로 생긴 누락 복구용.
    시작 시 마스터 엑셀 canonical 중복 행도 자동 정리.
    """
    _compact_master_excel()
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return 0
    try:
        df         = pd.read_excel(MASTER_EXCEL_PATH)
        cols       = _detect_columns(df)
        chroma_ids = _get_chroma_ids(collection)
        hyperlink_map = hyperlink_map or {}

        rows_by_id: dict = {}
        for _, row in df.iterrows():
            pat_id = _normalize_patent_id(row[cols["id"]])
            if pat_id:
                rows_by_id[pat_id] = row

        ids, docs, metas = [], [], []
        for pat_id, row in rows_by_id.items():
            if pat_id in chroma_ids:
                continue
            patent_url = hyperlink_map.get(pat_id, "")
            if not patent_url and cols["title"]:
                clean_title = str(row[cols["title"]]).strip().replace("-", "")
                patent_url  = hyperlink_map.get(clean_title, "")
            ids.append(pat_id)
            docs.append(_build_document(row, cols))
            metas.append(_build_metadata(row, cols, patent_url=patent_url))

        if not ids:
            return 0
        _upsert_patent_batches(collection, ids, docs, metas)
        return len(ids)
    except Exception as e:
        print(f"[ChromaDB 누락분 동기화 오류] {e}")
        return 0


def _excel_display_value(value) -> str:
    """Chroma/엑셀 공통 표시값 정리."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "없음"
    v = str(value).strip()
    if not v or v.lower() in ("nan", "none", "정보없음"):
        return "없음"
    return v


def _master_row_to_standard_dict(row, cols: dict) -> dict:
    """마스터 엑셀 행 → 표준 컬럼 dict."""
    url = ""
    if "URL" in row.index and pd.notna(row.get("URL")):
        url = str(row["URL"]).strip()
    return {
        "출원번호": str(row[cols["id"]]),
        "명칭":     _get_cell(row, cols["title"]) if cols["title"] else "없음",
        "요약":     _get_cell(row, cols["abstract"]) if cols["abstract"] else "없음",
        "청구항":   _get_cell(row, cols["claims"]) if cols["claims"] else "없음",
        "출원일":   _get_cell(row, cols["app_date"]),
        "등록일":   _get_cell(row, cols["reg_date"]),
        "IPC":      _get_cell(row, cols["ipc"]),
        "CPC":      _get_cell(row, cols["cpc"]),
        "발명자":   _get_cell(row, cols["inventor"]),
        "출원인":   _get_cell(row, cols["applicant"]),
        "URL":      url,
    }


def _chroma_record_to_row(chroma_id: str, meta: dict, doc: str) -> dict:
    """ChromaDB id·메타데이터·문서 → 마스터 엑셀 표준 행."""
    parsed = _parse_patent_doc(doc or "")

    title = _excel_display_value(meta.get("명칭", ""))
    if title == "없음":
        raw_title = parsed.get("title", "")
        if raw_title.startswith("특허명칭:"):
            raw_title = raw_title.split("특허명칭:", 1)[1].strip()
        title = _excel_display_value(raw_title)

    abstract = parsed.get("abstract", "").strip()
    claims   = parsed.get("claims", "").strip()
    if abstract in ("", "정보없음"):
        abstract = "없음"
    if claims in ("", "정보없음"):
        claims = "없음"

    app_num = meta.get("출원번호", chroma_id)
    if not str(app_num).strip() or str(app_num).strip() in ("없음", "nan"):
        app_num = chroma_id

    return {
        "출원번호": str(app_num),
        "명칭":     title,
        "요약":     abstract,
        "청구항":   claims,
        "출원일":   _excel_display_value(meta.get("출원일", "")),
        "등록일":   _excel_display_value(meta.get("등록일", "")),
        "IPC":      _excel_display_value(meta.get("IPC", "")),
        "CPC":      _excel_display_value(meta.get("CPC", "")),
        "발명자":   _excel_display_value(meta.get("발명자", "")),
        "출원인":   _excel_display_value(meta.get("출원인", "")),
        "URL":      str(meta.get("URL", "") or "").strip(),
    }


_MASTER_EXCEL_COLUMNS = (
    "출원번호", "명칭", "요약", "청구항", "출원일", "등록일",
    "IPC", "CPC", "발명자", "출원인", "URL",
)


def sync_master_from_chroma(collection, batch_size: int = 500) -> tuple[int, int, int]:
    """
    ChromaDB → 마스터 엑셀 역동기화 (재임베딩·재파싱 없음, 메타데이터+문서만 읽음).
    동일 canonical 출원번호는 Chroma 데이터가 우선(last-wins).
    마스터에만 있던 행은 유지(합집합).
    반환: (chroma보낸 건수, 동기화 전 마스터 고유 건수, 동기화 후 마스터 고유 건수)
    """
    master_before = _count_master_excel_patents() or 0
    chroma_total  = safe_count(collection)
    if chroma_total == 0:
        return 0, master_before, master_before

    rows_by_id: dict[str, dict] = {}

    # 기존 마스터 행 선적재 (Chroma에 없는 행 보존)
    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0:
        try:
            master_df   = pd.read_excel(MASTER_EXCEL_PATH)
            master_cols = _detect_columns(master_df)
            for _, row in master_df.iterrows():
                cid = _normalize_patent_id(row[master_cols["id"]])
                if cid:
                    rows_by_id[cid] = _master_row_to_standard_dict(row, master_cols)
        except Exception as e:
            print(f"[마스터 선적재 오류] {e}")

    # ChromaDB 배치 읽기 → 덮어쓰기 (재임베딩 없음)
    exported = 0
    try:
        all_ids = collection.get(include=[])["ids"]
    except Exception as e:
        print(f"[ChromaDB id 목록 조회 오류] {e}")
        return 0, master_before, master_before

    for i in range(0, len(all_ids), batch_size):
        batch_ids = all_ids[i:i + batch_size]
        try:
            batch = collection.get(ids=batch_ids, include=["metadatas", "documents"])
        except Exception as e:
            print(f"[ChromaDB 배치 조회 오류] {e}")
            continue
        for chroma_id, meta, doc in zip(
            batch.get("ids", []),
            batch.get("metadatas", []),
            batch.get("documents", []),
        ):
            cid = _normalize_patent_id(chroma_id) or _normalize_patent_id(meta.get("출원번호", ""))
            if not cid:
                continue
            rows_by_id[cid] = _chroma_record_to_row(chroma_id, meta or {}, doc or "")
            exported += 1

    if not rows_by_id:
        return 0, master_before, master_before

    df_out = pd.DataFrame(list(rows_by_id.values()), columns=list(_MASTER_EXCEL_COLUMNS))
    os.makedirs(os.path.dirname(MASTER_EXCEL_PATH), exist_ok=True)
    df_out.to_excel(MASTER_EXCEL_PATH, index=False)

    master_after = len(rows_by_id)
    return exported, master_before, master_after


def _dedupe_dataframe_by_patent_id(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    """마스터 엑셀 저장 전 canonical 출원번호 기준 중복 행 제거(마지막 행 유지)."""
    if df.empty:
        return df
    keep_idx: dict[str, int] = {}
    for idx, row in df.iterrows():
        cid = _normalize_patent_id(row[cols["id"]])
        if cid:
            keep_idx[cid] = idx
    if not keep_idx:
        return df
    return df.loc[sorted(keep_idx.values())].reset_index(drop=True)


def _build_chroma_batches(rows_by_id: dict, cols: dict, hyperlink_map: dict):
    """canonical id → row dict에서 ChromaDB upsert 배치 생성."""
    batch_ids, batch_docs, batch_metas = [], [], []
    for doc_id, row in rows_by_id.items():
        patent_url = hyperlink_map.get(doc_id, "")
        if not patent_url and cols["title"]:
            clean_title = str(row[cols["title"]]).strip().replace("-", "")
            patent_url  = hyperlink_map.get(clean_title, "")
        batch_ids.append(doc_id)
        batch_docs.append(_build_document(row, cols))
        batch_metas.append(_build_metadata(row, cols, patent_url=patent_url))
    return batch_ids, batch_docs, batch_metas


def reindex_from_master_excel(collection) -> int:
    """
    재시작 후 ChromaDB 재구성 전용 함수.
    process_and_update_db는 uploaded_file을 MASTER_EXCEL_PATH와 비교해
    전부 중복으로 처리하는 문제가 있어, 재인덱싱은 이 함수를 사용한다.
    """
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return 0
    try:
        df   = pd.read_excel(MASTER_EXCEL_PATH)
        cols = _detect_columns(df)

        rows_by_id: dict = {}
        for _, row in df.iterrows():
            pat_id = _normalize_patent_id(row[cols["id"]])
            if pat_id:
                rows_by_id[pat_id] = row

        if not rows_by_id:
            return 0

        ids, docs, metas = [], [], []
        for pat_id, row in rows_by_id.items():
            ids.append(pat_id)
            docs.append(_build_document(row, cols))
            metas.append(_build_metadata(row, cols, patent_url=""))

        _upsert_patent_batches(collection, ids, docs, metas)
        return len(ids)
    except Exception as e:
        print(f"재인덱싱 실패: {e}")
        return 0


def process_and_update_db(uploaded_file, collection):
    """
    업로드 파일 기준 누락 없는 적재 (canonical 출원번호 스키마).

    원칙:
      · 동일 canonical ID = 중복 1건 (표기 차이는 _normalize_patent_id로 통합)
      · 업로드 파일의 모든 고유 ID는 ChromaDB에 반드시 존재해야 함
        (마스터에만 있고 Chroma에 없으면 스킵하지 않고 upsert)
      · ChromaDB 적재 성공 후 마스터 엑셀 저장 + canonical 기준 dedupe

    반환: (chroma_신규적재, master_신규행, 파일내중복, 출원번호없음, gap추가복구)
    """
    file_bytes    = uploaded_file.read()
    hyperlink_map = extract_excel_hyperlinks(io.BytesIO(file_bytes))

    try:
        new_df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"엑셀 파일 로드 실패: {e}")
        return 0, 0, 0, 0, 0

    cols = _detect_columns(new_df)

    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0:
        try:
            master_df   = pd.read_excel(MASTER_EXCEL_PATH)
            master_cols = _detect_columns(master_df)
            master_ids  = {
                _normalize_patent_id(v)
                for v in master_df[master_cols["id"]]
                if _normalize_patent_id(v)
            }
        except Exception:
            master_df  = pd.DataFrame(columns=new_df.columns)
            master_ids = set()
    else:
        master_df  = pd.DataFrame(columns=new_df.columns)
        master_ids = set()

    chroma_ids = _get_chroma_ids(collection)

    file_rows: dict = {}
    dup_in_file   = 0
    skipped_empty = 0

    for _, row in new_df.iterrows():
        cid = _normalize_patent_id(row[cols["id"]])
        if not cid:
            skipped_empty += 1
            continue
        if cid in file_rows:
            dup_in_file += 1
        file_rows[cid] = row

    chroma_needed = {cid: row for cid, row in file_rows.items() if cid not in chroma_ids}
    master_needed = {cid: row for cid, row in file_rows.items() if cid not in master_ids}
    already_indexed = len(file_rows) - len(chroma_needed)

    chroma_ingested = 0
    if chroma_needed:
        batch_ids, batch_docs, batch_metas = _build_chroma_batches(chroma_needed, cols, hyperlink_map)
        try:
            _upsert_patent_batches(collection, batch_ids, batch_docs, batch_metas)
            chroma_ingested = len(batch_ids)
        except Exception as e:
            st.error(f"ChromaDB 적재 실패 — 마스터 엑셀은 갱신하지 않았습니다: {e}")
            return 0, 0, dup_in_file + already_indexed, skipped_empty, 0

    master_added = 0
    if master_needed:
        added_df = pd.DataFrame(list(master_needed.values()))
        if master_df.empty:
            updated_master_df = added_df
        else:
            updated_master_df = pd.concat([master_df, added_df], ignore_index=True)
        master_cols = _detect_columns(updated_master_df)
        updated_master_df = _dedupe_dataframe_by_patent_id(updated_master_df, master_cols)
        updated_master_df.to_excel(MASTER_EXCEL_PATH, index=False)
        master_added = len(master_needed)

    synced = sync_chroma_missing_from_master(collection, hyperlink_map)
    return chroma_ingested, master_added, dup_in_file + already_indexed, skipped_empty, synced


# --- 4. 메인 어플리케이션 인터페이스 구동 런타임 ---
def run_main_portal():
    chroma_client, collection, llm = load_permanent_infra_singleton()

    # ── 세션 최초 진입 시: GitHub → 로컬 전체 동기화 ──
    # github_synced는 세션 단위 플래그. 프로세스 재시작(컨테이너 재생성) 시 항상 False로 초기화됨.
    if not st.session_state.github_synced:
        download_logo_from_github()
        with st.spinner("🔄 GitHub 데이터 웨어하우스 동기화 중..."):
            excel_ok, registry_ok = sync_all_from_github()

        if excel_ok is True:
            st.toast("✅ GitHub에서 마스터 특허 데이터 복원 완료")
        if registry_ok is True:
            st.toast("✅ GitHub에서 회원 정보 복원 완료")
        if excel_ok is False or registry_ok is False:
            # 실제 오류 (토큰/네트워크 문제) — None(파일 없음)은 오류 아님
            st.toast("⚠️ GitHub 동기화 중 오류 발생 — 사이드바 'GitHub 연결 진단' 확인 권장")
        elif excel_ok is None and registry_ok is None:
            st.toast("ℹ️ GitHub에 저장된 데이터 없음 (포맷 후 초기 상태 또는 최초 배포)")

        st.session_state.github_synced = True
        # 엑셀 복원 여부와 무관하게 재인덱싱은 아래 조건에서 처리

    # ChromaDB가 비어 있으면(재시작) 엑셀 기반 자동 재인덱싱 — 세션당 1회만
    chroma_n = safe_count(collection)
    master_n = _count_master_excel_patents()
    if (
        master_n
        and chroma_n == 0
        and not st.session_state.auto_reindex_attempted
    ):
        st.session_state.auto_reindex_attempted = True
        try:
            with st.spinner("📦 벡터 DB 자동 재인덱싱 중... (특허 수에 따라 1~3분 소요)"):
                restored = reindex_from_master_excel(collection)
            st.session_state._last_chroma_count = safe_count(collection)
            st.toast(f"✅ 벡터 DB 복원 완료 ({restored}건)")
        except Exception as e:
            st.warning(f"자동 재인덱싱 오류: {e}")

    elif (
        master_n
        and chroma_n > 0
        and master_n > chroma_n
        and not st.session_state.chroma_gap_sync_done
    ):
        st.session_state.chroma_gap_sync_done = True
        gap = master_n - chroma_n
        try:
            with st.spinner(f"ChromaDB 누락분 자동 복구 중 ({chroma_n}→{master_n}, 약 {gap}건)..."):
                synced = sync_chroma_missing_from_master(collection)
            st.session_state._last_chroma_count = safe_count(collection)
            if synced > 0:
                st.toast(f"✅ ChromaDB 누락분 {synced}건 복구 완료")
        except Exception as e:
            st.warning(f"ChromaDB 누락분 자동 복구 오류: {e}")

    is_admin = st.session_state.get("is_admin", False)

    col_logo, col_title, col_logout = st.columns([1, 7, 2])
    with col_logo:
        logo_path = os.path.join(BASE_DIR, "atec_logo.png")
        if os.path.exists(logo_path):
            st.image(logo_path, width=110)
    with col_title:
        st.title("AI 경쟁사 특허 조사 분석")
        mode_label = "🔧 관리자" if is_admin else "👤 사용자"
        master_n = _count_master_excel_patents()
        count_line = f"적재 특허(ChromaDB): {safe_count(collection)}건"
        if master_n is not None:
            count_line += f" | 마스터 엑셀: {master_n}건"
        st.caption(f"{mode_label} | 접속 계정: {st.session_state.user_id} | {count_line}")
    with col_logout:
        if st.button("🔒 로그아웃"):
            st.session_state.logged_in   = False
            st.session_state.user_id     = None
            st.session_state.is_admin    = False
            st.rerun()

    # ── 사이드바: 관리자 전용 데이터 관리 센터 ──
    with st.sidebar:
        if is_admin:
            st.header("📂 데이터 관리 센터")

            # GitHub 수동 재동기화 버튼
            if st.button("🔄 GitHub 데이터 강제 재동기화", use_container_width=True):
                with st.spinner("GitHub → 로컬 전체 동기화 중..."):
                    excel_ok, reg_ok = sync_all_from_github()
                if excel_ok is True or reg_ok is True:
                    # 하나 이상 성공적으로 다운로드됨
                    st.session_state.github_synced = False
                    st.toast("✅ 동기화 완료 — 최신 데이터가 복원됩니다.")
                    st.rerun()
                elif excel_ok is False or reg_ok is False:
                    # 실제 연결/인증 오류
                    st.warning("⚠️ GitHub 동기화 실패 — 아래 '🔍 GitHub 연결 진단' 버튼으로 원인을 확인하세요.")
                else:
                    # 둘 다 None → GitHub 연결은 정상이나 파일이 없는 상태 (포맷 후 등)
                    st.info("ℹ️ GitHub에 저장된 데이터가 없습니다. "
                            "Excel 파일을 업로드하면 자동으로 GitHub에 백업됩니다.")

            # GitHub 연결 진단 버튼
            if st.button("🔍 GitHub 연결 진단", use_container_width=True):
                with st.spinner("진단 중..."):
                    diag = diagnose_github()

                if diag["repo_accessible"]:
                    st.success(f"✅ GitHub 연결 정상\n\n"
                               f"- 레포: `{diag.get('repo_name','')}`\n"
                               f"- 토큰: `{diag['token_prefix']}`")
                else:
                    st.error(f"❌ 연결 실패\n\n**원인:**\n{diag['error']}")
                    if diag["secret_ok"] and not diag["api_reachable"]:
                        st.info("💡 Streamlit Cloud 네트워크 문제일 수 있습니다. 잠시 후 재시도해 주세요.")
                    elif diag["secret_ok"]:
                        st.info("💡 Streamlit Cloud → **Manage app → Secrets**에서 토큰을 새로 발급한 값으로 교체해 주세요.")
                    else:
                        st.code(
                            "# Secrets 올바른 구조 (섹션 헤더 위에 위치)\n"
                            'GITHUB_TOKEN = "ghp_새토큰값"\n'
                            'GITHUB_REPO_URL = "https://github.com/ssahaga97-max/my-patent-rag.git"\n\n'
                            "[USER_CREDENTIALS]\n"
                            'admin = "1234!"',
                            language="toml"
                        )
            st.divider()

            uploaded_file = st.file_uploader("경쟁사 특허 엑셀 리스트 업로드 (.xlsx)", type=["xlsx"])
            if uploaded_file is not None:
                if st.button("🚀 신규 특허 무결성 적재"):
                    try:
                        preview_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        preview_df    = pd.read_excel(io.BytesIO(preview_bytes), nrows=3)
                        detected_cols = _detect_columns(preview_df)
                        total_rows    = pd.read_excel(io.BytesIO(preview_bytes)).shape[0]
                        uploaded_file.seek(0)

                        with st.expander("📋 업로드 파일 열 감지 결과 (클릭 확인)", expanded=True):
                            st.write(f"- **전체 행 수:** {total_rows}행")
                            st.write(f"- **감지된 출원번호 열:** `{detected_cols['id']}`")
                            st.write(f"- **감지된 명칭 열:** `{detected_cols['title']}`")
                            st.write(f"- **전체 열 목록:** {list(preview_df.columns)}")
                    except Exception as diag_e:
                        st.warning(f"파일 사전 진단 실패: {diag_e}")

                    with st.spinner("canonical 출원번호 정규화 및 ChromaDB 적재 중..."):
                        chroma_new, master_new, dup, skipped, synced = process_and_update_db(
                            uploaded_file, collection
                        )

                    total_now = safe_count(collection)
                    master_now = _count_master_excel_patents()
                    summary = (
                        f"📊 처리 결과: Chroma 신규 **{chroma_new}건** / 마스터 신규 **{master_new}건** / "
                        f"이미 색인됨 **{dup}건** / ChromaDB 총 **{total_now}건**"
                    )
                    if master_now is not None:
                        summary += f" / 마스터 엑셀 **{master_now}건**"
                    if skipped:
                        summary += f" / 출원번호 없음 **{skipped}건** 스킵"
                    if synced:
                        summary += f" / 마스터→Chroma 추가복구 **{synced}건**"

                    github_ok = None
                    if chroma_new > 0 or master_new > 0 or synced > 0:
                        with st.spinner("💾 GitHub 데이터 웨어하우스 영구 동기화 중... (대용량 파일은 최대 2분 소요)"):
                            github_ok = commit_and_push_data()
                        if github_ok:
                            st.session_state.upload_feedback = {
                                "level": "success",
                                "message": (
                                    f"{summary}\n\n"
                                    f"✅ 인덱싱 및 GitHub 백업 성공 — 재시작 후에도 데이터가 보존됩니다."
                                ),
                            }
                            st.toast(f"✅ 적재 완료 (ChromaDB {total_now}건)")
                        else:
                            st.session_state.upload_feedback = {
                                "level": "warning",
                                "message": (
                                    f"{summary}\n\n"
                                    f"⚠️ ChromaDB 인덱싱은 완료되었으나 GitHub 백업 실패. "
                                    f"'💾 GitHub 마스터 백업 재시도' 버튼으로 다시 업로드하세요."
                                ),
                            }
                    elif dup > 0:
                        st.session_state.upload_feedback = {
                            "level": "info",
                            "message": (
                                f"{summary}\n\n"
                                f"ℹ️ 업로드 파일의 특허는 이미 ChromaDB에 색인되어 있습니다. "
                                f"ChromaDB({total_now})와 마스터({master_now}) 건수가 다르면 "
                                f"아래 **ChromaDB 누락분 복구** 버튼을 실행하세요."
                            ),
                        }
                    else:
                        st.session_state.upload_feedback = {
                            "level": "error",
                            "message": (
                                f"{summary}\n\n"
                                f"❌ 처리된 데이터가 없습니다. 열 감지 결과에서 '출원번호' 열이 올바른지 확인하세요."
                            ),
                        }

            if st.session_state.upload_feedback:
                _render_feedback_box(st.session_state.upload_feedback)

            st.divider()
            chroma_n       = safe_count(collection)
            master_n       = _count_master_excel_patents()
            unique_chroma_n = _count_chroma_unique_ids(collection)
            st.markdown(_build_db_count_status(chroma_n, master_n, unique_chroma_n))

            if st.session_state.sync_feedback:
                _render_feedback_box(st.session_state.sync_feedback)

            needs_reverse_sync = (
                master_n is not None
                and unique_chroma_n > 0
                and master_n < unique_chroma_n
            )
            if needs_reverse_sync:
                if st.button("📥 ChromaDB → 마스터 엑셀 역동기화", use_container_width=True):
                    with st.spinner(
                        f"ChromaDB {chroma_n}건을 마스터 엑셀로 보내는 중 "
                        f"(재임베딩 없음, 약 1~3분)..."
                    ):
                        exported, before_n, after_n = sync_master_from_chroma(collection)
                    if exported > 0:
                        st.session_state.sync_feedback = {
                            "level": "success",
                            "message": (
                                f"✅ 역동기화 완료: Chroma **{exported}건** 반영 → "
                                f"마스터 **{before_n} → {after_n}건** (ChromaDB는 그대로 유지)  \n"
                                f"💾 **GitHub 마스터 백업**을 실행해 영속화하세요."
                            ),
                        }
                        st.rerun()
                    else:
                        st.session_state.sync_feedback = {
                            "level": "error",
                            "message": "❌ 역동기화 실패 — ChromaDB 데이터를 읽지 못했습니다.",
                        }
                        st.rerun()

            if master_n is not None and master_n > chroma_n:
                if st.button("🔧 ChromaDB 누락분 복구 (마스터 엑셀 기준)", use_container_width=True):
                    gap = master_n - chroma_n
                    with st.spinner(f"마스터 엑셀 → ChromaDB 누락분 복구 중 (약 {gap}건)..."):
                        synced = sync_chroma_missing_from_master(collection)
                    st.session_state._last_chroma_count = safe_count(collection)
                    if synced > 0:
                        st.session_state.sync_feedback = {
                            "level": "success",
                            "message": (
                                f"✅ ChromaDB 누락분 **{synced}건** 복구 완료 "
                                f"(현재 {safe_count(collection)}건). GitHub 백업을 권장합니다."
                            ),
                        }
                    else:
                        st.session_state.sync_feedback = {
                            "level": "info",
                            "message": "복구할 누락분이 없거나 이미 동기화되어 있습니다.",
                        }
                    st.rerun()

            if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0:
                file_kb = os.path.getsize(MASTER_EXCEL_PATH) // 1024
                st.caption(f"로컬 마스터: {file_kb} KB")
                if st.button("💾 GitHub 마스터 백업 재시도", use_container_width=True):
                    with st.spinner("GitHub에 마스터 엑셀 업로드 중... (최대 2분)"):
                        ok = upload_file_to_github_api(
                            MASTER_EXCEL_PATH, "my_patent_folder/master_patents.xlsx"
                        )
                    if ok:
                        master_now = _count_master_excel_patents()
                        st.session_state.sync_feedback = {
                            "level": "success",
                            "message": (
                                f"✅ GitHub 백업 성공! (마스터 엑셀 **{master_now}건** → GitHub 저장 완료)"
                            ),
                        }
                    else:
                        st.session_state.sync_feedback = {
                            "level": "error",
                            "message": (
                                "❌ 백업 실패. **🔍 GitHub 연결 진단**으로 원인 확인 후 재시도하세요.  \n"
                                "토큰 `repo` 쓰기 권한·만료 여부를 확인하세요."
                            ),
                        }
                    st.rerun()
            else:
                st.caption("로컬 마스터 파일 없음 (업로드 후 활성화)")

            also_clear_github = st.checkbox(
                "GitHub 백업도 함께 초기화 (master_patents.xlsx 삭제)",
                value=False,
                help="체크 시 GitHub에 저장된 master_patents.xlsx도 삭제합니다. "
                     "원본 Excel을 다시 업로드하여 완전히 새로 시작할 때 사용하세요."
            )
            if st.button("🚨 가상 데이터 웨어하우스 전체 포맷"):
                with st.spinner("⏳ 벡터 DB 및 마스터 데이터 완전 초기화 중..."):
                    try:
                        # 1. 로컬 마스터 엑셀 삭제
                        if os.path.exists(MASTER_EXCEL_PATH):
                            os.remove(MASTER_EXCEL_PATH)

                        # 2. (옵션) GitHub 백업도 삭제
                        github_cleared = False
                        if also_clear_github:
                            github_cleared = delete_file_from_github_api(
                                "my_patent_folder/master_patents.xlsx"
                            )
                            if not github_cleared:
                                st.warning("GitHub 삭제 실패 — GitHub 연결 진단 후 재시도하세요.")

                        reset_collection()

                        st.session_state.sync_feedback = None
                        st.session_state.upload_feedback = None

                        # github_synced = True: 포맷 직후 리런에서 GitHub 재다운로드 방지
                        # (이전에 False로 설정 시 GitHub의 기존 데이터가 즉시 복원되는 문제 해결)
                        # 컨테이너 재시작 시에는 session_state가 초기화되므로 정상적으로 GitHub 동기화됨
                        st.session_state.github_synced = True

                        if also_clear_github and github_cleared:
                            st.toast("✅ 로컬 + GitHub 데이터 완전 초기화 완료. 원본 Excel을 새로 업로드해 주세요.")
                        elif also_clear_github and not github_cleared:
                            st.toast("⚠️ 로컬 초기화 완료. GitHub 삭제 실패 — GitHub 연결 진단 후 재시도하세요.")
                        else:
                            st.toast("✅ 로컬 초기화 완료. 원본 Excel을 업로드하면 GitHub 이전 데이터와 무관하게 새로 적재됩니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"초기화 중 오류 발생: {e}")

            st.divider()
            st.subheader("👥 가입 회원 현황")
            registry = load_user_registry()
            if not registry:
                st.caption("등록된 일반 회원 없음")
            else:
                # 변경 상태를 session_state에 누적 후 한 번에 저장
                if "registry_dirty" not in st.session_state:
                    st.session_state.registry_dirty = False

                updated_registry = dict(registry)
                for uid, info in registry.items():
                    is_active   = info.get("active", True)
                    status_icon = "🟢" if is_active else "🔴"
                    btn_label   = "비활성화" if is_active else "활성화"
                    btn_type    = "secondary" if is_active else "primary"

                    col_info, col_btn = st.columns([3, 1])
                    with col_info:
                        st.markdown(
                            f"{status_icon} **{uid}**  \n"
                            f"<span style='font-size:12px;color:gray'>"
                            f"{info.get('name','')} · {info.get('department','부서없음')} · "
                            f"{info.get('registered_at','')[:10]}</span>",
                            unsafe_allow_html=True,
                        )
                    with col_btn:
                        if st.button(btn_label, key=f"toggle_{uid}", type=btn_type):
                            updated_registry[uid]["active"] = not is_active
                            save_user_registry(updated_registry)
                            upload_user_registry_to_github()
                            action = "활성화" if not is_active else "비활성화"
                            st.toast(f"✅ {uid} 계정을 {action}했습니다.")
                            st.rerun()
                    st.divider()
        else:
            st.caption("분석 기능 전용 접속 모드입니다.")

    st.subheader("⚙️ 1단계: AI 전문가 선택")
    analysis_mode = st.selectbox(
        "사용 목적에 맞는 전문가 관점을 선택해 주세요:",
        [
            "💡 단순 키워드 매칭 및 특허 검색",
            "🔬 특정 기술 관련 심층 특허 분석",
            "🛡 개발기술 침해 분석 & 진보성 회피 설계",
            "📊 출원정보 기반 다차원 통계조사 (출원인, 발명자, IPC, 일자 등)"
        ]
    )

    st.subheader("🔍 2단계: 검색 키워드 또는 질의 내용 입력")
    placeholders = {
        "💡 단순 키워드 매칭 및 특허 검색": "검색하고자 하는 핵심 키워드들을 입력하세요. (예: 카세트 도어 잠금장치)",
        "🔬 특정 기술 관련 심층 특허 분석": "동향을 파악할 타겟 기술이나 모듈명을 입력하세요. (예: 센서 기반 매체 지폐 잼 장애 예측 알고리즘)",
        "🛡 개발기술 침해 분석 & 진보성 회피 설계": "우리가 출원 예정이거나 개발한 기술 아이디어를 청구항 수준으로 상세히 입력하세요.",
        "📊 출원정보 기반 다차원 통계조사 (출원인, 발명자, IPC, 일자 등)": (
            "조건·질의 예: '2021년 이후 출원된 특허의 출원인별로 각각 중점적으로 개발한 기술 분석' "
            "(연도·출원인별 코호트 전체 집계 — 10건 제한 없음)"
        )
    }
    user_query = st.text_area("분석 대상 내용을 입력하세요:", height=110, placeholder=placeholders[analysis_mode])

    if "📊" not in analysis_mode:
        default_n = 7 if "🛡" in analysis_mode else 5
        n_results_user = st.slider(
            "🔢 3단계: 참조할 관련 특허 수",
            min_value=3, max_value=20, value=default_n, step=1,
            help="AI가 분석에 참조할 최대 특허 건수입니다. Groq API 입력 한도(6,000 TPM) 때문에 수가 많으면 본문이 자동 축약됩니다."
        )
        st.caption(
            "💡 **출원연도·출원인별 전체 동향** 분석(예: '2021년 이후 출원인별 기술')은 "
            "위 **📊 통계조사 모드**를 선택하세요. 10~20건 제한 없이 DB 전체를 집계합니다."
        )
    else:
        n_results_user = 10  # 통계 모드는 슬라이더 불필요 (전체 데이터 집계)

    if st.button("🧬 가상 전문가 엔진 구동"):
        if user_query.strip() == "":
            st.warning("분석 내용을 입력해 주세요.")
        elif safe_count(collection) == 0:
            st.error("서버 DB에 적재된 특허 소스가 없습니다. 좌측 메뉴에서 엑셀을 먼저 등록해 주세요.")
        else:
            with st.spinner("가상 전문가가 실시간 시맨틱 문헌 대조 및 클라우드 초고속 추론을 진행 중입니다..."):
                context_truncated = False

                # ── 통계 모드: 전체 메타데이터 pandas 집계 후 요약 컨텍스트 구성 ──
                if "📊" in analysis_mode:
                    all_data = collection.get(include=["metadatas"])
                    all_metas = all_data.get("metadatas", [])
                    total_count = len(all_metas)

                    if total_count == 0:
                        st.error("서버 DB에 적재된 특허 소스가 없습니다.")
                        st.stop()

                    stats_df = pd.DataFrame(all_metas)

                    min_year, max_year = _extract_year_filter_from_query(user_query)
                    use_year_filter = min_year is not None or max_year is not None
                    analysis_df = (
                        _filter_metadata_df_by_year(stats_df, min_year, max_year)
                        if use_year_filter else stats_df
                    )
                    filtered_count = len(analysis_df)
                    cohort_mode = _wants_applicant_cohort(user_query) or (
                        use_year_filter and filtered_count > 0
                    )

                    if use_year_filter:
                        yr_label = f"{min_year or '…'}년 ~ {max_year or '현재'}"
                        st.info(
                            f"📅 질의에서 추출한 출원일 조건: **{yr_label}** → "
                            f"**{filtered_count}건** / 전체 {total_count}건"
                        )
                        if filtered_count == 0:
                            st.warning("조건에 맞는 특허가 없습니다. 연도 표현을 확인해 주세요.")
                            st.stop()

                    applicant_df, applicant_stat = _top_counts_table(
                        analysis_df.get("출원인", pd.Series(dtype=str)),
                        n=15,
                        label="출원인(대표명)",
                        normalizer=_canonical_applicant_name,
                        explode=True,
                    )
                    ipc_df, ipc_stat = _top_counts_table(
                        analysis_df.get("IPC", pd.Series(dtype=str)),
                        n=10,
                        label="IPC",
                        explode=True,
                        allow_comma=True,
                    )
                    inventor_df, inventor_stat = _top_counts_table(
                        analysis_df.get("발명자", pd.Series(dtype=str)),
                        n=10,
                        label="발명자",
                        explode=True,
                        allow_comma=True,
                    )
                    year_df, year_stat = _year_counts_table(
                        analysis_df.get("출원일", pd.Series(dtype=str))
                    )

                    scope_label = (
                        f"조건 필터 ({filtered_count}건)" if use_year_filter else "전체 DB"
                    )
                    st.markdown(f"### 📊 사전 집계 통계 ({scope_label})")
                    stat_col1, stat_col2 = st.columns(2)
                    with stat_col1:
                        st.markdown("**출원인별 (대표명 통합)**")
                        st.dataframe(applicant_df, use_container_width=True, hide_index=True)
                        st.markdown("**IPC 분류별**")
                        st.dataframe(ipc_df, use_container_width=True, hide_index=True)
                    with stat_col2:
                        st.markdown("**발명자별**")
                        st.dataframe(inventor_df, use_container_width=True, hide_index=True)
                        st.markdown("**출원 연도별**")
                        st.dataframe(year_df, use_container_width=True, hide_index=True)

                    cohort_md = ""
                    if cohort_mode:
                        cohort_md, cohort_df = _build_applicant_cohort_context(analysis_df)
                        st.markdown("### 🏢 출원인별 기술 코호트 (전체 집계 · 10건 제한 없음)")
                        st.caption(
                            "각 출원인의 필터 범위 내 **전체 특허명·IPC**를 집계했습니다. "
                            "AI는 아래 코호트 전체를 기준으로 기술 동향을 분석합니다."
                        )
                        st.dataframe(cohort_df, use_container_width=True, hide_index=True)

                    filter_note = ""
                    if use_year_filter:
                        filter_note = (
                            f"※ 분석 대상: 출원일 {min_year or '…'}년~{max_year or '현재'} "
                            f"({filtered_count}건 / 전체 {total_count}건)\n"
                        )

                    context_text = f"""[DB 통계 요약] {filter_note}분석 범위 내 {filtered_count if use_year_filter else total_count}건
※ 출원인은 경쟁사 대표명으로 통합. 발명자·IPC는 복수값 분리 후 집계.

■ 출원인별 상위 현황 (대표명)
{applicant_stat}

■ IPC 분류별 상위 10 현황
{ipc_stat}

■ 주요 발명자 상위 10 현황
{inventor_stat}

■ 출원 연도별 건수 추이
{year_stat}
"""
                    if cohort_md:
                        context_text += (
                            f"\n■ 출원인별 기술 코호트 (필터 범위 전체 — 샘플 10건 아님)\n"
                            f"{cohort_md}\n"
                        )
                    elif not cohort_mode:
                        sem_results = collection.query(
                            query_texts=[user_query.strip()],
                            n_results=min(filtered_count if use_year_filter else total_count, 10),
                        )
                        if sem_results and sem_results["documents"][0]:
                            context_text += "\n■ 질의 관련 시맨틱 매칭 상위 특허\n"
                            for i, m in enumerate(sem_results["metadatas"][0]):
                                context_text += (
                                    f"  [{i+1}] {m.get('출원번호','')} | "
                                    f"{m.get('명칭','')} | {m.get('출원인','')}\n"
                                )

                    system_prompt = (
                        "당신은 특허 데이터 통계 전문 분석가입니다. "
                        "아래 집계 표의 숫자(건수)는 이미 확정된 값이므로 변경·재계산·추정하지 말고 그대로 인용하세요. "
                        "출원인은 경쟁사 대표명으로 통합된 결과입니다. "
                    )
                    if cohort_md:
                        system_prompt += (
                            "'출원인별 기술 코호트' 섹션에 각 출원인의 **필터 범위 내 전체 특허명·IPC**가 포함되어 있습니다. "
                            "10건 샘플이 아닌 코호트 전체를 근거로, 출원인별로 중점 개발 기술·기술 포트폴리오·IPC 기반 기술 분야를 "
                            "구체적으로 비교 분석하세요. "
                        )
                    system_prompt += (
                        "마크다운 표를 작성할 때 반드시 '항목/출원인/발명자/IPC/연도' 열과 '건수' 열을 구분하고, "
                        "건수 열에는 숫자만 넣으세요. "
                        "집계 통계를 기반으로 출원 동향, 핵심 출원인, 기술 분야 분포를 "
                        "체계적인 다차원 통계 리포트로 작성하세요."
                    )

                # ── 일반 모드: 시맨틱 RAG 검색 ──
                else:
                    n_results = min(n_results_user, safe_count(collection))

                    results = collection.query(
                        query_texts=[user_query.strip()],
                        n_results=n_results
                    )

                    if not (results and results["documents"] and results["documents"][0]):
                        st.error("관련 특허를 찾지 못했습니다. 다른 키워드로 시도해 보세요.")
                        st.stop()

                    retrieved_docs  = results["documents"][0]
                    retrieved_metas = results["metadatas"][0]

                    if "💡 단순 키워드" in analysis_mode:
                        system_prompt = (
                            "당신은 신속하고 정확하게 관련 문헌을 찾아내는 '수석 특허 검색 조사관'입니다. "
                            "관련 특허를 마크다운 링크 서식과 함께 요약 브리핑하세요."
                        )
                    elif "🔬 특정 기술" in analysis_mode:
                        system_prompt = (
                            "당신은 수석 기술 전문 분석가입니다. "
                            "마크다운 링크를 포함한 기술 동향 보고서를 체계적으로 작성하세요."
                        )
                    else:
                        system_prompt = (
                            "당신은 특허청 수석 심사관 및 특허법률 전문가 집단입니다. "
                            "관련 선행문헌들의 링크 주소를 명시하며 침해 가능성 및 "
                            "회피설계 가이드를 구성요소 완비 법칙에 근거하여 작성하세요."
                        )

                    context_budget = _max_prompt_input_tokens(system_prompt, user_query)
                    claims_first = "🛡" in analysis_mode
                    context_text, context_truncated = _build_rag_context_for_llm(
                        retrieved_docs, retrieved_metas, context_budget, claims_first=claims_first
                    )

                prompt, prompt_truncated, context_text = _assemble_llm_prompt(
                    system_prompt, context_text, user_query
                )
                truncated = context_truncated or prompt_truncated

                try:
                    if truncated:
                        if "🛡" in analysis_mode:
                            trunc_msg = (
                                "Groq API 입력 한도(6,000 TPM)에 맞추기 위해 참조 특허 본문을 자동 축약했습니다. "
                                "침해 분석 모드는 **청구항을 우선** 유지하고 요약을 먼저 줄입니다. "
                                "더 상세한 분석이 필요하면 '참조할 관련 특허 수'를 줄여 보세요."
                            )
                        else:
                            trunc_msg = (
                                "Groq API 입력 한도(6,000 TPM)에 맞추기 위해 참조 특허 본문을 자동 축약했습니다. "
                                "키워드·심층 분석 모드는 **요약을 우선** 유지하고 청구항을 먼저 생략·축소합니다. "
                                "더 상세한 분석이 필요하면 '참조할 관련 특허 수'를 줄여 보세요."
                            )
                        st.info(trunc_msg)
                    response = llm.invoke(prompt)
                    st.markdown(f"### 📊 AI {analysis_mode.split(' ')[1]} 결과 보고서")
                    st.write(response.content)
                    st.divider()
                    with st.expander("👁 시스템이 매칭한 원천 데이터 (클릭 시 원문 이동 가능)"):
                        st.markdown(context_text)
                except Exception as e:
                    st.error(f"AI 추론 중 오류: {e}")


if __name__ == "__main__":
    if check_authentication():
        run_main_portal()
