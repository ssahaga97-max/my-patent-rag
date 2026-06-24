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
import tempfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from urllib.request import Request, urlopen
from urllib.error import HTTPError

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
    from googleapiclient.errors import HttpError
    from google.cloud import storage as gcs_storage
    _GOOGLE_DRIVE_AVAILABLE = True
except ImportError:
    HttpError = Exception  # type: ignore
    gcs_storage = None  # type: ignore
    _GOOGLE_DRIVE_AVAILABLE = False

# --- 1. 클라우드 서버 전용 절대 경로 고정 및 초기화 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_EXCEL_PATH    = os.path.join(BASE_DIR, "my_patent_folder", "master_patents.xlsx")
USER_REGISTRY_PATH   = os.path.join(BASE_DIR, "my_patent_folder", "user_registry.json")
VECTOR_SNAPSHOT_PATH = os.path.join(BASE_DIR, "my_patent_folder", "vector_snapshot.parquet")
VECTOR_MANIFEST_PATH = os.path.join(BASE_DIR, "my_patent_folder", "vector_manifest.json")
VECTOR_SNAPSHOT_FILENAME = "vector_snapshot.parquet"  # GCS 고정명 — 업로드 시 덮어쓰기(용량 누수 방지)
VECTOR_MANIFEST_FILENAME = "vector_manifest.json"
VECTOR_SNAPSHOT_VERSION  = 1
_EMBED_DIM = 768  # paraphrase-multilingual-mpnet-base-v2
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
if "gap_sync_active" not in st.session_state:
    st.session_state.gap_sync_active = False
if "gap_sync_synced_session" not in st.session_state:
    st.session_state.gap_sync_synced_session = 0

_PATENT_COUNT_CACHE_KEY = "_patent_count_cache"
# Streamlit 연결 타임아웃 방지 — 누락분 복구 시 1회 실행당 임베딩 건수
_GAP_SYNC_CHUNK = 15


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
# Google Drive API (벡터 스냅샷 영속화)
# 개인 Gmail + 서비스 계정은 storageQuotaExceeded(403)가 자주 발생 → GCS 버킷 권장
# ==========================================
_GCP_SCOPES = (
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/devstorage.read_write",
)


def _drive_configured() -> bool:
    """Secrets에 Drive 폴더 ID + gcp_service_account가 설정되어 있는지."""
    if not _GOOGLE_DRIVE_AVAILABLE:
        return False
    try:
        folder_id = _get_google_drive_folder_id()
        return bool(folder_id and "gcp_service_account" in st.secrets)
    except Exception:
        return False


def _parse_google_drive_folder_id(value: str) -> str | None:
    """
    Drive 폴더 ID 추출. Secrets에 URL 전체를 넣어도 ID만 사용.
    예: https://drive.google.com/drive/folders/1abc... → 1abc...
    """
    raw = _clean_ascii(str(value or "")).strip()
    if not raw:
        return None
    if "/" not in raw and "?" not in raw:
        return raw
    m = re.search(r"/folders/([^/?&#]+)", raw)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([^&]+)", raw)
    if m:
        return m.group(1)
    return raw


def _get_google_drive_folder_id() -> str | None:
    try:
        return _parse_google_drive_folder_id(st.secrets.get("GOOGLE_DRIVE_FOLDER_ID", ""))
    except Exception:
        return None


def _get_gcp_credentials():
    if not _GOOGLE_DRIVE_AVAILABLE:
        return None
    try:
        sa_info = dict(st.secrets["gcp_service_account"])
        pk = sa_info.get("private_key", "")
        if isinstance(pk, str) and "\\n" in pk and "-----BEGIN" in pk:
            sa_info["private_key"] = pk.replace("\\n", "\n")
        creds = service_account.Credentials.from_service_account_info(
            sa_info, scopes=_GCP_SCOPES
        )
        # Workspace 도메인 위임(선택) — 개인 Gmail은 미지원
        delegate = _clean_ascii(str(st.secrets.get("GOOGLE_DRIVE_DELEGATE_EMAIL", "")))
        if delegate:
            creds = creds.with_subject(delegate)
        return creds
    except Exception as e:
        print(f"[GCP credentials 오류] {e}")
        return None


def _drive_format_error(exc: Exception) -> str:
    """Drive API HttpError → 사용자용 메시지."""
    if isinstance(exc, HttpError):
        detail = ""
        try:
            body = json.loads(exc.content.decode()) if exc.content else {}
            for err in body.get("error", {}).get("errors", []):
                detail = err.get("reason", "") or err.get("message", "")
                if detail:
                    break
            if not detail:
                detail = body.get("error", {}).get("message", "")
        except Exception:
            detail = ""
        base = f"HTTP {exc.resp.status}"
        if detail:
            return f"{base} ({detail})"
        return base
    return str(exc)[:400]


def _drive_folder_meta(service, folder_id: str) -> dict:
    return service.files().get(
        fileId=folder_id,
        fields="id,name,mimeType,capabilities,driveId,owners",
        supportsAllDrives=True,
    ).execute()


def _drive_sa_permission_role(service, folder_id: str, client_email: str) -> str:
    """폴더 permissions 목록에서 서비스 계정 역할 조회."""
    if not client_email:
        return "unknown"
    try:
        resp = service.permissions().list(
            fileId=folder_id,
            fields="permissions(emailAddress,role,type)",
            supportsAllDrives=True,
        ).execute()
        for perm in resp.get("permissions", []):
            if perm.get("emailAddress", "").lower() == client_email.lower():
                return perm.get("role", "unknown")
    except Exception:
        pass
    return "not_listed"


def _drive_two_step_create(
    service,
    folder_id: str,
    drive_filename: str,
    local_file_path: str,
    *,
    shared_drive: bool,
) -> None:
    """메타데이터만 먼저 생성 후 본문 업로드 — storageQuotaExceeded 우회 시도."""
    body = {"name": drive_filename, "parents": [folder_id]}
    create_kw: dict = {"body": body, "fields": "id"}
    if shared_drive:
        create_kw["supportsAllDrives"] = True
    created = service.files().create(**create_kw).execute()
    file_id = created["id"]

    size = os.path.getsize(local_file_path)
    media = MediaFileUpload(
        local_file_path,
        mimetype="application/octet-stream",
        resumable=size > 5 * 1024 * 1024,
    )
    update_kw: dict = {"fileId": file_id, "media_body": media}
    if shared_drive:
        update_kw["supportsAllDrives"] = True
    service.files().update(**update_kw).execute()


def _drive_create_or_update_file(
    service,
    folder_id: str,
    drive_filename: str,
    local_file_path: str,
    file_id: str | None,
    *,
    shared_drive: bool,
) -> None:
    """My Drive 공유 폴더 / 공유 드라이브 모두 대응하는 업로드."""
    size = os.path.getsize(local_file_path)
    media = MediaFileUpload(
        local_file_path,
        mimetype="application/octet-stream",
        resumable=size > 5 * 1024 * 1024,
    )
    strategies: list[bool] = [True] if shared_drive else [False, True]

    last_exc: Exception | None = None
    for use_shared_drive_flag in strategies:
        try:
            if file_id:
                kwargs: dict = {"fileId": file_id, "media_body": media}
                if use_shared_drive_flag:
                    kwargs["supportsAllDrives"] = True
                service.files().update(**kwargs).execute()
            else:
                kwargs = {
                    "body": {"name": drive_filename, "parents": [folder_id]},
                    "media_body": media,
                    "fields": "id",
                }
                if use_shared_drive_flag:
                    kwargs["supportsAllDrives"] = True
                service.files().create(**kwargs).execute()
            return
        except Exception as e:
            last_exc = e
            err = _drive_format_error(e)
            if (
                not file_id
                and ("storageQuota" in err or "storage quota" in err.lower())
            ):
                try:
                    _drive_two_step_create(
                        service, folder_id, drive_filename, local_file_path,
                        shared_drive=use_shared_drive_flag,
                    )
                    return
                except Exception as e2:
                    last_exc = e2
                    continue
            continue
    if last_exc:
        raise last_exc


@st.cache_resource
def _get_drive_service():
    creds = _get_gcp_credentials()
    if not creds:
        return None
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _find_drive_file_id(service, folder_id: str, filename: str) -> str | None:
    safe_name = filename.replace("'", "\\'")
    query = (
        f"name='{safe_name}' and '{folder_id}' in parents and trashed=false"
    )
    resp = service.files().list(
        q=query,
        fields="files(id,name)",
        pageSize=1,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def upload_file_to_drive(local_file_path: str, drive_filename: str) -> tuple[bool, str]:
    """로컬 파일을 Drive 폴더에 업로드(동일 이름이면 갱신). (성공 여부, 메시지)"""
    service = _get_drive_service()
    folder_id = _get_google_drive_folder_id()
    if not service:
        return False, "Drive API 서비스 초기화 실패 (gcp_service_account 확인)"
    if not folder_id:
        return False, "GOOGLE_DRIVE_FOLDER_ID 없음"
    if not os.path.exists(local_file_path):
        return False, f"로컬 파일 없음: {local_file_path}"
    try:
        sa_email = ""
        try:
            sa_email = st.secrets["gcp_service_account"].get("client_email", "")
        except Exception:
            pass

        folder_meta = _drive_folder_meta(service, folder_id)
        caps = folder_meta.get("capabilities", {})
        shared_drive = bool(folder_meta.get("driveId"))

        if not caps.get("canAddChildren", False):
            return False, (
                "폴더에 파일 추가 권한 없음 (canAddChildren=false). "
                f"Drive → 폴더 공유 → 아래 이메일을 **편집자**로 추가하세요:\n"
                f"`{sa_email}`"
            )

        file_id = _find_drive_file_id(service, folder_id, drive_filename)
        _drive_create_or_update_file(
            service, folder_id, drive_filename, local_file_path, file_id,
            shared_drive=shared_drive,
        )
        size_kb = os.path.getsize(local_file_path) // 1024
        return True, f"{drive_filename} ({size_kb} KB)"
    except Exception as e:
        err = _drive_format_error(e)
        print(f"[Drive upload 오류] {drive_filename}: {e}")
        sa_email = ""
        try:
            sa_email = st.secrets["gcp_service_account"].get("client_email", "")
        except Exception:
            pass
        if "storageQuota" in err or "storage quota" in err.lower():
            return False, (
                f"{drive_filename} 업로드 실패: {err}\n"
                "서비스 계정은 저장 용량이 없습니다. **내 Drive 폴더**를 서비스 계정에 "
                f"**편집자**로 공유해야 합니다 (공유 대상: `{sa_email}`). "
                "공유 드라이브(팀 드라이브) 멤버로 추가하는 방법도 있습니다."
            )
        if "403" in err or "forbidden" in err.lower():
            return False, (
                f"{drive_filename} 업로드 실패: {err}\n"
                f"폴더 공유 대상 이메일이 정확한지 확인: `{sa_email}` (편집자)"
            )
        return False, f"{drive_filename} 업로드 실패: {err}"


def download_file_from_drive(drive_filename: str, local_path: str):
    """
    Drive 폴더에서 파일 다운로드.
    반환: True=성공, None=파일 없음, False=오류
    """
    service = _get_drive_service()
    folder_id = _get_google_drive_folder_id()
    if not service or not folder_id:
        return False
    try:
        file_id = _find_drive_file_id(service, folder_id, drive_filename)
        if not file_id:
            return None
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        with open(local_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return True
    except Exception as e:
        print(f"[Drive download 오류] {drive_filename}: {e}")
        return False


def diagnose_google_drive() -> dict:
    """Google Drive 연결·스냅샷 파일 존재 여부 진단."""
    result = {
        "configured": False,
        "library_ok": _GOOGLE_DRIVE_AVAILABLE,
        "folder_id": "",
        "client_email": "",
        "service_ok": False,
        "folder_accessible": False,
        "manifest_on_drive": False,
        "snapshot_on_drive": False,
        "write_test_ok": False,
        "folder_name": "",
        "can_add_children": False,
        "sa_permission_role": "",
        "is_shared_drive": False,
        "write_test_detail": "",
        "error": "",
    }
    if not _GOOGLE_DRIVE_AVAILABLE:
        result["error"] = (
            "google-api-python-client / google-auth 패키지가 설치되지 않았습니다."
        )
        return result
    folder_id = _get_google_drive_folder_id()
    if not folder_id:
        result["error"] = (
            "GOOGLE_DRIVE_FOLDER_ID가 Streamlit Secrets에 없습니다."
        )
        return result
    result["configured"] = True
    result["folder_id"] = folder_id[:8] + "..."

    try:
        sa_info = dict(st.secrets.get("gcp_service_account", {}))
        result["client_email"] = sa_info.get("client_email", "")
    except Exception:
        pass

    service = _get_drive_service()
    if not service:
        result["error"] = (
            "[gcp_service_account] Secrets 구조 또는 private_key 형식을 확인하세요."
        )
        return result
    result["service_ok"] = True

    try:
        folder_meta = _drive_folder_meta(service, folder_id)
        if folder_meta.get("mimeType") != "application/vnd.google-apps.folder":
            result["error"] = "GOOGLE_DRIVE_FOLDER_ID가 폴더가 아닙니다 (파일 ID일 수 있음)."
            return result

        caps = folder_meta.get("capabilities", {})
        result["folder_accessible"] = True
        result["folder_name"] = folder_meta.get("name", "")
        result["can_add_children"] = bool(caps.get("canAddChildren", False))
        result["is_shared_drive"] = bool(folder_meta.get("driveId"))
        result["sa_permission_role"] = _drive_sa_permission_role(
            service, folder_id, result["client_email"]
        )

        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id,name)",
            pageSize=20,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        names = {f.get("name") for f in resp.get("files", [])}
        result["manifest_on_drive"] = VECTOR_MANIFEST_FILENAME in names
        result["snapshot_on_drive"] = VECTOR_SNAPSHOT_FILENAME in names

        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        ) as tf:
            tf.write("PatentRAG Drive write test")
            test_path = tf.name
        try:
            ok, msg = upload_file_to_drive(test_path, "_patentrag_write_test.txt")
            result["write_test_ok"] = ok
            result["write_test_detail"] = msg
            if not ok:
                result["error"] = msg
        finally:
            try:
                os.remove(test_path)
            except OSError:
                pass
    except Exception as e:
        err = str(e)
        if "drive.google.com" in str(st.secrets.get("GOOGLE_DRIVE_FOLDER_ID", "")):
            result["error"] = (
                "GOOGLE_DRIVE_FOLDER_ID에 폴더 URL 전체가 들어가 있습니다. "
                "ID만 넣으세요 (예: 1exsHPZIbCGooQdB3EWtdzkklYRGjLpf1). "
                "또는 app.py 최신 버전은 URL도 자동 변환합니다."
            )
        else:
            result["error"] = (
                f"Drive 폴더 접근 실패 — 서비스 계정을 폴더 '편집자'로 공유했는지, "
                f"GOOGLE_DRIVE_FOLDER_ID가 올바른지 확인: {e}"
            )
    return result


# ==========================================
# Google Cloud Storage (벡터 스냅샷 — 서비스 계정 권장 저장소)
# ==========================================
def _get_gcs_bucket_name() -> str | None:
    try:
        name = _clean_ascii(str(st.secrets.get("GCS_BUCKET_NAME", "")))
        return name if name else None
    except Exception:
        return None


def _gcs_configured() -> bool:
    if not _GOOGLE_DRIVE_AVAILABLE or gcs_storage is None:
        return False
    try:
        return bool(_get_gcs_bucket_name() and "gcp_service_account" in st.secrets)
    except Exception:
        return False


def _vector_storage_configured() -> bool:
    return _gcs_configured()


def _vector_storage_label() -> str:
    if _gcs_configured():
        return f"GCS (`{_get_gcs_bucket_name()}`)"
    return "미설정"


@st.cache_resource
def _get_gcs_client():
    creds = _get_gcp_credentials()
    if not creds or gcs_storage is None:
        return None
    try:
        project = st.secrets["gcp_service_account"].get("project_id", "")
        return gcs_storage.Client(credentials=creds, project=project or None)
    except Exception as e:
        print(f"[GCS client 오류] {e}")
        return None


def upload_file_to_gcs(local_file_path: str, blob_name: str) -> tuple[bool, str]:
    bucket_name = _get_gcs_bucket_name()
    client = _get_gcs_client()
    if not client or not bucket_name:
        return False, "GCS 미설정 (GCS_BUCKET_NAME)"
    if not os.path.exists(local_file_path):
        return False, f"로컬 파일 없음: {local_file_path}"
    try:
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        # 동일 blob 이름 → 덮어쓰기. 날짜별 누적 없음 (GCS 5GB 무료 한도 보호).
        blob.upload_from_filename(local_file_path)
        size_kb = os.path.getsize(local_file_path) // 1024
        return True, f"gs://{bucket_name}/{blob_name} ({size_kb} KB)"
    except Exception as e:
        print(f"[GCS upload 오류] {blob_name}: {e}")
        return False, (
            f"GCS 업로드 실패: {e}\n"
            "GCP Console → Cloud Storage → 버킷 → **권한** → **액세스 권한 부여** →\n"
            "주 구성원: 서비스 계정 이메일 → 역할: **Storage 관리자** "
            "(또는 Storage 객체 관리자)"
        )


def download_file_from_gcs(blob_name: str, local_path: str):
    """반환: True=성공, None=없음, False=오류"""
    bucket_name = _get_gcs_bucket_name()
    client = _get_gcs_client()
    if not client or not bucket_name:
        return False
    try:
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        if not blob.exists():
            return None
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        blob.download_to_filename(local_path)
        return True
    except Exception as e:
        print(f"[GCS download 오류] {blob_name}: {e}")
        return False


def diagnose_gcs() -> dict:
    result = {
        "configured": False,
        "bucket_name": "",
        "client_ok": False,
        "bucket_accessible": False,
        "manifest_on_gcs": False,
        "snapshot_on_gcs": False,
        "write_test_ok": False,
        "write_test_detail": "",
        "error": "",
    }
    if not _gcs_configured():
        result["error"] = "GCS_BUCKET_NAME 또는 gcp_service_account 미설정"
        return result
    result["configured"] = True
    result["bucket_name"] = _get_gcs_bucket_name() or ""

    client = _get_gcs_client()
    if not client:
        result["error"] = "GCS 클라이언트 초기화 실패"
        return result
    result["client_ok"] = True

    try:
        bucket = client.bucket(result["bucket_name"])
        # bucket.reload()는 storage.buckets.get 권한 필요 → 쓰기 테스트로 대체
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        ) as tf:
            tf.write("PatentRAG GCS write test")
            test_path = tf.name
        try:
            ok, msg = upload_file_to_gcs(test_path, "_patentrag_gcs_write_test.txt")
            result["write_test_ok"] = ok
            result["write_test_detail"] = msg
            result["bucket_accessible"] = ok
            if ok:
                result["manifest_on_gcs"] = bucket.blob(VECTOR_MANIFEST_FILENAME).exists()
                result["snapshot_on_gcs"] = bucket.blob(VECTOR_SNAPSHOT_FILENAME).exists()
            else:
                result["error"] = msg
        finally:
            try:
                os.remove(test_path)
            except OSError:
                pass
    except Exception as e:
        err = str(e)
        if "storage.buckets.get" in err or "does not have" in err:
            sa = ""
            try:
                sa = st.secrets["gcp_service_account"].get("client_email", "")
            except Exception:
                pass
            result["error"] = (
                f"버킷 IAM 권한 부족 — `{sa or '서비스 계정'}`에 "
                f"버킷 `{result['bucket_name']}` 권한이 없습니다.\n\n"
                "GCP Console → Cloud Storage → 해당 버킷 → **권한** → "
                "**액세스 권한 부여** → 역할 **Storage 관리자** "
                "(roles/storage.admin) 또는 **Storage 객체 관리자** "
                "(roles/storage.objectAdmin) 추가."
            )
        else:
            result["error"] = (
                f"GCS 버킷 접근 실패 — 버킷 이름·서비스 계정 Storage 권한 확인: {e}"
            )
    return result


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


def _patent_url_from_row(
    row,
    cols: dict,
    patent_id: str,
    hyperlink_map: dict | None = None,
) -> str:
    """마스터 행·하이퍼링크 맵에서 특허 URL 추출 (유효한 http URL만)."""
    hyperlink_map = hyperlink_map or {}

    if "URL" in getattr(row, "index", []):
        raw_url = row.get("URL")
        if pd.notna(raw_url):
            url = str(raw_url).strip()
            if url.startswith("http"):
                return url

    url = hyperlink_map.get(patent_id, "")
    if url and str(url).startswith("http"):
        return str(url)

    if cols.get("title"):
        clean_title = str(row[cols["title"]]).strip().replace("-", "")
        url = hyperlink_map.get(clean_title, "")
        if url and str(url).startswith("http"):
            return str(url)

    return ""


def _master_excel_mtime() -> float:
    """마스터 엑셀 mtime. 없거나 비어 있으면 -1."""
    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0:
        return os.path.getmtime(MASTER_EXCEL_PATH)
    return -1.0


def _invalidate_patent_count_cache() -> None:
    """마스터·Chroma 건수 캐시 무효화 (데이터 변경 직후 호출)."""
    st.session_state.pop(_PATENT_COUNT_CACHE_KEY, None)


def _get_patent_count_cache() -> dict:
    if _PATENT_COUNT_CACHE_KEY not in st.session_state:
        st.session_state[_PATENT_COUNT_CACHE_KEY] = {}
    return st.session_state[_PATENT_COUNT_CACHE_KEY]


def _count_master_excel_patents(force: bool = False) -> int | None:
    """마스터 엑셀의 고유 출원번호 건수. 파일 없으면 None."""
    mtime = _master_excel_mtime()
    cache = _get_patent_count_cache()
    if not force and cache.get("master_mtime") == mtime and "master_n" in cache:
        return cache["master_n"]

    if mtime < 0:
        cache["master_mtime"] = mtime
        cache["master_n"] = None
        return None

    try:
        df   = pd.read_excel(MASTER_EXCEL_PATH)
        cols = _detect_columns(df)
        ids  = {
            _normalize_patent_id(v)
            for v in df[cols["id"]]
            if _normalize_patent_id(v)
        }
        result = len(ids)
    except Exception as e:
        print(f"[마스터 엑셀 건수 조회 오류] {e}")
        result = None

    cache["master_mtime"] = mtime
    cache["master_n"] = result
    return result


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


def _count_chroma_unique_patents(
    collection, batch_size: int = 500, force: bool = False
) -> int:
    """
    ChromaDB 내 고유 출원번호(canonical) 수.
    마스터 엑셀 _count_master_excel_patents()와 동일 기준 — document id 중복과 무관.
    """
    chroma_n = safe_count(collection)
    cache = _get_patent_count_cache()
    cache_key = (chroma_n, _master_excel_mtime())
    if not force and cache.get("chroma_key") == cache_key and "chroma_unique_n" in cache:
        return cache["chroma_unique_n"]

    try:
        if chroma_n == 0:
            result = 0
        else:
            all_ids = collection.get(include=[])["ids"]
            canonical: set[str] = set()
            for i in range(0, len(all_ids), batch_size):
                batch = collection.get(
                    ids=all_ids[i:i + batch_size],
                    include=["metadatas"],
                )
                for chroma_id, meta in zip(
                    batch.get("ids", []),
                    batch.get("metadatas", []),
                ):
                    meta = meta or {}
                    cid = (
                        _normalize_patent_id(meta.get("출원번호", ""))
                        or _normalize_patent_id(chroma_id)
                    )
                    if cid:
                        canonical.add(cid)
            result = len(canonical)
    except Exception as e:
        print(f"[ChromaDB 고유 출원번호 조회 오류] {e}")
        result = 0

    cache["chroma_key"] = cache_key
    cache["chroma_unique_n"] = result
    return result


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


def _build_db_count_status(chroma_n: int, master_n: int | None, unique_patent_n: int) -> str:
    """Chroma·마스터 건수 표시. 정상 동기화 시 건수만, 불일치 시에만 안내 문구."""
    count_md = f"📊 **누적 적재 데이터 (ChromaDB):** `{chroma_n}` 건"
    if master_n is None:
        return count_md

    count_md += f"  \n📄 **마스터 엑셀 (고유 출원번호):** `{master_n}` 건"

    # 고유 출원번호 일치 = 정상 — 추가 경고·안내 없이 건수만 표시
    if master_n == unique_patent_n:
        return count_md

    # 불일치 시에만 상세·안내 (고유 출원번호 기준 비교)
    if unique_patent_n:
        count_md += f"  \n🔑 **ChromaDB 고유 출원번호:** `{unique_patent_n}` 건"

    if master_n < unique_patent_n:
        gap = unique_patent_n - master_n
        count_md += (
            f"  \n⚠️ 마스터가 Chroma 고유 출원번호보다 **{gap}건** 부족합니다. "
            f"**ChromaDB → 마스터 엑셀 역동기화** 후 GitHub 백업하세요."
        )
    elif master_n > unique_patent_n:
        gap = master_n - unique_patent_n
        count_md += (
            f"  \n⚠️ 마스터가 Chroma 고유 출원번호보다 **{gap}건** 많습니다. "
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
            _invalidate_patent_count_cache()
            return before - after
    except Exception as e:
        print(f"[마스터 compact 오류] {e}")
    return 0


def _collect_missing_patents_from_master(
    collection, hyperlink_map: dict | None = None
) -> tuple[list, list, list]:
    """마스터 엑셀에만 있고 Chroma에 없는 출원번호 → (ids, docs, metas)."""
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return [], [], []
    df = pd.read_excel(MASTER_EXCEL_PATH)
    cols = _detect_columns(df)
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
        patent_url = _patent_url_from_row(row, cols, pat_id, hyperlink_map)
        ids.append(pat_id)
        docs.append(_build_document(row, cols))
        metas.append(_build_metadata(row, cols, patent_url=patent_url))
    return ids, docs, metas


def run_gap_sync_step(collection, chunk_size: int = _GAP_SYNC_CHUNK) -> dict:
    """
    누락분 복구 1스텝(배치). Streamlit rerun 루프와 함께 사용.
    완료 시 GCS 스냅샷 업로드 시도.
    """
    result = {
        "synced_this_run": 0,
        "remaining": 0,
        "total_missing": 0,
        "done": True,
        "snapshot_ok": None,
        "snapshot_msg": "",
    }
    _compact_master_excel()
    try:
        ids, docs, metas = _collect_missing_patents_from_master(collection)
        total = len(ids)
        result["total_missing"] = total
        if total == 0:
            return result

        n = min(chunk_size, total)
        _upsert_patent_batches(
            collection, ids[:n], docs[:n], metas[:n], batch_size=10
        )
        _invalidate_patent_count_cache()
        result["synced_this_run"] = n
        result["remaining"] = total - n
        result["done"] = n >= total
        if result["done"]:
            ok_snap, snap_msg = maybe_upload_vector_snapshot(collection)
            result["snapshot_ok"] = ok_snap
            result["snapshot_msg"] = snap_msg
    except Exception as e:
        result["done"] = True
        result["snapshot_ok"] = False
        result["snapshot_msg"] = str(e)
        print(f"[ChromaDB 누락분 배치 오류] {e}")
    return result


def _start_gap_sync() -> None:
    st.session_state.gap_sync_active = True
    st.session_state.gap_sync_synced_session = 0


def _run_gap_sync_if_active(collection) -> None:
    """gap_sync_active이면 배치 복구 1스텝 실행 후 rerun."""
    if not st.session_state.get("gap_sync_active"):
        return

    master_n = _count_master_excel_patents()
    unique_n = (
        _count_chroma_unique_patents(collection)
        if safe_count(collection) > 0
        else 0
    )
    gap = max(0, (master_n or 0) - unique_n)

    with st.status(
        f"ChromaDB 누락분 복구 중 (고유 {unique_n}→{master_n}, 남음 약 {gap}건)...",
        expanded=True,
    ) as status:
        step = run_gap_sync_step(collection)
        synced_session = (
            st.session_state.get("gap_sync_synced_session", 0)
            + step["synced_this_run"]
        )
        st.session_state.gap_sync_synced_session = synced_session

        if step["synced_this_run"]:
            st.write(
                f"✓ 이번 배치 **+{step['synced_this_run']}건** "
                f"(이번 세션 누적 {synced_session}건)"
            )

        if not step["done"]:
            total = step["total_missing"] or gap
            done_est = max(0, total - step["remaining"])
            pct = min(0.99, done_est / max(1, total))
            st.progress(pct, text=f"남음 {step['remaining']}건")
            st.caption(
                "⚠️ **연결 끊김 방지**를 위해 15건씩 처리합니다. "
                "창을 닫거나 새로고침하지 마세요 — 자동으로 이어집니다."
            )
            status.update(label=f"누락분 복구 중… 남음 {step['remaining']}건")
            st.rerun()

        st.session_state.gap_sync_active = False
        st.session_state.chroma_gap_sync_done = True
        st.session_state._last_chroma_count = safe_count(collection)
        _invalidate_patent_count_cache()

        if synced_session > 0:
            msg = (
                f"✅ ChromaDB 누락분 **{synced_session}건** 복구 완료 "
                f"(Chroma {safe_count(collection)}건)"
            )
            if step.get("snapshot_ok") is True:
                msg += "\n\n☁️ GCS 벡터 스냅샷도 업데이트되었습니다."
            elif step.get("snapshot_ok") is False and step.get("snapshot_msg"):
                msg += f"\n\n⚠️ GCS 스냅샷: {step['snapshot_msg']}"
            st.session_state.sync_feedback = {"level": "success", "message": msg}
            status.update(label="누락분 복구 완료", state="complete")
        else:
            st.session_state.sync_feedback = {
                "level": "info",
                "message": "복구할 누락분이 없거나 이미 동기화되어 있습니다.",
            }
            status.update(label="동기화 완료", state="complete")
        st.rerun()


def sync_chroma_missing_from_master(
    collection,
    hyperlink_map: dict | None = None,
    *,
    max_items: int | None = None,
    upload_snapshot: bool = True,
) -> int:
    """
    마스터 엑셀에는 있으나 ChromaDB에 없는 출원번호만 upsert.
    max_items: None이면 전체(대량 시 Streamlit 타임아웃 위험). gap 복구 UI는 run_gap_sync_step 사용.
    """
    _compact_master_excel()
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return 0
    try:
        ids, docs, metas = _collect_missing_patents_from_master(
            collection, hyperlink_map
        )
        if not ids:
            return 0
        total_missing = len(ids)
        if max_items is not None:
            ids, docs, metas = ids[:max_items], docs[:max_items], metas[:max_items]
        _upsert_patent_batches(collection, ids, docs, metas)
        _invalidate_patent_count_cache()
        if upload_snapshot and (max_items is None or len(ids) >= total_missing):
            ok_snap, snap_msg = maybe_upload_vector_snapshot(collection)
            if not ok_snap:
                print(f"[gap 복구 후 스냅샷 업로드 스킵] {snap_msg}")
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
    _invalidate_patent_count_cache()

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
        patent_url = _patent_url_from_row(row, cols, doc_id, hyperlink_map)
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
            patent_url = _patent_url_from_row(row, cols, pat_id)
            metas.append(_build_metadata(row, cols, patent_url=patent_url))

        _upsert_patent_batches(collection, ids, docs, metas)
        _invalidate_patent_count_cache()
        ok_snap, snap_msg = maybe_upload_vector_snapshot(collection)
        if not ok_snap:
            print(f"[재인덱싱 후 스냅샷 업로드 스킵] {snap_msg}")
        return len(ids)
    except Exception as e:
        print(f"재인덱싱 실패: {e}")
        return 0


# ==========================================
# 벡터 스냅샷 (Google Drive 영속 복원)
# ==========================================
def _file_sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _upsert_embedding_batches(
    collection,
    ids: list,
    embeddings: list,
    docs: list,
    metas: list,
    batch_size: int = 100,
) -> None:
    for i in range(0, len(ids), batch_size):
        collection.upsert(
            ids=ids[i:i + batch_size],
            embeddings=embeddings[i:i + batch_size],
            documents=docs[i:i + batch_size],
            metadatas=metas[i:i + batch_size],
        )


def _validate_local_vector_manifest() -> tuple[bool, str]:
    """로컬 manifest·parquet 무결성 및 마스터 정합 검증."""
    if not os.path.exists(VECTOR_MANIFEST_PATH):
        return False, "manifest 파일 없음"
    if not os.path.exists(VECTOR_SNAPSHOT_PATH):
        return False, "snapshot parquet 없음"
    try:
        with open(VECTOR_MANIFEST_PATH, "r", encoding="utf-8") as f:
            manifest = json.load(f)
        if manifest.get("version", 0) != VECTOR_SNAPSHOT_VERSION:
            return False, "스냅샷 버전 불일치"
        if manifest.get("embed_model") != _EMBED_MODEL:
            return False, "임베딩 모델 불일치"
        if manifest.get("embed_dim") != _EMBED_DIM:
            return False, "임베딩 차원 불일치"
        master_n = _count_master_excel_patents(force=True)
        if master_n is None:
            return False, "마스터 엑셀 없음"
        if manifest.get("master_unique_count") != master_n:
            return False, (
                f"마스터 건수 불일치 (manifest={manifest.get('master_unique_count')}, "
                f"local={master_n})"
            )
        mtime = _master_excel_mtime()
        if abs(float(manifest.get("master_mtime", -999)) - mtime) > 1.0:
            return False, "마스터 mtime 불일치"
        sha = _file_sha256(VECTOR_SNAPSHOT_PATH)
        if manifest.get("snapshot_sha256") != sha:
            return False, "snapshot SHA256 불일치"
        return True, ""
    except Exception as e:
        return False, str(e)


def export_vector_snapshot(collection, batch_size: int = 500) -> tuple[bool, str]:
    """ChromaDB → 로컬 parquet + manifest 생성 (재임베딩 없음)."""
    total = safe_count(collection)
    if total == 0:
        return False, "ChromaDB가 비어 있습니다."
    try:
        all_ids = collection.get(include=[])["ids"]
        records: list[dict] = []
        skipped_no_emb = 0
        for i in range(0, len(all_ids), batch_size):
            batch = collection.get(
                ids=all_ids[i:i + batch_size],
                include=["embeddings", "documents", "metadatas"],
            )
            for cid, emb, doc, meta in zip(
                batch.get("ids", []),
                batch.get("embeddings", []),
                batch.get("documents", []),
                batch.get("metadatas", []),
            ):
                if emb is None:
                    skipped_no_emb += 1
                    continue
                records.append({
                    "chroma_id": cid,
                    "document": doc or "",
                    "embedding": [float(x) for x in emb],
                    "metadata_json": json.dumps(meta or {}, ensure_ascii=False),
                })
        if not records:
            return False, (
                f"임베딩 추출 0건 (Chroma 문서 {total}건, 임베딩 없음 {skipped_no_emb}건). "
                f"재인덱싱 완료 후 다시 시도하세요."
            )

        os.makedirs(os.path.dirname(VECTOR_SNAPSHOT_PATH), exist_ok=True)
        pd.DataFrame(records).to_parquet(VECTOR_SNAPSHOT_PATH, index=False)

        manifest = {
            "version": VECTOR_SNAPSHOT_VERSION,
            "embed_model": _EMBED_MODEL,
            "embed_dim": _EMBED_DIM,
            "master_unique_count": _count_master_excel_patents(force=True),
            "master_mtime": _master_excel_mtime(),
            "snapshot_record_count": len(records),
            "chroma_document_count": total,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "snapshot_sha256": _file_sha256(VECTOR_SNAPSHOT_PATH),
        }
        with open(VECTOR_MANIFEST_PATH, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        size_mb = os.path.getsize(VECTOR_SNAPSHOT_PATH) / (1024 * 1024)
        note = ""
        if skipped_no_emb:
            note = f" (임베딩 없음 스킵 {skipped_no_emb}건)"
        return True, f"로컬 생성 완료: {len(records)}건, {size_mb:.1f} MB{note}"
    except Exception as e:
        print(f"[벡터 스냅샷 export 오류] {e}")
        return False, f"parquet 생성 오류: {e}"


def restore_chroma_from_snapshot(collection, batch_size: int = 100) -> int:
    """로컬 parquet → ChromaDB (embeddings 포함, 재임베딩 없음)."""
    valid, reason = _validate_local_vector_manifest()
    if not valid:
        print(f"[스냅샷 복원 검증 실패] {reason}")
        return 0
    try:
        df = pd.read_parquet(VECTOR_SNAPSHOT_PATH)
        if df.empty:
            return 0
        restored = 0
        for i in range(0, len(df), batch_size):
            chunk = df.iloc[i:i + batch_size]
            ids = chunk["chroma_id"].astype(str).tolist()
            embeddings = [
                [float(x) for x in row]
                for row in chunk["embedding"].tolist()
            ]
            documents = chunk["document"].astype(str).tolist()
            metadatas = [
                json.loads(m) if isinstance(m, str) else (m or {})
                for m in chunk["metadata_json"].tolist()
            ]
            _upsert_embedding_batches(
                collection, ids, embeddings, documents, metadatas, batch_size=batch_size
            )
            restored += len(ids)
        _invalidate_patent_count_cache()
        return restored
    except Exception as e:
        print(f"[스냅샷 복원 오류] {e}")
        return 0


def download_vector_snapshot_from_storage() -> bool:
    """GCS(우선) 또는 Drive → 로컬 manifest + parquet."""
    if _gcs_configured():
        man = download_file_from_gcs(VECTOR_MANIFEST_FILENAME, VECTOR_MANIFEST_PATH)
        snap = download_file_from_gcs(VECTOR_SNAPSHOT_FILENAME, VECTOR_SNAPSHOT_PATH)
        return man is True and snap is True
    if _drive_configured():
        man = download_file_from_drive(VECTOR_MANIFEST_FILENAME, VECTOR_MANIFEST_PATH)
        snap = download_file_from_drive(VECTOR_SNAPSHOT_FILENAME, VECTOR_SNAPSHOT_PATH)
        return man is True and snap is True
    return False


def upload_vector_snapshot_to_storage() -> tuple[bool, str]:
    """로컬 manifest + parquet → GCS(우선) 또는 Drive."""
    if not os.path.exists(VECTOR_SNAPSHOT_PATH) or not os.path.exists(VECTOR_MANIFEST_PATH):
        return False, "로컬 스냅샷 파일 없음 — export 먼저 실행"

    if _gcs_configured():
        ok_m, msg_m = upload_file_to_gcs(VECTOR_MANIFEST_PATH, VECTOR_MANIFEST_FILENAME)
        if not ok_m:
            return False, msg_m
        ok_s, msg_s = upload_file_to_gcs(VECTOR_SNAPSHOT_PATH, VECTOR_SNAPSHOT_FILENAME)
        if not ok_s:
            return False, f"manifest는 GCS 업로드됨. parquet 실패: {msg_s}"
        return True, f"GCS 업로드 완료 — {msg_m}, {msg_s}"

    if _drive_configured():
        ok_m, msg_m = upload_file_to_drive(VECTOR_MANIFEST_PATH, VECTOR_MANIFEST_FILENAME)
        if not ok_m:
            return False, msg_m
        ok_s, msg_s = upload_file_to_drive(VECTOR_SNAPSHOT_PATH, VECTOR_SNAPSHOT_FILENAME)
        if not ok_s:
            return False, f"manifest는 업로드됨. parquet 실패: {msg_s}"
        return True, f"Drive 업로드 완료 — {msg_m}, {msg_s}"

    return False, "스냅샷 저장소 미설정 (GCS_BUCKET_NAME 또는 GOOGLE_DRIVE_FOLDER_ID)"


def maybe_upload_vector_snapshot(collection) -> tuple[bool, str]:
    """Chroma 변경 후 스냅샷 export + 클라우드 업로드."""
    if not _vector_storage_configured():
        return False, (
            "스냅샷 저장소 미설정 — Secrets에 GCS_BUCKET_NAME을 설정하세요."
        )
    if safe_count(collection) == 0:
        return False, "ChromaDB 비어 있음"
    ok, msg = export_vector_snapshot(collection)
    if not ok:
        return False, f"export 실패: {msg}"
    ok2, msg2 = upload_vector_snapshot_to_storage()
    if not ok2:
        return False, f"upload 실패: {msg2}"
    return True, f"[{_vector_storage_label()}] {msg} | {msg2}"


def try_restore_chroma_from_storage(collection) -> tuple[int, str]:
    """GCS/Drive 스냅샷 다운로드 → 검증 → Chroma 복원."""
    if not _vector_storage_configured():
        return 0, "스냅샷 저장소 미설정"
    if not download_vector_snapshot_from_storage():
        return 0, f"{_vector_storage_label()}에 스냅샷 없거나 다운로드 실패"
    valid, reason = _validate_local_vector_manifest()
    if not valid:
        return 0, f"검증 실패: {reason}"
    restored = restore_chroma_from_snapshot(collection)
    if restored > 0:
        return restored, f"{_vector_storage_label()} 스냅샷 {restored}건 복원"
    return 0, "복원 0건"


# 하위 호환 별칭
def download_vector_snapshot_from_drive() -> bool:
    return download_vector_snapshot_from_storage()


def upload_vector_snapshot_to_drive() -> tuple[bool, str]:
    return upload_vector_snapshot_to_storage()


def try_restore_chroma_from_drive(collection) -> tuple[int, str]:
    return try_restore_chroma_from_storage(collection)


def load_vector_manifest_summary() -> dict | None:
    """로컬 또는 클라우드 manifest 요약."""
    if not os.path.exists(VECTOR_MANIFEST_PATH) and _vector_storage_configured():
        if _gcs_configured():
            download_file_from_gcs(VECTOR_MANIFEST_FILENAME, VECTOR_MANIFEST_PATH)
        elif _drive_configured():
            download_file_from_drive(VECTOR_MANIFEST_FILENAME, VECTOR_MANIFEST_PATH)
    if not os.path.exists(VECTOR_MANIFEST_PATH):
        return None
    try:
        with open(VECTOR_MANIFEST_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


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
        _invalidate_patent_count_cache()
        master_added = len(master_needed)

    synced = sync_chroma_missing_from_master(
        collection, hyperlink_map, upload_snapshot=False
    )
    if chroma_ingested or master_added or synced:
        _invalidate_patent_count_cache()
        if chroma_ingested or synced:
            ok_snap, snap_msg = maybe_upload_vector_snapshot(collection)
            if not ok_snap:
                print(f"[적재 후 스냅샷 업로드 스킵] {snap_msg}")
    return chroma_ingested, master_added, dup_in_file + already_indexed, skipped_empty, synced


# --- 4. 메인 어플리케이션 인터페이스 구동 런타임 ---
def run_main_portal():
    _, collection, llm = load_permanent_infra_singleton()

    # ── 세션 최초 진입 시: GitHub → 로컬 전체 동기화 ──
    # github_synced는 세션 단위 플래그. 프로세스 재시작(컨테이너 재생성) 시 항상 False로 초기화됨.
    if not st.session_state.github_synced:
        download_logo_from_github()
        with st.spinner("🔄 GitHub 데이터 웨어하우스 동기화 중..."):
            excel_ok, registry_ok = sync_all_from_github()

        if excel_ok is True:
            _invalidate_patent_count_cache()
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
    unique_patent_n = _count_chroma_unique_patents(collection) if chroma_n > 0 else 0
    if (
        master_n
        and chroma_n == 0
        and not st.session_state.auto_reindex_attempted
    ):
        st.session_state.auto_reindex_attempted = True
        restored = 0
        snapshot_msg = ""
        try:
            if _vector_storage_configured():
                with st.spinner(
                    f"⚡ {_vector_storage_label()} 벡터 스냅샷 복원 시도 중... (재임베딩 생략)"
                ):
                    restored, snapshot_msg = try_restore_chroma_from_storage(collection)
            if restored <= 0:
                with st.spinner(
                    "📦 벡터 DB 자동 재인덱싱 중... (스냅샷 없음·검증 실패, 1~3분 소요)"
                ):
                    restored = reindex_from_master_excel(collection)
                if restored > 0:
                    st.toast(f"✅ Excel 재인덱싱 완료 ({restored}건)")
                elif snapshot_msg:
                    st.caption(f"스냅샷: {snapshot_msg}")
            else:
                st.toast(f"✅ {snapshot_msg}")
            st.session_state._last_chroma_count = safe_count(collection)
            _invalidate_patent_count_cache()
        except Exception as e:
            st.warning(f"벡터 DB 복원 오류: {e}")

    _run_gap_sync_if_active(collection)

    if (
        master_n
        and chroma_n > 0
        and unique_patent_n > 0
        and master_n > unique_patent_n
        and not st.session_state.chroma_gap_sync_done
        and not st.session_state.get("gap_sync_active")
    ):
        _start_gap_sync()
        st.rerun()

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

            st.checkbox(
                "⚠️ 로컬 마스터·회원 데이터를 GitHub 버전으로 덮어쓰기 (확인)",
                value=False,
                key="confirm_github_overwrite",
                help="체크하지 않으면 강제 재동기화를 실행할 수 없습니다. "
                     "로컬에만 있는 미백업 데이터는 사라질 수 있습니다.",
            )
            if st.button(
                "🔄 GitHub 데이터 강제 재동기화",
                use_container_width=True,
                disabled=not st.session_state.get("confirm_github_overwrite", False),
            ):
                with st.spinner("GitHub → 로컬 전체 동기화 중..."):
                    excel_ok, reg_ok = sync_all_from_github()
                if excel_ok is True or reg_ok is True:
                    _invalidate_patent_count_cache()
                    st.session_state.github_synced = True
                    st.session_state.confirm_github_overwrite = False
                    st.toast("✅ 동기화 완료 — 최신 GitHub 데이터가 로컬에 반영되었습니다.")
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
            st.subheader("☁️ 벡터 스냅샷 (GCS)")
            storage_label = _vector_storage_label()
            if not _GOOGLE_DRIVE_AVAILABLE:
                st.caption("GCP 패키지 미설치 — requirements.txt 확인 후 재배포하세요.")
            elif not _gcs_configured():
                st.caption(
                    "Secrets에 `GCS_BUCKET_NAME` + `[gcp_service_account]`를 설정하세요."
                )
            else:
                st.caption(
                    f"저장소: **{storage_label}** · "
                    f"파일 `vector_snapshot.parquet` + `vector_manifest.json` "
                    f"**고정명 덮어쓰기** (버전 누적 없음, ~30MB)"
                )
                manifest = load_vector_manifest_summary()
                if manifest:
                    st.caption(
                        f"manifest: **{manifest.get('snapshot_record_count', '?')}건** · "
                        f"모델 `{manifest.get('embed_model', '')}` · "
                        f"생성 `{manifest.get('created_at', '')}`"
                    )
                else:
                    st.caption("GCS에 스냅샷 없음 — 아래 버튼으로 첫 스냅샷을 생성하세요.")

            if _gcs_configured() and st.button("🔍 GCS 연결 진단", use_container_width=True):
                with st.spinner("GCS 진단 중..."):
                    gd = diagnose_gcs()
                if gd["bucket_accessible"]:
                    write_line = (
                        "✅ 쓰기 테스트 통과"
                        if gd.get("write_test_ok")
                        else "⚠️ 쓰기 테스트 실패"
                    )
                    st.success(
                        f"✅ GCS 버킷 접근 정상\n\n"
                        f"- 버킷: `{gd['bucket_name']}`\n"
                        f"- manifest: {'있음' if gd['manifest_on_gcs'] else '없음'}\n"
                        f"- snapshot: {'있음' if gd['snapshot_on_gcs'] else '없음'}\n"
                        f"- {write_line}"
                    )
                    if not gd.get("write_test_ok"):
                        st.error(gd.get("write_test_detail") or gd.get("error", ""))
                        st.info(
                            "**GCP Console 설정 (5분)**\n"
                            "1. [Cloud Storage](https://console.cloud.google.com/storage/browser) "
                            f"→ 버킷 `{gd['bucket_name']}` 클릭\n"
                            "2. **권한** 탭 → **액세스 권한 부여**\n"
                            "3. 새 주 구성원:\n"
                            "   `patent-rag-drive@patentrag.iam.gserviceaccount.com`\n"
                            "4. 역할: **Storage 관리자** (Storage Admin)\n"
                            "5. 저장 후 1~2분 뒤 **GCS 연결 진단** 재실행"
                        )
                else:
                    st.error(f"❌ GCS 연결 실패\n\n**원인:**\n{gd['error']}")

            chroma_for_snap = safe_count(collection)
            if st.button(
                f"💾 벡터 스냅샷 수동 생성 · 업로드 ({storage_label})",
                use_container_width=True,
                disabled=chroma_for_snap == 0 or not _gcs_configured(),
                help="ChromaDB 현재 상태를 parquet로 내보내 클라우드에 저장합니다. "
                     "첫 마이그레이션·재배포 전 필수 1회 실행.",
            ):
                with st.spinner(
                    f"ChromaDB {chroma_for_snap}건 스냅샷 생성 및 {storage_label} 업로드 중..."
                ):
                    ok, snap_detail = maybe_upload_vector_snapshot(collection)
                if ok:
                    st.session_state.sync_feedback = {
                        "level": "success",
                        "message": (
                            f"✅ 벡터 스냅샷 업로드 완료 "
                            f"({chroma_for_snap}건). 재시작 시 재임베딩 없이 복원됩니다.\n\n"
                            f"{snap_detail}"
                        ),
                    }
                else:
                    st.session_state.sync_feedback = {
                        "level": "error",
                        "message": (
                            f"❌ 스냅샷 생성·업로드 실패\n\n"
                            f"**상세:** {snap_detail}\n\n"
                            f"👉 **GCS 연결 진단**을 실행하세요."
                        ),
                    }
                st.rerun()

            if st.button(
                f"📥 클라우드 스냅샷 → ChromaDB 수동 복원 ({storage_label})",
                use_container_width=True,
                disabled=not _gcs_configured(),
                help="재시작 없이 클라우드 스냅샷으로 Chroma를 덮어씁니다. manifest 검증 통과 시에만 실행.",
            ):
                with st.spinner(f"{storage_label} 스냅샷 다운로드 및 Chroma 복원 중..."):
                    restored, msg = try_restore_chroma_from_storage(collection)
                st.session_state._last_chroma_count = safe_count(collection)
                _invalidate_patent_count_cache()
                if restored > 0:
                    st.session_state.sync_feedback = {
                        "level": "success",
                        "message": f"✅ {msg} (Chroma {safe_count(collection)}건)",
                    }
                else:
                    st.session_state.sync_feedback = {
                        "level": "warning",
                        "message": f"⚠️ {msg}",
                    }
                st.rerun()

            st.divider()

            uploaded_file = st.file_uploader("경쟁사 특허 엑셀 리스트 업로드 (.xlsx)", type=["xlsx"])
            if uploaded_file is not None:
                if st.button("🚀 신규 특허 무결성 적재"):
                    try:
                        preview_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        full_df       = pd.read_excel(io.BytesIO(preview_bytes))
                        preview_df    = full_df.head(3)
                        detected_cols = _detect_columns(full_df)
                        total_rows    = len(full_df)
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
                    _invalidate_patent_count_cache()
                    master_now = _count_master_excel_patents(force=True)
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
                    drive_ok = None
                    drive_detail = ""
                    if chroma_new > 0 or master_new > 0 or synced > 0:
                        with st.spinner("💾 GitHub 데이터 웨어하우스 영구 동기화 중... (대용량 파일은 최대 2분 소요)"):
                            github_ok = commit_and_push_data()
                        if chroma_new > 0 or synced > 0:
                            with st.spinner(
                                f"☁️ 벡터 스냅샷 업로드 중 ({_vector_storage_label()})..."
                            ):
                                drive_ok, drive_detail = maybe_upload_vector_snapshot(collection)
                        drive_note = ""
                        if drive_ok is True:
                            drive_note = f" {_vector_storage_label()} 벡터 스냅샷도 업데이트되었습니다."
                        elif drive_ok is False and _vector_storage_configured():
                            drive_note = (
                                f" 스냅샷 업로드 실패 — {drive_detail} "
                                f"수동 생성 버튼을 실행하세요."
                            )
                        if github_ok:
                            st.session_state.upload_feedback = {
                                "level": "success",
                                "message": (
                                    f"{summary}\n\n"
                                    f"✅ 인덱싱 및 GitHub 백업 성공 — 재시작 후에도 데이터가 보존됩니다."
                                    f"{drive_note}"
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
            chroma_n          = safe_count(collection)
            master_n          = _count_master_excel_patents()
            unique_patent_n   = _count_chroma_unique_patents(collection)
            st.markdown(_build_db_count_status(chroma_n, master_n, unique_patent_n))

            if st.session_state.sync_feedback:
                _render_feedback_box(st.session_state.sync_feedback)
                if st.button("✕ 알림 닫기", key="dismiss_sync_feedback", use_container_width=True):
                    st.session_state.sync_feedback = None
                    st.rerun()

            needs_reverse_sync = (
                master_n is not None
                and unique_patent_n > 0
                and master_n < unique_patent_n
            )
            if needs_reverse_sync:
                if st.button("📥 ChromaDB → 마스터 엑셀 역동기화", use_container_width=True):
                    with st.spinner(
                        f"ChromaDB {chroma_n}건을 마스터 엑셀로 보내는 중 "
                        f"(재임베딩 없음, 약 1~3분)..."
                    ):
                        exported, before_n, after_n = sync_master_from_chroma(collection)
                    if exported > 0:
                        _invalidate_patent_count_cache()
                        backup_msg = ""
                        with st.spinner("💾 GitHub 마스터 자동 백업 중..."):
                            github_ok = upload_file_to_github_api(
                                MASTER_EXCEL_PATH, "my_patent_folder/master_patents.xlsx"
                            )
                        if github_ok:
                            backup_msg = "  \n✅ GitHub 마스터 백업 자동 완료"
                        else:
                            backup_msg = (
                                "  \n⚠️ GitHub 자동 백업 실패 — "
                                "**💾 GitHub 마스터 백업 재시도** 버튼을 실행하세요."
                            )
                        st.session_state.sync_feedback = {
                            "level": "success" if github_ok else "warning",
                            "message": (
                                f"✅ 역동기화 완료: Chroma **{exported}건** 반영 → "
                                f"마스터 **{before_n} → {after_n}건** (ChromaDB는 그대로 유지)"
                                f"{backup_msg}"
                            ),
                        }
                        st.rerun()
                    else:
                        st.session_state.sync_feedback = {
                            "level": "error",
                            "message": "❌ 역동기화 실패 — ChromaDB 데이터를 읽지 못했습니다.",
                        }
                        st.rerun()

            needs_gap_sync = (
                master_n is not None
                and unique_patent_n > 0
                and master_n > unique_patent_n
            )
            if needs_gap_sync:
                if st.button("🔧 ChromaDB 누락분 복구 (마스터 엑셀 기준)", use_container_width=True):
                    _start_gap_sync()
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
                        _invalidate_patent_count_cache()
                        master_now = _count_master_excel_patents(force=True)
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
                        _invalidate_patent_count_cache()
                        st.session_state._last_chroma_count = 0

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
