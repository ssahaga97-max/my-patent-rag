import streamlit as st
import pandas as pd
import os

os.environ.setdefault("ANONYMIZED_TELEMETRY", "False")

import io
import re
import json
import time
import hashlib
import hmac
import shutil
import sqlite3
import tarfile
import smtplib
import chromadb
from chromadb.utils import embedding_functions
from langchain_groq import ChatGroq
import openpyxl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# boto3: Cloudflare R2 (S3 호환) 클라이언트
try:
    import boto3
    from botocore.config import Config as _BotoConfig
    _BOTO_AVAILABLE = True
except ImportError:
    boto3 = None  # type: ignore
    _BOTO_AVAILABLE = False

# Google Gemini 임베딩
try:
    import google.generativeai as genai
    _GENAI_AVAILABLE = True
except ImportError:
    genai = None  # type: ignore
    _GENAI_AVAILABLE = False


# ==========================================
# 1. 경로·상수
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "my_patent_folder")
CHROMA_DIR = os.path.join(DATA_DIR, "chroma_db")          # PersistentClient 디렉토리
USER_REGISTRY_PATH = os.path.join(DATA_DIR, "user_registry.json")
LOGO_PATH = os.path.join(BASE_DIR, "atec_logo.png")
os.makedirs(DATA_DIR, exist_ok=True)

# R2 오브젝트 키 (버킷 내 고정 경로)
R2_SNAPSHOT_KEY = "chroma_snapshot.tar.gz"
R2_REGISTRY_KEY = "user_registry.json"
R2_LOGO_KEY = "atec_logo.png"

_EMBED_DIM = 768                      # gemini-embedding-001을 768차원으로 축소 사용
_COLLECTION_NAME = "competitor_patents"

# ── Groq 무료 한도 (2026-06 기준: llama-3.3-70b-versatile = 12K TPM / 100K TPD) ──
GROQ_TPM_LIMIT = 12000
GROQ_MAX_OUTPUT_TOKENS = 1024
GROQ_REQUEST_MARGIN = 800             # 추정 오차·API 오버헤드 여유
GROQ_DAILY_TOKEN_LIMIT = 100000       # TPD — 일일 누적 추적용

st.set_page_config(page_title="AI 경쟁사 특허 조사 분석", layout="wide", page_icon="🔬")


# ==========================================
# 2. 세션 상태 초기화
# ==========================================
def _init_session():
    defaults = {
        "logged_in": False,
        "user_id": None,
        "is_admin": False,
        "restored": False,            # R2 → 로컬 복원 1회 플래그
        "upload_feedback": None,
        "infra_initialized": False,
        "_daily_token_date": "",
        "_daily_token_used": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_session()

_REGISTRY_REFRESH_TTL_SEC = 30


# ==========================================
# 3. 공통 유틸리티
# ==========================================
def _clean_ascii(value: str) -> str:
    """비가시적 유니코드 문자(BOM, Zero-Width Space 등) 제거."""
    return value.encode("ascii", errors="ignore").decode("ascii").strip()


def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _is_sha256_hex(value: str) -> bool:
    s = str(value or "").strip().lower()
    return len(s) == 64 and all(c in "0123456789abcdef" for c in s)


def _verify_password(plain: str, stored_hash: str) -> bool:
    """타이밍 공격 완화 — stored_hash는 SHA-256 hex만 허용."""
    if not plain or not stored_hash:
        return False
    if not _is_sha256_hex(stored_hash):
        return False
    return hmac.compare_digest(stored_hash.strip().lower(), _hash_pw(plain))


def _normalize_patent_id(value) -> str:
    """
    출원번호 정규화(canonical ID) — 표기 차이를 하나의 키로 통일.
      1) Excel float/지수표기 → 정수 문자열
      2) 구분자(하이픈·공백·슬래시·점) 제거
      3) KR/kr 선두 접두사 제거
      4) 영숫자만 유지(대문자)
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


# ==========================================
# 4. 통계 집계: 복수값 분리 · 경쟁사 대표명화
# ==========================================
_MULTI_VALUE_SPLIT_RE = re.compile(r"\s*[|;/]\s*")

# lookup key(대문자·기호·접미사 제거) → 대표명.
# 효성·히타치처럼 변형이 많은 경쟁사는 모든 변형이 하나의 대표명으로 수렴하도록 매핑.
# 한글 변형(노틸러스효성, 효성티앤에스 등)도 키로 등록.
_APPLICANT_CANONICAL_KEYS: dict = {
    # ── 효성 계열 → "효성" 으로 통합 ──
    "HYOSUNG": "효성",
    "HYOSUNGNAUTILUS": "효성",
    "NAUTILUSHYOSUNG": "효성",
    "HYOSUNGTNS": "효성",
    "효성": "효성",
    "효성티앤에스": "효성",
    "효성티엔에스": "효성",
    "노틸러스효성": "효성",
    "효성노틸러스": "효성",
    "나우테크놀로지": "효성",          # 효성 계열사(필요 시 조정)
    # ── 히타치 계열 → "히타치" 로 통합 ──
    "HITACHI": "히타치",
    "HITACHIOMRON": "히타치",
    "HITACHIOMRONTERMINALSOLUTIONS": "히타치",
    "HITACHICHANNELSOLUTIONS": "히타치",
    "HITACHITERMINALSOLUTIONS": "히타치",
    "히타치": "히타치",
    "히타치옴론": "히타치",
    # ── Diebold/Wincor 계열 ──
    "DIEBOLD": "DIEBOLD NIXDORF",
    "DIEBOLDNIXDORF": "DIEBOLD NIXDORF",
    "DIEBOLDNIXDORFINCORPORATED": "DIEBOLD NIXDORF",
    "DIEBOLDNIXDORFSYSTEMSGMBH": "DIEBOLD NIXDORF",
    "WINCORNIXDORF": "DIEBOLD NIXDORF",
    "WINCORNIXDORFINTERNATIONALGMBH": "DIEBOLD NIXDORF",
    "디볼드": "DIEBOLD NIXDORF",
    "디볼드닉스도르프": "DIEBOLD NIXDORF",
    # ── NCR ──
    "NCR": "NCR",
    "NCRCORPORATION": "NCR",
    "NCRVOYIX": "NCR",
    "NCRATLEOS": "NCR",
    # ── 기타 경쟁사 ──
    "HOTS": "HOTS",
    "GLORY": "GLORY",
    "GLORYLTD": "GLORY",
    "OKIELECTRIC": "OKI ELECTRIC",
    "OKIELECTRICINDUSTRY": "OKI ELECTRIC",
    "OKIELECTRICINDUSTRYCOLTD": "OKI ELECTRIC",
    "OKI": "OKI ELECTRIC",
    "FTEC": "FTEC",
    "GRGBANKING": "GRG BANKING",
    "GRGBANKINGEQUIPMENT": "GRG BANKING",
    "GRGBANKINGEQUIPMENTCOLTD": "GRG BANKING",
    "GRG": "GRG BANKING",
    "SHENZHENYIHUA": "SHENZHEN YIHUA",
    "SHENZHENYIHUACOMPCOLTD": "SHENZHEN YIHUA",
    "SHENZHENYIHUATIMETECHNOLOGY": "SHENZHEN YIHUA",
    "SHENZHENYIHUAFINANCIALINTELLIGENTRESINST": "SHENZHEN YIHUA",
    "YIHUA": "SHENZHEN YIHUA",
    "CASHWAY": "CASHWAY",
    "CASHWAYTECHNOLOGY": "CASHWAY",
    "GUARDIAN": "GUARDIAN",
    "GUARDIANANALYTICS": "GUARDIAN",
    "FUJITSU": "FUJITSU",
    "후지쯔": "FUJITSU",
    "TOSHIBA": "TOSHIBA",
    "도시바": "TOSHIBA",
    "RICOH": "RICOH",
    "CANON": "CANON",
    "PANASONIC": "PANASONIC",
    "CUMMINSALLISON": "CUMMINS ALLISON",
    "DELARUE": "DE LA RUE",
    "GIESSECKE": "GIECKE+DEVRIENT",
    "GIESECKE": "GIECKE+DEVRIENT",
    "GIECKE": "GIECKE+DEVRIENT",
}

# lookup key가 아래 접두사로 시작하면 대표명으로 통합 (매핑표에 없는 변형 포착).
# 효성·히타치는 어떤 꼬리표가 붙어도 대표명으로 수렴.
_APPLICANT_PREFIX_RULES: list = [
    ("HYOSUNG", "효성"),
    ("NAUTILUS", "효성"),
    ("효성", "효성"),
    ("노틸러스효성", "효성"),
    ("HITACHI", "히타치"),
    ("히타치", "히타치"),
    ("DIEBOLD", "DIEBOLD NIXDORF"),
    ("WINCOR", "DIEBOLD NIXDORF"),
    ("디볼드", "DIEBOLD NIXDORF"),
    ("NCR", "NCR"),
    ("GLORY", "GLORY"),
    ("OKIELECTRIC", "OKI ELECTRIC"),
    ("GRGBANKING", "GRG BANKING"),
    ("SHENZHENYIHUA", "SHENZHEN YIHUA"),
    ("CASHWAY", "CASHWAY"),
    ("GUARDIAN", "GUARDIAN"),
    ("CUMMINSALLISON", "CUMMINS ALLISON"),
    ("GIESSECKE", "GIECKE+DEVRIENT"),
    ("GIESECKE", "GIECKE+DEVRIENT"),
    ("GIECKE", "GIECKE+DEVRIENT"),
    ("DELARUE", "DE LA RUE"),
]

_CORP_SUFFIX_PATTERNS = (
    r"incorporated", r"inc\.?", r"corp\.?", r"corporation", r"ltd\.?", r"limited",
    r"gmbh", r"co\.?", r"company", r"llc", r"plc", r"pte\.?", r"ag", r"sa", r"bv",
    r"주식회사", r"\(주\)", r"㈜", r"유한회사", r"\(유\)", r"coltd", r"col\.?",
    r"holdings?", r"group", r"international", r"systems?", r"equipment",
    r"technology", r"technologies", r"industry", r"industries", r"financial",
    r"intelligent", r"research", r"inst(?:itute)?", r"res", r"inst",
    # 한글 접미사(효성티앤에스 → 효성, 히타치옴론 → 히타치 등 수렴 지원)
    r"티앤에스", r"티엔에스", r"앤에스", r"테크놀로지", r"테크놀러지", r"전자",
    r"솔루션즈?", r"시스템즈?", r"인더스트리", r"옴론",
)
# 결합 정규식: 접미사를 1패스로 제거 (성능 개선)
# 영문은 단어경계(\b), 한글 접미사는 경계가 없으므로 위치 무관 제거.
_ASCII_SUFFIXES = tuple(p for p in _CORP_SUFFIX_PATTERNS if not re.search(r"[가-힣]", p))
_HANGUL_SUFFIXES = tuple(p for p in _CORP_SUFFIX_PATTERNS if re.search(r"[가-힣]", p))
# 한글 회사 표현(주식회사/(주)/㈜ 등)은 어디에 있든 제거
_HANGUL_CORP_TOKENS = (r"주식회사", r"\(주\)", r"㈜", r"유한회사", r"\(유\)")
_CORP_SUFFIX_RE = re.compile(
    r"\b(?:" + "|".join(_ASCII_SUFFIXES) + r")\b", flags=re.IGNORECASE
)
_HANGUL_CORP_RE = re.compile("(?:" + "|".join(_HANGUL_CORP_TOKENS) + ")")
_CORP_SUFFIX_HANGUL_RE = re.compile("(?:" + "|".join(_HANGUL_SUFFIXES) + r")")
_CORP_SEP_RE = re.compile(r"[\s\-_./\\()（）\[\]]+")


def _applicant_lookup_key(name: str) -> str:
    """출원인 문자열을 alias 조회용 키로 변환."""
    s = str(name).strip()
    # 한글 회사 표현(주식회사/(주)/㈜)을 먼저 위치 무관 제거
    s = _HANGUL_CORP_RE.sub("", s)
    for _ in range(3):
        prev = s
        s = s.replace(",", " ")
        s = _CORP_SUFFIX_RE.sub("", s)
        s = _CORP_SUFFIX_HANGUL_RE.sub("", s)  # 한글 접미사(티앤에스 등) 제거
        s = _CORP_SEP_RE.sub("", s)
        if s == prev:
            break
    return s.upper()


def _canonical_applicant_name(name: str) -> str:
    """경쟁사 출원인 표기를 대표명으로 통합 (집계·필터 전용, 메타데이터 원본 유지)."""
    raw = str(name).strip()
    if not raw or raw.lower() in ("nan", "none", "없음", "정보없음", "미기재"):
        return ""

    key = _applicant_lookup_key(raw)
    if not key:
        return raw

    # 1) 정확 매칭 (영문 키 + 한글 변형 키 모두 등록됨)
    if key in _APPLICANT_CANONICAL_KEYS:
        return _APPLICANT_CANONICAL_KEYS[key]

    # 2) 접두사 매칭 (변형 흡수)
    for prefix, canonical in _APPLICANT_PREFIX_RULES:
        if key.startswith(prefix):
            return canonical

    # 3) 매핑 없는 한글 출원인: 회사 접미사만 정리해 표시
    if re.search(r"[가-힣]", raw):
        cleaned = re.sub(r"\s*(\(주\)|주식회사|㈜|유한회사|\(유\))\s*", "", raw).strip()
        return cleaned if cleaned else raw

    return raw


def _explode_multi_values(value, *, allow_comma: bool = False) -> list:
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


def _flatten_for_counts(series, *, normalizer=None, explode=False, allow_comma=False):
    """Series를 value_counts 가능한 flat Series로 변환."""
    items: list = []
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


def _top_counts_table(series, n=10, *, label="항목", normalizer=None,
                      explode=False, allow_comma=False):
    """상위 N건 집계표(DataFrame)와 LLM용 마크다운 표 문자열 반환."""
    flat = _flatten_for_counts(series, normalizer=normalizer, explode=explode,
                               allow_comma=allow_comma)
    if flat.empty:
        empty = pd.DataFrame(columns=[label, "건수"])
        return empty, f"| {label} | 건수 |\n| --- | ---: |\n| (데이터 없음) | 0 |"

    counts = flat.value_counts().head(n)
    df = pd.DataFrame({label: counts.index, "건수": counts.values})

    lines = [f"| {label} | 건수 |", "| --- | ---: |"]
    for _, row in df.iterrows():
        lines.append(f"| {row[label]} | {int(row['건수'])} |")
    return df, "\n".join(lines)


def _year_counts_table(series):
    """출원 연도별 건수 집계표."""
    years = (
        series.astype(str).str[:4]
        .replace("", pd.NA).replace("없음", pd.NA).dropna()
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


def _extract_year_filter_from_query(query: str):
    """질의문에서 출원연도 필터 추출. (min_year, max_year)."""
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


def _filter_metadata_df_by_year(df, min_year, max_year=None):
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


def _metadata_df_with_applicant_rows(df):
    """복수 출원인(구분자) 행을 출원인 단위로 펼침."""
    records: list = []
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


def _build_applicant_cohort_context(df, max_applicants=25, titles_per=15, ipc_per=8):
    """필터된 메타데이터에서 출원인별 IPC·특허명 코호트 요약 생성."""
    exploded = _metadata_df_with_applicant_rows(df)
    if exploded.empty:
        return "(조건에 맞는 데이터 없음)", pd.DataFrame(columns=["출원인(대표명)", "특허건수"])

    parts: list = []
    summary_rows: list = []
    app_counts = exploded["_대표출원인"].value_counts()

    for app, cnt in app_counts.head(max_applicants).items():
        sub = exploded[exploded["_대표출원인"] == app]
        _, ipc_md = _top_counts_table(
            sub.get("IPC", pd.Series(dtype=str)),
            n=ipc_per, label="IPC", explode=True, allow_comma=True,
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
# 5. Cloudflare R2 스토리지 (S3 호환, egress 무료)
# ==========================================
def _r2_configured() -> bool:
    if not _BOTO_AVAILABLE:
        return False
    need = ("R2_ACCOUNT_ID", "R2_ACCESS_KEY_ID", "R2_SECRET_ACCESS_KEY", "R2_BUCKET")
    return all(st.secrets.get(k) for k in need)


@st.cache_resource
def _get_r2_client():
    """R2 S3 호환 클라이언트 (프로세스당 1회)."""
    account_id = _clean_ascii(st.secrets.get("R2_ACCOUNT_ID", ""))
    access_key = _clean_ascii(st.secrets.get("R2_ACCESS_KEY_ID", ""))
    secret_key = _clean_ascii(st.secrets.get("R2_SECRET_ACCESS_KEY", ""))
    endpoint = f"https://{account_id}.r2.cloudflarestorage.com"
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        config=_BotoConfig(signature_version="s3v4", retries={"max_attempts": 3}),
        region_name="auto",
    )


def _r2_bucket() -> str:
    return _clean_ascii(st.secrets.get("R2_BUCKET", ""))


def r2_upload_bytes(key: str, data: bytes, content_type: str = "application/octet-stream") -> bool:
    if not _r2_configured():
        return False
    try:
        _get_r2_client().put_object(
            Bucket=_r2_bucket(), Key=key, Body=data, ContentType=content_type
        )
        return True
    except Exception as e:
        print(f"[R2 업로드 실패] {key}: {e}")
        return False


def r2_download_bytes(key: str):
    """R2에서 오브젝트 바이트 반환. 없거나 오류 시 None."""
    if not _r2_configured():
        return None
    try:
        obj = _get_r2_client().get_object(Bucket=_r2_bucket(), Key=key)
        return obj["Body"].read()
    except Exception as e:
        msg = str(e)
        if "NoSuchKey" not in msg and "Not Found" not in msg and "404" not in msg:
            print(f"[R2 다운로드 실패] {key}: {e}")
        return None


def r2_delete(key: str) -> bool:
    if not _r2_configured():
        return False
    try:
        _get_r2_client().delete_object(Bucket=_r2_bucket(), Key=key)
        return True
    except Exception as e:
        print(f"[R2 삭제 실패] {key}: {e}")
        return False


def diagnose_r2() -> dict:
    """R2 연결 진단."""
    result = {"configured": False, "reachable": False, "bucket": "", "error": ""}
    if not _BOTO_AVAILABLE:
        result["error"] = "boto3 미설치 — requirements.txt에 boto3 추가 필요"
        return result
    if not _r2_configured():
        result["error"] = (
            "R2 Secrets 누락 — R2_ACCOUNT_ID, R2_ACCESS_KEY_ID, "
            "R2_SECRET_ACCESS_KEY, R2_BUCKET을 설정하세요."
        )
        return result
    result["configured"] = True
    result["bucket"] = _r2_bucket()
    try:
        _get_r2_client().head_bucket(Bucket=_r2_bucket())
        result["reachable"] = True
    except Exception as e:
        result["error"] = f"R2 버킷 접근 실패: {e}"
    return result


def list_r2_objects() -> list:
    """
    R2 버킷의 객체 목록 반환 [(key, size_bytes), ...].
    스냅샷이 실제로 존재하는지/용량은 얼마인지 진단용.
    """
    if not _r2_configured():
        return []
    try:
        resp = _get_r2_client().list_objects_v2(Bucket=_r2_bucket())
        return [(o["Key"], o["Size"]) for o in resp.get("Contents", [])]
    except Exception as e:
        print(f"[R2 목록 조회 실패] {e}")
        return []


# ── ChromaDB 디렉토리 ↔ R2 (tar.gz 통째 백업/복원) ──
def backup_chroma_to_r2() -> bool:
    """CHROMA_DIR 전체를 tar.gz로 압축해 R2에 업로드."""
    if not os.path.isdir(CHROMA_DIR):
        return False
    try:
        buf = io.BytesIO()
        with tarfile.open(fileobj=buf, mode="w:gz") as tar:
            tar.add(CHROMA_DIR, arcname="chroma_db")
        buf.seek(0)
        return r2_upload_bytes(R2_SNAPSHOT_KEY, buf.getvalue(), "application/gzip")
    except Exception as e:
        print(f"[Chroma 백업 실패] {e}")
        return False


def restore_chroma_from_r2() -> bool:
    """R2의 tar.gz를 받아 CHROMA_DIR로 복원. 성공 시 True.
    (상세 진단이 필요하면 restore_chroma_from_r2_verbose 사용)"""
    info = restore_chroma_from_r2_verbose()
    return bool(info["extracted"]) and not info["error"]


def restore_chroma_from_r2_verbose() -> dict:
    """복원 + 상세 진단 정보 반환. 복원의 단일 진입점."""
    info = {"downloaded_mb": 0, "extracted": False, "files": [], "sqlite_mb": 0,
            "tar_members": [], "error": ""}
    data = r2_download_bytes(R2_SNAPSHOT_KEY)
    if not data:
        info["error"] = "스냅샷 다운로드 실패 (R2에 파일 없음)"
        return info
    info["downloaded_mb"] = round(len(data) / (1024 * 1024), 2)
    try:
        if os.path.isdir(CHROMA_DIR):
            shutil.rmtree(CHROMA_DIR, ignore_errors=True)
        with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as tar:
            info["tar_members"] = tar.getnames()[:10]
            tar.extractall(path=DATA_DIR)
        info["extracted"] = os.path.isdir(CHROMA_DIR)
        # 풀린 파일 목록·sqlite 크기
        for root, dirs, files in os.walk(CHROMA_DIR):
            for f in files:
                fp = os.path.join(root, f)
                rel = os.path.relpath(fp, CHROMA_DIR)
                info["files"].append(rel)
                if f == "chroma.sqlite3":
                    info["sqlite_mb"] = round(os.path.getsize(fp) / (1024 * 1024), 2)
        info["files"] = info["files"][:15]
    except Exception as e:
        info["error"] = str(e)
    return info


def recover_from_sqlite_directly(collection) -> tuple:
    """
    chromadb 버전 불일치로 컬렉션이 0건일 때 최후 복구.
    복원된 chroma.sqlite3에서 문서·메타데이터를 직접 SELECT하여
    현재 버전·현재 임베딩 함수로 재적재(재임베딩)한다.

    반환: (복구건수, 메시지)
    """
    sqlite_path = os.path.join(CHROMA_DIR, "chroma.sqlite3")
    if not os.path.exists(sqlite_path):
        return 0, "chroma.sqlite3 파일이 없습니다."

    try:
        conn = sqlite3.connect(sqlite_path)
        cur = conn.cursor()

        # 0.5.x 스키마: embedding_metadata(id, key, string_value, ...) + embeddings(embedding_id, id...)
        # 문서 본문은 embedding_metadata에서 key='chroma:document' 로 저장됨
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = {r[0] for r in cur.fetchall()}
        if "embedding_metadata" not in tables:
            conn.close()
            return 0, f"예상한 테이블 구조가 아닙니다. 발견된 테이블: {sorted(tables)[:8]}"

        # 각 embedding의 메타데이터를 key-value로 모음
        cur.execute("SELECT id, key, string_value FROM embedding_metadata")
        rows = cur.fetchall()
        conn.close()

        records: dict = {}
        for emb_id, key, sval in rows:
            if emb_id not in records:
                records[emb_id] = {}
            if sval is not None:
                records[emb_id][key] = sval

        if not records:
            return 0, "sqlite에 메타데이터 레코드가 없습니다."

        # 재적재 배치 구성
        ids, docs, metas = [], [], []
        for emb_id, kv in records.items():
            doc = kv.get("chroma:document", "")
            pat_id = kv.get("출원번호", "")
            cid = _normalize_patent_id(pat_id) if pat_id else ""
            if not cid:
                continue
            _rep_list = _canonical_applicants_list(kv.get("출원인", ""))
            meta = {
                "출원번호": kv.get("출원번호", ""),
                "명칭": kv.get("명칭", "정보없음"),
                "출원일": kv.get("출원일", "없음"),
                "등록일": kv.get("등록일", "없음"),
                "IPC": kv.get("IPC", "없음"),
                "CPC": kv.get("CPC", "없음"),
                "발명자": kv.get("발명자", "없음"),
                "출원인": kv.get("출원인", "없음"),
                "대표출원인": "|".join(_rep_list),
                "대표출원인목록": _rep_list,
                "URL": kv.get("URL", ""),
            }
            ids.append(cid)
            docs.append(doc if doc else f"특허명칭: {meta['명칭']}")
            metas.append(meta)

        if not ids:
            return 0, "복구 가능한 출원번호가 없습니다."

        # 현재 임베딩 함수로 재적재 (재임베딩)
        _upsert_patent_batches(collection, ids, docs, metas, batch_size=50)
        return len(ids), f"{len(ids)}건을 sqlite에서 직접 읽어 재임베딩·복구했습니다."

    except Exception as e:
        return 0, f"sqlite 직접 복구 오류: {e}"


# ==========================================
# 6. 회원 레지스트리 (R2 백업)
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


def upload_user_registry_to_r2() -> bool:
    try:
        with open(USER_REGISTRY_PATH, "rb") as f:
            return r2_upload_bytes(R2_REGISTRY_KEY, f.read(), "application/json")
    except Exception:
        return False


def download_user_registry_from_r2() -> bool:
    data = r2_download_bytes(R2_REGISTRY_KEY)
    if data is None:
        return False
    try:
        os.makedirs(os.path.dirname(USER_REGISTRY_PATH), exist_ok=True)
        with open(USER_REGISTRY_PATH, "wb") as f:
            f.write(data)
        return True
    except Exception:
        return False


def refresh_user_registry_from_r2(*, force: bool = False) -> dict:
    """짧은 TTL 캐시로 R2 레지스트리 갱신."""
    now = time.time()
    last = st.session_state.get("_registry_refresh_ts", 0)
    if force or (now - last) > _REGISTRY_REFRESH_TTL_SEC:
        download_user_registry_from_r2()
        st.session_state["_registry_refresh_ts"] = now
    return load_user_registry()


def download_logo_from_r2() -> bool:
    if os.path.exists(LOGO_PATH):
        return True
    data = r2_download_bytes(R2_LOGO_KEY)
    if not data:
        return False
    try:
        with open(LOGO_PATH, "wb") as f:
            f.write(data)
        return True
    except Exception:
        return False


# ==========================================
# 7. 관리자 알림 이메일
# ==========================================
def send_admin_signup_email(user_info: dict) -> bool:
    """신규 회원 가입 시 관리자 이메일로 알림. Secrets: SMTP_EMAIL/SMTP_PASSWORD/ADMIN_EMAIL."""
    try:
        smtp_email = st.secrets.get("SMTP_EMAIL", "")
        smtp_password = st.secrets.get("SMTP_PASSWORD", "")
        admin_email = st.secrets.get("ADMIN_EMAIL", "")
        if not smtp_email or not smtp_password or not admin_email:
            return False

        msg = MIMEMultipart("alternative")
        msg["From"] = smtp_email
        msg["To"] = admin_email
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
<p><b>👉 관리자 포털</b> → 사이드바 <b>가입 회원 현황</b> → 해당 계정 <b>활성화</b> 버튼을 눌러 승인해 주세요.</p>
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


# ==========================================
# 8. 인증 시스템
# ==========================================
def _get_admin_credentials() -> dict:
    """Streamlit Secrets의 [USER_CREDENTIALS]에서 관리자 자격증명 반환."""
    if "USER_CREDENTIALS" in st.secrets:
        return dict(st.secrets["USER_CREDENTIALS"])
    return {}


def check_authentication():
    if st.session_state.logged_in:
        return True

    if not os.path.exists(USER_REGISTRY_PATH):
        download_user_registry_from_r2()

    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=160)
    st.title("AI 경쟁사 특허 조사 분석")

    tab_login, tab_register = st.tabs(["🔑 로그인", "📝 신규 회원 가입"])

    with tab_login:
        st.subheader("사내 연구원 로그인")
        with st.form("login_form"):
            username = st.text_input("계정 ID", key="login_id")
            password = st.text_input("비밀번호", type="password", key="login_pw")
            submitted = st.form_submit_button("접속")

        if submitted:
            admin_creds = _get_admin_credentials()
            registry = refresh_user_registry_from_r2(force=True)

            if username in admin_creds:
                stored = admin_creds[username]
                if not _is_sha256_hex(stored):
                    st.error(
                        "⛔ 관리자 비밀번호는 Streamlit Secrets에 SHA-256 해시(64자)로만 "
                        "설정하세요. 평문 비밀번호는 허용되지 않습니다."
                    )
                elif _verify_password(password, stored):
                    st.session_state.logged_in = True
                    st.session_state.user_id = username
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    st.error("❌ 비밀번호가 올바르지 않습니다.")
            elif username in registry:
                user_rec = registry[username]
                if user_rec.get("active") is False:
                    st.error("⛔ 계정 승인 대기 중입니다. 관리자가 **활성화**한 뒤 로그인해 주세요.")
                elif _verify_password(password, user_rec.get("password_hash", "")):
                    st.session_state.logged_in = True
                    st.session_state.user_id = username
                    st.session_state.is_admin = False
                    st.rerun()
                else:
                    st.error("❌ 비밀번호가 올바르지 않습니다.")
            else:
                st.error("❌ 등록되지 않은 계정입니다. '신규 회원 가입' 탭을 이용해 주세요.")

    with tab_register:
        st.subheader("신규 연구원 계정 등록")
        st.caption(
            "가입 신청 후 **관리자 승인(활성화)** 이 완료되면 로그인할 수 있습니다. "
            "신청 시 관리자에게 이메일로 알림이 발송됩니다."
        )
        with st.form("register_form"):
            r_id = st.text_input("사용자 ID (영문·숫자, 4자 이상)")
            r_name = st.text_input("이름 *")
            r_email = st.text_input("이메일 *")
            r_dept = st.text_input("부서 (선택)")
            r_pw = st.text_input("비밀번호 (6자 이상)", type="password")
            r_pw2 = st.text_input("비밀번호 확인", type="password")
            reg_submitted = st.form_submit_button("가입 신청")

        if reg_submitted:
            admin_creds = _get_admin_credentials()
            registry = load_user_registry()
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
                    "username": r_id,
                    "name": r_name,
                    "email": r_email,
                    "department": r_dept,
                    "password_hash": _hash_pw(r_pw),
                    "registered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "active": False,
                }
                registry[r_id] = user_info
                save_user_registry(registry)

                with st.spinner("💾 계정 정보를 저장하는 중..."):
                    upload_user_registry_to_r2()

                email_ok = send_admin_signup_email(user_info)

                st.success(
                    f"✅ '{r_id}' 가입 신청이 접수되었습니다. "
                    f"관리자 승인 후 로그인 탭에서 접속해 주세요."
                )
                if email_ok:
                    st.info("📧 관리자에게 가입 알림 이메일이 발송되었습니다.")
                else:
                    st.caption("(이메일 발송 미설정 — Secrets에 SMTP_EMAIL / SMTP_PASSWORD / ADMIN_EMAIL 추가 시 활성화)")

    return False


# ==========================================
# 9. Gemini 임베딩 함수 (ChromaDB EmbeddingFunction 인터페이스)
# ==========================================
class GeminiEmbeddingFunction(embedding_functions.EmbeddingFunction):
    """
    Google Gemini embedding (gemini-embedding-001, 768차원으로 축소 사용).
    text-embedding-004는 2026-01-14 deprecated → gemini-embedding-001로 마이그레이션.
    기본 3072차원이지만 output_dimensionality=768로 기존 데이터와 차원 호환.
    비대칭 임베딩: 문서는 retrieval_document, 질의는 retrieval_query로 인코딩.
    """

    def __init__(self, api_key: str, model: str = "models/gemini-embedding-001"):
        if not _GENAI_AVAILABLE:
            raise RuntimeError("google-generativeai 미설치 — requirements.txt에 추가 필요")
        genai.configure(api_key=api_key)
        self._model = model

    def _embed_one(self, text: str, task_type: str, *, max_retries: int = 5):
        text = (text or "").strip() or " "
        # 2048토큰 ≈ 한글 수천 자. 청구항 포함 대비 여유 있게 12000자까지 허용
        if len(text) > 12000:
            text = text[:12000]
        delay = 2.0
        last_err = None
        for attempt in range(max_retries):
            try:
                resp = genai.embed_content(
                    model=self._model, content=text, task_type=task_type,
                    output_dimensionality=_EMBED_DIM,   # 768 — 기존 데이터 차원 유지
                )
                emb = resp["embedding"]
                # gemini-embedding-001은 축소 차원 사용 시 정규화 권장
                norm = sum(x * x for x in emb) ** 0.5
                if norm > 0:
                    emb = [x / norm for x in emb]
                return emb
            except Exception as e:
                last_err = e
                msg = str(e).lower()
                # 429/rate limit/quota → 지수 백오프 후 재시도
                if any(k in msg for k in ("429", "rate", "quota", "resource", "exhaust")):
                    if attempt < max_retries - 1:
                        time.sleep(delay)
                        delay = min(delay * 2, 30)  # 2→4→8→16→30초
                        continue
                # 그 외 오류는 짧게 한 번만 재시도
                if attempt < 1:
                    time.sleep(1.0)
                    continue
                break
        print(f"[Gemini 임베딩 최종 실패/{task_type}] {last_err}")
        # 재시도 다 소진 → 0벡터 (검색에서 제외됨, 나중에 재임베딩 대상)
        return [0.0] * _EMBED_DIM

    def __call__(self, input):
        """문서 저장용 — retrieval_document."""
        if isinstance(input, str):
            input = [input]
        return [self._embed_one(t, "retrieval_document") for t in input]

    def embed_query(self, text: str):
        """검색 질의용 — retrieval_query (비대칭 검색 정확도 핵심)."""
        return self._embed_one(text, "retrieval_query")


# ==========================================
# 10. ChromaDB (PersistentClient) / LLM 인프라
# ==========================================
@st.cache_resource
def _get_embedding_fn():
    """
    임베딩 함수 선택:
      · GEMINI_API_KEY가 있으면 Gemini API (권장 — RAM 절약 + 긴 컨텍스트)
      · 없으면 로컬 SentenceTransformer로 폴백
    """
    gemini_key = _clean_ascii(st.secrets.get("GEMINI_API_KEY", ""))
    if gemini_key and _GENAI_AVAILABLE:
        return GeminiEmbeddingFunction(gemini_key)
    # 폴백: 로컬 모델 (RAM 수백 MB 사용)
    return embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name="sentence-transformers/paraphrase-multilingual-mpnet-base-v2"
    )


@st.cache_resource
def _get_chroma_client():
    """PersistentClient — CHROMA_DIR 디스크 영속. 프로세스당 1회."""
    os.makedirs(CHROMA_DIR, exist_ok=True)
    return chromadb.PersistentClient(path=CHROMA_DIR)


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


def _embed_query_vector(query: str):
    """
    검색 질의를 retrieval_query task_type으로 임베딩.
    Gemini(비대칭)면 query 전용 벡터를 반환 → query_embeddings로 사용.
    로컬 모델(대칭)이면 None 반환 → 호출부에서 query_texts 사용.
    """
    fn = _get_embedding_fn()
    if isinstance(fn, GeminiEmbeddingFunction):
        try:
            return fn.embed_query(query)
        except Exception as e:
            print(f"[질의 임베딩 오류] {e}")
            return None
    return None


def _get_collection():
    """
    기존 컬렉션이 있으면 그대로 열고(메타 충돌 회피), 없으면 생성.
    복원된 DB의 컬렉션을 metadata 충돌 없이 안전하게 인식.
    """
    client = _get_chroma_client()
    try:
        # 기존 컬렉션 우선 — 임베딩 함수만 연결, metadata는 재지정하지 않음
        return client.get_collection(
            name=_COLLECTION_NAME,
            embedding_function=_get_embedding_fn(),
        )
    except Exception:
        # 없으면 새로 생성
        return client.get_or_create_collection(
            name=_COLLECTION_NAME,
            embedding_function=_get_embedding_fn(),
            metadata={"hnsw:space": "cosine"},
        )


def reset_collection():
    """포맷 전용: 컬렉션 삭제 후 빈 컬렉션 재생성."""
    client = _get_chroma_client()
    try:
        client.delete_collection(_COLLECTION_NAME)
    except Exception:
        pass
    return client.get_or_create_collection(
        name=_COLLECTION_NAME,
        embedding_function=_get_embedding_fn(),
        metadata={"hnsw:space": "cosine"},
    )


def load_permanent_infra_singleton():
    """캐시된 client, collection, llm 반환. 최초 1회 초기화 spinner 표시."""
    if not st.session_state.infra_initialized:
        with st.spinner("📦 AI 임베딩 엔진 및 벡터 커널 초기화 중..."):
            _get_chroma_client()
            _get_llm()
        st.session_state.infra_initialized = True
    return _get_chroma_client(), _get_collection(), _get_llm()


def chroma_count(collection) -> int:
    """collection.count() — 순수 호출."""
    return collection.count()


def safe_count(collection) -> int:
    """collection.count() 안전 호출. 일시 오류 시 마지막 정상값 반환."""
    try:
        n = chroma_count(collection)
        st.session_state["_last_chroma_count"] = n
        return n
    except Exception as e:
        print(f"[ChromaDB safe_count 오류] {e}")
        return st.session_state.get("_last_chroma_count", 0)


# ==========================================
# 11. Groq 토큰 예산 + 일일 한도(TPD) 추적
# ==========================================
def _estimate_tokens(text: str) -> int:
    """한국어 특허 텍스트 보수적 토큰 추정 (과소 추정 방지)."""
    if not text:
        return 0
    # 한국어는 실제로 1자당 1토큰 이상인 경우가 많음 → 1.1자/토큰으로 보수적 추정
    return max(1, int(len(text) / 1.1) + 5)


def _chars_for_token_budget(tokens: int) -> int:
    return max(100, int(tokens * 1.1))


def _today_key() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def _daily_tokens_used() -> int:
    """오늘 누적 사용 토큰. 날짜 바뀌면 리셋."""
    if st.session_state.get("_daily_token_date") != _today_key():
        st.session_state["_daily_token_date"] = _today_key()
        st.session_state["_daily_token_used"] = 0
    return st.session_state.get("_daily_token_used", 0)


def _add_daily_tokens(n: int):
    _daily_tokens_used()  # 날짜 갱신 보장
    st.session_state["_daily_token_used"] = st.session_state.get("_daily_token_used", 0) + max(0, n)


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


def _format_patent_meta_line(i: int, m: dict):
    app_num = m.get("출원번호", "번호없음")
    p_name = m.get("명칭", "제목없음")
    applicant = m.get("출원인", "미기재")
    inventor = m.get("발명자", "미기재")
    ipc = m.get("IPC", "없음")
    cpc = m.get("CPC", "없음")
    app_date = m.get("출원일", "없음")
    patent_url = m.get("URL", "")

    if patent_url and patent_url.startswith("http"):
        display_num = f"[{app_num}]({patent_url})"
        display_name = f"[{p_name}]({patent_url})"
    else:
        display_num = app_num
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


def _compress_patent_doc_for_llm(doc: str, max_chars: int, *, claims_first: bool = False):
    """
    토큰 예산 내로 특허 본문 축약.
    claims_first=False: 요약 우선 — 청구항 먼저 생략·축소.
    claims_first=True : 침해 분석용 — 청구항 우선, 요약 먼저 축소.
    """
    parsed = _parse_patent_doc(doc)
    title = parsed["title"] or "특허명칭: (미기재)"
    abstract = parsed["abstract"]
    claims = parsed["claims"]

    def _join(title_line, abs_text, claims_text):
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
        abstract_only = _join(title, abstract, omitted_claims if claims else None)
        if len(abstract_only) <= max_chars:
            return abstract_only, truncated
        prefix = f"{title}\n특허요약: "
        suffix = f"\n특허청구항: {omitted_claims}" if claims else ""
        abs_budget = max_chars - len(prefix) - len(suffix)
        if abs_budget >= 80:
            compressed = prefix + _truncate_text(abstract, abs_budget, suffix="…") + suffix
            if len(compressed) <= max_chars:
                return compressed, truncated
    else:
        if claims:
            claims_only = _join(title, omitted_abstract, claims)
            if len(claims_only) <= max_chars:
                return claims_only, truncated
            prefix = f"{title}\n특허요약: {omitted_abstract}\n특허청구항: "
            claims_budget = max_chars - len(prefix)
            if claims_budget >= 80:
                compressed = prefix + _truncate_text(claims, claims_budget, suffix="…")
                if len(compressed) <= max_chars:
                    return compressed, truncated
        abstract_only = _join(title, abstract, None)
        if len(abstract_only) <= max_chars:
            return abstract_only, truncated

    return _truncate_text(full, max_chars), truncated


def _build_rag_context_for_llm(docs, metas, token_budget, *, claims_first=False):
    """검색된 특허 문서를 Groq 입력 토큰 예산 내로 축소."""
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


def _assemble_llm_prompt(system_prompt: str, context_text: str, user_query: str):
    instruction = "답변 시 참고한 특허 번호·명칭은 [번호](URL) 마크다운 링크 형식을 그대로 유지하세요."
    context_budget = _max_prompt_input_tokens(system_prompt, user_query)
    truncated = False

    if _estimate_tokens(context_text) > context_budget:
        context_text = _truncate_text(context_text, _chars_for_token_budget(context_budget))
        truncated = True

    max_input = GROQ_TPM_LIMIT - GROQ_MAX_OUTPUT_TOKENS - GROQ_REQUEST_MARGIN

    for _ in range(8):
        prompt = (
            f"[SYSTEM] {system_prompt}\n{instruction}\n\n"
            f"[참고 데이터]\n{context_text}\n\n"
            f"[사용자 요청]\n{user_query}\n\n"
            "보고서는 마크다운 양식으로 한국어로 작성하세요."
        )
        if _estimate_tokens(prompt) <= max_input:
            return prompt, truncated, context_text
        truncated = True
        context_text = _truncate_text(context_text, max(200, int(len(context_text) * 0.82)))

    prompt = (
        f"[SYSTEM] {system_prompt}\n{instruction}\n\n"
        f"[참고 데이터]\n{context_text}\n\n"
        f"[사용자 요청]\n{user_query}\n\n"
        "보고서는 마크다운 양식으로 한국어로 작성하세요."
    )
    return prompt, truncated, context_text


# ==========================================
# 12. 엑셀 파싱 · 메타데이터 빌드
# ==========================================
def _find_column(col_map: dict, *keywords: str):
    for kw in keywords:
        for k, v in col_map.items():
            if kw in k:
                return v
    return None


def _detect_columns(df: pd.DataFrame) -> dict:
    """DataFrame 컬럼명을 분석해 각 필드에 해당하는 실제 컬럼명 반환."""
    col_map = {str(c).strip().replace(" ", "").upper(): c for c in df.columns}

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
        "id": id_col,
        "title": _find_column(col_map, "명칭", "제목", "특허명", "INVENTIONTITLE") or None,
        "abstract": _find_column(col_map, "요약", "초록", "ABSTRACT") or None,
        "claims": _find_column(col_map, "청구항", "청구", "범위", "CLAIM") or None,
        "app_date": _find_column(col_map, "출원일", "출원일자", "APPDATE", "FILINGDATE") or None,
        "reg_date": _find_column(col_map, "등록일", "등록일자", "REGDATE") or None,
        "ipc": _find_column(col_map, "IPC") or None,
        "cpc": _find_column(col_map, "CPC") or None,
        "inventor": _find_column(col_map, "발명자", "발명인", "INVENTOR") or None,
        "applicant": _find_column(col_map, "출원인", "권리자", "APPLICANT", "ASSIGNEE") or None,
    }


def _get_cell(row, col, default: str = "없음") -> str:
    if col and pd.notna(row.get(col)):
        val = str(row[col]).strip()
        return val if val else default
    return default


def _get_info_cell(row, col) -> str:
    return _get_cell(row, col, default="정보없음")


def _canonical_applicants_for_meta(applicant_raw: str) -> str:
    """
    출원인 원본 → 대표명 문자열. 복수 출원인은 '|'로 결합.
    (표시·하위호환용. 필터는 _canonical_applicants_list 배열 필드 사용)
    """
    lst = _canonical_applicants_list(applicant_raw)
    return "|".join(lst) if lst else "기타"


def _canonical_applicants_list(applicant_raw: str) -> list:
    """
    출원인 원본 → 대표명 리스트 (중복 제거).
    ChromaDB 배열 메타데이터($contains 필터)용. chromadb>=1.5 필요.
    """
    apps = _explode_multi_values(applicant_raw, allow_comma=False)
    if not apps:
        c = _canonical_applicant_name(applicant_raw)
        return [c] if c else ["기타"]
    canon = []
    for a in apps:
        c = _canonical_applicant_name(a)
        if c and c not in canon:
            canon.append(c)
    return canon if canon else ["기타"]


def _build_metadata(row, cols: dict, patent_url: str = "") -> dict:
    title = _get_info_cell(row, cols["title"]) if cols["title"] else "정보없음"
    applicant_raw = _get_cell(row, cols["applicant"])
    rep_list = _canonical_applicants_list(applicant_raw)
    return {
        "출원번호": str(row[cols["id"]]),
        "명칭": title,
        "출원일": _get_cell(row, cols["app_date"]),
        "등록일": _get_cell(row, cols["reg_date"]),
        "IPC": _get_cell(row, cols["ipc"]),
        "CPC": _get_cell(row, cols["cpc"]),
        "발명자": _get_cell(row, cols["inventor"]),
        "출원인": applicant_raw,
        "대표출원인": "|".join(rep_list),          # 표시·하위호환용 문자열
        "대표출원인목록": rep_list,                 # $contains 필터용 배열 (chromadb>=1.5)
        "URL": patent_url,
    }


def _build_document(row, cols: dict) -> str:
    title = _get_info_cell(row, cols["title"])
    abstract = _get_info_cell(row, cols["abstract"])
    claims = _get_info_cell(row, cols["claims"])
    return f"특허명칭: {title}\n특허요약: {abstract}\n특허청구항: {claims}"


def _patent_url_from_row(row, cols: dict, patent_id: str, hyperlink_map=None) -> str:
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


def extract_excel_hyperlinks(uploaded_file) -> dict:
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


# ==========================================
# 13. ChromaDB 적재
# ==========================================
def _get_chroma_ids(collection) -> set:
    try:
        if collection.count() == 0:
            return set()
        return set(collection.get(include=[])["ids"])
    except Exception as e:
        print(f"[ChromaDB id 조회 오류] {e}")
        return set()


def _upsert_patent_batches(collection, ids, docs, metas, batch_size: int = 100):
    for i in range(0, len(ids), batch_size):
        collection.upsert(
            ids=ids[i:i + batch_size],
            documents=docs[i:i + batch_size],
            metadatas=metas[i:i + batch_size],
        )


def _build_chroma_batches(rows_by_id: dict, cols: dict, hyperlink_map: dict):
    batch_ids, batch_docs, batch_metas = [], [], []
    for doc_id, row in rows_by_id.items():
        patent_url = _patent_url_from_row(row, cols, doc_id, hyperlink_map)
        batch_ids.append(doc_id)
        batch_docs.append(_build_document(row, cols))
        batch_metas.append(_build_metadata(row, cols, patent_url=patent_url))
    return batch_ids, batch_docs, batch_metas


def process_and_update_db(uploaded_file, collection):
    """
    업로드 엑셀 기준 누락 없는 적재 (canonical 출원번호 스키마).
    ChromaDB가 유일한 진실 — 마스터 엑셀 비교/저장 단계 제거.
    적재 성공 후 R2 백업까지 한 번에.

    반환: (신규적재, 파일내중복, 이미존재, 출원번호없음, 백업성공bool)
    """
    file_bytes = uploaded_file.read()
    hyperlink_map = extract_excel_hyperlinks(io.BytesIO(file_bytes))

    try:
        new_df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"엑셀 파일 로드 실패: {e}")
        return 0, 0, 0, 0, False

    cols = _detect_columns(new_df)
    chroma_ids = _get_chroma_ids(collection)

    file_rows: dict = {}
    dup_in_file = 0
    skipped_empty = 0

    for _, row in new_df.iterrows():
        cid = _normalize_patent_id(row[cols["id"]])
        if not cid:
            skipped_empty += 1
            continue
        if cid in file_rows:
            dup_in_file += 1
        file_rows[cid] = row

    needed = {cid: row for cid, row in file_rows.items() if cid not in chroma_ids}
    already_indexed = len(file_rows) - len(needed)

    ingested = 0
    if needed:
        ids, docs, metas = _build_chroma_batches(needed, cols, hyperlink_map)
        try:
            _upsert_patent_batches(collection, ids, docs, metas)
            ingested = len(ids)
        except Exception as e:
            st.error(f"ChromaDB 적재 실패 — 백업하지 않았습니다: {e}")
            return 0, dup_in_file, already_indexed, skipped_empty, False

    # 적재 → 즉시 R2 백업 (트랜잭션처럼 묶음, 실패 시 2회 재시도)
    backup_ok = False
    if ingested:
        with st.spinner("💾 R2에 벡터 DB 백업 중... (재시작 후에도 데이터 보존)"):
            for attempt in range(3):
                backup_ok = backup_chroma_to_r2()
                if backup_ok:
                    break
                if attempt < 2:
                    time.sleep(2)

    return ingested, dup_in_file, already_indexed, skipped_empty, backup_ok


def export_collection_to_excel_bytes(collection) -> bytes:
    """ChromaDB 전체를 엑셀 바이트로 내보내기 (원본 재생성용)."""
    try:
        if collection.count() == 0:
            return b""
        data = collection.get(include=["metadatas", "documents"])
        metas = data.get("metadatas", [])
        docs = data.get("documents", []) or [""] * len(metas)
        rows = []
        for meta, doc in zip(metas, docs):
            meta = meta or {}
            parsed = _parse_patent_doc(doc or "")
            _title = meta.get("명칭", "") or parsed.get("title", "").replace("특허명칭:", "").strip()
            rows.append({
                "출원번호": meta.get("출원번호", ""),
                "명칭": _title,
                "요약": parsed.get("abstract", ""),
                "청구항": parsed.get("claims", ""),
                "출원인": meta.get("출원인", ""),
                "발명자": meta.get("발명자", ""),
                "IPC": meta.get("IPC", ""),
                "CPC": meta.get("CPC", ""),
                "출원일": meta.get("출원일", ""),
                "등록일": meta.get("등록일", ""),
                "URL": meta.get("URL", ""),
            })
        df = pd.DataFrame(rows)
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        return buf.getvalue()
    except Exception as e:
        print(f"[엑셀 내보내기 오류] {e}")
        return b""


# ==========================================
# 13-b. 대표출원인 메타 마이그레이션 + 검색 재순위
# ==========================================
def count_zero_vector_patents(collection) -> tuple:
    """0벡터 특허 수를 센다. 반환: (0벡터 건수, 전체 건수)."""
    try:
        total = collection.count()
        if total == 0:
            return 0, 0
        data = collection.get(include=["embeddings"])
        ids = data.get("ids", [])
        embs = data.get("embeddings", None)
        if embs is None:
            return 0, len(ids)
        zero = 0
        for emb in embs:
            if emb is None or sum(1 for x in emb if abs(x) > 1e-9) == 0:
                zero += 1
        return zero, len(ids)
    except Exception as e:
        print(f"[0벡터 카운트 오류] {e}")
        return 0, 0


def reembed_zero_vector_patents(collection, batch_size: int = 50, max_process: int = 300,
                                progress_cb=None) -> tuple:
    """
    0벡터 특허를 찾아 재임베딩. 이번 호출에서 최대 max_process건만 처리
    (Streamlit ~20분 WebSocket 타임아웃 회피 — 여러 번 나눠 실행).
    반환: (이번에 처리한 건수, 남은 0벡터 건수, 전체 건수)
    """
    try:
        total = collection.count()
        if total == 0:
            return 0, 0, 0
        data = collection.get(include=["documents", "metadatas", "embeddings"])
        ids = data.get("ids", [])
        docs = data.get("documents", [])
        metas = data.get("metadatas", [])
        embs = data.get("embeddings", None)
        if embs is None:
            return 0, 0, len(ids)

        zero_ids, zero_docs, zero_metas = [], [], []
        for cid, doc, meta, emb in zip(ids, docs, metas, embs):
            if emb is None or sum(1 for x in emb if abs(x) > 1e-9) == 0:
                zero_ids.append(cid)
                zero_docs.append(doc if doc else f"특허명칭: {(meta or {}).get('명칭','')}")
                zero_metas.append(meta or {})

        total_zero = len(zero_ids)
        if total_zero == 0:
            return 0, 0, len(ids)

        # 이번 호출에서 처리할 만큼만 자름
        proc_ids = zero_ids[:max_process]
        proc_docs = zero_docs[:max_process]
        proc_metas = zero_metas[:max_process]

        done = 0
        for i in range(0, len(proc_ids), batch_size):
            _upsert_patent_batches(
                collection,
                proc_ids[i:i + batch_size],
                proc_docs[i:i + batch_size],
                proc_metas[i:i + batch_size],
                batch_size=batch_size,
            )
            done += len(proc_ids[i:i + batch_size])
            if progress_cb:
                progress_cb(done, len(proc_ids))
            if i + batch_size < len(proc_ids):
                time.sleep(1.0)

        remaining = total_zero - done
        return done, remaining, len(ids)
    except Exception as e:
        print(f"[0벡터 재임베딩 오류] {e}")
        return 0, -1, 0
    except Exception as e:
        print(f"[0벡터 재임베딩 오류] {e}")
        return 0, 0


def migrate_add_representative_applicant(collection, batch_size: int = 200) -> int:
    """
    기존 적재분(대표출원인 필드 없음)에 대표출원인 메타를 채워 넣음.
    재임베딩 불필요 — 메타데이터만 update. 반환: 갱신 건수.
    """
    try:
        total = collection.count()
        if total == 0:
            return 0
        data = collection.get(include=["metadatas"])
        ids = data.get("ids", [])
        metas = data.get("metadatas", [])

        upd_ids, upd_metas = [], []
        for cid, meta in zip(ids, metas):
            meta = dict(meta or {})
            rep_list = _canonical_applicants_list(meta.get("출원인", ""))
            rep_str = "|".join(rep_list)
            # 문자열 또는 배열 필드가 최신이 아니면 갱신
            if meta.get("대표출원인") != rep_str or meta.get("대표출원인목록") != rep_list:
                meta["대표출원인"] = rep_str
                meta["대표출원인목록"] = rep_list
                upd_ids.append(cid)
                upd_metas.append(meta)

        for i in range(0, len(upd_ids), batch_size):
            collection.update(
                ids=upd_ids[i:i + batch_size],
                metadatas=upd_metas[i:i + batch_size],
            )
        return len(upd_ids)
    except Exception as e:
        print(f"[대표출원인 마이그레이션 오류] {e}")
        return 0


@st.cache_data(show_spinner=False)
def _applicant_options_cached(chroma_n: int, collection_id: str) -> list:
    """드롭다운용 대표출원인 목록 (건수 내림차순). chroma_n 변동 시 갱신."""
    try:
        col = _get_collection()
        if col.count() == 0:
            return []
        metas = col.get(include=["metadatas"]).get("metadatas", [])
        counter: dict = {}
        for m in metas:
            rep = (m or {}).get("대표출원인", "")
            if not rep:
                rep = _canonical_applicants_for_meta((m or {}).get("출원인", ""))
            for name in str(rep).split("|"):
                name = name.strip()
                if name and name != "기타":
                    counter[name] = counter.get(name, 0) + 1
        return [n for n, _ in sorted(counter.items(), key=lambda x: -x[1])]
    except Exception as e:
        print(f"[출원인 목록 조회 오류] {e}")
        return []


def _extract_keywords(query: str) -> list:
    """질의에서 키워드 가산용 토큰 추출 (2자 이상 한글/영문 단어)."""
    tokens = re.findall(r"[가-힣A-Za-z0-9]{2,}", query)
    # 너무 흔한 조사·일반어 제거
    stop = {"특허", "기술", "관련", "분석", "대한", "조사", "검색", "출원", "모듈", "장치", "방법", "시스템"}
    out = []
    for t in tokens:
        # 한글 토큰 끝의 조사(의/을/를/은/는/이/가/와/과/에) 1글자 제거 시도
        if re.search(r"[가-힣]$", t) and len(t) >= 3 and t[-1] in "의을를은는이가와과에":
            t = t[:-1]
        if t and t not in stop:
            out.append(t)
    return out


def _keyword_score(doc: str, meta: dict, keywords: list) -> int:
    """문서·명칭에 키워드가 포함된 정도(가산점). 명칭 매칭에 가중치."""
    if not keywords:
        return 0
    title = str(meta.get("명칭", ""))
    score = 0
    for kw in keywords:
        if kw in title:
            score += 3          # 명칭 매칭 강한 신호
        if kw in doc:
            score += 1          # 본문 매칭
    return score


def search_patents(
    collection,
    query: str,
    n_results: int,
    *,
    applicant_filter: str = "",
    keyword_boost: bool = True,
):
    """
    시맨틱 검색 + (선택)출원인 필터 + 키워드 우선 재순위.

    출원인 필터: chromadb>=1.5의 배열 메타 $contains로 DB 레벨에서 정확히 필터.
      대표출원인목록 = ["효성", "푸른기술"] 형태이므로, 단독·공동출원 모두 매칭됨.
      (구버전 데이터로 배열 필드가 없으면 문자열 부분일치로 폴백)
    반환: (docs, metas) — n_results 건.
    """
    total = max(collection.count(), 1)
    # 질의를 retrieval_query로 임베딩 (비대칭 검색 — 정확도 핵심)
    q_vec = _embed_query_vector(query)

    def _query(where=None, k=None):
        k = min(k or n_results, total)
        try:
            kwargs = {"n_results": k}
            if q_vec is not None:
                kwargs["query_embeddings"] = [q_vec]
            else:
                kwargs["query_texts"] = [query]
            if where:
                kwargs["where"] = where
            r = collection.query(**kwargs)
        except Exception as e:
            print(f"[query 오류] {e}")
            return [], [], []
        if not (r and r["documents"] and r["documents"][0]):
            return [], [], []
        d = r.get("distances", [[None] * len(r["documents"][0])])[0]
        return r["documents"][0], r["metadatas"][0], d

    if applicant_filter:
        triples = []
        # 1순위: 배열 메타 $contains (정확·효율, 단독+공동출원 모두 포착)
        #  후보를 넉넉히(최소 100, n_results*10) 확보해 정확도 손실 방지
        docs, metas, dists = _query(
            where={"대표출원인목록": {"$contains": applicant_filter}},
            k=min(max(n_results * 10, 100), total),
        )
        if docs:
            triples = list(zip(docs, metas, dists))
        else:
            # 폴백: 배열 필드 없는 구버전 데이터 → 대량 fetch 후 문자열 부분일치
            of = min(max(n_results * 20, 300), total)
            d2, m2, ds2 = _query(k=of)
            for doc, meta, dist in zip(d2, m2, ds2):
                lst = meta.get("대표출원인목록")
                if isinstance(lst, list):
                    match = applicant_filter in lst
                else:
                    rep = str(meta.get("대표출원인", "")) or _canonical_applicants_for_meta(meta.get("출원인", ""))
                    match = applicant_filter in {r.strip() for r in rep.split("|")}
                if match:
                    triples.append((doc, meta, dist))
    else:
        of = min(max(n_results * 3, n_results), total)
        docs, metas, dists = _query(k=of)
        triples = list(zip(docs, metas, dists))

    if not triples:
        return [], []

    # 키워드 우선 재순위 (키워드 점수 desc, 그다음 거리 asc)
    if keyword_boost:
        keywords = _extract_keywords(query)
        if keywords:
            def _rank_key(t):
                doc, meta, dist = t
                kw = _keyword_score(doc, meta, keywords)
                d = dist if dist is not None else 1.0
                return (-kw, d)
            triples.sort(key=_rank_key)
    else:
        triples.sort(key=lambda t: t[2] if t[2] is not None else 1.0)

    triples = triples[:n_results]
    return [t[0] for t in triples], [t[1] for t in triples]


# ==========================================
# 14. 메인 포털
# ==========================================
def run_main_portal():
    # ── 세션 최초 진입: 인프라 초기화 전에 R2 복원 먼저 ──
    # (빈 컬렉션이 디스크에 먼저 생기는 것을 방지)
    if not st.session_state.restored:
        download_logo_from_r2()
        # 로컬에 chroma_db가 없을 때만 복원 (있으면 그대로 사용)
        local_empty = not os.path.isdir(CHROMA_DIR) or not os.listdir(CHROMA_DIR)
        if local_empty:
            with st.spinner("🔄 R2에서 벡터 DB 복원 중... (egress 무료)"):
                restored_ok = restore_chroma_from_r2()
            # 복원 후 모든 캐시 초기화 (이전에 만들어진 빈 클라이언트 제거)
            try:
                _get_chroma_client.clear()
                _get_embedding_fn.clear()
            except Exception:
                pass
            st.session_state.infra_initialized = False
            if restored_ok:
                try:
                    n = _get_collection().count()
                    if n > 0:
                        st.toast(f"✅ R2에서 벡터 DB 복원 완료 ({n}건)")
                except Exception as e:
                    print(f"[복원 후 카운트 오류] {e}")
        st.session_state.restored = True

    client, collection, llm = load_permanent_infra_singleton()

    is_admin = st.session_state.is_admin

    # ── 헤더 ──
    col_logo, col_title, col_logout = st.columns([1, 7, 2])
    with col_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=110)
    with col_title:
        st.title("AI 경쟁사 특허 조사 분석")
        mode_label = "🔧 관리자" if is_admin else "👤 사용자"
        used = _daily_tokens_used()
        st.caption(
            f"{mode_label} | 접속: {st.session_state.user_id} | "
            f"적재 특허: {safe_count(collection)}건 | "
            f"오늘 토큰: {used:,}/{GROQ_DAILY_TOKEN_LIMIT:,}"
        )
    with col_logout:
        if st.button("🔒 로그아웃"):
            st.session_state.logged_in = False
            st.session_state.user_id = None
            st.session_state.is_admin = False
            st.rerun()

    # ── 사이드바: 관리자 데이터 관리 ──
    with st.sidebar:
        if is_admin:
            st.header("📂 데이터 관리 센터")

            # R2 연결 진단 (버킷 객체 목록 포함)
            if st.button("🔍 R2 연결 진단", use_container_width=True):
                with st.spinner("진단 중..."):
                    diag = diagnose_r2()
                    objs = list_r2_objects() if diag["reachable"] else []
                if diag["reachable"]:
                    st.success(f"✅ R2 연결 정상\n\n- 버킷: `{diag['bucket']}`")
                    if objs:
                        st.markdown("**📦 버킷 내 파일 목록:**")
                        snapshot_found = False
                        for key, size in objs:
                            mb = size / (1024 * 1024)
                            flag = ""
                            if key == R2_SNAPSHOT_KEY:
                                snapshot_found = True
                                flag = " ← 벡터 스냅샷"
                            st.markdown(f"- `{key}` ({mb:.2f} MB){flag}")
                        if not snapshot_found:
                            st.error(
                                f"⚠️ 벡터 스냅샷(`{R2_SNAPSHOT_KEY}`)이 버킷에 **없습니다**. "
                                f"→ 어제 적재 후 R2 백업이 실패했을 가능성이 큽니다. "
                                f"엑셀을 다시 적재하거나, 로컬에 데이터가 남아 있으면 '💾 R2 백업 재시도'를 누르세요."
                            )
                    else:
                        st.error(
                            "⚠️ 버킷이 **비어 있습니다**. 벡터 스냅샷이 한 번도 저장되지 않았습니다. "
                            "→ 엑셀을 다시 적재하면 적재 직후 자동 백업됩니다."
                        )
                else:
                    st.error(f"❌ 연결 실패\n\n{diag['error']}")

            # R2 → 로컬 강제 복원 (상세 진단)
            if st.button("🔄 R2에서 벡터 DB 강제 복원", use_container_width=True):
                with st.spinner("R2 → 로컬 복원 + 진단 중..."):
                    info = restore_chroma_from_r2_verbose()

                if info["error"]:
                    st.error(f"❌ 복원 실패: {info['error']}")
                elif not info["extracted"]:
                    st.error("❌ 압축 해제 후 chroma_db 디렉토리가 생성되지 않았습니다.")
                else:
                    # 캐시 완전 초기화 후 컬렉션 새로 읽기
                    try:
                        _get_chroma_client.clear()
                        _get_embedding_fn.clear()
                    except Exception:
                        pass
                    st.session_state.infra_initialized = False
                    st.session_state.restored = True

                    # 컬렉션 직접 카운트
                    try:
                        new_col = _get_collection()
                        n = new_col.count()
                    except Exception as e:
                        n = -1
                        st.error(f"컬렉션 읽기 오류: {e}")

                    st.info(
                        f"📦 **복원 진단**\n\n"
                        f"- 다운로드: {info['downloaded_mb']} MB\n"
                        f"- 압축 해제: {'성공' if info['extracted'] else '실패'}\n"
                        f"- chroma.sqlite3 크기: {info['sqlite_mb']} MB\n"
                        f"- 풀린 파일 수: {len(info['files'])}개\n"
                        f"- **로드된 특허 건수: {n}건**"
                    )
                    with st.expander("🔍 tar 내부 구조 / 풀린 파일"):
                        st.write("**tar 멤버:**", info["tar_members"])
                        st.write("**풀린 파일:**", info["files"])

                    if n > 0:
                        st.success(f"✅ {n}건 정상 로드. 새로고침하면 반영됩니다.")
                        st.rerun()
                    elif info["sqlite_mb"] > 1:
                        st.error(
                            "⚠️ sqlite에 데이터는 있는데 0건으로 읽힙니다 "
                            "(chromadb 버전 불일치 가능성). "
                            "→ **sqlite에서 직접 복구**를 시도합니다."
                        )
                        with st.spinner("🛠 sqlite에서 직접 읽어 재임베딩 복구 중... (시간이 걸릴 수 있음)"):
                            try:
                                col2 = _get_collection()
                                rec_n, rec_msg = recover_from_sqlite_directly(col2)
                            except Exception as e:
                                rec_n, rec_msg = 0, str(e)
                        if rec_n > 0:
                            with st.spinner("💾 복구분 R2 재백업 중..."):
                                backup_chroma_to_r2()
                            st.success(f"✅ {rec_msg} R2 재백업 완료. 새로고침하면 반영됩니다.")
                            st.rerun()
                        else:
                            st.error(
                                f"❌ sqlite 직접 복구 실패: {rec_msg}\n\n"
                                "→ requirements.txt에 `chromadb==0.5.23`이 적용됐는지 확인하고, "
                                "그래도 안 되면 원본 엑셀을 다시 적재해야 합니다."
                            )
                    else:
                        st.warning("sqlite가 비어 있습니다 — 백업 시점에 데이터가 없었을 수 있습니다.")

            st.divider()

            # 엑셀 업로드 및 적재
            uploaded_file = st.file_uploader("경쟁사 특허 엑셀 리스트 업로드 (.xlsx)", type=["xlsx"])
            if uploaded_file is not None:
                if st.button("🚀 신규 특허 무결성 적재"):
                    try:
                        preview_bytes = uploaded_file.read()
                        uploaded_file.seek(0)
                        preview_df = pd.read_excel(io.BytesIO(preview_bytes), nrows=3)
                        detected = _detect_columns(preview_df)
                        total_rows = pd.read_excel(io.BytesIO(preview_bytes)).shape[0]
                        uploaded_file.seek(0)
                        with st.expander("📋 업로드 파일 열 감지 결과", expanded=True):
                            st.write(f"- **전체 행 수:** {total_rows}행")
                            st.write(f"- **감지된 출원번호 열:** `{detected['id']}`")
                            st.write(f"- **감지된 명칭 열:** `{detected['title']}`")
                            st.write(f"- **전체 열 목록:** {list(preview_df.columns)}")
                    except Exception as diag_e:
                        st.warning(f"파일 사전 진단 실패: {diag_e}")

                    with st.spinner("중복 제거 및 실시간 인덱싱 중... (임베딩 호출, 건수에 따라 시간 소요)"):
                        ingested, dup, already, skipped, backup_ok = process_and_update_db(
                            uploaded_file, collection
                        )

                    st.info(
                        f"📊 처리 결과: 신규 **{ingested}건** / 파일내 중복 **{dup}건** / "
                        f"이미 존재 **{already}건** / 번호없음 **{skipped}건** / "
                        f"DB 총 **{safe_count(collection)}건**"
                    )
                    if ingested > 0:
                        if backup_ok:
                            st.success(
                                f"✅ 신규 {ingested}건 인덱싱 + R2 백업 성공 — "
                                f"재시작·슬립 후에도 데이터가 보존됩니다."
                            )
                        else:
                            st.error(
                                f"⚠️ 신규 {ingested}건 인덱싱은 됐으나 **R2 백업 실패**. "
                                f"'🔍 R2 연결 진단' 후 다시 적재하거나 아래 백업 버튼을 누르세요."
                            )
                    elif dup > 0 or already > 0:
                        st.warning("업로드 특허가 이미 DB에 존재합니다. 새 데이터가 없습니다.")
                    else:
                        st.error("처리된 데이터가 없습니다. 출원번호 열 감지 결과를 확인하세요.")
                    st.rerun()

            st.divider()
            st.markdown(f"📊 **누적 적재 데이터:** `{safe_count(collection)}` 건")

            # R2 수동 백업
            if st.button("💾 R2 백업 재시도", use_container_width=True):
                with st.spinner("R2에 벡터 DB 업로드 중..."):
                    ok = backup_chroma_to_r2()
                if ok:
                    st.success("✅ R2 백업 성공!")
                else:
                    st.error("❌ 백업 실패. R2 연결 진단을 확인하세요.")

            # 대표출원인 메타 마이그레이션 (기존 적재분에 검색 필터 필드 추가)
            if st.button("🏷 대표출원인 필드 갱신 (검색 필터용)", use_container_width=True,
                         help="기존 적재 특허에 '대표출원인' 메타를 채웁니다. 재임베딩 없이 메타만 갱신 → 빠름. "
                              "효성·히타치 등 변형 통합 규칙이 바뀐 경우에도 다시 누르세요."):
                with st.spinner("대표출원인 메타 갱신 중... (재임베딩 없음)"):
                    n_upd = migrate_add_representative_applicant(collection)
                    backup_chroma_to_r2()
                _applicant_options_cached.clear()
                st.success(f"✅ {n_upd}건 대표출원인 갱신 + R2 백업 완료. 출원인 드롭다운에 반영됩니다.")
                st.rerun()

            # 출원인 필터 진단 (특정 출원인이 검색 안 될 때 원인 파악)
            with st.expander("🔬 출원인 필터 진단"):
                diag_name = st.text_input("진단할 대표출원인명", value="효성", key="diag_applicant")
                if st.button("진단 실행", key="run_applicant_diag"):
                    try:
                        # 1) 전체 메타에서 대표출원인 값 분포
                        all_metas = collection.get(include=["metadatas"]).get("metadatas", [])
                        exact = 0        # 대표출원인 == diag_name
                        contains = 0     # '|' 결합 포함
                        raw_samples = set()
                        for m in all_metas:
                            rep = str((m or {}).get("대표출원인", ""))
                            reps = {r.strip() for r in rep.split("|")}
                            if diag_name in reps:
                                contains += 1
                                if rep.strip() == diag_name:
                                    exact += 1
                            # 원본 출원인 표기 샘플 수집 (해당 대표명 관련)
                            raw = str((m or {}).get("출원인", ""))
                            if diag_name in _canonical_applicants_for_meta(raw):
                                if len(raw_samples) < 8:
                                    raw_samples.add(raw)

                        st.write(f"**'{diag_name}' 대표출원인 통계:**")
                        st.write(f"- 단독 대표출원인 일치: **{exact}건**")
                        st.write(f"- 복수출원인('|') 포함: **{contains}건**")

                        # 2) $contains 배열 필터 직접 테스트
                        try:
                            wres = collection.get(
                                where={"대표출원인목록": {"$contains": diag_name}}, limit=10000
                            )
                            n_contains = len(wres.get("ids", []))
                            st.write(f"- **$contains 배열 필터 매칭: {n_contains}건** ← 실제 검색에 쓰이는 방식")
                            if n_contains == 0 and contains > 0:
                                st.warning(
                                    "⚠️ 문자열엔 있으나 배열 필터가 0건 → **배열 필드가 아직 없습니다**. "
                                    "'🏷 대표출원인 필드 갱신'을 누르면 배열 필드가 채워집니다."
                                )
                        except Exception as we:
                            st.warning(
                                f"$contains 필터 오류: {we}\n\n"
                                "→ chromadb가 1.5 미만이거나 배열 필드가 없습니다. "
                                "requirements.txt의 `chromadb==1.5.9` 확인 후 '대표출원인 필드 갱신'을 누르세요."
                            )

                        # 3) 원본 출원인 표기 샘플
                        if raw_samples:
                            st.write("**원본 출원인 표기 샘플:**")
                            for s in raw_samples:
                                st.write(f"  - `{s}` → 대표명: `{_canonical_applicants_for_meta(s)}`")

                        if contains == 0:
                            st.error(
                                f"⚠️ '{diag_name}'로 저장된 특허가 0건입니다. "
                                f"대표출원인 필드가 아직 안 채워졌을 수 있습니다 → "
                                f"'🏷 대표출원인 필드 갱신'을 먼저 누르세요."
                            )
                        elif exact == 0 and contains > 0:
                            st.warning(
                                f"'{diag_name}'가 모두 복수출원인('|' 결합)으로만 존재합니다. "
                                f"where 정확매칭은 0건이지만 하이브리드 검색으로 조회됩니다."
                            )
                    except Exception as e:
                        st.error(f"진단 오류: {e}")

            # 검색 경로 진단 (특정 특허가 검색 안 될 때 원인 추적)
            with st.expander("🔎 특허 검색 경로 진단"):
                st.caption("검색이 안 되는 특허의 출원번호를 넣으면 저장·임베딩·검색을 단계별로 추적합니다.")
                diag_pid = st.text_input("출원번호 (마스터 엑셀의 값)", key="diag_pid")
                if st.button("검색 경로 추적", key="run_search_diag") and diag_pid.strip():
                    try:
                        cid = _normalize_patent_id(diag_pid.strip())
                        st.write(f"정규화된 ID: `{cid}`")

                        # 1) 해당 ID가 컬렉션에 있는지 + 문서/메타/임베딩 확인
                        got = collection.get(ids=[cid], include=["documents", "metadatas", "embeddings"])
                        if not got.get("ids"):
                            st.error(
                                f"❌ ID `{cid}`가 컬렉션에 없습니다. "
                                f"정규화 규칙 때문에 저장된 ID와 다를 수 있습니다. "
                                f"엑셀의 출원번호 원본 표기를 확인하세요."
                            )
                        else:
                            doc = got["documents"][0]
                            emb = got["embeddings"][0] if got.get("embeddings") is not None else None
                            st.success(f"✅ ID `{cid}` 존재")
                            st.write(f"- 문서 길이: {len(doc)}자")
                            st.write(f"- 문서 앞 120자: `{doc[:120]}`")

                            # 2) 임베딩 유효성 (0벡터면 적재 시 임베딩 실패한 것)
                            if emb is not None:
                                import math
                                norm = math.sqrt(sum(x * x for x in emb))
                                nonzero = sum(1 for x in emb if abs(x) > 1e-9)
                                st.write(f"- 임베딩 차원: {len(emb)}, 비영(非零) 성분: {nonzero}, 노름: {norm:.4f}")
                                if norm < 1e-6 or nonzero == 0:
                                    st.error(
                                        "🚨 **이 특허의 임베딩이 0벡터입니다.** "
                                        "적재 시 Gemini 임베딩 호출이 실패해 0벡터로 저장됐습니다. "
                                        "→ 이 특허(및 유사 케이스)를 재임베딩해야 합니다. "
                                        "아래 '🔧 0벡터 특허 재임베딩'을 실행하세요."
                                    )
                                else:
                                    st.info("임베딩은 정상(비영 벡터)입니다.")

                            # 3) 문서 자체로 검색 시 몇 위에 나오는지 (retrieval_query)
                            st.write("---")
                            st.write("**이 특허의 문서 내용으로 검색 시 순위:**")
                            qtext = doc[:500]
                            qv = _embed_query_vector(qtext)
                            if qv is not None:
                                rq = collection.query(query_embeddings=[qv], n_results=10)
                            else:
                                rq = collection.query(query_texts=[qtext], n_results=10)
                            found_rank = None
                            for rank, m in enumerate(rq["metadatas"][0], 1):
                                if _normalize_patent_id(m.get("출원번호", "")) == cid:
                                    found_rank = rank
                                    break
                            if found_rank:
                                st.success(f"✅ 자기 문서로 검색 시 **{found_rank}위**에 나옵니다. 검색 경로 정상.")
                            else:
                                st.error(
                                    "🚨 자기 문서로 검색해도 상위 10위 안에 안 나옵니다. "
                                    "임베딩이 0벡터이거나, 저장 임베딩과 질의 임베딩 방식이 다릅니다."
                                )
                                # 상위 10위가 뭔지 표시
                                st.write("상위 10위 특허:")
                                for rank, m in enumerate(rq["metadatas"][0], 1):
                                    st.write(f"  {rank}. {m.get('출원번호','')} | {m.get('명칭','')[:30]}")
                    except Exception as e:
                        st.error(f"진단 오류: {e}")

            # 0벡터 특허 재임베딩 (임베딩 실패분 복구) — 청크 처리로 타임아웃 회피
            st.markdown("**🔧 0벡터 특허 복구**")
            st.caption(
                "적재 시 rate limit으로 임베딩 실패한 특허를 재임베딩합니다. "
                "Streamlit 20분 제한 때문에 한 번에 300건씩 처리 → 여러 번 나눠 누르세요."
            )

            # 임베딩 엔진 실시간 진단 (전량 0벡터일 때 원인 파악)
            if st.button("🩺 임베딩 엔진 진단 (먼저 실행)", use_container_width=True):
                st.write("**1. 설정 확인**")
                gkey = _clean_ascii(st.secrets.get("GEMINI_API_KEY", ""))
                st.write(f"- GEMINI_API_KEY 존재: {'✅ 예 (길이 ' + str(len(gkey)) + ')' if gkey else '❌ 없음'}")
                st.write(f"- google-generativeai 설치: {'✅' if _GENAI_AVAILABLE else '❌'}")

                fn = _get_embedding_fn()
                fn_type = type(fn).__name__
                st.write(f"- 현재 임베딩 함수: `{fn_type}`")
                if fn_type != "GeminiEmbeddingFunction":
                    st.error(
                        "⚠️ Gemini가 아닌 폴백(로컬) 임베딩이 선택됐습니다. "
                        "GEMINI_API_KEY 또는 google-generativeai 설치를 확인하세요."
                    )

                st.write("**2. 실제 임베딩 호출 테스트**")
                try:
                    if fn_type == "GeminiEmbeddingFunction":
                        # 재시도 없이 1회만 직접 호출해 raw 에러 확인
                        import google.generativeai as _genai
                        _genai.configure(api_key=gkey)
                        resp = _genai.embed_content(
                            model="models/gemini-embedding-001",
                            content="지폐 계수 장치 테스트",
                            task_type="retrieval_document",
                            output_dimensionality=_EMBED_DIM,
                        )
                        vec = resp["embedding"]
                        nonzero = sum(1 for x in vec if abs(x) > 1e-9)
                        st.write(f"- 반환 벡터 차원: {len(vec)}, 비영 성분: {nonzero}")
                        if nonzero > 0:
                            st.success("✅ Gemini 임베딩 정상 작동! 이제 재임베딩하면 0벡터가 채워집니다.")
                        else:
                            st.error("❌ 호출은 됐으나 0벡터 반환. 모델/키 상태 이상.")
                    else:
                        test_vec = fn(["지폐 계수 장치 테스트"])[0]
                        nonzero = sum(1 for x in test_vec if abs(x) > 1e-9)
                        st.write(f"- 로컬 임베딩 비영 성분: {nonzero}")
                        st.info("로컬 임베딩 사용 중. Gemini로 바꾸려면 GEMINI_API_KEY 설정 필요.")
                except Exception as e:
                    st.error(
                        f"❌ **임베딩 호출 실패 (이게 전량 0벡터의 원인):**\n\n```\n{e}\n```\n\n"
                        "→ 에러 내용에 따라: API 키 만료·무효, 할당량 소진(quota), "
                        "또는 google-generativeai 버전 문제일 수 있습니다."
                    )

            col_chk, col_fix = st.columns(2)
            with col_chk:
                if st.button("🔍 0벡터 개수 확인", use_container_width=True):
                    with st.spinner("스캔 중..."):
                        zn, tn = count_zero_vector_patents(collection)
                    if zn == 0:
                        st.success(f"✅ 0벡터 특허 없음 (전체 {tn}건 정상)")
                    else:
                        st.warning(f"⚠️ 0벡터 특허 **{zn}건** / 전체 {tn}건. 아래 재임베딩을 실행하세요.")
            with col_fix:
                chunk = st.number_input("한 번에 처리할 건수", min_value=100, max_value=1000,
                                        value=400, step=100,
                                        help="20분 타임아웃 내에 끝날 만큼. 400~500 권장.")
                if st.button("🔧 0벡터 재임베딩", use_container_width=True):
                    prog = st.progress(0.0, text="재임베딩 준비 중...")

                    def _cb(done, tot):
                        prog.progress(min(done / max(tot, 1), 1.0),
                                      text=f"재임베딩 {done}/{tot}건...")

                    done, remaining, tn = reembed_zero_vector_patents(
                        collection, max_process=int(chunk), progress_cb=_cb
                    )
                    prog.empty()
                    if done > 0:
                        with st.spinner("💾 R2 백업 중..."):
                            backup_chroma_to_r2()
                        _applicant_options_cached.clear()
                        if remaining > 0:
                            st.warning(
                                f"✅ 이번에 {done}건 재임베딩 완료. "
                                f"**아직 {remaining}건 남음** → 버튼을 다시 눌러 계속하세요."
                            )
                        else:
                            st.success(f"🎉 {done}건 재임베딩 완료. 0벡터 특허가 모두 복구됐습니다!")
                        st.rerun()
                    elif remaining == -1:
                        st.error("재임베딩 중 오류. 로그를 확인하세요.")
                    else:
                        st.info("0벡터 특허가 없습니다. 모두 정상입니다.")

            st.divider()

            # 엑셀 내보내기 (원본 재생성)
            if safe_count(collection) > 0:
                if st.button("📥 마스터 엑셀 준비", use_container_width=True):
                    with st.spinner("엑셀 생성 중..."):
                        xlsx_bytes = export_collection_to_excel_bytes(collection)
                    if xlsx_bytes:
                        st.download_button(
                            "⬇️ 다운로드 (요약·청구항 포함)",
                            data=xlsx_bytes,
                            file_name="master_patents.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )

            st.divider()

            # 전체 포맷
            also_clear_r2 = st.checkbox(
                "R2 백업도 함께 삭제 (snapshot 제거)",
                value=False,
                help="체크 시 R2의 벡터 스냅샷도 삭제합니다. 완전히 새로 시작할 때 사용하세요.",
            )
            if st.button("🚨 데이터 웨어하우스 전체 포맷"):
                with st.spinner("⏳ 벡터 DB 완전 초기화 중..."):
                    try:
                        reset_collection()
                        if also_clear_r2:
                            r2_delete(R2_SNAPSHOT_KEY)
                        st.session_state["_last_chroma_count"] = 0
                        st.session_state.restored = True  # 포맷 직후 R2 재복원 방지
                        if also_clear_r2:
                            st.toast("✅ 로컬 + R2 초기화 완료. 엑셀을 새로 업로드하세요.")
                        else:
                            st.toast("✅ 로컬 초기화 완료. (R2 백업은 유지)")
                        st.rerun()
                    except Exception as e:
                        st.error(f"초기화 중 오류: {e}")

            st.divider()
            st.subheader("👥 가입 회원 현황")
            st.caption("🔴 승인 대기 계정은 **활성화** 버튼으로 로그인을 허용합니다.")
            registry = load_user_registry()
            if not registry:
                st.caption("등록된 일반 회원 없음")
            else:
                updated_registry = dict(registry)
                for uid, info in registry.items():
                    is_active = info.get("active", True)
                    status_icon = "🟢" if is_active else "🔴"
                    status_note = "" if is_active else " · **승인 대기**"
                    btn_label = "비활성화" if is_active else "활성화 (승인)"
                    btn_type = "secondary" if is_active else "primary"
                    col_info, col_btn = st.columns([3, 1])
                    with col_info:
                        st.markdown(
                            f"{status_icon} **{uid}**  \n"
                            f"<span style='font-size:12px;color:gray'>"
                            f"{info.get('name','')} · {info.get('department','부서없음')} · "
                            f"{info.get('registered_at','')[:10]}{status_note}</span>",
                            unsafe_allow_html=True,
                        )
                    with col_btn:
                        if st.button(btn_label, key=f"toggle_{uid}", type=btn_type):
                            updated_registry[uid]["active"] = not is_active
                            save_user_registry(updated_registry)
                            upload_user_registry_to_r2()
                            refresh_user_registry_from_r2(force=True)
                            action = "활성화(승인)" if not is_active else "비활성화"
                            st.toast(f"✅ {uid} 계정을 {action}했습니다.")
                            st.rerun()
                    st.divider()
        else:
            st.caption("분석 기능 전용 접속 모드입니다.")

    # ── 메인 분석 UI ──
    st.subheader("⚙️ 1단계: AI 전문가 선택")
    analysis_mode = st.selectbox(
        "사용 목적에 맞는 전문가 관점을 선택해 주세요:",
        [
            "💡 단순 키워드 매칭 및 특허 검색",
            "🔬 특정 기술 관련 심층 특허 분석",
            "🛡 개발기술 침해 분석 & 진보성 회피 설계",
            "📊 출원정보 기반 다차원 통계조사 (출원인, 발명자, IPC, 일자 등)",
        ],
    )

    st.subheader("🔍 2단계: 검색 키워드 또는 질의 내용 입력")
    placeholders = {
        "💡 단순 키워드 매칭 및 특허 검색": "검색하고자 하는 핵심 키워드들을 입력하세요. (예: 카세트 도어 잠금장치)",
        "🔬 특정 기술 관련 심층 특허 분석": "동향을 파악할 타겟 기술이나 모듈명을 입력하세요. (예: 센서 기반 매체 지폐 잼 장애 예측 알고리즘)",
        "🛡 개발기술 침해 분석 & 진보성 회피 설계": "우리가 출원 예정이거나 개발한 기술 아이디어를 청구항 수준으로 상세히 입력하세요.",
        "📊 출원정보 기반 다차원 통계조사 (출원인, 발명자, IPC, 일자 등)": (
            "통계 요약 조건을 입력하세요. (예: '2021년 이후 출원인별 기술 동향', '전체 통계 요약')"
        ),
    }
    user_query = st.text_area("분석 대상 내용을 입력하세요:", height=110, placeholder=placeholders[analysis_mode])

    if "📊" not in analysis_mode:
        default_n = 7 if "🛡" in analysis_mode else 5
        n_results_user = st.slider(
            "🔢 3단계: 참조할 관련 특허 수",
            min_value=3, max_value=20, value=default_n, step=1,
            help="AI가 분석에 참조할 최대 특허 건수. Groq 입력 한도(12,000 TPM) 때문에 수가 많으면 본문이 자동 축약됩니다.",
        )
        # 출원인 한정 드롭다운 (경쟁사 특허 조사 정확도 향상)
        applicant_opts = _applicant_options_cached(safe_count(collection), _COLLECTION_NAME)
        applicant_filter = st.selectbox(
            "🏢 4단계: 출원인 한정 (선택)",
            ["(전체)"] + applicant_opts,
            help="특정 경쟁사를 고르면 그 출원인의 특허 안에서만 검색합니다. "
                 "예: '효성'을 고르고 '현금 수표 통합 모듈' 입력 → 효성 특허만 조사.",
        )
        applicant_filter = "" if applicant_filter == "(전체)" else applicant_filter
        keyword_boost = st.checkbox(
            "🔑 키워드 우선 매칭 (질의 단어가 명칭·본문에 포함된 특허를 상위로)",
            value=True,
            help="예: '수표' 입력 시 '수표'가 실제 포함된 특허를 의미만 비슷한 특허보다 우선합니다.",
        )
    else:
        n_results_user = 10
        applicant_filter = ""
        keyword_boost = False
        st.caption(
            "💡 **출원연도·출원인별 전체 동향** 분석(예: '2021년 이후 출원인별 기술')은 "
            "10~20건 제한 없이 DB 전체를 집계합니다."
        )

    # 일일 토큰 경고
    used = _daily_tokens_used()
    if used >= GROQ_DAILY_TOKEN_LIMIT * 0.8:
        st.warning(
            f"⚠️ 오늘 사용 토큰 {used:,} / {GROQ_DAILY_TOKEN_LIMIT:,} (80% 초과). "
            f"무료 한도 소진이 가까워 분석이 곧 제한될 수 있습니다."
        )

    if st.button("🧬 가상 전문가 엔진 구동"):
        if user_query.strip() == "":
            st.warning("분석 내용을 입력해 주세요.")
        elif safe_count(collection) == 0:
            st.error("서버 DB에 적재된 특허 소스가 없습니다. 좌측 메뉴에서 엑셀을 먼저 등록해 주세요.")
        else:
            with st.spinner("가상 전문가가 시맨틱 문헌 대조 및 추론을 진행 중입니다..."):
                context_truncated = False

                # ── 통계 모드: 전체 메타데이터 집계 + 복합 프롬프트 ──
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
                        n=15, label="출원인(대표명)",
                        normalizer=_canonical_applicant_name, explode=True,
                    )
                    ipc_df, ipc_stat = _top_counts_table(
                        analysis_df.get("IPC", pd.Series(dtype=str)),
                        n=10, label="IPC", explode=True, allow_comma=True,
                    )
                    inventor_df, inventor_stat = _top_counts_table(
                        analysis_df.get("발명자", pd.Series(dtype=str)),
                        n=10, label="발명자", explode=True, allow_comma=True,
                    )
                    year_df, year_stat = _year_counts_table(
                        analysis_df.get("출원일", pd.Series(dtype=str))
                    )

                    scope_label = f"조건 필터 ({filtered_count}건)" if use_year_filter else "전체 DB"
                    st.markdown(f"### 📊 사전 집계 통계 ({scope_label})")
                    sc1, sc2 = st.columns(2)
                    with sc1:
                        st.markdown("**출원인별 (대표명 통합)**")
                        st.dataframe(applicant_df, use_container_width=True, hide_index=True)
                        st.markdown("**IPC 분류별**")
                        st.dataframe(ipc_df, use_container_width=True, hide_index=True)
                    with sc2:
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
                            f"\n■ 출원인별 기술 코호트 (필터 범위 전체 — 샘플 10건 아님)\n{cohort_md}\n"
                        )
                    elif not cohort_mode:
                        _qv = _embed_query_vector(user_query.strip())
                        _sk = min(filtered_count if use_year_filter else total_count, 10)
                        if _qv is not None:
                            sem_results = collection.query(query_embeddings=[_qv], n_results=_sk)
                        else:
                            sem_results = collection.query(query_texts=[user_query.strip()], n_results=_sk)
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

                # ── 일반 모드: 시맨틱 RAG (+ 출원인 필터 + 키워드 우선) ──
                else:
                    n_results = min(n_results_user, safe_count(collection))
                    retrieved_docs, retrieved_metas = search_patents(
                        collection,
                        user_query.strip(),
                        n_results,
                        applicant_filter=applicant_filter,
                        keyword_boost=keyword_boost,
                    )
                    if not retrieved_docs:
                        if applicant_filter:
                            st.error(
                                f"'{applicant_filter}' 출원인 특허 중 관련 결과를 찾지 못했습니다. "
                                f"출원인 한정을 '(전체)'로 바꾸거나 키워드를 조정해 보세요."
                            )
                        else:
                            st.error("관련 특허를 찾지 못했습니다. 다른 키워드로 시도해 보세요.")
                        st.stop()

                    if applicant_filter:
                        st.info(f"🏢 '{applicant_filter}' 출원인으로 한정해 {len(retrieved_docs)}건을 조사했습니다.")

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
                            st.info(
                                "Groq 입력 한도(12,000 TPM)에 맞춰 본문을 축약했습니다. "
                                "침해 분석 모드는 **청구항 우선** 유지, 요약을 먼저 줄입니다. "
                                "더 상세하려면 '참조 특허 수'를 줄여 보세요."
                            )
                        else:
                            st.info(
                                "Groq 입력 한도(12,000 TPM)에 맞춰 본문을 축약했습니다. "
                                "키워드·심층 분석 모드는 **요약 우선** 유지, 청구항을 먼저 줄입니다. "
                                "더 상세하려면 '참조 특허 수'를 줄여 보세요."
                            )
                    # 일일 토큰 추정 누적
                    _add_daily_tokens(_estimate_tokens(prompt) + GROQ_MAX_OUTPUT_TOKENS)

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
