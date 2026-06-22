import streamlit as st
import pandas as pd
import os
import chromadb
from chromadb.utils import embedding_functions
from langchain_groq import ChatGroq
import openpyxl
import json
import base64
import shutil
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
DB_PATH = "/tmp/my_patent_vector_db"

os.makedirs(os.path.join(BASE_DIR, "my_patent_folder"), exist_ok=True)
os.makedirs(DB_PATH, exist_ok=True)

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


# ==========================================
# [인프라 무결성 안착] GitHub API 강제 업로드 엔진
# ==========================================
def _get_github_secrets():
    """
    TOML 최상위 키로 설정된 GITHUB_TOKEN / GITHUB_REPO_URL을 안전하게 반환.
    섹션 헤더([USER_CREDENTIALS]) 뒤에 배치하면 해당 섹션 하위에 귀속되어
    st.secrets["GITHUB_TOKEN"]으로 접근이 불가능해지므로, 반드시 TOML 최상위에 위치해야 함.
    """
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_url = st.secrets["GITHUB_REPO_URL"]
        if not token or token.startswith("ghp_본인의"):
            return None, None
        # https:// 누락 방어
        if not repo_url.startswith("https://"):
            repo_url = "https://" + repo_url.lstrip("http://")
        return token, repo_url
    except Exception:
        return None, None


def upload_file_to_github_api(local_file_path, github_target_path):
    token, repo_url = _get_github_secrets()
    if not token:
        return False
        
    raw_url = repo_url.replace(".git", "")
    repo_path = raw_url.split("github.com/")[-1]

    if not os.path.exists(local_file_path):
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
            with urlopen(req_get) as response:
                res_data = json.loads(response.read().decode())
                sha = res_data.get("sha")
        except HTTPError as e:
            if e.code != 404:
                print(f"[API Warning] SHA 획득 생략")

        payload = {
            "message": f"🔄 [Automated API Warehouse Sync] {github_target_path}",
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
        
        with urlopen(req_put) as response:
            if response.status in [200, 201]:
                return True
    except Exception as e:
        print(f"GitHub API 통신 중 예외 제어: {e}")
    return False

def commit_and_push_data():
    """마스터 엑셀을 GitHub에 업로드하여 컨테이너 재시작 후에도 데이터가 복원될 수 있도록 보존."""
    excel_status = upload_file_to_github_api(MASTER_EXCEL_PATH, "my_patent_folder/master_patents.xlsx")
    if excel_status:
        st.toast("💾 GitHub 데이터 웨어하우스 동기화 완료 — 재시작 후 자동 복원 보장!")
    else:
        st.toast("⚠️ GitHub 동기화 실패 — Secrets에서 GITHUB_TOKEN을 TOML 최상위(섹션 헤더 위)에 배치했는지 확인하세요.")


# ==========================================
# [회원 관리] 사용자 레지스트리 엔진
# ==========================================
def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


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


def download_user_registry_from_github() -> bool:
    token, repo_url = _get_github_secrets()
    if not token:
        return False
    try:
        raw_url   = repo_url.replace(".git", "")
        repo_path = raw_url.split("github.com/")[-1]
        api_url   = f"https://api.github.com/repos/{repo_path}/contents/my_patent_folder/user_registry.json"
        req = Request(api_url, headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json"
        })
        with urlopen(req) as resp:
            data = json.loads(resp.read().decode())
        if data.get("content"):
            raw = base64.b64decode(data["content"])
        elif data.get("download_url"):
            with urlopen(Request(data["download_url"], headers={"Authorization": f"Bearer {token}"})) as r:
                raw = r.read()
        else:
            return False
        os.makedirs(os.path.dirname(USER_REGISTRY_PATH), exist_ok=True)
        with open(USER_REGISTRY_PATH, "wb") as f:
            f.write(raw)
        return True
    except Exception as e:
        print(f"사용자 레지스트리 복원 실패: {e}")
        return False


def upload_user_registry_to_github() -> bool:
    return upload_file_to_github_api(USER_REGISTRY_PATH, "my_patent_folder/user_registry.json")


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
    token, repo_url = _get_github_secrets()
    if not token:
        return False
    try:
        raw_url   = repo_url.replace(".git", "")
        repo_path = raw_url.split("github.com/")[-1]
        api_url   = f"https://api.github.com/repos/{repo_path}/contents/atec_logo.png"
        req = Request(api_url, headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json"
        })
        with urlopen(req) as resp:
            data = json.loads(resp.read().decode())
        if data.get("content"):
            raw = base64.b64decode(data["content"])
        elif data.get("download_url"):
            with urlopen(Request(data["download_url"], headers={"Authorization": f"Bearer {token}"})) as r:
                raw = r.read()
        else:
            return False
        with open(logo_path, "wb") as f:
            f.write(raw)
        return True
    except Exception as e:
        print(f"로고 복원 실패: {e}")
        return False


def download_master_excel_from_github():
    """
    컨테이너 재시작으로 로컬 파일이 소실된 경우 GitHub에서 마스터 엑셀을 내려받아 복원.
    성공 여부를 bool로 반환.
    """
    token, repo_url = _get_github_secrets()
    if not token:
        return False
    try:
        raw_url = repo_url.replace(".git", "")
        repo_path = raw_url.split("github.com/")[-1]
        api_url = f"https://api.github.com/repos/{repo_path}/contents/my_patent_folder/master_patents.xlsx"
        req = Request(api_url, headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json"
        })
        with urlopen(req) as response:
            data = json.loads(response.read().decode())

        # GitHub Contents API는 1MB 초과 파일의 content 필드를 비워 반환
        # → download_url로 폴백하여 대용량 엑셀도 안전하게 처리
        if data.get("content"):
            raw_content = base64.b64decode(data["content"])
        elif data.get("download_url"):
            dl_req = Request(data["download_url"], headers={
                "Authorization": f"Bearer {token}"
            })
            with urlopen(dl_req) as dl_resp:
                raw_content = dl_resp.read()
        else:
            return False

        os.makedirs(os.path.dirname(MASTER_EXCEL_PATH), exist_ok=True)
        with open(MASTER_EXCEL_PATH, "wb") as f:
            f.write(raw_content)
        return True
    except Exception as e:
        print(f"GitHub 마스터 엑셀 복원 실패: {e}")
        return False


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

            # 관리자 계정 확인 (Secrets 기반 평문 비교)
            if username in admin_creds and admin_creds[username] == password:
                st.session_state.logged_in = True
                st.session_state.user_id   = username
                st.session_state.is_admin  = True
                st.rerun()
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
@st.cache_resource
def _build_infra():
    """
    @st.cache_resource: 프로세스당 단 한 번만 실행되어 동일 인스턴스를 모든 세션·리런에서 재사용.
    PersistentClient를 중복 생성하려 할 때 발생하는
    'An instance of Chroma already exists' ValueError를 구조적으로 원천 차단.
    """
    # ChromaDB 클라이언트 — 실패 시 디렉토리 정화 후 재시도
    try:
        chroma_client = chromadb.PersistentClient(path=DB_PATH)
    except Exception:
        if os.path.exists(DB_PATH):
            shutil.rmtree(DB_PATH)
        os.makedirs(DB_PATH, exist_ok=True)
        chroma_client = chromadb.PersistentClient(path=DB_PATH)

    # Ko-SRoBERTa 임베딩 함수 및 컬렉션
    sentence_transformer_ef = embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name="jhgan/ko-sroberta-multitask"
    )
    collection = chroma_client.get_or_create_collection(
        name="competitor_patents",
        embedding_function=sentence_transformer_ef
    )

    # Groq Cloud API 기반 Llama 3.3 엔진 (키는 반드시 Streamlit Secrets에서 로드)
    groq_api_key = st.secrets.get("GROQ_API_KEY", "")
    if not groq_api_key:
        raise ValueError("Streamlit Secrets에 GROQ_API_KEY가 설정되어 있지 않습니다.")
    llm = ChatGroq(
        model="llama-3.3-70b-versatile",
        groq_api_key=groq_api_key,
        temperature=0.1
    )

    return chroma_client, collection, llm


def load_permanent_infra_singleton():
    # 최초 1회만 스피너 표시 — 이후 캐시 히트 시 즉시 반환
    if "infra_initialized" not in st.session_state:
        with st.spinner("📦 가상 특허 가동 커널 및 AI 전문 임베딩 엔진 초기화 중..."):
            result = _build_infra()
        st.session_state.infra_initialized = True
        return result
    return _build_infra()


# --- 3. 엑셀 파싱 및 무결성 메타데이터 적재 로직 ---
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


def process_and_update_db(uploaded_file, collection):
    import copy
    file_for_links = copy.deepcopy(uploaded_file)
    hyperlink_map = extract_excel_hyperlinks(file_for_links)
    
    try:
        new_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"엑셀 파일 로드 실패: {e}")
        return 0, 0
    
    columns_map = {str(col).strip().replace(" ", "").upper(): col for col in new_df.columns}
    
    id_col = next((v for k, v in columns_map.items() if "출원번호" in k or "번호" in k), new_df.columns[0])
    title_col = next((v for k, v in columns_map.items() if "명칭" in k or "제목" in k or "특허명" in k), None)
    abstract_col = next((v for k, v in columns_map.items() if "요약" in k or "초록" in k), None)
    claims_col = next((v for k, v in columns_map.items() if "청구" in k or "범위" in k or "청구항" in k), None)
    
    app_date_col = next((v for k, v in columns_map.items() if "출원일" in k or "출원일자" in k), None)
    reg_date_col = next((v for k, v in columns_map.items() if "등록일" in k or "등록일자" in k), None)
    ipc_col = next((v for k, v in columns_map.items() if "IPC" in k), None)
    cpc_col = next((v for k, v in columns_map.items() if "CPC" in k), None)
    inventor_col = next((v for k, v in columns_map.items() if "발명자" in k or "발명인" in k), None)
    applicant_col = next((v for k, v in columns_map.items() if "출원인" in k or "권리자" in k), None)

    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0:
        try:
            master_df = pd.read_excel(MASTER_EXCEL_PATH)
            existing_numbers = set(master_df[id_col].astype(str).str.replace("-", "").str.strip().tolist())
        except Exception:
            master_df = pd.DataFrame(columns=new_df.columns)
            existing_numbers = set()
    else:
        master_df = pd.DataFrame(columns=new_df.columns)
        existing_numbers = set()

    new_records = []
    duplicate_count = 0

    for idx, row in new_df.iterrows():
        current_number = str(row[id_col]).replace("-", "").strip()
        if current_number == "" or current_number == "nan":
            continue
        if current_number in existing_numbers:
            duplicate_count += 1
            continue
        new_records.append(row)
        existing_numbers.add(current_number)

    if new_records:
        added_df = pd.DataFrame(new_records)
        updated_master_df = added_df if master_df.empty else pd.concat([master_df, added_df], ignore_index=True)
        updated_master_df.to_excel(MASTER_EXCEL_PATH, index=False)
        
        for idx, row in added_df.iterrows():
            title = str(row[title_col]).strip() if title_col and pd.notna(row[title_col]) else "정보없음"
            abstract = str(row[abstract_col]).strip() if abstract_col and pd.notna(row[abstract_col]) else "정보없음"
            claims = str(row[claims_col]).strip() if claims_col and pd.notna(row[claims_col]) else "정보없음"
            
            search_context = f"특허명칭: {title}\n특허요약: {abstract}\n특허청구항: {claims}"
            doc_id = str(row[id_col]).replace("-", "").strip()
            
            patent_url = hyperlink_map.get(doc_id, "")
            if patent_url == "" and title_col:
                clean_title = str(row[title_col]).strip().replace("-", "")
                patent_url = hyperlink_map.get(clean_title, "")

            # 경쟁사 특허 RAG 목적에 부합하는 시맨틱 인덱싱 벡터화 실행
            collection.add(
                documents=[search_context],
                metadatas=[{
                    "출원번호": str(row[id_col]),
                    "명칭": title,
                    "출원일": str(row[app_date_col]) if app_date_col and pd.notna(row[app_date_col]) else "없음",
                    "등록일": str(row[reg_date_col]) if reg_date_col and pd.notna(row[reg_date_col]) else "없음",
                    "IPC": str(row[ipc_col]) if ipc_col and pd.notna(row[ipc_col]) else "없음",
                    "CPC": str(row[cpc_col]) if cpc_col and pd.notna(row[cpc_col]) else "없음",
                    "발명자": str(row[inventor_col]) if inventor_col and pd.notna(row[inventor_col]) else "없음",
                    "출원인": str(row[applicant_col]) if applicant_col and pd.notna(row[applicant_col]) else "없음",
                    "URL": patent_url
                }],
                ids=[doc_id]
            )
        return len(new_records), duplicate_count
    else:
        return 0, duplicate_count


# --- 4. 메인 어플리케이션 인터페이스 구동 런타임 ---
def run_main_portal():
    chroma_client, collection, llm = load_permanent_infra_singleton()

    # 컨테이너 재시작 시 로고 파일 자동 복원
    download_logo_from_github()

    # [데이터 휘발 방지] 컨테이너 재시작 후 로컬 엑셀이 없으면 GitHub에서 즉시 복원
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        with st.spinner("🔄 GitHub 데이터 웨어하우스에서 마스터 엑셀 복원 중..."):
            restored = download_master_excel_from_github()
            if restored:
                st.toast("✅ GitHub로부터 마스터 데이터 복원 완료!")

    # 엑셀은 있지만 /tmp ChromaDB가 비어 있으면(재시작) 자동 재인덱싱
    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0 and collection.count() == 0:
        try:
            with st.spinner("📦 마스터 데이터 기반 벡터 DB 자동 재인덱싱 중... (첫 접속 시 1~2분 소요)"):
                process_and_update_db(MASTER_EXCEL_PATH, collection)
            st.toast(f"✅ 벡터 DB 복원 완료! ({collection.count()}건 인덱싱됨)")
        except Exception as e:
            st.warning(f"자동 재인덱싱 중 오류 발생: {e}")

    is_admin = st.session_state.get("is_admin", False)

    col_logo, col_title, col_logout = st.columns([1, 7, 2])
    with col_logo:
        logo_path = os.path.join(BASE_DIR, "atec_logo.png")
        if os.path.exists(logo_path):
            st.image(logo_path, width=110)
    with col_title:
        st.title("AI 경쟁사 특허 조사 분석")
        mode_label = "🔧 관리자" if is_admin else "👤 사용자"
        st.caption(f"{mode_label} | 접속 계정: {st.session_state.user_id} | 적재 특허: {collection.count()}건")
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
            uploaded_file = st.file_uploader("경쟁사 특허 엑셀 리스트 업로드 (.xlsx)", type=["xlsx"])
            if uploaded_file is not None:
                if st.button("🚀 신규 특허 무결성 적재"):
                    with st.spinner("중복 제거 및 실시간 인덱싱 중..."):
                        added, dup = process_and_update_db(uploaded_file, collection)
                    with st.spinner("💾 GitHub 데이터 웨어하우스 영구 동기화 중..."):
                        commit_and_push_data()
                    st.success(f"처리 완료! (신규: {added}건 / 중복 제외: {dup}건) — GitHub 백업 완료")
                    st.rerun()

            st.divider()
            st.markdown(f"📊 **누적 적재 데이터:** `{collection.count()}` 건")

            if st.button("🚨 가상 데이터 웨어하우스 전체 포맷"):
                with st.spinner("⏳ 벡터 DB 및 마스터 데이터 완전 초기화 중..."):
                    try:
                        if os.path.exists(MASTER_EXCEL_PATH):
                            os.remove(MASTER_EXCEL_PATH)
                        if os.path.exists(DB_PATH):
                            shutil.rmtree(DB_PATH)
                        os.makedirs(DB_PATH, exist_ok=True)
                        _build_infra.clear()
                        if "infra_initialized" in st.session_state:
                            del st.session_state["infra_initialized"]
                        st.toast("✅ 데이터 웨어하우스 전체 초기화 완료. 새 엑셀을 업로드해 주세요.")
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
        "📊 출원정보 기반 다차원 통계조사 (출원인, 발명자, IPC, 일자 등)": "통계 요약을 보고 싶은 조건이나 '전체 통계 요약해줘'라고 입력하세요."
    }
    user_query = st.text_area("분석 대상 내용을 입력하세요:", height=110, placeholder=placeholders[analysis_mode])

    if st.button("🧬 가상 전문가 엔진 구동"):
        if user_query.strip() == "":
            st.warning("분석 내용을 입력해 주세요.")
        elif collection.count() == 0:
            st.error("서버 DB에 적재된 특허 소스가 없습니다. 좌측 메뉴에서 엑셀을 먼저 등록해 주세요.")
        else:
            with st.spinner("가상 전문가가 실시간 시맨틱 문헌 대조 및 클라우드 초고속 추론을 진행 중입니다..."):

                # ── 통계 모드: 전체 메타데이터 pandas 집계 후 요약 컨텍스트 구성 ──
                if "📊" in analysis_mode:
                    all_data = collection.get(include=["metadatas"])
                    all_metas = all_data.get("metadatas", [])
                    total_count = len(all_metas)

                    if total_count == 0:
                        st.error("서버 DB에 적재된 특허 소스가 없습니다.")
                        st.stop()

                    stats_df = pd.DataFrame(all_metas)

                    def top_counts(series, n=10):
                        return series.replace("없음", pd.NA).dropna().value_counts().head(n).to_string()

                    applicant_stat = top_counts(stats_df.get("출원인", pd.Series(dtype=str)))
                    ipc_stat       = top_counts(stats_df.get("IPC", pd.Series(dtype=str)))
                    inventor_stat  = top_counts(stats_df.get("발명자", pd.Series(dtype=str)))

                    year_series = stats_df.get("출원일", pd.Series(dtype=str)).str[:4]
                    year_stat   = year_series.replace("", pd.NA).dropna().value_counts().sort_index().to_string()

                    context_text = f"""[전체 DB 통계 요약] 총 {total_count}건

■ 출원인별 상위 10 현황
{applicant_stat}

■ IPC 분류별 상위 10 현황
{ipc_stat}

■ 주요 발명자 상위 10 현황
{inventor_stat}

■ 출원 연도별 건수 추이
{year_stat}
"""
                    # 질의 관련 시맨틱 매칭 상위 특허도 추가
                    sem_results = collection.query(
                        query_texts=[user_query.strip()],
                        n_results=min(total_count, 10)
                    )
                    if sem_results and sem_results["documents"][0]:
                        context_text += "\n■ 질의 관련 시맨틱 매칭 상위 특허\n"
                        for i, m in enumerate(sem_results["metadatas"][0]):
                            context_text += f"  [{i+1}] {m.get('출원번호','')} | {m.get('명칭','')} | {m.get('출원인','')}\n"

                    system_prompt = (
                        "당신은 특허 데이터 통계 전문 분석가입니다. "
                        "아래 집계 통계를 기반으로 출원 동향, 핵심 출원인, 기술 분야 분포를 "
                        "마크다운 표와 함께 체계적인 다차원 통계 리포트로 작성하세요."
                    )

                # ── 일반 모드: 시맨틱 RAG 검색 ──
                else:
                    # 침해 분석은 7건, 나머지는 5건으로 품질 향상
                    n_results = 7 if "🛡" in analysis_mode else 5
                    n_results = min(n_results, collection.count())

                    results = collection.query(
                        query_texts=[user_query.strip()],
                        n_results=n_results
                    )

                    if not (results and results["documents"] and results["documents"][0]):
                        st.error("관련 특허를 찾지 못했습니다. 다른 키워드로 시도해 보세요.")
                        st.stop()

                    retrieved_docs  = results["documents"][0]
                    retrieved_metas = results["metadatas"][0]

                    context_text = ""
                    for i, (doc, m) in enumerate(zip(retrieved_docs, retrieved_metas)):
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

                        context_text += (
                            f"[특허 {i+1}] 번호: {display_num} | 명칭: {display_name} | "
                            f"출원인: {applicant} | 발명자: {inventor} | "
                            f"IPC: {ipc} | CPC: {cpc} | 출원일: {app_date}\n{doc}\n\n"
                        )

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

                prompt = (
                    f"[SYSTEM] {system_prompt}\n"
                    "답변 시 참고한 특허 번호·명칭은 [번호](URL) 마크다운 링크 형식을 그대로 유지하세요.\n\n"
                    f"[참고 데이터]\n{context_text}\n\n"
                    f"[사용자 요청]\n{user_query}\n\n"
                    "보고서는 마크다운 양식으로 한국어로 작성하세요."
                )

                try:
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
