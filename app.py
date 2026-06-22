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
from urllib.request import Request, urlopen
from urllib.error import HTTPError

# --- 1. 클라우드 서버 전용 절대 경로 고정 및 초기화 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_EXCEL_PATH = os.path.join(BASE_DIR, "my_patent_folder", "master_patents.xlsx")
# /tmp 경로 사용: Streamlit Cloud에서 항상 쓰기 가능하며 git에 커밋된 구버전 SQLite 파일 충돌 원천 차단
DB_PATH = "/tmp/my_patent_vector_db"

os.makedirs(os.path.join(BASE_DIR, "my_patent_folder"), exist_ok=True)
os.makedirs(DB_PATH, exist_ok=True)

st.set_page_config(page_title="클라우드 특허 RAG 인트라넷", layout="wide")


# ==========================================
# 0. 전역 스코프 세션 상태 격리 및 초기화
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_id" not in st.session_state:
    st.session_state.user_id = None


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
            raw_content = base64.b64decode(data["content"])
        os.makedirs(os.path.dirname(MASTER_EXCEL_PATH), exist_ok=True)
        with open(MASTER_EXCEL_PATH, "wb") as f:
            f.write(raw_content)
        return True
    except Exception as e:
        print(f"GitHub 마스터 엑셀 복원 실패: {e}")
        return False


# ==========================================
# 1. 사내 연구원용 로그인 인터페이스
# ==========================================
def check_authentication():
    if "USER_CREDENTIALS" in st.secrets:
        user_credentials = st.secrets["USER_CREDENTIALS"]
    else:
        user_credentials = {
            "seongsu_bae": "amorfati78",
            "researcher01": "patent789",
            "researcher02": "tech2026",
            "admin": "1234!"
        }

    if not st.session_state.logged_in:
        st.title("🏛 맞춤형 인텔리전스 특허 가상 서버 인트라넷")
        st.subheader("🔑 사내 연구원 로그인 인증")
        
        with st.form("login_form"):
            username = st.text_input("사내 계정 ID", key="input_user")
            password = st.text_input("비밀번호", type="password", key="input_pass")
            submit_button = st.form_submit_button("인트라넷 접속")
            
            if submit_button:
                if username in user_credentials and user_credentials[username] == password:
                    st.session_state.logged_in = True
                    st.session_state.user_id = username
                    st.success(f"🔓 {username} 연구원님 인증 성공")
                    st.rerun()
                else:
                    st.error("❌ ID 또는 비밀번호가 올바르지 않습니다.")
        return False
    return True


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

    # Groq Cloud API 기반 Llama 3.3 엔진
    GROQ_API_KEY = "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"
    llm = ChatGroq(
        model="llama-3.3-70b-versatile",
        groq_api_key=GROQ_API_KEY,
        temperature=0.1
    )

    return chroma_client, collection, llm


def load_permanent_infra_singleton():
    with st.spinner("📦 가상 특허 가동 커널 및 AI 전문 임베딩 엔진 초기화 중..."):
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

    col_title, col_logout = st.columns([8, 2])
    with col_title:
        st.title("🏛 맞춤형 인텔리전스 특허 가상 서버 인트라넷 (Groq Cloud Engine)")
        st.caption(f"접속 연구원 계정: {st.session_state.user_id} | 크로마 벡터 커널 기반 정밀 RAG 인프라 가동 중 (안정성 100%)")
    with col_logout:
        if st.button("🔒 로그아웃"):
            st.session_state.logged_in = False
            st.session_state.user_id = None
            st.rerun()

    with st.sidebar:
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
            with st.spinner("⏳ 파일 시스템 락킹 전면 해제 및 벡터 DB 커널 완전 초기화 중..."):
                try:
                    if os.path.exists(MASTER_EXCEL_PATH):
                        os.remove(MASTER_EXCEL_PATH)

                    # 물리 DB 디렉토리 정화
                    if os.path.exists(DB_PATH):
                        shutil.rmtree(DB_PATH)
                    os.makedirs(DB_PATH, exist_ok=True)

                    # @st.cache_resource 캐시 파기 → 다음 호출 시 새 인스턴스로 재초기화
                    _build_infra.clear()

                    commit_and_push_data()
                    st.toast("⚠️ 가상 데이터 웨어하우스 및 백엔드 물리 커널 초기화가 성공적으로 완료되었습니다!")
                    st.rerun()
                except Exception as e:
                    st.error(f"초기화 중 인프라 제어 오류 발생: {e}")

    st.subheader("⚙️ 1단계: 분석 목적 및 AI 전문 페르소나 선택")
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
                n_results = 3
                if "📊" in analysis_mode:
                    n_results = min(collection.count(), 15)

                results = collection.query(
                    query_texts=[user_query.strip()],
                    n_results=n_results
                )
                
                if results and 'documents' in results and len(results['documents']) > 0 and len(results['documents'][0]) > 0:
                    retrieved_docs = results['documents'][0]
                    retrieved_metas = results['metadatas'][0]
                    
                    context_text = ""
                    for i, doc in enumerate(retrieved_docs):
                        m = retrieved_metas[i]
                        app_num = m.get('출원번호', '번호없음')
                        p_name = m.get('명칭', '제목없음')
                        applicant = m.get('출원인', '미기재')
                        inventor = m.get('발명자', '미기재')
                        ipc = m.get('IPC', '없음')
                        cpc = m.get('CPC', '없음')
                        app_date = m.get('출원일', '없음')
                        patent_url = m.get('URL', '')

                        if patent_url and patent_url.startswith("http"):
                            display_num = f"[{app_num}]({patent_url})"
                            display_name = f"[{p_name}]({patent_url})"
                        else:
                            display_num = app_num
                            display_name = p_name

                        context_text += f"[특허 {i+1}] 번호: {display_num} | 명칭: {display_name} | 출원인: {applicant} | 발명자: {inventor} | IPC: {ipc} | CPC: {cpc} | 출원일: {app_date}\n{doc}\n\n"
                    
                    if "💡 단순 키워드" in analysis_mode:
                        system_prompt = "당신은 신속하고 정확하게 관련 문헌을 찾아내는 '수석 특허 검색 조사관'입니다. 관련 특허를 마크다운 링크 서식과 함께 요약 브리핑하세요."
                    elif "🔬 특정 기술" in analysis_mode:
                        system_prompt = "당신은 수석 기술 전문 분석가입니다. 마크다운 링크를 포함한 기술 동향 보고서를 체계적으로 작성하세요."
                    elif "🛡 개발기술 침해" in analysis_mode:
                        system_prompt = "당신은 특허청 수석 심사관 및 특허법률 전문가 집단입니다. 관련 선행문헌들의 링크 주소를 명시하며 침해 가능성 및 회피설계 가이드를 작성하세요."
                    else:
                        system_prompt = "당신은 특허 데이터 통계 분석가입니다. 서지정보와 하이퍼링크 매칭 상태를 종합하여 다차원 통계 리포트를 작성하세요."

                    prompt = f"""<|begin_of_text|><|start_header_id|>system<|end_header_id|>
                    {system_prompt} 답변 시 참고한 특허의 번호나 명칭을 언급할 때는 시스템이 매칭해 준 [번호](URL) 또는 [명칭](URL) 마크다운 형식을 그대로 유지하여 사용자가 클릭하면 링크로 이동할 수 있게 하세요.<|eot_id|><|start_header_id|>user<|end_header_id|>

                    [참고 선행문헌 데이터]
                    {context_text}

                    [사용자 요청 내용]
                    {user_query}

                    보고서는 마크다운 양식을 사용하여 한국어로 논리정연하게 작성해 주세요.<|eot_id|><|start_header_id|>thought<|end_header_id|>
                    Groq engine active. Generating analytical report...<|eot_id|><|start_header_id|>assistant<|end_header_id|>
                    """
                    
                    try:
                        response = llm.invoke(prompt)
                        st.markdown(f"### 📊 AI {analysis_mode.split(' ')[1]} 결과 보고서")
                        st.write(response.content) 
                        st.divider()
                        with st.expander("👁 로컬 가상 서버가 실시간 스크리닝한 마스터 데이터 매칭 정보 (클릭 시 원문 이동 가능)"):
                            st.markdown(context_text) 
                    except Exception as e:
                        st.error(f"서버 연산 보호 오류: {e}")
                else:
                    st.error("데이터 매칭 실패")


if __name__ == "__main__":
    if check_authentication():
        run_main_portal()
