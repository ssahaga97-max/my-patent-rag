import streamlit as st
import pandas as pd
import os
from langchain_groq import ChatGroq
import openpyxl 
import json
import base64
from urllib.request import Request, urlopen
from urllib.error import HTTPError
import numpy as np

# --- 1. 클라우드 서버 전용 절대 경로 고정 및 초기화 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_EXCEL_PATH = os.path.join(BASE_DIR, "my_patent_folder", "master_patents.xlsx")

os.makedirs(os.path.join(BASE_DIR, "my_patent_folder"), exist_ok=True)

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
def upload_file_to_github_api(local_file_path, github_target_path):
    """
    Streamlit 내부 파일 감시자 간섭 및 I/O 교착을 완벽히 우회하여
    GitHub REST API를 통해 Private 저장소에 데이터를 Direct 적재하는 함수
    """
    if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO_URL" not in st.secrets:
        return False
        
    token = st.secrets["GITHUB_TOKEN"]
    raw_url = st.secrets["GITHUB_REPO_URL"].replace(".git", "")
    repo_path = raw_url.split("github.com/")[-1]
    
    if not os.path.exists(local_file_path):
        # 포맷으로 인해 파일이 없는 경우, 깃허브 원격지 파일도 삭제 스트림 처리
        api_url = f"https://api.github.com/repos/{repo_path}/contents/{github_target_path}"
        try:
            sha = None
            req_get = Request(api_url, headers={
                "Authorization": f"Bearer {token}", 
                "Accept": "application/vnd.github.v3+json"
            })
            with urlopen(req_get) as response:
                res_data = json.loads(response.read().decode())
                sha = res_data.get("sha")
            
            if sha:
                payload = {
                    "message": "🗑️ [Automated API Warehouse Sync] 마스터 데이터베이스 전체 포맷 반영",
                    "sha": sha,
                    "branch": "main"
                }
                req_del = Request(
                    api_url,
                    data=json.dumps(payload).encode("utf-8"),
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json",
                        "Accept": "application/vnd.github.v3+json"
                    },
                    method="DELETE"
                )
                with urlopen(req_del) as response:
                    return True
        except Exception:
            pass
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
    """가상 컨테이너 리부팅 대응용 마스터 엑셀 백업 스트림 엔진"""
    excel_status = upload_file_to_github_api(MASTER_EXCEL_PATH, "my_patent_folder/master_patents.xlsx")
    if excel_status:
        st.toast("💾 사내 가상 데이터 웨어하우스(GitHub) 마스터 엑셀 영구 동기화 완료!")


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
# 2. [초고도화 정공법] 라이브러리 충돌 프리 AI 검색 커널 팩토리
# ==========================================
def load_permanent_infra_singleton():
    """
    ChromaDB의 세션 및 테넌트 결함을 완전히 도려내어 부팅 크래시를 원천 방어하고,
    Groq Cloud API 기반 Llama 3.3 엔진만을 싱글톤 격리 구동하는 무결성 함수입니다.
    """
    if "llm_engine" not in st.session_state:
        GROQ_API_KEY = "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"
        st.session_state.llm_engine = ChatGroq(
            model="llama-3.3-70b-versatile", 
            groq_api_key=GROQ_API_KEY,
            temperature=0.1 
        )
    return st.session_state.llm_engine


# --- 3. 초경량 무결성 텍스트 매칭 검색 엔진 내부 로직 ---
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


def process_and_update_db(uploaded_file):
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
            
        # URL 메타데이터 강제 매칭 매핑
        patent_url = hyperlink_map.get(current_number, "")
        if patent_url == "" and title_col:
            clean_title = str(row[title_col]).strip().replace("-", "")
            patent_url = hyperlink_map.get(clean_title, "")
            
        row_dict = row.to_dict()
        row_dict['MATCHED_URL'] = patent_url
        new_records.append(row_dict)
        existing_numbers.add(current_number)

    if new_records:
        added_df = pd.DataFrame(new_records)
        updated_master_df = added_df if master_df.empty else pd.concat([master_df, added_df], ignore_index=True)
        updated_master_df.to_excel(MASTER_EXCEL_PATH, index=False)
        return len(new_records), duplicate_count
    else:
        return 0, duplicate_count


def local_intelligence_search(query_text, n_results=3):
    """메모리 내 파싱 완료된 데이터프레임 기반 고속 다차원 검색 스크리닝 매칭 커널"""
    if not os.path.exists(MASTER_EXCEL_PATH) or os.path.getsize(MASTER_EXCEL_PATH) == 0:
        return []
        
    df = pd.read_excel(MASTER_EXCEL_PATH)
    columns_map = {str(col).strip().replace(" ", "").upper(): col for col in df.columns}
    
    id_col = next((v for k, v in columns_map.items() if "출원번호" in k or "번호" in k), df.columns[0])
    title_col = next((v for k, v in columns_map.items() if "명칭" in k or "제목" in k or "특허명" in k), None)
    abstract_col = next((v for k, v in columns_map.items() if "요약" in k or "초록" in k), None)
    claims_col = next((v for k, v in columns_map.items() if "청구" in k or "범위" in k or "청구항" in k), None)
    
    app_date_col = next((v for k, v in columns_map.items() if "출원일" in k or "출원일자" in k), None)
    reg_date_col = next((v for k, v in columns_map.items() if "등록일" in k or "등록일자" in k), None)
    ipc_col = next((v for k, v in columns_map.items() if "IPC" in k), None)
    cpc_col = next((v for k, v in columns_map.items() if "CPC" in k), None)
    inventor_col = next((v for k, v in columns_map.items() if "발명자" in k or "발명인" in k), None)
    applicant_col = next((v for k, v in columns_map.items() if "출원인" in k or "권리자" in k), None)

    search_pool = []
    for idx, row in df.iterrows():
        title = str(row[title_col]) if title_col and pd.notna(row[title_col]) else ""
        abstract = str(row[abstract_col]) if abstract_col and pd.notna(row[abstract_col]) else ""
        claims = str(row[claims_col]) if claims_col and pd.notna(row[claims_col]) else ""
        
        # 형태소 단어 매칭 스코어링 가중치 기법 연산
        full_text = f"{title} {abstract} {claims}"
        score = sum(1 for word in query_text.split() if word.lower() in full_text.lower())
        
        search_pool.append({
            "score": score,
            "출원번호": str(row[id_col]),
            "명칭": title if title else "정보없음",
            "요약": abstract if abstract else "정보없음",
            "청구항": claims if claims else "정보없음",
            "출원일": str(row[app_date_col]) if app_date_col and pd.notna(row[app_date_col]) else "없음",
            "등록일": str(row[reg_date_col]) if reg_date_col and pd.notna(row[reg_date_col]) else "없음",
            "IPC": str(row[ipc_col]) if ipc_col and pd.notna(row[ipc_col]) else "없음",
            "CPC": str(row[cpc_col]) if cpc_col and pd.notna(row[cpc_col]) else "없음",
            "발명자": str(row[inventor_col]) if inventor_col and pd.notna(row[inventor_col]) else "없음",
            "출원인": str(row[applicant_col]) if applicant_col and pd.notna(row[applicant_col]) else "없음",
            "URL": str(row['MATCHED_URL']) if 'MATCHED_URL' in row and pd.notna(row['MATCHED_URL']) else ""
        })
        
    # 가중치 점수 정렬 후 상위 n개 반환
    search_pool = sorted(search_pool, key=lambda x: x['score'], reverse=True)
    return search_pool[:n_results]


# --- 4. 메인 어플리케이션 인터페이스 구동 런타임 ---
def run_main_portal():
    llm = load_permanent_infra_singleton()

    # 데이터 카운팅 동적 맵 계측
    total_count = 0
    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getsize(MASTER_EXCEL_PATH) > 0:
        try:
            total_count = len(pd.read_excel(MASTER_EXCEL_PATH))
        except Exception:
            pass

    col_title, col_logout = st.columns([8, 2])
    with col_title:
        st.title("🏛 맞춤형 인텔리전스 특허 가상 서버 인트라넷 (Groq Cloud Engine)")
        st.caption(f"접속 연구원 계정: {st.session_state.user_id} | 무결성 가상 데이터 처리 엔진 작동 중 (안정성 100%)")
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
                    added, dup = process_and_update_db(uploaded_file)
                    commit_and_push_data()
                    st.success(f"처리 완료! (신규: {added}건 / 중복 제외: {dup}건)")
                    st.rerun()
                    
        st.divider()
        st.markdown(f"📊 **누적 적재 데이터:** `{total_count}` 건")
        
        # ==========================================
        # 하드웨어 레벨 강제 포맷 엔진 (완전 정상화)
        # ==========================================
        if st.button("🚨 가상 데이터 웨어하우스 전체 포맷"):
            with st.spinner("⏳ 사내 가상 데이터 웨어하우스 초기화 스트림 제어 중..."):
                try:
                    # 물리 마스터 백업 파일 즉각 소거
                    if os.path.exists(MASTER_EXCEL_PATH): 
                        os.remove(MASTER_EXCEL_PATH)
                    
                    # 공백 상태(초기화 상태)를 GitHub 원격지 API 가상 웨어하우스로 플러시 전송
                    commit_and_push_data()
                    
                    st.toast("⚠️ 가상 데이터 웨어하우스 및 백엔드 물리 파일 초기화가 완벽하게 완료되었습니다!")
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
        elif total_count == 0:
            st.error("서버 DB에 적재된 특허 소스가 없습니다. 좌측 메뉴에서 엑셀을 먼저 등록해 주세요.")
        else:
            with st.spinner("가상 전문가가 실시간 문헌 대조 및 클라우드 초고속 추론을 진행 중입니다..."):
                n_results = 3
                if "📊" in analysis_mode:
                    n_results = min(total_count, 15)

                matched_records = local_intelligence_search(user_query.strip(), n_results=n_results)
                
                if matched_records:
                    context_text = ""
                    for i, m in enumerate(matched_records):
                        app_num = m.get('출원번호', '번호없음')
                        p_name = m.get('명칭', '제목없음')
                        applicant = m.get('출원인', '미기재')
                        inventor = m.get('발명자', '미기재')
                        ipc = m.get('IPC', '없음')
                        cpc = m.get('CPC', '없음')
                        app_date = m.get('출원일', '없음')
                        patent_url = m.get('URL', '')
                        doc_context = f"특허명칭: {p_name}\n특허요약: {m.get('요약')}\n특허청구항: {m.get('청구항')}"

                        if patent_url and patent_url.startswith("http"):
                            display_num = f"[{app_num}]({patent_url})"
                            display_name = f"[{p_name}]({patent_url})"
                        else:
                            display_num = app_num
                            display_name = p_name

                        context_text += f"[특허 {i+1}] 번호: {display_num} | 명칭: {display_name} | 출원인: {applicant} | 발명자: {inventor} | IPC: {ipc} | CPC: {cpc} | 출원일: {app_date}\n{doc_context}\n\n"
                    
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
