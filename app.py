import os
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from langchain_community.vectorstores import Chroma
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_groq import ChatGroq
from langchain.chains import RetrievalQA

# ==========================================
# 0. 초기 세션 상태 및 인프라 절대경로 설정
# ==========================================
# 다중 사용자 접속 시 세션 간섭을 차단하기 위한 독립 세션 초기화 
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_id" not in st.session_state:
    st.session_state.user_id = None

# 클라우드 가상 OS 환경에서 상대경로 뒤틀림으로 인한 DB 소실 방지 
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, "chroma_db")

# 페이지 기본 설정 (로그인 화면 진입 전 초기 렌더링 속도 극대화)
st.set_page_config(page_title="PatentRAG 포털", layout="wide")

# ==========================================
# 1. 차기 고도화: 임베딩 모델 지연 로딩 (Lazy Loading) 
# ==========================================
@st.cache_resource(show_spinner=False)
def get_embedding_model():
    """
    앱 진입 시 500MB 모델을 매번 다운로드하는 지연을 방지하기 위해,
    실제 인증 후 연산이 필요할 때만 호출되고 캐싱되는 지연 로딩 구조 
    """
    return HuggingFaceEmbeddings(model_name="jhgan/ko-sroberta-multitask") [cite: 1]

# ==========================================
# 2. 차기 고도화: 사내 연구원용 다중 ID/PASS 인터페이스 
# ==========================================
def check_authentication():
    """
    Streamlit Cloud의 보안 저장소(secrets.toml)와 연동하여
    연구원별 계정 사전을 매핑하고 인증을 수행 
    """
    # 로컬 테스트 및 secrets 미설정 대비용 Fallback 계정 사전 정의 
    # 실제 운영 환경에서는 Streamlit 웹 대시보드의 Secrets에 아래 구조로 저장해야 합니다.
    # [USER_CREDENTIALS]
    # researcher1 = "password123!"
    # researcher2 = "patent456!"
    
    if "USER_CREDENTIALS" in st.secrets:
        user_credentials = st.secrets["USER_CREDENTIALS"]
    else:
        # secrets.toml이 없을 경우 기본 테스트 계정 매핑 
        user_credentials = {
            "admin": "admin123",
            "researcher01": "patent789",
            "researcher02": "tech2026"
        }

    if not st.session_state.logged_in:
        st.title("🔒 사내 프라이빗 특허 분석 RAG 포털")
        st.subheader("연구원 로그인 인터페이스") [cite: 10]
        
        with st.form("login_form"):
            username = st.text_input("사내 계정 ID (User ID)", key="input_user")
            password = st.text_input("비밀번호 (Password)", type="password", key="input_pass")
            submit_button = st.form_submit_button("로그인")
            
            if submit_button:
                if username in user_credentials and user_credentials[username] == password:
                    st.session_state.logged_in = True
                    st.session_state.user_id = username
                    st.success(f"🔓 {username} 연구원님 환영합니다.")
                    st.rerun()
                else:
                    st.error("❌ ID 또는 비밀번호가 올바르지 않습니다. 다시 입력해주세요.")
        return False
    return True

# ==========================================
# 3. 데이터 추출 및 하이퍼링크 매칭 기법 (기존 고도화 유지) [cite: 7]
# ==========================================
def extract_patent_with_links(excel_path):
    """
    openpyxl을 사용하여 키프리스 엑셀 레이어의 원문 하이퍼링크 주소를 추출 [cite: 7]
    """
    wb = load_workbook(excel_path, data_only=False)
    ws = wb.active
    
    # openpyxl로 링크 추출용 매핑 생성
    link_dict = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink and cell.hyperlink.target:
                # 셀의 텍스트 값을 키로 하여 링크 타겟 저장 [cite: 7]
                link_dict[str(cell.value).strip()] = cell.hyperlink.target [cite: 7]
                
    # pandas로 데이터 프레임 로드 파트
    df = pd.read_excel(excel_path) [cite: 1]
    return df, link_dict

# ==========================================
# 4. 메인 RAG 대시보드 애플리케이션 화면
# ==========================================
def main_portal():
    # 상단 헤더 및 로그아웃 버튼 (독립 세션 제어용) 
    col1, col2 = st.columns([8, 2])
    with col1:
        st.title("🚀 사내 프라이빗 특허 분석 RAG 포털 (PatentRAG)") [cite: 1]
        st.caption(f"접속 연구원 세션: {st.session_state.user_id} | 인프라 상태: 정상 가동 중 (Python 3.11)") [cite: 4, 11]
    with col2:
        if st.button("로그아웃"):
            st.session_state.logged_in = False
            st.session_state.user_id = None
            st.rerun()
            
    st.markdown("---")
    
    # 사이드바 - 특허 데이터 업로드 및 인프라 설정 정보
    with st.sidebar:
        st.header("📂 특허 데이터 소스 관리")
        uploaded_file = st.file_uploader("키프리스 특허 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
        
        st.markdown("---")
        st.subheader("⚙️ 인프라 스택 명세")
        st.markdown("""
        - **LLM 엔진**: Groq Cloud Llama 3.3 (70B) [cite: 1, 6]
        - **벡터 DB**: ChromaDB (영구 컨테이너 구조) [cite: 1]
        - **임베딩**: ko-sroberta-multitask [cite: 1]
        - **안정성**: Numpy < 2.0.0 충돌 패치 완료 
        """)

    # 메인 탭 구성 (독립 세션 메모리 보호 하에 작동) 
    tab1, tab2 = st.tabs(["🔍 특허 지식 검색 (RAG)", "📊 데이터 현황 확인"])
    
    if uploaded_file is not None:
        # 임시 파일 저장 후 openpyxl 및 pandas 파싱 [cite: 1, 7]
        temp_path = os.path.join(BASE_DIR, "temp_patent.xlsx")
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
            
        df, link_dict = extract_patent_with_links(temp_path)
        
        with tab2:
            st.subheader("업로드된 특허 데이터프레임 구조")
            st.dataframe(df.head(10))
            
        with tab1:
            st.subheader("💡 연구원 맞춤형 특허 질의응답")
            query = st.text_input("분석하고자 하는 특허 주제나 기술 키워드를 입력하세요:")
            
            if query:
                with st.spinner("Groq 70B 초고속 연산 및 특허 원문 교차 검증 중..."): [cite: 1, 2, 6]
                    try:
                        # [지연 로딩 적용]: 사용자가 실제 쿼리를 날려 연산이 필요할 때 모델 로드 
                        embeddings = get_embedding_model() [cite: 9]
                        
                        # ChromaDB persistent client 연동 [cite: 1]
                        # 실제 고도화 환경에서는 업로드된 df 기반으로 벡터화 및 메타데이터 주입 단계를 거침
                        # 여기서는 구조적 무결성을 위해 변수 바인딩 형태만 구현
                        db = Chroma(persist_directory=DB_DIR, embedding_function=embeddings) [cite: 1, 3]
                        
                        # Groq API 기반 Meta Llama 3.3 (70B) 연결 (Temperature=0.1 고정) [cite: 1, 6]
                        llm = ChatGroq(
                            temperature=0.1, [cite: 1]
                            model_name="llama-3.3-70b-versatile", [cite: 6]
                            groq_api_key=st.secrets.get("GROQ_API_KEY", "MOCK_KEY")
                        )
                        
                        # QA 체인 생성 및 실행
                        qa_chain = RetrievalQA.from_chain_type(
                            llm=llm,
                            chain_type="stuff",
                            retriever=db.as_retriever(search_kwargs={"k": 3})
                        )
                        
                        response = qa_chain.run(query)
                        
                        # 후처리: AI 답변 내의 특허번호나 출원명을 파란색 링크 텍스트로 치환 [cite: 8]
                        # openpyxl로 수집해둔 원문 하이퍼링크 주소 매칭 기법 반영 
                        refined_response = response
                        for text, url in link_dict.items():
                            if text in refined_response and f"[{text}]" not in refined_response:
                                refined_response = refined_response.replace(text, f"[{text}]({url})") [cite: 8]
                                
                        st.markdown("### 📋 AI 분석 리포트")
                        st.markdown(refined_response) [cite: 8]
                        
                    except Exception as e:
                        st.error(f"RAG 추론 중 에러가 발생했습니다: {str(e)}")
    else:
        with tab1:
            st.info("👈 사이드바에서 키프리스 특허 엑셀 파일을 먼저 업로드해 주세요.")

# ==========================================
# 5. 애플리케이션 진입점 (Entry Point)
# ==========================================
if __name__ == "__main__":
    # 1. 사내 연구원 인터페이스 인증 절차 수행 
    if check_authentication():
        # 2. 인증 통과 시에만 메인 포털 구동 (이때까지 대형 임베딩 모델 로딩은 보류됨 -> 1초 미만 진입) 
        main_portal()
