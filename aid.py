import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
import os
import time
import requests
import tempfile
import io
from docx import Document # 🌟 워드 파일 생성을 위해 추가된 라이브러리

# ---------------------------------------------------------
# 1. 초기 설정 및 라이브러리 폴더 준비
# ---------------------------------------------------------
st.set_page_config(page_title="사회복지사 강의 요약기 PRO", page_icon="📚", layout="wide")

os.makedirs("library/pdfs", exist_ok=True)
os.makedirs("library/videos", exist_ok=True)

st.title("📚 사회복지사 강의 요약 웹서비스 (Library 버전)")

def get_optimal_model():
    available_models = []
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            available_models.append(m.name)
            
    for name in available_models:
        if 'gemini-1.5-pro' in name: return name
    for name in available_models:
        if 'gemini-1.5-flash' in name: return name
    for name in available_models:
        if 'gemini' in name: return name
    raise ValueError("사용 가능한 모델을 찾을 수 없습니다.")

# 🌟 워드 파일 생성 함수 (새로 추가됨)
def generate_word_file(title, content):
    doc = Document()
    doc.add_heading(title, level=1) # 문서 맨 위에 큰 제목 추가
    doc.add_paragraph(content)      # AI가 요약한 내용 추가
    
    # 다운로드를 위해 파일을 메모리(버퍼)에 임시 저장
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

saved_api_key = st.secrets["GEMINI_API_KEY"] if "GEMINI_API_KEY" in st.secrets else ""
with st.sidebar:
    st.header("⚙️ 기본 설정")
    api_key = st.text_input("Gemini API Key", value=saved_api_key, type="password")
    if api_key:
        genai.configure(api_key=api_key)

tab1, tab2 = st.tabs(["📂 1단계: 교안 라이브러리 등록", "🎬 2단계: 동영상 링크 분석 및 요약"])

# ---------------------------------------------------------
# [탭 1] 교안 PDF 업로드 및 라이브러리 저장
# ---------------------------------------------------------
with tab1:
    st.header("1. 교안 PDF 등록하기")
    course_name_input = st.text_input("🏷️ 강의 이름 입력 (예: 사회복지학개론)")
    pdf_file = st.file_uploader("📄 교안 PDF 파일 업로드", type=['pdf'])
    
    if st.button("💾 교안 라이브러리에 저장"):
        if not course_name_input or not pdf_file:
            st.warning("강의 이름과 PDF 파일을 모두 입력해 주세요.")
        else:
            with st.spinner("PDF 문서 텍스트를 추출하여 저장 중입니다..."):
                doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
                full_text = "".join([page.get_text() for page in doc])
                
                with open(f"library/pdfs/{course_name_input}.txt", "w", encoding="utf-8") as f:
                    f.write(full_text)
                st.success(f"🎉 '{course_name_input}' 교안이 저장되었습니다!")

# ---------------------------------------------------------
# [탭 2] 동영상 분석 및 최종 요약
# ---------------------------------------------------------
with tab2:
    st.header("2. 동영상 분석 및 최종 요약본 생성")
    
    saved_pdfs = [f.replace(".txt", "") for f in os.listdir("library/pdfs") if f.endswith(".txt")]
    
    if not saved_pdfs:
        st.info("👈 먼저 '1단계' 탭에서 교안을 등록해 주세요.")
    else:
        selected_course = st.selectbox("📚 분석할 강의 선택 (라이브러리)", saved_pdfs)
        col1, col2 = st.columns(2)
        with col1: week = st.number_input("주차 (예: 1)", min_value=1, step=1)
        with col2: session = st.number_input("강 (예: 1)", min_value=1, step=1)
            
        video_url = st.text_input("🔗 MP4 동영상 링크 입력 (https://...mp4)")
        
        if st.button("🚀 분석 및 워드(Word) 요약본 생성 시작"):
            if not api_key:
                st.error("사이드바에 API 키를 입력해 주세요.")
            elif not video_url:
                st.warning("동영상 링크를 입력해 주세요.")
            else:
                try:
                    with st.spinner("1/4. 인터넷에서 동영상을 임시 다운로드 중입니다... ⏳"):
                        response = requests.get(video_url, stream=True)
                        response.raise_for_status()
                        
                        tmp_video = tempfile.NamedTemporaryFile(delete=False, suffix='.mp4')
                        for chunk in response.iter_content(chunk_size=8192):
                            tmp_video.write(chunk)
                        tmp_video.close()
                        temp_video_path = tmp_video.name
                        
                    st.write("2/4. AI 서버에 영상을 올리고 처리하는 중입니다. (수 분 소요) 🧠")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    optimal_model_name = get_optimal_model()
                    model = genai.GenerativeModel(optimal_model_name)
                    video_upload = genai.upload_file(path=temp_video_path)
                    
                    start_time = time.time()
                    while True:
                        file_info = genai.get_file(video_upload.name)
                        state = file_info.state.name
                        elapsed = int(time.time() - start_time)
                        
                        if state == "PROCESSING":
                            status_text.text(f"⏳ 서버에서 영상 분석 준비 중... (현재 {elapsed}초 경과)")
                            progress_bar.progress(min(elapsed, 95))
                            time.sleep(5)
                        elif state == "ACTIVE":
                            progress_bar.progress(100)
                            status_text.success(f"✅ AI 영상 분석 준비 완료! (총 {elapsed}초 소요)")
                            break
                        elif state == "FAILED":
                            progress_bar.empty()
                            status_text.empty()
                            raise Exception("AI 서버에서 동영상 처리에 실패했습니다.")
                        else:
                            time.sleep(5)

                    with st.spinner("3/4. 영상의 핵심 내용을 추출하고 있습니다... ✍️"):
                        video_prompt = """
                        이 강의 영상에서 다음 사항을 추출해 주세요:
                        1. 강사가 음성으로 "시험에 나온다", "중요하다" 등으로 특별히 강조한 내용
                        2. 화면 상에서 하이라이트(밑줄, 별표, 빨간 글씨 등) 처리된 핵심 내용
                        """
                        video_response = model.generate_content([video_upload, video_prompt])
                        video_key_points = video_response.text
                        
                        genai.delete_file(video_upload.name)
                        os.unlink(temp_video_path)

                    with st.spinner("4/4. 최종 노트 필기본을 작성 중입니다... 📝"):
                        with open(f"library/pdfs/{selected_course}.txt", "r", encoding="utf-8") as f:
                            course_full_text = f.read()
                            
                        # 🌟 프롬프트 수정: 마크다운 기호를 빼고 워드 문서용으로 깔끔하게 작성하도록 지시
                        final_prompt = f"""
                        당신은 훌륭한 사회복지사 학습 조교입니다.
                        아래 두 가지 자료를 대조하여 '{selected_course}'의 {week}주차 {session}강에 대한 노트 필기 정리본을 작성해 주세요.

                        [자료 1: 영상 핵심 내용]
                        {video_key_points}

                        [자료 2: 전체 교안 텍스트]
                        {course_full_text}

                        [지시 사항]
                        1. [자료 1]의 핵심 내용을 바탕으로, [자료 2]에서 자세한 설명을 찾아내세요.
                        2. 컴퓨터 프로그래밍 기호(마크다운의 **, # 등)는 절대 사용하지 마세요.
                        3. 어르신들이 읽기 편하도록 숫자(1, 2, 3), 기호(-, ※, ▶)를 활용해 들여쓰기와 줄바꿈을 아주 깔끔하게 정리한 '일반 텍스트'로 요약해 주세요.
                        """
                        final_response = model.generate_content(final_prompt)
                        final_summary = final_response.text

                    st.success("🎉 모든 분석 및 정리가 완료되었습니다!")
                    st.divider()
                    
                    final_title = f"[{selected_course}] - {week}주차 {session}강 요약 노트"
                    
                    # 미리보기 화면 제공
                    st.subheader("👀 요약본 미리보기")
                    st.text(final_summary) # 마크다운 렌더링 대신 원본 텍스트로 보여줌
                    
                    st.divider()
                    
                    # 🌟 다운로드 버튼 (워드 파일 제공)
                    word_buffer = generate_word_file(final_title, final_summary)
                    
                    st.download_button(
                        label="📄 워드 파일(.docx) 다운로드 - 클릭하여 저장하세요",
                        data=word_buffer,
                        file_name=f"{final_title}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                except requests.exceptions.RequestException as e:
                    st.error(f"동영상 링크 에러: {e}")
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")