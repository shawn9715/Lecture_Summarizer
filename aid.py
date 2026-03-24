import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
import os
import time
import requests
import tempfile
import io
import yt_dlp
import re  # 🌟 정규표현식(텍스트 검색)을 위한 라이브러리 추가
from docx import Document
from docx.enum.text import WD_COLOR_INDEX  # 🌟 워드 형광펜 색상 기능을 위해 추가

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

# ---------------------------------------------------------
# 🌟 보강된 기능: 워드 파일 서식(볼드, 띄어쓰기, 형광펜) 자동 적용
# ---------------------------------------------------------
def generate_word_file(title, content):
    doc = Document()
    doc.add_heading(title, level=1)
    
    lines = content.split('\n')
    first_subheading_done = False # 첫 소제목 위에는 빈 줄을 넣지 않기 위한 장치
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # 1. 소제목 확인 (줄이 "1.", "2." 처럼 숫자와 마침표로 시작하는지 검사)
        is_subheading = bool(re.match(r'^\d+\.', line))
        
        # 소제목일 경우 위로 한 줄 띄우고 문단 생성
        if is_subheading:
            if first_subheading_done:
                doc.add_paragraph() # 빈 줄(단락) 추가
            first_subheading_done = True
            p = doc.add_paragraph()
        else:
            p = doc.add_paragraph() # 일반 문단 생성

        # 2. [강조] 태그를 찾아 형광펜 칠하기
        # 텍스트를 [강조] 태그 기준으로 조각조각 자릅니다.
        parts = re.split(r'(\[강조\].*?\[/강조\])', line)
        
        for part in parts:
            if part.startswith('[강조]') and part.endswith('[/강조]'):
                clean_text = part[4:-5] # '[강조]'와 '[/강조]' 글자만 깔끔하게 제거
                run = p.add_run(clean_text)
                if is_subheading:
                    run.bold = True
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW # 💛 노란색 형광펜 칠하기
            else:
                if part: # 빈 텍스트가 아니면 추가
                    run = p.add_run(part)
                    if is_subheading:
                        run.bold = True # 소제목 줄의 나머지 텍스트도 모두 굵게 처리
                        
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
# [탭 1] 교안 PDF 업로드 및 라이브러리 저장 (변경 없음)
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
    
    saved_pdfs_cursor = pdf_collection.find({}, {"course_name": 1, "_id": 0})
    saved_pdfs = [doc["course_name"] for doc in saved_pdfs_cursor]
    
    if not saved_pdfs:
        st.info("👈 먼저 '1단계' 탭에서 교안을 등록해 주세요.")
    else:
        selected_course = st.selectbox("📚 분석할 강의 선택 (DB 보관함)", saved_pdfs)
        col1, col2 = st.columns(2)
        with col1: week = st.number_input("주차 (예: 1)", min_value=1, step=1)
        with col2: session = st.number_input("강 (예: 1)", min_value=1, step=1)
            
        st.divider()
        st.subheader("🎬 동영상 입력 방식 선택")
        
        # 🌟 투트랙 입력 UI 구성
        input_method = st.radio(
            "어떤 방식으로 동영상을 분석할까요?", 
            ["🔗 1. 웹페이지 링크(URL) 자동 추출 시도", "📁 2. 내 컴퓨터에서 직접 파일 업로드 (확실한 방법)"]
        )
        
        video_url = ""
        video_file = None
        
        if "링크(URL)" in input_method:
            st.info("💡 유튜브나 보안이 낮은 웹페이지의 주소를 입력하면 자동으로 영상을 찾아냅니다.")
            video_url = st.text_input("🔗 웹페이지 주소 입력 (예: https://...)")
        else:
            st.info("💡 유료 인강처럼 로그인이 필요한 사이트는 보안상 직접 업로드해야 합니다.")
            video_file = st.file_uploader("📁 MP4 동영상 파일 업로드", type=['mp4', 'avi', 'mov'])
        
        if st.button("🚀 분석 및 워드(Word) 요약본 생성 시작"):
            if not api_key:
                st.error("사이드바에 API 키를 입력해 주세요.")
            elif ("링크(URL)" in input_method) and (not video_url):
                st.warning("웹페이지 링크를 입력해 주세요.")
            elif ("업로드" in input_method) and (not video_file):
                st.warning("동영상 파일을 업로드해 주세요.")
            else:
                try:
                    # 임시 파일 경로 미리 생성
                    tmp_video = tempfile.NamedTemporaryFile(delete=False, suffix='.mp4')
                    temp_video_path = tmp_video.name
                    tmp_video.close()

                    # 🌟 1. 동영상 확보 (선택한 방식에 따라 다르게 처리)
                    with st.spinner("1/4. 영상을 준비하고 있습니다... ⏳"):
                        if "링크(URL)" in input_method:
                            # yt-dlp를 이용해 웹페이지에서 영상 추출 시도
                            ydl_opts = {
                                'format': 'best',
                                'outtmpl': temp_video_path,
                                'quiet': True,
                                'no_warnings': True
                            }
                            try:
                                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                                    ydl.download([video_url])
                            except Exception as e:
                                raise Exception(f"보안 설정이나 로그인 문제로 해당 웹페이지에서 영상을 추출하지 못했습니다. 플랜 B(파일 직접 업로드)를 사용해 주세요. (상세 에러: {e})")
                        else:
                            # 업로드된 파일을 임시 경로에 저장
                            with open(temp_video_path, 'wb') as f:
                                f.write(video_file.read())

                    # ==========================================================
                    # 아래 부분은 기존 코드와 100% 동일하게 진행됩니다!
                    # ==========================================================
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
                        os.unlink(temp_video_path) # 사용이 끝난 로컬 임시 파일 삭제
                        
                        video_collection.update_one(
                            {"course_name": selected_course, "week": week, "session": session},
                            {"$set": {"content": video_key_points}},
                            upsert=True
                        )

                    with st.spinner("4/4. 최종 노트 필기본을 작성 중입니다... 📝"):
                        course_data = pdf_collection.find_one({"course_name": selected_course})
                        course_full_text = course_data["content"]
                            
                        final_prompt = f"""
                        당신은 훌륭한 사회복지사 학습 조교입니다.
                        아래 두 자료를 대조하여 '{selected_course}'의 {week}주차 {session}강 노트 필기본을 작성해 주세요.

                        [자료 1: 영상 핵심 내용]
                        {video_key_points}

                        [자료 2: 전체 교안 텍스트]
                        {course_full_text}

                        [작성 지시 사항 - 매우 중요]
                        1. 내용 대조: [자료 1]의 핵심 내용을 바탕으로 [자료 2]에서 상세한 설명을 찾아 요약하세요.
                        2. 공통 강조점 하이라이트: 영상과 교안 양쪽 모두에서 중요하게 다뤄진 핵심 단어나 문장은 반드시 앞뒤에 [강조] 와 [/강조] 태그를 붙이세요. 
                        3. 소제목 규칙: 내용의 큰 주제가 바뀔 때는 반드시 "1. 주제명" 처럼 숫자와 마침표로 시작하는 소제목을 적어주세요.
                        4. 포맷: 마크다운 기호(**, # 등)는 절대 쓰지 말고, 번호와 기호(-, ※)만 사용하여 어르신들이 읽기 편한 일반 텍스트로 깔끔하게 정리해 주세요.
                        """
                        final_response = model.generate_content(final_prompt)
                        final_summary = final_response.text

                    st.success("🎉 모든 분석 및 정리가 완료되었습니다!")
                    st.divider()
                    
                    final_title = f"[{selected_course}] - {week}주차 {session}강 요약 노트"
                    preview_text = final_summary.replace("[강조]", "").replace("[/강조]", "")
                    st.subheader("👀 요약본 미리보기")
                    st.text(preview_text)
                    st.divider()
                    
                    word_buffer = generate_word_file(final_title, final_summary)
                    st.download_button(
                        label="📄 굵은 글씨 및 형광펜이 적용된 워드 파일 다운로드",
                        data=word_buffer,
                        file_name=f"{final_title}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                except requests.exceptions.RequestException as e:
                    st.error(f"인터넷 연결 에러: {e}")
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")
