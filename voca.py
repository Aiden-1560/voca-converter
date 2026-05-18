import streamlit as st
from google import genai
from google.genai import types
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import io
import asyncio
import time

# ==========================================
# 1. 워드 파일 디자인 & 생성 헬퍼 함수
# ==========================================
def set_cell_borders(cell, color="D9D9D9", sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_cell_shading(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def create_word_document(all_word_data):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style.font.size = Pt(10.5)
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("통합 본문 단어장")
    title_run.font.bold = True
    title_run.font.size = Pt(18)
    title_p.paragraph_format.space_after = Pt(24)
    
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    col_widths = [Inches(1.5), Inches(2.0), Inches(4.0)]
    
    hdr_cells = table.rows[0].cells
    headers = ["본문 단어", "우리말 뜻", "영영 풀이"]
    for i, text in enumerate(headers):
        hdr_cells[i].text = text
        hdr_cells[i].width = col_widths[i]
        set_cell_shading(hdr_cells[i], "4F81BD")
        set_cell_borders(hdr_cells[i], color="A6A6A6")
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = docx.shared.RGBColor(255, 255, 255)
            
    for row_idx, item in enumerate(all_word_data):
        row_cells = table.add_row().cells
        row_data = [item.get("word", ""), item.get("meaning", ""), item.get("definition", "")]
        
        for i, text in enumerate(row_data):
            row_cells[i].text = text
            row_cells[i].width = col_widths[i]
            if row_idx % 2 == 1:
                set_cell_shading(row_cells[i], "F2F5F8")
            set_cell_borders(row_cells[i], color="D9D9D9")
            
            p = row_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            for run in p.runs:
                run.font.size = Pt(10)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 2. 비동기 병렬 처리 함수
# ==========================================
async def analyze_single_image(client, file, api_key):
    loop = asyncio.get_event_loop()
    image_bytes = file.read()
    
    prompt = """
    이 이미지에서 영어 단어, 우리말 뜻, 영영 풀이를 추출해서 정확한 JSON 배열 형식으로 출력해줘.
    필기구로 수정한 흔적이나 추가로 적은 필기는 무시하고, 원래 인쇄되어 있던 텍스트만 추출해줘.
    결과는 오직 아래 구조를 가진 JSON 데이터만 반환해야 해:
    [
      {"word": "단어", "meaning": "품사 및 뜻", "definition": "영영 풀이 내용"}
    ]
    """
    
    def call_api():
        return client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[
                types.Part.from_bytes(data=image_bytes, mime_type=file.type),
                prompt
            ],
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        
    response = await loop.run_in_executor(None, call_api)
    return json.loads(response.text)

async def run_smooth_progress(progress_bar, status_text, stop_event, total_files):
    """0%부터 95%까지 부드럽게 숫자를 올려주는 애니메이션 함수"""
    current_percent = 0
    # AI 응답이 오기 전까지 점진적으로 상승 (95%에서 대기)
    while not stop_event.is_set() and current_percent < 95:
        # 초반에는 조금 빠르게, 후반으로 갈수록 신중하게 상승
        if current_percent < 50:
            current_percent += 1.5
        elif current_percent < 80:
            current_percent += 0.8
        else:
            current_percent += 0.3
            
        display_percent = min(int(current_percent), 95)
        progress_bar.progress(display_percent / 100)
        status_text.markdown(f"**⚡ 초고속 비동기 엔진 가동 중... {display_percent}%** (총 {total_files}장의 사진 병렬 분석 중)")
        await asyncio.sleep(0.05) # 0.05초 간격으로 매우 부드럽게 갱신

async def process_all_images(client, uploaded_files, api_key, progress_bar, status_text):
    """API 요청과 부드러운 게이지 상승을 동시에 처리"""
    tasks = [analyze_single_image(client, file, api_key) for file in uploaded_files]
    
    # 게이지 애니메이션을 제어할 이벤트 플래그
    stop_event = asyncio.Event()
    
    # 애니메이션 태스크 시작
    progress_task = asyncio.create_task(
        run_smooth_progress(progress_bar, status_text, stop_event, len(uploaded_files))
    )
    
    # 실제 AI 요청 돌리기 (전체 동시 실행)
    try:
        results = await asyncio.gather(*tasks)
    finally:
        # AI 분석이 모두 끝나면 애니메이션 루프 종료 명령
        stop_event.set()
        await progress_task
        
    # 완료 시 즉시 100%로 채우기
    progress_bar.progress(1.0)
    status_text.markdown("**🎉 단어 분석 진행률: 100% (완료!)**")
    
    all_data = []
    for page_data in results:
        all_data.extend(page_data)
    return all_data

# ==========================================
# 3. Streamlit 메인 UI 대시보드
# ==========================================
st.set_page_config(page_title="Voca 변환기 초고속", layout="centered", page_icon="📝")

st.title("📝 멀티 이미지 단어장 워드 변환기")
st.write("비동기 병렬 분석 엔진을 도입하여 여러 장의 사진도 눈 깜짝할 새에 하나의 워드 파일로 통합합니다.")

if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    st.error("❌ Streamlit Cloud 설정의 Secrets에 GEMINI_API_KEY가 등록되지 않았습니다.")
    st.stop()

uploaded_files = st.file_uploader(
    "단어장 사진 파일들을 선택하세요 (여러 장 동시 선택 가능)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📂 **선택된 파일 수:** {len(uploaded_files)}개")
    
    if st.button("🚀 초고속 병렬 분석 및 Word 파일 생성"):
        client = genai.Client(api_key=api_key)
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        # 병렬 프로세스 실행
        all_word_data = asyncio.run(
            process_all_images(client, uploaded_files, api_key, progress_bar, status_text)
        )
        
        if all_word_data:
            st.success("🎉 모든 사진의 단어 통합 완료!")
            st.write("### 🔍 통합 추출 데이터 미리보기")
            st.dataframe(all_word_data, use_container_width=True)
            
            word_file_buffer = create_word_document(all_word_data)
            
            st.download_button(
                label="📥 깔끔한 Word 문서 다운로드 (.docx)",
                data=word_file_buffer,
                file_name="🔮_통합_영어단어장.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
