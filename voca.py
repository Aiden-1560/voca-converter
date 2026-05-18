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
    tcPr.append(tcPr)

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

def process_images_safely(client, uploaded_files, api_key, progress_bar, status_text):
    all_data = []
    total_files = len(uploaded_files)
    
    prompt = """
    이 이미지에서 영어 단어, 우리말 뜻, 영영 풀이를 추출해서 정확한 JSON 배열 형식으로 출력해줘.
    필기구로 수정한 흔적이나 추가로 적은 필기는 무시하고, 원래 인쇄되어 있던 텍스트만 추출해줘.
    결과는 오직 아래 구조를 가진 JSON 데이터만 반환해야 해:
    [
      {"word": "단어", "meaning": "품사 및 뜻", "definition": "영영 풀이 내용"}
    ]
    """
    
    current_displayed_percent = 0
    
    for idx, file in enumerate(uploaded_files):
        target_percent = int(((idx + 1) / total_files) * 100)
        
        pre_target = target_percent - 3 if target_percent > 3 else 0
        while current_displayed_percent < pre_target:
            current_displayed_percent += 1
            progress_bar.progress(current_displayed_percent / 100)
            status_text.info(f"🌱 단어 아카이브 분석 중... {current_displayed_percent}% (`{file.name}` 판독 중)")
            time.sleep(0.01)
            
        page_data = None
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                file.seek(0)
                image_bytes = file.read()
                
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[
                        types.Part.from_bytes(data=image_bytes, mime_type=file.type),
                        prompt
                    ],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                page_data = json.loads(response.text)
                all_data.extend(page_data)
                break
                
            except Exception as e:
                if attempt < max_retries - 1:
                    status_text.warning(f"⏳ 구글 AI 서버에 동시 접속자가 많아 통신을 재시도하고 있습니다... ({attempt + 1}/{max_retries})")
                    time.sleep(2.5)
                else:
                    # [수정] 하루 무료 사용량 초과(분당 트래픽 차단 포함) 상황에 맞는 직관적인 안내문구로 전면 개편
                    st.error(f"🔋 **오늘 제공된 구글 AI 분석 에너지를 모두 소모했습니다.**\n\n구글 무료 서버 정책상 일일 업로드 제한(또는 단시간 과다 요청 수 제한)에 도달하여 `{file.name}` 처리 단계에서 멈췄습니다. 내일 다시 시도하시거나, 다른 구글 계정의 API 키로 교체하시면 즉시 이어서 사용이 가능합니다.")
            
        while current_displayed_percent < target_percent:
            current_displayed_percent += 1
            progress_bar.progress(current_displayed_percent / 100)
            status_text.info(f"✨ 단어 아카이브 분석 중... {current_displayed_percent}% (`{file.name}` 정제 완료)")
            time.sleep(0.01)
            
        time.sleep(0.2)
        
    progress_bar.progress(1.0)
    status_text.success("🌿 정제 프로세스가 성공적으로 완료되었습니다. (100%)")
    return all_data

# ==========================================
# 3. Streamlit 메인 UI 대시보드
# ==========================================
st.set_page_config(page_title="Voca-converter", layout="centered", page_icon="📝")

# [수정] 눈이 부시던 형광 빨간색 버튼을 차분하고 신뢰감을 주는 '고급 딥 네이비' 테마로 변경하는 스타일 태그 추가
st.markdown("""
    <style>
    div.stButton > button:first-child {
        background-color: #1E3A8A !important;
        color: white !important;
        border: none !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #172554 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("### 📝 Voca-converter")
    st.caption("MADE BY MANJU · 스마트 교재 연구 솔루션")
    st.divider() 
    
    st.info("💡 **안내**\n\n교재나 유인물 사진을 업로드하시면, 수업 및 제본에 즉시 활용할 수 있는 단정하고 정돈된 **표 형태의 워드 문서(.docx)**로 통합 변환해 드립니다.")
    st.write("") 

    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    else:
        st.error("❌ Streamlit Cloud 설정의 Secrets에 GEMINI_API_KEY가 등록되지 않았습니다.")
        st.stop()

    uploaded_files = st.file_uploader(
        "교재 및 단어장 사진 파일을 선택하세요 (여러 장 동시 선택 가능)", 
        type=["jpg", "jpeg", "png"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        st.write("")
        st.markdown(f"📂 **선택된 아카이브:** 총 `{len(uploaded_files)}개` 파일 변환 대기 중")
        
        # 버튼을 누르면 위에서 정의한 CSS 덕분에 고급스러운 딥 네이비 색상으로 출력됩니다.
        if st.button("✨ 업로드된 문서 분석 및 Word 파일 생성", type="primary"):
            client = genai.Client(api_key=api_key)
            
            status_text = st.empty()
            progress_bar = st.progress(0)
            
            all_word_data = process_images_safely(client, uploaded_files, api_key, progress_bar, status_text)
            
            if all_word_data:
                st.toast("단어 데이터 정제가 완료되었습니다!")
                st.write("---")
                st.write("### 🔍 데이터 통합 미리보기")
                st.dataframe(all_word_data, use_container_width=True)
                
                word_file_buffer = create_word_document(all_word_data)
                
                st.write("")
                st.download_button(
                    label="📥 정제된 수업용 Word 문서 다운로드 (.docx)",
                    data=word_file_buffer,
                    file_name="🔮_통합_영어단어장.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
