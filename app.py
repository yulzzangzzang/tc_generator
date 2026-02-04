import streamlit as st
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
from datetime import datetime
from google import genai
import time
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# --- [1. 설정 및 API 키] ---
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)


# --- [2. 데이터 추출 함수] ---
def get_pdf_text_from_upload(uploaded_files):
    all_text = ""
    for uploaded_file in uploaded_files:
        try:
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            for page in doc:
                all_text += page.get_text()
            doc.close()
        except Exception as e:
            st.error(f"❌ [PDF 오류] {uploaded_file.name}: {e}")
    return all_text


def get_old_excel_data(uploaded_excel):
    if uploaded_excel:
        try:
            df = pd.read_excel(uploaded_excel)
            # 데이터가 너무 많을 경우를 대비해 텍스트로 변환
            return df.to_string(index=False)
        except Exception as e:
            st.error(f"❌ [기존 엑셀 읽기 오류]: {e}")
            return None
    return None


# --- [3. 미리보기 스타일링 함수] ---
def highlight_tc_rows(row):
    note = str(row.비고)
    if '[수정]' in note:
        return ['background-color: #FFFF00'] * len(row)
    elif '[신규]' in note:
        return ['background-color: #CCEEFF'] * len(row)
    elif '[삭제]' in note:
        return ['background-color: #D3D3D3'] * len(row)
    return [''] * len(row)


# --- [4. 웹 화면 구성] ---
st.set_page_config(page_title="QA TC Generator Pro", layout="wide")
st.title("🚀 테스트 케이스 생성 및 업데이트")

st.info("💡 **신규 생성**: 기획서 PDF만 업로드\n\n💡 **업데이트**: 기획서 PDF와 이전에 생성한 엑셀 파일을 함께 업로드")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 1. 기획서 업로드 (필수)")
    uploaded_files = st.file_uploader("기획서 PDF (여러 개 가능)", type="pdf", accept_multiple_files=True)
with col2:
    st.subheader("📂 2. 기존 TC 업로드 (선택/업데이트용)")
    old_excel = st.file_uploader("이전에 다운받은 TC 엑셀 파일", type="xlsx")

if uploaded_files:
    is_update = old_excel is not None
    button_label = "🪄 변경 사항 분석 및 TC 업데이트" if is_update else "🪄 테스트 케이스 신규 생성"

    if st.button(button_label, type="primary"):
        with st.spinner("기획서 분석 및 TC 생성 중..."):

            plan_content = get_pdf_text_from_upload(uploaded_files)
            old_data_text = get_old_excel_data(old_excel)

            # [모드별 특화 지시문]
            mode_instruction = ""
            if is_update:
                mode_instruction = f"""
                ### [업데이트 모드 (중요)]
                1. 제공된 [기존 데이터]와 새로운 [기획서 내용]을 정밀 비교하라.
                2. 변경된 내용이 있는 행은 '비고' 컬럼에 [수정]이라 표기하고 내용을 갱신하라.
                3. 새로운 기능이나 UI 요소는 [신규]라 표기하고 추가하라.
                4. 기존에는 있었으나 새 기획서에서 사라진 요구사항은 [삭제 대상]이라 표기하라.
                5. 기존의 TC ID 및 전반적인 구조를 최대한 유지하며 업데이트하라.
                """
            else:
                mode_instruction = f"""
                ### [신규 생성 모드]
                1. 기획서 내용을 바탕으로 화면 요구사항 및 테스트 케이스를 처음부터 상세히 추출하라.
                """

            # [마스터 프롬프트: 기존 지침 100% 보존]
            prompt = f"""
            너는 QA 엔지니어이며 TC 작성 전문가이다.
            기획서에 작성된 UI 요소 및 Description에 따라 TC를 작성해라.
            출력은 반드시 '|'로 구분된 13개 컬럼 표 형식이어야 한다.

            {mode_instruction}

            ### [핵심 미션]
            - 기획서에 명시된 모든 UI 요소(아이콘 / 버튼 / 인풋박스 / 필터 등)를 빠짐없이 도출하라.

            ### [ISTQB 기반 테스트 설계 규칙]
            1. **경계값 분석 (Boundary Value Analysis)**: 
               - 입력란(숫자, 글자 수 등)에 제한이 있는 경우, [최솟값-1, 최솟값, 최솟값+1, 최댓값-1, 최댓값, 최댓값+1] 등 경계값을 확인하는 케이스를 반드시 포함한다.
            2. **동등 분할 (Equivalence Partitioning)**: 
               - 유효한 입력 값(Pass)뿐만 아니라 유효하지 않은 입력 값(Fail) 군집을 정의하여 각각 최소 1개 이상의 케이스를 작성한다.
            3. **에러 추측 (Error Guessing)**: 
               - 기획서에 명시되지 않았더라도 '특수문자 입력', '공백 입력', '중복 클릭', '뒤로가기 시 데이터 유지' 등 시니어 QA로서 예상되는 결함 시나리오를 추가한다.
            4. **결정 테이블 (Decision Table)**: 
               - 여러 조건이 복합적으로 얽힌 로직(예: 권한별 접근 제어, 조건별 할인 등)은 조건의 조합에 따른 결과 값을 각각 별개의 행으로 작성한다.

            ### [TC 구성 및 위계]
            1. 화면 진입 및 전체 레이아웃 확인 케이스를 최상단에 배치하라.
            2. **Label 위계**: 
               - Label 1: 대분류 영역 명칭 (예: 로고 영역 / 검색 영역)
               - Label 2: 구체적 확인 대상 (예: 아이디 인풋박스, 로그인 버튼)
               - Label 3: 확인 성격 (UI 확인 / 기능 확인 / 밸리데이션 확인)
               - Label 3에 작성될 항목이나 기능명이 없을 경우 '-로 표기한다.
            3. **구성 요소별 단계적 시나리오**: 아래 순서로 행 생성
               ① UI 확인 -> ② 기능 확인 -> ③ 밸리데이션 확인

            ### [TC 작성 문체 통일]
            1. **수행 절차**: "조건에 맞게 입력한다." 또는 "항목별 노출 여부를 확인한다." 문구 통일.
            2. **기대 결과**: 특정 데이터에 대한 '단 하나의 예상 결과'만 기술.
            3. **구분자**: '/' 기호로 항목 구분.

            ### [작성 예시 컬럼]
            | TC ID | 프로그램명(화면명) | 화면 ID | 요구사항 ID | Label 1 | Label 2 | Label 3 | 사전 조건 / 참고 | 수행 절차 | 기대 결과 | 결과 | 수행자 | 비고 |

            [기존 데이터]
            {old_data_text if old_data_text else "없음"}

            [기획서 내용]
            {plan_content}
            """

            # --- API 호출 및 재시도 로직 ---
            response = None
            max_retries = 3
            for i in range(max_retries):
                try:
                    model_list = list(client.models.list())
                    target_model = next((m.name for m in model_list if 'gemini' in m.name.lower()), "gemini-1.5-flash")
                    response = client.models.generate_content(model=target_model, contents=prompt)
                    break
                except Exception as e:
                    if "503" in str(e) and i < max_retries - 1:
                        time.sleep(5)
                        continue
                    st.error(f"❌ 에러 발생: {e}");
                    st.stop()

            if response:
                raw_data = response.text.strip()
                lines = [line.strip() for line in raw_data.split('\n') if '|' in line]
                lines = [line for line in lines if not all(c in '| -:' for c in line)]

                table_data = []
                for line in lines:
                    cells = [cell.strip() for cell in line.split('|')]
                    if cells and cells[0] == "": cells.pop(0)
                    if cells and cells[-1] == "": cells.pop()
                    if len(cells) >= 10 and "TC ID" not in cells[0]:
                        if len(cells) < 13: cells.extend([""] * (13 - len(cells)))
                        table_data.append(cells[:13])

                if table_data:
                    columns = ["TC ID", "프로그램명(화면명)", "화면 ID", "요구사항 ID", "Label 1", "Label 2", "Label 3", "사전 조건 / 참고",
                               "수행 절차", "기대 결과", "결과", "수행자", "비고"]
                    df = pd.DataFrame(table_data, columns=columns)

                    # 데이터 정리
                    df.loc[df['Label 2'] == df['Label 3'], 'Label 3'] = '-'
                    if not is_update:  # 신규 생성 시 초기값
                        df['TC ID'] = "";
                        df['결과'] = "Not Tested";
                        df['수행자'] = ""

                    # --- [엑셀 스타일링 및 저장] ---
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Test Case')
                        ws = writer.sheets['Test Case']

                        # 스타일 설정
                        yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        blue = PatternFill(start_color='CCEEFF', end_color='CCEEFF', fill_type='solid')
                        gray = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
                        header_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                        header_font = Font(name='맑은 고딕', size=9, bold=True)
                        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                                             bottom=Side(style='thin'))

                        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=len(df) + 1), 1):
                            # '비고' 컬럼은 13번째(M열)
                            note = str(ws.cell(row=r_idx, column=13).value)
                            for c_idx, cell in enumerate(row, 1):
                                cell.border = thin_border
                                if r_idx == 1:
                                    cell.fill = header_fill;
                                    cell.font = header_font
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                else:
                                    cell.font = Font(name='맑은 고딕', size=9)
                                    # 업데이트 상태에 따른 색상 적용
                                    if "[수정]" in note:
                                        cell.fill = yellow
                                    elif "[신규]" in note:
                                        cell.fill = blue
                                    elif "[삭제]" in note:
                                        cell.fill = gray

                                    align = 'left' if c_idx in [9, 10, 13] else 'center'
                                    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)

                        # 열 너비 설정
                        column_widths = [10, 15, 12, 12, 15, 18, 15, 18, 35, 35, 10, 10, 25]
                        for i, width in enumerate(column_widths, 1):
                            ws.column_dimensions[chr(64 + i)].width = width

                    # --- [최종 결과 표시] ---
                    st.balloons()
                    st.success(f"✅ 분석 완료! 총 {len(df)}개의 케이스가 준비되었습니다.")

                    st.subheader("📝 추출 결과 미리보기")
                    st.dataframe(df.style.apply(highlight_tc_rows, axis=1), use_container_width=True)

                    st.download_button(
                        label="📥 테스트 케이스 다운로드 (Excel)",
                        data=output.getvalue(),
                        file_name=f"TC_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )