import streamlit as st
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
from datetime import datetime
from google import genai
import time
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# --- [1. 기존 설정 그대로 유지] ---
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)


# --- [2. PDF 텍스트 추출 함수] ---
def get_pdf_text_from_upload(uploaded_files):
    all_text = ""
    for uploaded_file in uploaded_files:
        try:
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            for page in doc:
                all_text += page.get_text()
            doc.close()
        except Exception as e:
            st.error(f"❌ [파일 읽기 오류] {uploaded_file.name}: {e}")
    return all_text


# --- [3. 웹 화면 구성] ---
st.set_page_config(page_title="QA TC Generator", layout="wide")
st.title("🚀 테스트 케이스 생성기")

uploaded_files = st.file_uploader("기획서 PDF 파일을 선택하세요", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🪄 테스트 케이스 생성 시작"):
        with st.spinner("기획서 분석 및 데이터 정밀 추출 중..."):

            plan_content = get_pdf_text_from_upload(uploaded_files)

            if plan_content:
                # 🚨 [503 에러 대응 재시도 로직 추가]
                response = None
                max_retries = 3
                for i in range(max_retries):
                    try:
                        # 모델 선정
                        model_list = list(client.models.list())
                        target_model = next((m.name for m in model_list if 'gemini' in m.name.lower()), "gemini-1.5-flash")

                        # [마스터 프롬프트: 사용자 지침 100% 반영 - 생략 절대 없음]
                        prompt = f"""
                        너는 QA 엔지니어이며 TC 작성 전문가이다.
                        기획서에 작성된 UI 요소 및 Description에 따라 TC를 작성해라.
                        출력은 반드시 '|'로 구분된 13개 컬럼 표 형식이어야 한다.

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
                           - Label 1: 대분류 영역 명칭
                                - 레이아웃 확인 케이스에서 나열한 항목명을 그대로 사용하라.(예 : 로고 영역 / 아이디 영역 / 비밀번호 영역 / 검색 영역 / 로그인 버튼)
                                - Label 1의 첫 번째 행은 반드시 해당 영역의 전체 구성을 확인하는 케이스여야 한다.
                           - Label 2: 구체적 확인 대상 **영역별 첫 번째 케이스 (Layout Check)**:
                                - Label 1에서 나열된 항목의 세부 테스트 케이스를 작성하며, Label 2 : "UI 확인", Label 3: "-", 수행 절차: "검색 영역의 구성을 확인한다.", 기대결과는 "항목명 / 검색 인풋박스 / 검색 버튼으로 구성되어 노출된다.")
                                - 구체적인 컴포넌트 명칭 사용(예: 아이디 인풋박스, 로그인 버튼)
                                - 다음 케이스 부터는 앞서 나열한 항목의 UI 확인 / 기능 확인 순으로 작성된다.
                           - Label 3: 확인 성격 (UI 확인 / 기능 확인 / 밸리데이션 확인)
                                - Label 2와 Label 3의 명칭이 동일한 경우, Label 3는 반드시 '-'로 표기한다.
                                - Label 3에 작성될 항목이나 기능명이 없을 경우 '-로 표기한다.
                                **구성 요소별 단계적 시나리오 (Sequential Order)**:
                                    아래의 순서대로 **반드시 행을 생성**하라.
                                        ① **UI 확인**: 컴포넌트의 노출 상태, 레이블, Placeholder 등을 확인.
                                        ② **기능 확인**: 클릭, 입력 값 표시, 마스킹 처리 등 기본 동작 확인.
                                        ③ **밸리데이션 확인**: 경계값(29/30/31자), 특수문자, 공백 등 유효성 체크.
                        4. **문구 규칙**: 사전 조건 / 참고 컬럼은 명사형 간결체로 작성하고 수행 절차와 내용이 중복되지 않게 하라. (예: 30자 미만 입력(ex.abcdefg), 미입력 상태)
                            - 해당 행에서 테스트할 **단 하나의 구체적인 입력 조건 또는 데이터**만 명시하라.

                        ### [TC 작성 문체 통일]
                        1. **수행 절차 (간결화)**:
                        - 사용자 스타일인 **"조건에 맞게 입력한다."** 또는 **"항목별 노출 여부를 확인한다."**라는 문구로 통일하되, 각 행의 맥락에 맞게 작성하라.
                        - 예시 :
                            - **로그인 관련 영역**: "조건에 맞게 입력한 뒤 [로그인] 버튼을 클릭한다."
                            - **검색 관련 영역**: "조건에 맞게 입력한 뒤 [검색] 버튼을 클릭한다."
                            - **일반 입력/선택**: "조건에 맞게 입력/선택 후 결과를 확인한다."
                            - **UI 확인**: "해당 영역의 구성 요소 및 노출 상태를 확인한다."

                        2. **기대 결과 (단일 결과)**:
                        - 해당 행의 사전 조건에 명시된 특정 데이터에 대한 **단 하나의 예상 결과**만 정확히 기술하라.
                        - 예시: 29자가 정상적으로 입력되고 표시된다. / 30자까지만 입력되고 31자부터는 입력되지 않는다.

                        3. 구분자
                        / 기호로 항목의 구분을 한다.


                        ### [작성 예시]
                        | TC ID | 프로그램명 | 화면 ID | 요구사항 ID | Label 1 | Label 2 | Label 3 | 사전 조건 / 참고 | 수행 절차 | 기대 결과 | 결과 | 수행자 | 비고 |
                        |---|---|---|---|---|---|---|---|---|---|---|---|---|
                        | | 로그인 | - | - | 대분류 영역 명칭 | 구체적 확인 대상 | 확인 성격 | 아이디 미입력 | 아이디 입력 영역을 확인한다. | 가이드 문구가 노출된다. | | | |

                        [기획서 내용]
                        {plan_content}
                        """

                        response = client.models.generate_content(model=target_model, contents=prompt)
                        break  # 성공 시 루프 탈출
                    except Exception as e:
                        if "503" in str(e) and i < max_retries - 1:
                            st.warning(f"⚠️ 서버 과부하로 재시도 중입니다... ({i+1}/{max_retries})")
                            time.sleep(5)
                        else:
                            st.error(f"❌ 에러 발생: {e}")
                            st.stop()

                if response:
                    raw_data = response.text.strip()

                    # 데이터 파싱 로직 (원본 보존)
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
                        columns = ["TC ID", "프로그램명(화면명)", "화면 ID", "요구사항 ID", "Label 1", "Label 2", "Label 3", "사전 조건 / 참고", "수행 절차", "기대 결과", "결과", "수행자", "비고"]
                        df = pd.DataFrame(table_data, columns=columns)
                        df = df.replace(r'<br\s*/?>', '\n', regex=True)
                        df.loc[df['Label 2'] == df['Label 3'], 'Label 3'] = '-'
                        df['TC ID'] = ""; df['결과'] = "Not Tested"; df['수행자'] = ""

                        # --- [엑셀 스타일링 및 병합 로직 - 원본 보존] ---
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Test Case')
                            ws = writer.sheets['Test Case']

                            pt_font = Font(name='Pretendard', size=9)
                            header_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

                            column_widths = {'A': 10, 'B': 15, 'C': 15, 'D': 10, 'E': 15, 'F': 18, 'G': 15, 'H': 18, 'I': 35, 'J': 35, 'K': 10, 'L': 10, 'M': 30}
                            for i, (col, width) in enumerate(column_widths.items(), 1):
                                ws.column_dimensions[col].width = width
                                for row in range(1, len(df) + 2):
                                    cell = ws.cell(row=row, column=i)
                                    cell.font = pt_font; cell.border = thin_border
                                    if row == 1:
                                        cell.alignment = center_align; cell.fill = header_fill; cell.font = Font(name='Pretendard', size=9, bold=True)
                                    else:
                                        cell.alignment = left_align if i in [9, 10, 13] else center_align

                            # 계층 병합 로직 (원본 보존)
                            for col_idx in [5, 6, 7]:
                                start_row = 2
                                for row in range(3, len(df) + 3):
                                    prev_val = str(ws.cell(row=row - 1, column=col_idx).value or "").strip()
                                    curr_val = str(ws.cell(row=row, column=col_idx).value or "").strip()
                                    upper_changed = any(str(ws.cell(row=row - 1, column=u).value) != str(ws.cell(row=row, column=u).value) for u in range(5, col_idx))
                                    if curr_val != prev_val or upper_changed or row == len(df) + 2:
                                        if row - 1 > start_row:
                                            merge_val = ws.cell(row=start_row, column=col_idx).value
                                            if col_idx == 5 or (str(merge_val).strip() != '-' and merge_val):
                                                try: ws.merge_cells(start_row=start_row, start_column=col_idx, end_row=row - 1, end_column=col_idx)
                                                except: pass
                                        start_row = row

                        # --- [결과 안내 및 다운로드 버튼] ---
                        st.balloons()
                        st.success(f"✅ 완료! 총 {len(df)}개의 테스트 케이스가 생성되었습니다.")
                        st.download_button(
                            label="📥 엑셀 파일 다운로드",
                            data=output.getvalue(),
                            file_name=f"TC_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        # --- 엑셀을 받기 전, 어떤 내용이 생성되었는지 화면에서 표 형식으로 미리 확인 ---
                        st.dataframe(df)