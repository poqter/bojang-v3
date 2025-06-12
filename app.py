import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime

# ✅ 페이지 설정
st.set_page_config(page_title="보장 분석 도우미", layout="centered")

# ✅ 기본 템플릿 파일 로드 (다운로드 버튼용)
with open("print.xlsx", "rb") as f:
    default_template_data = f.read()

# ✅ 사이드바 안내문 + 다운로드 버튼 + 제작자 정보
st.sidebar.markdown("### 📘 사용 방법 안내")
st.sidebar.markdown("""
1. **비밀번호를 입력**하여 인증합니다.  
2. `컨설팅보장분석.xlsx` 파일을 업로드하세요.  
3. (선택) `개인용 보장분석 폼.xlsx` 파일도 업로드할 수 있어요.  
4. 분석이 완료되면 결과 파일을 다운로드할 수 있습니다.

📌 참고:  
- `print.xlsx` 파일을 업로드하지 않으면 **기본 폼**이 사용됩니다.  
- 지원 파일 형식: `.xlsx` (엑셀 전용)
""")

st.sidebar.download_button(
    label="📥 기본 폼(print.xlsx) 다운로드",
    data=default_template_data,
    file_name="print.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("---")
st.sidebar.markdown("👨‍💻 **제작자:** 비전본부 드림지점 박병선 팀장")  
st.sidebar.markdown("🗓️ **버전:** v1.0.0")  
st.sidebar.markdown("📅 **최종 업데이트:** 2025-06-13")

# ✅ 비밀번호 인증
PASSWORD = st.secrets["PASSWORD"]

st.title("🔐 보장 분석 도우미")
password_input = st.text_input("비밀번호를 입력하세요", type="password")

if password_input != PASSWORD:
    st.warning("비밀번호가 올바르지 않습니다. 올바른 비밀번호를 입력하세요.")
    st.stop()

st.success("✅ 인증 성공! 계속 진행하세요.")
st.write("아래에 `컨설팅보장분석.xlsx` 파일을 업로드하면 자동으로 결과 분석 파일이 생성됩니다.")

# ✅ 엑셀 파일 업로드
uploaded_main = st.file_uploader("⬆️ 컨설팅보장분석.xlsx 파일을 업로드하세요", type=["xlsx"])
uploaded_print = st.file_uploader("🖨️ (선택) 개인용 보장분석 폼.xlsx 파일을 업로드하세요", type=["xlsx"])

# ✅ print.xlsx 로드
try:
    if uploaded_print:
        print_wb = openpyxl.load_workbook(uploaded_print)
        st.info("✅ 업로드한 print.xlsx를 사용합니다.")
    else:
        print_wb = openpyxl.load_workbook("print.xlsx")
        st.info("📌 기본 내장된 print.xlsx를 사용합니다.")
    print_ws = print_wb.active
except Exception as e:
    st.error(f"❌ print.xlsx 파일을 열 수 없습니다: {e}")
    st.stop()

# ✅ main.xlsx 처리
if uploaded_main:
    try:
        main_wb = openpyxl.load_workbook(uploaded_main, data_only=True)
        main_ws1 = main_wb["계약사항"]
        main_ws2 = main_wb["보장사항"]

        for idx in range(27):
            print_ws.cell(row=10, column=4 + idx).value = main_ws1[f"J{9+idx}"].value

        for row_offset, col in enumerate(['K', 'L']):
            for idx in range(27):
                print_ws.cell(row=8 + row_offset, column=4 + idx).value = main_ws1[f"{col}{9+idx}"].value

        for row in range(2, 8):
            for col in range(6, 30):
                print_ws.cell(row=row, column=col - 2).value = main_ws2.cell(row=row, column=col).value

        for row in range(9, 46):
            for col in range(6, 30):
                print_ws.cell(row=row + 3, column=col - 2).value = main_ws2.cell(row=row, column=col).value

        name_prefix = (main_ws1["B2"].value or "고객")[:3]
        detail_text = main_ws1["D2"].value or ""
        print_ws["A1"] = f"{name_prefix}님의 기존 보험 보장 분석 {detail_text}"

        today_str = datetime.today().strft_
