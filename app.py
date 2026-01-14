import streamlit as st
import pandas as pd

from parser import parse_icecream_excel
from template_writer import fill_template

st.set_page_config(page_title="아이스크림몰 장바구니 → 양식 자동 채움", layout="wide")

st.title("아이스크림몰 장바구니 엑셀 → 양식 자동 채움")
st.caption("장바구니/견적서 엑셀을 업로드하면, 품목/규격/단가/수량을 추출해서 '아이스크림 장바구니 양식.xlsx' 형태로 자동 작성합니다.")

with st.expander("사용 방법", expanded=True):
    st.markdown(
        """
1) 아이스크림몰에서 **장바구니(또는 견적서) 엑셀**을 다운로드  
2) 아래에 업로드  
3) 변환 결과를 웹에서 확인  
4) **양식 채움 엑셀**을 다운로드
        """.strip()
    )

col1, col2 = st.columns([1, 1])

with col1:
    uploaded = st.file_uploader("① 아이스크림몰 장바구니/견적서 엑셀 업로드", type=["xlsx"])

with col2:
    template_up = st.file_uploader("② (선택) 다른 양식 파일 업로드", type=["xlsx"], help="업로드하지 않으면 기본 제공 양식(template.xlsx)을 사용합니다.")

if uploaded is None:
    st.info("엑셀 파일을 업로드하면 변환이 시작됩니다.")
    st.stop()

try:
    items_df = parse_icecream_excel(uploaded)
except Exception as e:
    st.error("엑셀을 읽는 중 오류가 발생했습니다. 파일 형식이 다른지 확인해 주세요.")
    st.exception(e)
    st.stop()

if items_df.empty:
    st.warning("엑셀에서 상품 정보를 찾지 못했습니다. (파일이 다른 형식이거나, 상품 행이 비어 있을 수 있어요)")
    st.stop()

st.subheader("추출 결과 미리보기 (정규화)")
st.dataframe(items_df, use_container_width=True, hide_index=True)

# Load template bytes
if template_up is not None:
    template_bytes = template_up.getvalue()
else:
    with open("template.xlsx", "rb") as f:
        template_bytes = f.read()

filled_bytes = fill_template(template_bytes, items_df)

st.subheader("다운로드")
st.download_button(
    label="양식 채움 엑셀 다운로드",
    data=filled_bytes,
    file_name="아이스크림_양식_자동작성.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("Tip: 만약 아이스크림몰 엑셀 형식이 바뀌어서 인식이 안 되면, 그 엑셀을 1개만 더 보내주시면 컬럼 탐지 규칙을 추가로 보강해 드릴게요.")
