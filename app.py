import streamlit as st
import pandas as pd
from parser import parse_icecream_excel
from exporter import build_output_excel

st.set_page_config(page_title="아이스크림몰 장바구니 엑셀 → 자동 변환", layout="wide")

st.title("아이스크림몰 장바구니 엑셀 → 자동 변환")
st.caption("아이스크림몰 장바구니/견적서 엑셀을 업로드하면, 학습준비물 정리용 표(품목/규격/수량/단가/금액/상품코드/사이트)로 자동 변환하고 엑셀로 내려받을 수 있습니다.")

with st.expander("사용 방법", expanded=True):
    st.markdown(
        """
1. 아이스크림몰에서 **장바구니(또는 견적서)** 엑셀을 다운로드합니다.
2. 아래에 엑셀 파일(.xlsx)을 업로드합니다.
3. 변환 결과를 웹에서 확인합니다.
4. **결과 엑셀 다운로드** 버튼으로 내려받습니다.
        """.strip()
    )

uploaded = st.file_uploader("① 아이스크림몰 장바구니/견적서 엑셀 업로드", type=["xlsx"])

if not uploaded:
    st.info("엑셀 파일을 업로드하면 변환이 시작됩니다.")
    st.stop()

try:
    items_df = parse_icecream_excel(uploaded.getvalue())
except Exception as e:
    st.error("엑셀을 읽는 중 오류가 발생했습니다. (파일 형식/헤더가 다른 경우일 수 있어요)")
    st.exception(e)
    st.stop()

st.subheader("추출 결과 미리보기 (정규화)")
st.dataframe(items_df, use_container_width=True, hide_index=True)

# Build downloadable excel
out_bytes = build_output_excel(items_df)

st.download_button(
    "✅ 결과 엑셀 다운로드",
    data=out_bytes,
    file_name=f"아이스크림_변환결과_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
