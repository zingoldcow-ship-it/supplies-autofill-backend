import io
import streamlit as st

from cart_parser import parse_iscreammall_cart_xlsx
from excel_builder import build_output_workbook, workbook_to_bytes

st.set_page_config(page_title="장바구니 엑셀 자동 변환", layout="wide")

st.title("🛒 아이스크림몰 장바구니 엑셀 → 신청서 자동 변환")
st.caption("아이스크림몰 장바구니/견적서 엑셀(.xlsx)을 업로드하면, 신청서에 바로 붙여넣기 좋은 형식으로 자동 변환해드립니다.")

with st.expander("✅ 사용 방법", expanded=True):
    st.markdown(
        """
1) 아이스크림몰에서 **장바구니(견적서) 엑셀**을 다운로드  
2) 아래에서 **.xlsx 파일 업로드**  
3) 변환 결과를 확인한 뒤 **엑셀 다운로드**
        """.strip()
    )

uploaded = st.file_uploader("📎 아이스크림몰 장바구니/견적서 엑셀 업로드", type=["xlsx"])

st.divider()

if uploaded is None:
    st.info("엑셀 파일을 업로드하면 변환 결과 미리보기와 다운로드 버튼이 나타납니다.")
    st.stop()

try:
    # parse (file-like)
    with st.spinner("엑셀에서 상품 정보를 추출 중..."):
        items = parse_iscreammall_cart_xlsx(io.BytesIO(uploaded.getvalue()))

    # preview table
    preview_rows = [
        {
            "품명": it.name,
            "규격": it.spec,
            "수량": it.qty,
            "단가(정가)": it.unit_price_list,
            "단가(할인)": it.unit_price_sale,
            "금액(정가)": it.qty * it.unit_price_list,
            "최종금액": it.qty * it.unit_price_sale,
            "상품코드": it.product_code,
            "사이트": "아이스크림몰",
        }
        for it in items
    ]

    st.success(f"추출 완료! 총 {len(items)}개 품목을 찾았습니다.")
    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    wb = build_output_workbook(items)
    out_bytes = workbook_to_bytes(wb)

    st.download_button(
        label="⬇️ 변환된 엑셀 다운로드",
        data=out_bytes,
        file_name="아이스크림몰_장바구니_변환결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("⚙️ 변환 규칙(참고)"):
        st.markdown(
            """
- **품명/규격**: 상품명에 `( )`, `[ ]`, ` / ` 형태로 규격이 붙어 있으면 자동 분리합니다.  
- **금액(정가) / 최종금액**: 엑셀에 수식이 들어가도록 `=수량*단가`로 계산합니다.  
- **상품코드**: 원본 엑셀에 코드가 없으면 빈칸으로 남겨둡니다.  
- 형식이 다른 엑셀이라면, **헤더(상품명/수량/정가/할인가)** 줄을 자동으로 찾아 최대한 맞춰 읽습니다.
            """.strip()
        )

except Exception as e:
    st.error(f"변환 중 오류가 발생했습니다: {type(e).__name__}: {e}")
    st.write("가능하면 원본 엑셀(개인정보 제거)을 예시로 공유해주시면, 헤더 인식 규칙을 더 튼튼하게 맞춰드릴게요.")
