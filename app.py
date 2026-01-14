import io
import streamlit as st
from cart_parser import parse_iscreammall_cart_xlsx
from excel_builder import build_output_workbook

st.set_page_config(page_title="장바구니 엑셀 → 학습준비물 신청서 자동 변환", layout="wide")

st.title("🛒 아이스크림몰 장바구니 엑셀 → 학습준비물 신청서 자동 변환")
st.caption("장바구니(견적서) 엑셀을 업로드하면 품명/규격/정가·할인가/수량을 자동 정리하고, 금액 계산 수식이 포함된 신청서 엑셀을 생성합니다.")

with st.expander("사용 방법", expanded=True):
    st.markdown(
        """
1. 아이스크림몰에서 장바구니(견적서) 엑셀을 다운로드합니다.  
2. 아래에서 파일을 업로드합니다.  
3. 변환 버튼을 누르면 ‘신청서(할인가 기준)’ 및 ‘가격정보(정가-할인가)’ 시트가 포함된 엑셀을 다운로드할 수 있습니다.  
        """.strip()
    )

col1, col2 = st.columns([1, 1])

with col1:
    uploaded = st.file_uploader("📎 아이스크림몰 장바구니/견적서 엑셀 업로드", type=["xlsx"])
    school_title = st.text_input("신청서 제목(선택)", value="■ 학습준비물 신청서 ■")
    term_title = st.text_input("학년도/학기(선택)", value="2026학년도 1학기")
    grade_info = st.text_input("학년 정보(선택)", value="(  )학년 부장 교사 : (인)")

with col2:
    st.markdown("### 출력 안내")
    st.markdown("- **신청서(할인가 기준)**: 기존 신청서 형식에 맞춰 `단가=할인가`로 입력하고 `금액=수량×단가` 수식이 자동으로 들어갑니다.")
    st.markdown("- **가격정보(정가-할인가)**: 정가/할인가를 모두 확인할 수 있도록 별도 시트로 정리합니다.")
    st.info("상품코드는 장바구니 엑셀에 포함되지 않는 경우가 많아, 기본적으로 빈 칸으로 출력됩니다. (필요 시 수동 입력)")

if uploaded is not None:
    try:
        # Parse
        with st.spinner("엑셀에서 상품 정보를 추출 중..."):
            # streamlit uploader -> bytes -> temp in memory
            data = uploaded.getvalue()
            tmp = io.BytesIO(data)
            # openpyxl requires a filename or file-like; file-like ok
            # But our parser expects path; so write to temp file
            import tempfile, os
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
                f.write(data)
                tmp_path = f.name

            items = parse_iscreammall_cart_xlsx(tmp_path)
            os.unlink(tmp_path)

        st.success(f"✅ 추출 완료: {len(items)}개 품목")

        # Preview table (minimal)
        import pandas as pd
        preview = pd.DataFrame([
            {"품명(원문)": it.name_raw, "수량": it.qty, "단가(정가)": it.unit_price_list, "단가(할인)": it.unit_price_sale}
            for it in items
        ])
        st.dataframe(preview, use_container_width=True, hide_index=True)

        if st.button("📄 신청서 엑셀로 변환 & 다운로드 준비", type="primary"):
            with st.spinner("출력 엑셀 생성 중..."):
                wb = build_output_workbook(
                    items,
                    school_title=school_title,
                    term_title=term_title,
                    grade_info=grade_info,
                )
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

            st.download_button(
                label="⬇️ 변환된 엑셀 다운로드",
                data=out,
                file_name="학습준비물_신청서_변환결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"변환 중 오류가 발생했습니다: {type(e).__name__}: {e}")
else:
    st.warning("먼저 장바구니/견적서 엑셀(.xlsx)을 업로드해 주세요.")
