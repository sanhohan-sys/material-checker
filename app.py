import streamlit as st
import pandas as pd
import math
import io

# 1. 웹페이지 설정 (타이틀, 레이아웃)
st.set_page_config(
    page_title="자재 청구 리스트 자동 점검 시스템",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. 스타일 및 타이틀
st.title("📋 자재 청구 리스트 자동 점검 시스템")
st.markdown("자재 청구 엑셀 파일을 업로드하시면 검증 규칙에 맞춰 자동으로 점검하고 분석 결과를 보여줍니다.")
st.markdown("---")

# 사이드바 - 규칙 설명
with st.sidebar:
    st.header("⚙️ 검증 규칙 정보")
    st.markdown("""
    - **A열**: 자재코드
    - **H열**: 포장사양 (예: `P0804` $\rightarrow$ `08` 파싱)
    - **J열**: 양산청구수량 (필요수량과 1:1 일치 여부)
    - **L열**: 공용청구수량 (포장사양 `08` 기준 검증)
        - *필요수량 500개 이하*: **30개** 고정 추가 청구
        - *필요수량 500개 초과*: **필요수량의 5%** 추가 청구 (올림 처리)
    - **O열**: 필요수량
    """)
    st.info("💡 엑셀 파일은 첫 번째 행에 열 이름(헤더)이 있고, 실 데이터는 2번째 행부터 시작하는 구조여야 합니다.")

# 3. 파일 업로더 구성
uploaded_file = st.file_uploader("검증할 자재 청구 엑셀 파일(.xlsx, .xls)을 선택하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 엑셀 파일 읽기 (헤더 포함)
        df = pd.read_excel(uploaded_file, header=0)
        
        # 행 개수 확인
        total_rows = len(df)
        st.success(f"📂 파일 업로드 성공! 총 {total_rows}개의 데이터 행을 감지했습니다.")
        
        # 검증 결과 및 오류 내역 리스트 초기화
        error_list = []
        verified_data = [] # 웹 화면 표기용 리스트
        
        # 행별 검증 시작
        for idx, row in df.iterrows():
            excel_row_num = idx + 2  # 엑셀 기준 실제 행 번호 (헤더가 1행이므로 +2)
            row_status = "정상"
            error_reason = []
            
            # 1. 열 개수 확보 여부 확인 (최소 O열이 존재해야 하므로 15개 열 필요)
            if len(row) < 15:
                error_list.append({
                    "행 번호": excel_row_num,
                    "자재코드": "N/A",
                    "오류 유형": "형식 오류",
                    "상세 내용": "엑셀 열 개수가 부족합니다. A열부터 O열까지 올바르게 채워져 있는지 확인해 주세요."
                })
                continue
                
            # 데이터 추출 (안전하게 문자열/숫자 처리)
            material_code = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else "N/A"
            packing_spec = str(row.iloc[7]).strip() if not pd.isna(row.iloc[7]) else ""
            qty_mass = row.iloc[9]
            qty_common = row.iloc[11]
            qty_required = row.iloc[14]
            
            # 빈 값(NaN) 처리
            if pd.isna(qty_required) or pd.isna(qty_mass) or pd.isna(qty_common):
                row_status = "오류"
                error_reason.append("필수 입력 수량 값(양산청구/공용청구/필요수량) 중 일부가 비어 있습니다.")
            else:
                try:
                    qty_required = float(qty_required)
                    qty_mass = float(qty_mass)
                    qty_common = float(qty_common)
                except ValueError:
                    row_status = "오류"
                    error_reason.append("수량 데이터가 올바른 숫자 형식이 아닙니다.")
                    qty_required, qty_mass, qty_common = 0, 0, 0
            
            if row_status != "오류":
                # [규칙 1] 포장사양 파싱 (P0804 -> '08' 파싱)
                parsed_spec = ""
                if len(packing_spec) >= 3:
                    parsed_spec = packing_spec[1:3]
                else:
                    row_status = "오류"
                    error_reason.append(f"포장사양 형식('{packing_spec}')이 부적합합니다. (최소 3자리 필요)")
                
                # [규칙 2] 양산청구수량(J열) == 필요수량(O열) 검증
                if qty_mass != qty_required:
                    row_status = "오류"
                    error_reason.append(f"양산청구수량 불일치 (필요: {int(qty_required)}개 / 청구: {int(qty_mass)}개)")
                
                # [규칙 3] 공용청구수량(L열) 검증 (포장사양이 '08'인 경우)
                if parsed_spec == "08":
                    expected_common = 30 if qty_required <= 500 else math.ceil(qty_required * 0.05)
                    if qty_common != expected_common:
                        row_status = "오류"
                        error_reason.append(
                            f"공용청구수량 불일치 (필요가 {int(qty_required)}개이므로 기준치는 {expected_common}개이나, 현재 {int(qty_common)}개)"
                        )
            
            # 오류 내역 누적
            if row_status == "오류":
                for reason in error_reason:
                    error_list.append({
                        "행 번호": excel_row_num,
                        "자재코드": material_code,
                        "오류 유형": "수량/규칙 불일치" if "불일치" in reason else "데이터 누락/형식",
                        "상세 내용": reason
                    })
            
            # 검증 결과를 데이터프레임 기록용으로 보관
            verified_data.append({
                "행 번호": excel_row_num,
                "자재코드": material_code,
                "포장사양": packing_spec,
                "양산청구": qty_mass,
                "공용청구": qty_common,
                "필요수량": qty_required,
                "결과": "정상" if row_status == "정상" else "❌ 불일치/에러",
                "비고": ", ".join(error_reason) if error_reason else "정상 검증 완료"
            })
            
        # 4. 검증 결과 시각화
        error_df = pd.DataFrame(error_list)
        verified_df = pd.DataFrame(verified_data)
        
        # 요약 통계 정보 표시
        col1, col2, col3 = st.columns(3)
        total_items = len(verified_df)
        error_items = len(error_df["행 번호"].unique()) if not error_df.empty else 0
        success_items = total_items - error_items
        
        with col1:
            st.metric("총 검증 자재 수", f"{total_items} 건")
        with col2:
            st.metric("검증 성공 (정상)", f"{success_items} 건", delta=f"{success_items}정상")
        with col3:
            st.metric("검증 실패 (불일치)", f"{error_items} 건", delta=f"-{error_items}건" if error_items > 0 else "0건", delta_color="inverse")
            
        # 웹용 팝업(Modal) 알림 효과를 위한 Streamlit dialog 기능 적용
        @st.dialog("🚨 검증 결과 알림")
        def show_popup_results(err_count):
            if err_count == 0:
                st.balloons()
                st.success("🎉 모든 자재 청구 수량이 완벽하게 점검 규칙과 일치합니다!")
            else:
                st.error(f"⚠️ 총 {err_count}건의 불일치 혹은 데이터 오류가 발견되었습니다.")
                st.dataframe(error_df, use_container_width=True)
            if st.button("확인 완료"):
                st.rerun()

        # 최초 점검 완료 시 팝업을 띄우기 위한 세션 상태 관리
        if "checked_file" not in st.session_state or st.session_state.checked_file != uploaded_file.name:
            st.session_state.checked_file = uploaded_file.name
            show_popup_results(len(error_df))

        # 5. 상세 결과 테이블 제공
        st.subheader("🔍 전체 행별 상세 검증 내역")
        
        # 결과에 따라 행 색상 하이라이트 함수 정의
        def highlight_errors(row):
            return ['background-color: #ffcccc' if row['결과'] == "❌ 불일치/에러" else 'background-color: #e6f4ea' for _ in row]
            
        st.dataframe(
            verified_df.style.apply(highlight_errors, axis=1),
            use_container_width=True,
            height=400
        )
        
        # 6. 결과 내역 다운로드 (엑셀 파일로 재내보내기)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            verified_df.to_excel(writer, index=False, sheet_name="검증결과_리포트")
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 검증 리포트 엑셀 다운로드",
            data=processed_data,
            file_name="자재청구_점검리포트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"파일을 읽는 과정에서 에러가 발생했습니다: {e}")