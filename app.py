import streamlit as st
import pandas as pd
import math
import io

# 1. 웹페이지 설정
st.set_page_config(
    page_title="자재 청구 리스트 자동 점검 시스템",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. 세션 상태(Session State) 초기화
# 관리자 모드용 자재별 포장사양 데이터베이스 예시 데이터 구축
if "admin_db" not in st.session_state:
    st.session_state.admin_db = pd.DataFrame([
        {"자재코드": "MTRL-001", "포장사양": "P0804"},
        {"자재코드": "MTRL-002", "포장사양": "P1201"},
        {"자재코드": "MTRL-003", "포장사양": "P0802"},
        {"자재코드": "MTRL-004", "포장사양": "P1205"}
    ])

# 3. 스타일 및 타이틀
st.title("📋 자재 청구 리스트 자동 점검 시스템")
st.markdown("자재 청구 엑셀 파일을 업로드하고 **[점검 시작]** 버튼을 누르면 점검 규칙에 맞춰 데이터를 분석합니다.")
st.markdown("---")

# 4. 탭 구성 (점검 화면과 관리자 모드를 분리)
tab_check, tab_admin = st.tabs(["📊 자재 청구 점검", "⚙️ 관리자 모드 (포장사양 관리)"])

# =========================================================================
# TAB 1: 자재 청구 점검 화면
# =========================================================================
with tab_check:
    st.subheader("📁 점검용 엑셀 파일 업로드")
    uploaded_file = st.file_uploader("자재 청구 엑셀 파일(.xlsx, .xls)을 선택하세요.", type=["xlsx", "xls"], key="checker_upload")

    if uploaded_file is not None:
        try:
            # 엑셀 파일 읽기 (헤더가 있는 1번째 행 기준)
            df = pd.read_excel(uploaded_file, header=0)
            total_rows = len(df)
            st.info(f"📂 파일명: **{uploaded_file.name}** | 총 **{total_rows}**개의 행이 감지되었습니다.")
            
            # 🚨 시작 버튼 추가
            start_btn = st.button("▶️ 점검 시작 (Start)", type="primary", use_container_width=True)
            
            if start_btn:
                # 검증 결과 및 오류 내역 리스트 초기화
                error_list = []
                verified_data = []

                # 행별 검증 시작
                for idx, row in df.iterrows():
                    excel_row_num = idx + 2  # 엑셀 기준 실제 행 번호 (헤더가 1행이므로 +2)
                    row_status = "정상"
                    error_reason = []
                    is_mapped_from_admin = False

                    # 최소 15개 열(O열)이 존재하는지 확인
                    if len(row) < 15:
                        error_list.append({
                            "행 번호": excel_row_num,
                            "자재코드": "N/A",
                            "오류 유형": "형식 오류",
                            "상세 내용": "엑셀 열 개수가 부족합니다. A열부터 O열까지 올바르게 채워져 있는지 확인해 주세요."
                        })
                        continue

                    # 데이터 안전하게 추출
                    material_code = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else "N/A"
                    packing_spec = str(row.iloc[7]).strip() if not pd.isna(row.iloc[7]) else ""
                    qty_mass = row.iloc[9]
                    qty_common = row.iloc[11]
                    qty_required = row.iloc[14]

                    # 💡 관리자 모드 연동: 파일에 포장사양이 비어있는 경우 관리자 DB에서 자동 매핑
                    if packing_spec == "" or packing_spec.lower() == "nan":
                        # 관리자 DB에서 자재코드 일치하는 항목 조회
                        match_row = st.session_state.admin_db[st.session_state.admin_db["자재코드"] == material_code]
                        if not match_row.empty:
                            packing_spec = str(match_row.iloc[0]["포장사양"]).strip()
                            is_mapped_from_admin = True
                        else:
                            row_status = "오류"
                            error_reason.append("엑셀 내 포장사양이 누락되었으며, 관리자 모드 DB에도 등록되어 있지 않습니다.")

                    # 빈 값(NaN) 처리 및 숫자 형식 검증
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

                    # 수량 규칙 검증 진행
                    if row_status != "오류":
                        # [규칙 1] 포장사양 파싱 (P0804 -> '08' / P1201 -> '12')
                        parsed_spec = ""
                        if len(packing_spec) >= 3:
                            parsed_spec = packing_spec[1:3]
                        else:
                            row_status = "오류"
                            error_reason.append(f"포장사양 형식('{packing_spec}')이 부적합합니다. (최소 3자리 필요)")

                        # [규칙 2] 양산청구수량(J열) == 필요수량(O열) 일치 검증
                        if qty_mass != qty_required:
                            row_status = "오류"
                            error_reason.append(f"양산청구수량 불일치 (필요: {int(qty_required)}개 / 청구: {int(qty_mass)}개)")

                        # [규칙 3] 공용청구수량(L열) 조건 검증
                        # ① 포장사양이 '08'인 경우
                        if parsed_spec == "08":
                            expected_common = 30 if qty_required <= 500 else math.ceil(qty_required * 0.05)
                            if qty_common != expected_common:
                                row_status = "오류"
                                error_reason.append(
                                    f"공용청구수량 불일치 (기준: 필요 {int(qty_required)}개 일 때 {expected_common}개 필요 / 현재 {int(qty_common)}개)"
                                )
                        # ② 포장사양이 '12'인 경우 (3개 고정 추가 청구 규칙 적용)
                        elif parsed_spec == "12":
                            expected_common = 3
                            if qty_common != expected_common:
                                row_status = "오류"
                                error_reason.append(
                                    f"공용청구수량 불일치 (포장사양 '12'는 3개 고정 추가 필요하나 현재 {int(qty_common)}개)"
                                )

                    # 오류 발생 시 오류 리스트에 상세 기록
                    if row_status == "오류":
                        for reason in error_reason:
                            error_list.append({
                                "행 번호": excel_row_num,
                                "자재코드": material_code,
                                "오류 유형": "수량/규칙 불일치" if "불일치" in reason else "데이터 누락/형식",
                                "상세 내용": reason
                            })

                    # 테이블 표시용 데이터 저장
                    note_msg = "정상 검증 완료"
                    if is_mapped_from_admin:
                        note_msg += " (관리자 DB에서 포장사양 연동됨)"
                    if error_reason:
                        note_msg = ", ".join(error_reason)

                    verified_data.append({
                        "행 번호": excel_row_num,
                        "자재코드": material_code,
                        "포장사양": packing_spec,
                        "양산청구": qty_mass,
                        "공용청구": qty_common,
                        "필요수량": qty_required,
                        "결과": "정상" if row_status == "정상" else "❌ 불일치/에러",
                        "비고": note_msg
                    })

                # 데이터프레임 변환
                error_df = pd.DataFrame(error_list)
                verified_df = pd.DataFrame(verified_data)

                # 점검 결과 대시보드 시각화
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

                # 웹용 팝업(Modal Dialog) 알림 발생
                @st.dialog("🚨 검증 완료 안내")
                def show_popup_results(err_count):
                    if err_count == 0:
                        st.balloons()
                        st.success("🎉 모든 자재 청구 항목이 설정된 규칙과 완벽하게 일치합니다!")
                    else:
                        st.error(f"⚠️ 총 {err_count}개의 불일치 및 오류 항목이 발견되었습니다.")
                        st.dataframe(error_df, use_container_width=True)
                    if st.button("확인"):
                        st.rerun()

                # 결과 팝업 트리거
                show_popup_results(len(error_df))

                # 상세 검증 리스트 화면 출력
                st.subheader("🔍 점검 상세 리스트")
                def highlight_errors(row):
                    return ['background-color: #ffcccc' if row['결과'] == "❌ 불일치/에러" else 'background-color: #e6f4ea' for _ in row]

                st.dataframe(
                    verified_df.style.apply(highlight_errors, axis=1),
                    use_container_width=True,
                    height=450
                )

                # 점검 리포트 엑셀 다운로드 생성
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    verified_df.to_excel(writer, index=False, sheet_name="검증결과_리포트")
                processed_data = output.getvalue()

                st.download_button(
                    label="📥 점검 결과 리포트 다운로드 (Excel)",
                    data=processed_data,
                    file_name="자재청구_점검결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"엑셀 파일을 점검하는 중 에러가 발생했습니다: {e}")

# =========================================================================
# TAB 2: 관리자 모드 (자재별 포장사양 데이터베이스 관리)
# =========================================================================
with tab_admin:
    st.subheader("⚙️ 자재별 포장사양 등록 및 수정 (데이터베이스)")
    st.markdown("""
    여기서 등록하고 편집한 데이터는 업로드된 자재 청구 리스트에 포장사양($H$열)이 공백이거나 누락되었을 때, 
    **자재코드**를 바탕으로 매핑하여 분석하는 데 사용됩니다. 
    """)

    # Streamlit의 데이터 에디터(Data Editor)를 통해 엑셀처럼 자유로운 수정 지원
    st.info("💡 아래 테이블을 더블클릭하여 자재코드 및 포장사양을 직접 수정하거나 행을 추가/삭제할 수 있습니다.")
    
    # 세션 상태 데이터프레임을 직접 편집하도록 설정
    edited_df = st.data_editor(
        st.session_state.admin_db, 
        num_rows="dynamic",  # 사용자가 직접 행 추가/삭제 가능하도록 지원
        use_container_width=True,
        key="admin_editor"
    )

    # 저장 버튼 제공
    if st.button("💾 데이터베이스 변경 내용 저장", type="primary"):
        st.session_state.admin_db = edited_df
        st.success("🎉 관리자용 포장사양 데이터베이스가 성공적으로 업데이트되었습니다!")
        st.balloons()

    # 데이터 백업용 다운로드/업로드 기능 추가
    st.write("---")
    st.markdown("### 📤 관리자 데이터 백업 및 가져오기")
    col_dl, col_ul = st.columns(2)

    with col_dl:
        # 현재 테이블을 CSV로 백업 다운로드
        csv_data = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 현재 관리자 데이터 백업 받기 (CSV)",
            data=csv_data,
            file_name="admin_packing_spec_backup.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col_ul:
        # 파일 업로드를 통해 대량 데이터 로드
        backup_file = st.file_uploader("백업된 CSV 파일을 업로드하여 덮어쓰기 합니다.", type=["csv"], key="backup_upload")
        if backup_file is not None:
            try:
                uploaded_admin = pd.read_csv(backup_file)
                # 컬럼 유효성 검증
                if "자재코드" in uploaded_admin.columns and "포장사양" in uploaded_admin.columns:
                    st.session_state.admin_db = uploaded_admin[["자재코드", "포장사양"]]
                    st.success("📥 대량 데이터 업로드 완료! 변경 내용을 확인하고 위 저장 버튼을 눌러주세요.")
                else:
                    st.error("파일 형식이 맞지 않습니다. '자재코드'와 '포장사양' 열이 있어야 합니다.")
            except Exception as e:
                st.error(f"백업 파일을 불러오는 과정에서 오류가 발생했습니다: {e}")
