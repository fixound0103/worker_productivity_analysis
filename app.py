import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

# 1. 웹페이지 레이아웃 및 타이틀 설정
st.set_page_config(page_title="입고 입력 생산성 분석기", layout="centered")

st.title("🏭 입고 입력 생산성 분석 프로그램")
st.write("엑셀 파일을 업로드하고 기준 초를 지정하여 생산성을 분석해 보세요.")

# 2. 파일 업로드 및 기준 초 입력 섹션
uploaded_file = st.file_uploader("입고내역 엑셀 파일(xlsx)을 업로드해주세요.", type=["xlsx"])

# 사용자가 기준 초를 직접 입력할 수 있는 숫자 입력창 추가 (기본값은 60초)
target_seconds = st.number_input(
    "🔄 작업시간 기준(초)를 입력해 주세요.",
    min_value=1,
    max_value=3600,
    value=60,
    step=1,
    help="해당 기준 초를 초과하는 작업은 (쉬는 시간 or 특이)로 간주하여 생산성 계산식 내 (작업시간)에서 제외합니다."
)

if uploaded_file is not None:
    st.success("파일이 정상적으로 로드되었습니다!")

    # 원본 파일명에서 확장자 제거한 순수 이름 추출 (예: '26-06-23')
    base_file_name = os.path.splitext(uploaded_file.name)[0]

    # 3. 분석 시작 버튼
    if st.button("입고 입력 생산성 분석 시작"):
        try:
            # 로딩 애니메이션 구현
            with st.spinner(f'데이터 정제 및 {target_seconds}초 기준 생산성 지표를 계산 중입니다...'):

                # =====================================================================
                # [단계 1] 데이터 로드 및 전처리 (필터링, 중복제거, 정렬)
                # =====================================================================
                df = pd.read_excel(uploaded_file)

                # 조건 필터링: 최초입고처리자와 최종입고처리자가 같은 행만 추출
                filtered_df = df[df['최초입고처리자'] == df['최종입고처리자']].copy()

                # 중복 제거 (동일인, 동일시간 로그 중 첫 행만 보존)
                unique_df = filtered_df.drop_duplicates(subset=['최초입고처리자', '최초입고처리시간'], keep='first')

                # 정렬 및 시간 타입 변환
                unique_df['최초입고처리시간'] = pd.to_datetime(unique_df['최초입고처리시간'])

                # [시트 2 데이터 정의]
                sheet2_df = unique_df.sort_values(by=['최초입고처리자', '최초입고처리시간'], ascending=[True, True]).reset_index(
                    drop=True)

                # =====================================================================
                # [단계 2] 작업자별 상세간격 및 생산성 통계 가공
                # =====================================================================
                processors = sheet2_df['최초입고처리자'].unique()

                columns_to_combine = []  # 시트 3용
                stat_records = []  # 시트 1용

                for processor in processors:
                    p_df = sheet2_df[sheet2_df['최초입고처리자'] == processor].copy()

                    # 직전 작업과의 간격(초) 계산
                    p_df['작업간격_초'] = p_df['최초입고처리시간'].diff().dt.total_seconds()

                    # --- [시트 3용 가공] ---
                    p_df_sheet3 = p_df.copy()
                    p_df_sheet3['작업간격_초'] = p_df_sheet3['작업간격_초'].fillna(0).astype(int)
                    p_df_sheet3['최초입고처리시간_str'] = p_df_sheet3['최초입고처리시간'].dt.strftime('%Y-%m-%d %H:%M:%S')

                    barcode_col_name = f"{processor}_사입바코드"
                    time_col_name = f"{processor}_입고처리시간"
                    diff_col_name = f"{processor}_작업간격_초"

                    p_res = p_df_sheet3[['사입바코드', '최초입고처리시간_str', '작업간격_초']].rename(columns={
                        '사입바코드': barcode_col_name,
                        '최초입고처리시간_str': time_col_name,
                        '작업간격_초': diff_col_name
                    }).reset_index(drop=True)
                    columns_to_combine.append(p_res)

                    # --- [시트 1용 가공] 사용자가 입력한 target_seconds 기준으로 필터링 ---
                    df_under_target = p_df[(p_df['작업간격_초'] >= 0) & (p_df['작업간격_초'] <= target_seconds)]
                    count_under_target = df_under_target.shape[0]
                    sum_time_under_target = df_under_target['작업간격_초'].sum()

                    df_over_target = p_df[p_df['작업간격_초'] > target_seconds]
                    count_over_target = df_over_target.shape[0]
                    sum_time_over_target = df_over_target['작업간격_초'].sum()

                    # 작업수 보정 수식 (+1)
                    job_count = count_under_target + count_over_target + 1

                    # 생산성 단위 교정 공식 적용
                    if sum_time_under_target > 0:
                        productivity_sec = job_count / sum_time_under_target
                        productivity_hour = productivity_sec * 3600
                    else:
                        productivity_sec = 0
                        productivity_hour = 0

                    # 사용자가 입력한 초에 맞춰 동적으로 열 이름을 정의하여 딕셔너리 생성
                    stat_records.append({
                        '작업자명': processor,
                        '작업수': job_count,
                        f'0~{target_seconds}초 작업 수': count_under_target,
                        f'{target_seconds}초이후 작업 수': count_over_target,
                        f'0~{target_seconds}초 작업시간 총합': int(sum_time_under_target),
                        f'{target_seconds}초이후 작업시간 총합': int(sum_time_over_target),
                        '생산성(초)': round(productivity_sec, 4),
                        '생산성(시간)': round(productivity_hour, 1)
                    })

                sheet1_df = pd.DataFrame(stat_records)  # 시트 1 데이터
                sheet3_df = pd.concat(columns_to_combine, axis=1)  # 시트 3 데이터

                # =====================================================================
                # [단계 3] 지정된 시트 순서대로 메모리 버퍼에 저장
                # =====================================================================
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    sheet1_df.to_excel(writer, sheet_name='생산성분석', index=False)
                    sheet2_df.to_excel(writer, sheet_name='작업자별 정렬', index=False)
                    sheet3_df.to_excel(writer, sheet_name='작업자별 상세 작업시간', index=False)

                processed_data = output.getvalue()

            # 성공 이펙트 및 안내
            st.balloons()
            st.success("분석 성공")

            # 요청하신 정확한 저장명 포맷 생성
            # 예: 26-06-23작업자별 생산성분석_60초기준 결과.xlsx
            final_download_name = f"{base_file_name}작업자별 생산성분석_{target_seconds}초기준 결과.xlsx"

            # 다운로드 버튼 제공
            st.download_button(
                label="📥 가공된 엑셀 파일 다운로드",
                data=processed_data,
                file_name=final_download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"⚠️처리 중 에러가 발생했습니다: {e}")
            st.info("엑셀 파일의 열 이름들을 다시 확인해 주세요.")
