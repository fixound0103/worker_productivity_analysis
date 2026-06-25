import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

# 1. 웹페이지 레이아웃 및 타이틀 설정
st.set_page_config(page_title="입고 입력 생산성 분석기", layout="centered")

st.title("🏭 입고 입력 생산성 분석 프로그램")
st.write("엑셀 파일을 업로드하고 기준 초를 지정하여 5개 시트로 구성된 종합 생산성 분석 리포트를 만들어 보세요.")

# 2. 파일 다중 업로드 및 기준 초 입력 섹션
uploaded_files = st.file_uploader(
    "입고내역 엑셀 파일(xlsx)을 업로드해주세요. (여러 파일 동시 업로드 가능)",
    type=["xlsx"],
    accept_multiple_files=True
)

# 사용자가 기준 초를 직접 입력할 수 있는 숫자 입력창 (기본값은 60초)
target_seconds = st.number_input(
    "🔄 작업시간 기준(초)를 입력해 주세요.",
    min_value=1,
    max_value=3600,
    value=60,
    step=1,
    help="해당 기준 초를 초과하는 작업은 (쉬는 시간 or 특이)로 간주하여 생산성 계산식 내 (작업시간)에서 제외합니다."
)

if uploaded_files:
    st.success(f"총 {len(uploaded_files)}개의 파일이 정상적으로 로드되었습니다!")

    # 3. 분석 시작 버튼
    if st.button("입고 입력 생산성 분석 시작"):
        try:
            # 로딩 애니메이션 구현
            with st.spinner(f'데이터 정제 및 {target_seconds}초 기준 5개 시트 분석 지표를 계산 중입니다...'):

                # 메모리에 최종 저장할 통합 엑셀 작성기 객체 생성
                output = BytesIO()

                # 파일 이름 순서대로 정렬 (날짜 정렬 보장)
                sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                file_names_summary = []

                # 요청하셨던 촘촘한 구간(Bins)과 라벨 정의
                bins = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 90, 120, 150, 180, 360, 540, 720, np.inf]
                labels = [
                    '0~5초', '5~10초', '10~15초', '15~20초', '20~25초', '25~30초',
                    '30~35초', '35~40초', '40~45초', '45~50초', '50~55초', '55~60초',
                    '60~90초', '90~120초', '120~150초', '150~180초', '180~360초', '360~540초', '540~720초', '720초~'
                ]

                with pd.ExcelWriter(output, engine='openpyxl') as writer:

                    for uploaded_file in sorted_files:
                        # 원본 파일명에서 확장자 제거 (예: '26-06-23')
                        base_file_name = os.path.splitext(uploaded_file.name)[0]
                        file_names_summary.append(base_file_name)

                        # =====================================================================
                        # [단계 1] 데이터 로드 및 시트 3, 시트 4 정의 (정렬 및 유니크)
                        # =====================================================================
                        df = pd.read_excel(uploaded_file)
                        df['최초입고처리시간'] = pd.to_datetime(df['최초입고처리시간'])

                        # 1-1. [시트 3 데이터] 작업자별 ➔ 최초입고처리시간 오름차순 정렬 (중복 제거 없음)
                        sheet3_df = df.sort_values(by=['최초입고처리자', '최초입고처리시간'], ascending=[True, True]).reset_index(
                            drop=True)

                        # 1-2. [시트 4 데이터] 작업자별 ➔ 최초입고처리시간 오름차순 정렬 후 사입바코드 유니크 추출
                        sheet4_df = sheet3_df.sort_values(by=['최초입고처리자', '최초입고처리시간'], ascending=[True, True])
                        sheet4_df = sheet4_df.drop_duplicates(subset=['최초입고처리자', '사입바코드'],
                                                              keep='first').copy().reset_index(drop=True)

                        # =====================================================================
                        # [단계 2] 유니크 데이터(시트4)를 기준으로 통계 및 상세작업시간 가공
                        # =====================================================================
                        processors = sheet4_df['최초입고처리자'].unique()

                        columns_to_combine = []  # 시트 5용 (상세작업시간)
                        stat_records = []  # 시트 1용 (생산성분석 요약)
                        detailed_records = []  # 시트 2용 (생산성_상세 구간 통계)

                        for processor in processors:
                            p_df = sheet4_df[sheet4_df['최초입고처리자'] == processor].copy().sort_values('최초입고처리시간',
                                                                                                   ascending=True).reset_index(
                                drop=True)

                            # 직전 행(전)과 현재 행(후) 데이터 매칭
                            p_df['사입바코드_전'] = p_df['사입바코드'].shift(1)
                            p_df['최초입고처리시간_전'] = p_df['최초입고처리시간'].shift(1)

                            p_df = p_df.rename(columns={
                                '사입바코드': '사입바코드_후',
                                '최초입고처리시간': '최초입고처리시간_후'
                            })

                            # 유니크 바코드 간의 작업간격(초) 계산
                            p_df['작업간격_초'] = (p_df['최초입고처리시간_후'] - p_df['최초입고처리시간_전']).dt.total_seconds()
                            p_df['작업간격_초'] = p_df['작업간격_초'].fillna(0).astype(int)

                            # 날짜 포맷 문자열 변환
                            p_df['최초입고처리시간_전_str'] = p_df['최초입고처리시간_전'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('-')
                            p_df['최초입고처리시간_후_str'] = p_df['최초입고처리시간_후'].dt.strftime('%Y-%m-%d %H:%M:%S')
                            p_df['사입바코드_전'] = p_df['사입바코드_전'].fillna('-')

                            # --- [시트 5용 가공] 작업자별 5개 열 구성 ---
                            col_prev_barcode = f"{processor}_사입바코드_전"
                            col_next_barcode = f"{processor}_사입바코드_후"
                            col_prev_time = f"{processor}_입고처리시각_전"
                            col_next_time = f"{processor}_입고처리시간_후"
                            col_diff_sec = f"{processor}_작업간격_초"

                            p_res = p_df[['사입바코드_전', '사입바코드_후', '최초입고처리시간_전_str', '최초입고처리시간_후_str', '작업간격_초']].rename(
                                columns={
                                    '사입바코드_전': col_prev_barcode,
                                    '사입바코드_후': col_next_barcode,
                                    '최초입고처리시간_전_str': col_prev_time,
                                    '최초입고처리시간_후_str': col_next_time,
                                    '작업간격_초': col_diff_sec
                                }).reset_index(drop=True)
                            columns_to_combine.append(p_res)

                            # --- [시트 1용 가공] 생산성분석 요약 지표 계산 ---
                            df_under_target = p_df[(p_df['작업간격_초'] >= 0) & (p_df['작업간격_초'] <= target_seconds)]
                            count_under_target = df_under_target.shape[0]
                            sum_time_under_target = df_under_target['작업간격_초'].sum()

                            df_over_target = p_df[p_df['작업간격_초'] > target_seconds]
                            count_over_target = df_over_target.shape[0]
                            sum_time_over_target = df_over_target['작업간격_초'].sum()

                            # 총 작업수 (+1 보정)
                            job_count = count_under_target + count_over_target + 1

                            if sum_time_under_target > 0:
                                productivity_sec = job_count / sum_time_under_target
                                productivity_hour = productivity_sec * 3600
                            else:
                                productivity_sec = 0
                                productivity_hour = 0

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

                            # --- [시트 2용 가공] 생산성_상세 구간 빈도수 집계 ---
                            p_df['구간'] = pd.cut(p_df['작업간격_초'], bins=bins, labels=labels, include_lowest=True)
                            counts = p_df['구간'].value_counts().reindex(labels, fill_value=0)

                            detailed_record = {'작업자명': processor}
                            for label in labels:
                                detailed_record[label] = counts[label]

                            detailed_record['총수량'] = len(p_df)
                            detailed_records.append(detailed_record)

                        sheet1_df = pd.DataFrame(stat_records)
                        sheet2_df = pd.DataFrame(detailed_records)
                        sheet5_df = pd.concat(columns_to_combine, axis=1) if columns_to_combine else pd.DataFrame()

                        # =====================================================================
                        # [단계 3] 엑셀 파일 내 시트 이름 동적 생성 후 내보내기 (최대 31자 제한 고려)
                        # =====================================================================
                        # 여러 파일 업로드 시 파일명을 구별할 수 있게 접두사를 붙여줍니다.
                        # 파일이 딱 1개일 경우 깔끔하게 접두사 없이 기본 이름으로 시트명을 생성합니다.
                        prefix = f"{base_file_name}_" if len(sorted_files) > 1 else ""

                        sheet1_df.to_excel(writer, sheet_name=f'{prefix}생산성분석', index=False)
                        sheet2_df.to_excel(writer, sheet_name=f'{prefix}생산성_상세', index=False)
                        sheet3_df.to_excel(writer, sheet_name=f'{prefix}작업자별정렬', index=False)
                        sheet4_df.to_excel(writer, sheet_name=f'{prefix}작업자별사입유니크정렬', index=False)
                        sheet5_df.to_excel(writer, sheet_name=f'{prefix}상세작업시간', index=False)

                processed_data = output.getvalue()

            # 성공 효과 및 알림
            st.balloons()
            st.success("분석 성공")

            # 다운로드 파일 이름 빌드 규칙
            if len(file_names_summary) > 1:
                display_name = f"{file_names_summary[0]}_외_{len(file_names_summary) - 1}개파일_"
            else:
                display_name = f"{file_names_summary[0]}"

            final_download_name = f"{display_name}작업자별 생산성분석_{target_seconds}초기준 결과.xlsx"

            # 다운로드 버튼 제공
            st.download_button(
                label="📥 가공된 종합 분석 엑셀 파일 다운로드",
                data=processed_data,
                file_name=final_download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"⚠️처리 중 에러가 발생했습니다: {e}")
            st.info("엑셀 파일의 열 이름(최초입고처리자, 사입바코드, 최초입고처리시간 등)을 다시 확인해 주세요.")
