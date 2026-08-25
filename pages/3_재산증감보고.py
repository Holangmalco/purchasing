import streamlit as st
import pandas as pd
import zipfile
import io

st.set_page_config(page_title="rERP 자산내역 증감보고서 자동화", layout="wide")

st.title("⚡ rERP 재산증감보고서 자동 생성기")
st.write("자산등재일 기준으로 월별 데이터를 발라내어 재산증감보고서에 들어갈 내용을 생성합니다.")

uploaded_file = st.file_uploader("rERP 자산내역 엑셀 파일을 업로드하세요 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    with st.spinner("데이터를 분석하고 보고서를 생성하는 중입니다..."):
        try:
            # 1. 데이터 로드 및 전처리
            df = pd.read_excel(uploaded_file, header=0)
            
            if '선택' in df.iloc[0].values or '자산관리번호' in df.iloc[0].values:
                df = df.iloc[1:].copy()
            
            df = df.dropna(subset=['자산관리번호']).copy()
            df = df[~df['자산관리번호'].astype(str).str.contains('건수|합계')]
            
            df['취득금액'] = pd.to_numeric(df['취득금액'], errors='coerce').fillna(0)
            
            # 자산등재일 날짜형 변환
            df['자산등재일_DT'] = pd.to_datetime(df['자산등재일'], errors='coerce')
            
            # 에러나서 변환 안 된 날짜(공란 등) 날려버리기
            df_valid = df.dropna(subset=['자산등재일_DT']).copy()
            
            # (이전에 있던 특정 기간 제한 코드 삭제 완료! 이제 엑셀에 있는 모든 데이터를 월별로 알아서 나눕니다)
            
            # 2. 통계 계산 (등재연월 기준)
            df_valid['RegMonth'] = df_valid['자산등재일_DT'].dt.to_period('M').astype(str)
            target_assets = ['기계기구', '집기비품', '소프트웨어']
            df_filtered = df_valid[df_valid['자산구분'].isin(target_assets)].copy()
            
            count_summary = df_filtered.groupby(['RegMonth', '자산구분']).size().unstack(fill_value=0)
            sum_summary = df_filtered.groupby(['RegMonth', '자산구분'])['취득금액'].sum().unstack(fill_value=0)
            
            for col in target_assets:
                if col not in count_summary.columns: count_summary[col] = 0
                if col not in sum_summary.columns: sum_summary[col] = 0
                    
            count_summary = count_summary[target_assets]
            sum_summary = sum_summary[target_assets]
            count_summary['총수량'] = count_summary.sum(axis=1)
            sum_summary['총금액'] = sum_summary.sum(axis=1)
            count_summary.columns = [f"{col}(수량)" for col in count_summary.columns]
            sum_summary.columns = [f"{col}(금액)" for col in sum_summary.columns]
            
            final_summary = pd.concat([count_summary, sum_summary], axis=1)
            cols_order = ['기계기구(수량)', '기계기구(금액)', '집기비품(수량)', '집기비품(금액)', '소프트웨어(수량)', '소프트웨어(금액)', '총수량(수량)', '총금액(금액)']
            final_summary = final_summary[cols_order]
            final_summary.rename(columns={'총수량(수량)': '총수량', '총금액(금액)': '총취득금액'}, inplace=True)
            final_summary.reset_index(inplace=True)
            final_summary.rename(columns={'RegMonth': '등재연월'}, inplace=True)

            # --- 엑셀 저장 도우미 함수 ---
            def save_to_excel_buffer(stat_df, detail_df=None):
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    
                    num_format = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter'})
                    money_format = workbook.add_format({'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter'})
                    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                    
                    stat_sheet_name = "월별 등재통계" if detail_df is None else f"{stat_df['등재연월'].iloc[0]} 증감통계"
                    stat_df.to_excel(writer, sheet_name=stat_sheet_name, index=False)
                    worksheet = writer.sheets[stat_sheet_name]
                    
                    col_settings = [{'header': c} for c in stat_df.columns]
                    worksheet.add_table(0, 0, len(stat_df), len(stat_df.columns)-1, {
                        'columns': col_settings,
                        'style': 'Table Style Light 9'
                    })
                    worksheet.set_column('A:A', 15, center_format)
                    worksheet.set_column('B:B', 15, num_format)
                    worksheet.set_column('C:C', 18, money_format)
                    worksheet.set_column('D:D', 15, num_format)
                    worksheet.set_column('E:E', 18, money_format)
                    worksheet.set_column('F:F', 15, num_format)
                    worksheet.set_column('G:G', 18, money_format)
                    worksheet.set_column('H:H', 15, num_format)
                    worksheet.set_column('I:I', 20, money_format)

                    if detail_df is not None:
                        detail_sheet_name = f"{stat_df['등재연월'].iloc[0]} 증감내역"
                        detail_df.to_excel(writer, sheet_name=detail_sheet_name, index=False)
                        worksheet_detail = writer.sheets[detail_sheet_name]
                        col_settings_dt = [{'header': str(c)} for c in detail_df.columns]
                        worksheet_detail.add_table(0, 0, len(detail_df), len(detail_df.columns)-1, {
                            'columns': col_settings_dt,
                            'style': 'Table Style Light 8'
                        })
                        worksheet_detail.set_column('A:AZ', 15)
                        
                return buffer.getvalue()

            # 3. ZIP 파일 묶기
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # 종합 파일 만들기
                main_excel_data = save_to_excel_buffer(final_summary)
                zip_file.writestr("1. 종합_자산내역_등재일기준_전체.xlsx", main_excel_data)
                
                # 월별 개별 파일 만들기
                months = final_summary['등재연월'].tolist()
                for month in months:
                    stat_m = final_summary[final_summary['등재연월'] == month]
                    detail_m = df_filtered[df_filtered['RegMonth'] == month].drop(columns=['RegMonth', '자산등재일_DT'], errors='ignore')
                    
                    if '자산등재일' in detail_m.columns:
                        detail_m['자산등재일'] = pd.to_datetime(detail_m['자산등재일'], errors='coerce').dt.strftime('%Y-%m-%d')
                    if '등록일자' in detail_m.columns:
                        detail_m['등록일자'] = pd.to_datetime(detail_m['등록일자'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
                    if '취득일자' in detail_m.columns:
                        detail_m['취득일자'] = pd.to_datetime(detail_m['취득일자'], errors='coerce').dt.strftime('%Y-%m-%d')
                        
                    detail_m = detail_m.astype(str).replace(['nan', 'NaT', 'None'], '')
                    
                    month_excel_data = save_to_excel_buffer(stat_m, detail_m)
                    zip_file.writestr(f"{month} 증감보고.xlsx", month_excel_data)

            st.success("✅ 보고서 생성이 깔끔하게 완료되었습니다!")

            # 다운로드 버튼
            st.download_button(
                label="📥 월별 증감보고서 ZIP 다운로드",
                data=zip_buffer.getvalue(),
                file_name="월별_재산증감보고서.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")