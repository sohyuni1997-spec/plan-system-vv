# streamlit_app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="FAN/FLANGE 배분 계산기", layout="wide")
st.title("🏭 FAN/FLANGE 배분 계산기 (조립1 라인)")

# 사이드바 설정
st.sidebar.header("⚙️ 설정")
DAILY_CAPA = st.sidebar.number_input("일일 CAPA", min_value=1000, max_value=10000, value=4000, step=100)
target_line = st.sidebar.selectbox("대상 라인", ["조립1", "조립2", "조립3"])

# 토요일, 일요일 판별 함수
def is_weekend(date_str):
    if pd.isna(date_str):
        return True
    date_str = str(date_str)
    return '(토)' in date_str or '(일)' in date_str

# 1️⃣ 엑셀 파일 업로드
uploaded_file = st.file_uploader("📁 엑셀 파일 업로드 (0차계획.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 날짜 행 읽기 (8행)
        df_dates = pd.read_excel(uploaded_file, header=None, skiprows=7, nrows=1)
        dates = df_dates.iloc[0, 6:34]
        
        # 주말 확인
        weekend_mask = [is_weekend(d) for d in dates]
        
        # 데이터 읽기 (12행부터)
        df = pd.read_excel(uploaded_file, header=None, skiprows=11, nrows=50)
        
        # FAN/FLANGE + 조립1 필터링
        df_filtered = df[
            (df[0].astype(str).str.contains('FAN|FLANGE', case=True, na=False)) &
            (df[5].astype(str) == target_line)
        ].copy()
        
        # G~AH열 복사
        numbers = df_filtered.iloc[:, 6:34].copy()
        numbers.columns = dates.values
        
        # 원본 합계가 0보다 큰 제품만
        row_sums = numbers.sum(axis=1)
        df_filtered = df_filtered[row_sums > 0].copy()
        numbers = numbers[row_sums > 0].copy()
        
        st.success(f"✅ {target_line} 라인 제품 {len(df_filtered)}개 로드 완료")
        
        # 원본 데이터 표시
        with st.expander("📋 원본 데이터 보기"):
            original_display = df_filtered[[0, 1, 2, 3, 4, 5]].copy()
            original_display.columns = ['구분', '제품코드', 'PLT', '제품명', '생산합계', 'LINE']
            st.dataframe(original_display, use_container_width=True)
        
        # 2️⃣ 배분 로직
        result = pd.DataFrame(0.0, index=numbers.index, columns=numbers.columns)
        
        for row_idx in numbers.index:
            unit = df_filtered.loc[row_idx, 2]
            if pd.isna(unit) or unit == 0:
                unit = 1
            
            col_list = list(numbers.columns)
            for col_idx in range(len(col_list)):
                col = col_list[col_idx]
                value = numbers.loc[row_idx, col]
                
                if isinstance(value, pd.Series):
                    value = value.iloc[0]
                
                if pd.isna(value) or value == 0:
                    continue
                
                # 왼쪽 평일 4개 찾기
                target_cols = []
                for i in range(col_idx + 1):
                    check_idx = col_idx - i
                    if check_idx < 0:
                        break
                    if not weekend_mask[check_idx]:
                        target_cols.append(check_idx)
                    if len(target_cols) == 4:
                        break
                
                if len(target_cols) == 0:
                    continue
                
                # 배분
                remaining = value
                while remaining >= unit:
                    max_space = -1
                    max_space_idx = -1
                    
                    for i, tc_idx in enumerate(target_cols):
                        tc = col_list[tc_idx]
                        current_sum = result[tc].sum()
                        available = DAILY_CAPA - current_sum
                        
                        if available >= unit and available > max_space:
                            max_space = available
                            max_space_idx = i
                    
                    if max_space_idx == -1:
                        break
                    
                    tc = col_list[target_cols[max_space_idx]]
                    current_sum = result[tc].sum()
                    available = DAILY_CAPA - current_sum
                    
                    can_add = min(unit, remaining, available)
                    can_add = int(can_add / unit) * unit
                    
                    if can_add > 0:
                        result.loc[row_idx, tc] += can_add
                        remaining -= can_add
                    else:
                        break
                
                # 남은 양 처리
                if remaining > 0:
                    for i, tc_idx in enumerate(target_cols):
                        tc = col_list[tc_idx]
                        current_sum = result[tc].sum()
                        available = DAILY_CAPA - current_sum
                        
                        if available >= remaining:
                            result.loc[row_idx, tc] += remaining
                            remaining = 0
                            break
        
        # 3️⃣ 결과 표
        st.subheader("📊 배분 결과")
        result_display = result.copy()
        result_display.insert(0, '제품명', df_filtered[3].values)
        result_display.insert(0, 'PLT', df_filtered[2].values)
        result_display.insert(0, '제품코드', df_filtered[1].values)
        result_display.insert(0, '구분', df_filtered[0].values)
        
        st.dataframe(result_display, use_container_width=True)
        
        # 4️⃣ 날짜별 합계 (날짜+요일 포함)
        st.subheader("📅 날짜별 합계")
        col_sums = result.sum(axis=0)
        
        date_summary = pd.DataFrame({
            '날짜(요일)': dates.values,
            '배분량': col_sums.values,
            'CAPA': DAILY_CAPA,
            '여유': DAILY_CAPA - col_sums.values,
            '가동률(%)': (col_sums.values / DAILY_CAPA * 100).round(1),
            '상태': ['✅' if v <= DAILY_CAPA else '❌' for v in col_sums.values]
        })
        
        # 주말 표시
        date_summary['주말여부'] = ['주말' if is_weekend(d) else '평일' for d in dates]
        
        st.dataframe(date_summary, use_container_width=True)
        
        # 초과 확인
        over_capa = date_summary[date_summary['배분량'] > DAILY_CAPA]
        if not over_capa.empty:
            st.error(f"⚠️ {DAILY_CAPA} 초과 날짜: {len(over_capa)}개")
            st.dataframe(over_capa)
        else:
            st.success(f"✅ 모든 날짜가 {DAILY_CAPA} 이하입니다!")
        
        # 5️⃣ 시각화 (날짜+요일 포함)
        st.subheader("📈 시각화")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**일별 생산량 추이**")
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=dates.values,
                y=col_sums.values,
                name='배분량',
                marker_color=['lightgray' if is_weekend(d) else 'steelblue' for d in dates]
            ))
            fig1.add_hline(y=DAILY_CAPA, line_dash="dash", line_color="red", 
                          annotation_text=f"CAPA {DAILY_CAPA}")
            fig1.update_layout(
                xaxis_title="날짜(요일)",
                yaxis_title="생산량",
                height=400,
                xaxis_tickangle=-45
            )
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            st.markdown("**CAPA 가동률**")
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                x=dates.values,
                y=(col_sums.values / DAILY_CAPA * 100),
                name='가동률(%)',
                marker_color=['lightgray' if is_weekend(d) else 'lightgreen' for d in dates]
            ))
            fig2.add_hline(y=100, line_dash="dash", line_color="red", 
                          annotation_text="100%")
            fig2.update_layout(
                xaxis_title="날짜(요일)",
                yaxis_title="가동률 (%)",
                height=400,
                xaxis_tickangle=-45
            )
            st.plotly_chart(fig2, use_container_width=True)
        
        # 제품별 합계
        st.markdown("**제품별 생산량**")
        product_sums = result.sum(axis=1)
        product_names = df_filtered[3].values
        
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=product_names,
            y=product_sums.values,
            marker_color='coral'
        ))
        fig3.update_layout(
            xaxis_title="제품명",
            yaxis_title="총 생산량",
            height=400,
            xaxis_tickangle=-45
        )
        st.plotly_chart(fig3, use_container_width=True)
        
        # 6️⃣ 통계 요약
        st.subheader("📊 통계 요약")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("총 원본 합계", f"{numbers.sum().sum():.0f}")
        with col2:
            st.metric("총 배분 합계", f"{result.sum().sum():.0f}")
        with col3:
            achievement = result.sum().sum() / numbers.sum().sum() * 100
            st.metric("달성률", f"{achievement:.1f}%")
        with col4:
            unallocated = numbers.sum().sum() - result.sum().sum()
            st.metric("미배분량", f"{unallocated:.0f}")
        
        # 제품별 비교
        comparison = pd.DataFrame({
            '제품코드': df_filtered[1].values,
            '제품명': df_filtered[3].values,
            'Unit': df_filtered[2].values,
            '원본합계': numbers.sum(axis=1).values,
            '배분후합계': result.sum(axis=1).values,
            '차이': (result.sum(axis=1) - numbers.sum(axis=1)).values,
            '달성률(%)': ((result.sum(axis=1) / numbers.sum(axis=1)) * 100).round(1).values
        })
        
        st.subheader("📋 제품별 합계 비교")
        st.dataframe(comparison, use_container_width=True)
        
        # 7️⃣ 엑셀 다운로드
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_display.to_excel(writer, index=False, sheet_name='배분결과')
            date_summary.to_excel(writer, index=False, sheet_name='날짜별합계')
            comparison.to_excel(writer, index=False, sheet_name='제품별비교')
        
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 배분 결과 다운로드 (Excel)",
            data=excel_data,
            file_name=f"{target_line}_배분결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 오류 발생: {str(e)}")
        st.exception(e)

else:
    st.info("👈 엑셀 파일을 업로드해주세요")
