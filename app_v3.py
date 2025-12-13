# streamlit_app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("FAN/FLANGE 배분 계산기 (테스트버전)")

# 1️⃣ 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])
if uploaded_file is not None:
    # 데이터 읽기 (6행 고정)
    df = pd.read_excel(uploaded_file, header=None, skiprows=11, nrows=6)

    # FAN 또는 FLANGE 행 필터링
    df_filtered = df[df[0].astype(str).str.contains('FAN|FLANGE', case=True, na=False)].copy()
    st.subheader("원본 데이터 (FAN/FLANGE 필터링)")
    st.dataframe(df_filtered)

    # G~AH 열 복사 + 숫자 변환
    numbers = df_filtered.iloc[:, 6:34].copy()
    numbers = numbers.apply(pd.to_numeric, errors='coerce').fillna(0)

    # 배분 결과 저장용
    result = pd.DataFrame(0, index=numbers.index, columns=numbers.columns)

    # 2️⃣ 배분 로직 (원본 그대로 유지)
    for row_idx in numbers.index:
        unit = df_filtered.loc[row_idx, 2]  # C열 단위
        if pd.isna(unit) or unit == 0:
            unit = 1
        for col_idx, col in enumerate(numbers.columns):
            value = numbers.loc[row_idx, col]
            if pd.isna(value) or value == 0:
                continue
            # 왼쪽 3칸까지 (총 4칸)
            target_cols = [col_idx - i for i in range(4) if col_idx - i >= 0]
            per_col = value / len(target_cols)
            per_col = int(per_col / unit) * unit
            for tc_idx in target_cols:
                tc = numbers.columns[tc_idx]
                available = 3300 - result[tc].sum()
                add_value = min(per_col, available)
                add_value = int(add_value / unit) * unit
                result.loc[row_idx, tc] += add_value

    # 3️⃣ 결과 표 (원본 구조 유지)
    result_display = df_filtered.copy()
    result_display.iloc[:, 6:34] = result
    st.subheader("배분 결과")
    st.dataframe(result_display)

    # 4️⃣ 열 합계 확인
    col_sums = result.sum(axis=0)
    st.subheader("열 합계")
    st.dataframe(col_sums)

    over_3300 = col_sums[col_sums > 3300]
    if not over_3300.empty:
        st.warning(f"3300 초과 열: {over_3300.index.tolist()}")
    else:
        st.success("모든 열 합계가 3300 이하입니다.")

    # 5️⃣ 시각화
    st.subheader("📊 제품별 합계")
    st.bar_chart(result.sum(axis=1))

    st.subheader("📉 일별 생산량 추이")
    st.line_chart(result.sum(axis=0))

    st.subheader("🎯 CAPA 활용률 (%)")
    capa_percent = result.sum(axis=0) / 3300 * 100
    st.bar_chart(capa_percent)

# 6️⃣ 엑셀 다운로드
    output = BytesIO()
    
    # with 구문 시작 (들여쓰기 주의)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # writer 객체를 사용하는 코드는 반드시 이 블록 안에 들여쓰기 되어야 합니다.
        df_filtered.to_excel(writer, index=False, sheet_name='배분 결과') 
        result_display.to_excel(writer, index=False, sheet_name='결과 요약') 
    
    # with 블록이 끝나면 자동으로 저장(save) 및 닫기(close)가 됩니다.
    
    # st.download_button에 전달할 데이터 준비
    excel_data = output.getvalue()
    
    st.download_button(
        label="수정된 파일 다운받기",  # 버튼에 표시될 텍스트
        data=excel_data,               # 다운로드할 데이터 (바이트 형태)
        file_name="배분_결과_다운로드.xlsx",    # 다운로드될 파일 이름
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
