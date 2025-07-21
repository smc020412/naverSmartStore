# excel_merge_app.py

import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 0) 상품목록 파일 업로드 (선택) 및 배송비 매핑
shipping_map = {}
item_file = st.sidebar.file_uploader(
    "아산율림_상품목록 파일 업로드 (선택)",
    type=["xlsx"],
    accept_multiple_files=False,
    key="item_file"
)
if item_file:
    try:
        item_df = pd.read_excel(item_file, engine='openpyxl')
        shipping_map = dict(zip(item_df['상품명'], item_df['배송비']))
    except Exception as e:
        st.sidebar.error(f"상품목록 파일 처리 중 오류: {e}")
else:
    st.sidebar.info("상품목록을 업로드하지 않으면 택배비는 기본 0으로 설정됩니다.")

# 1) 네이버스토어 엑셀 파일 업로드 (여러 개 가능)
uploaded_files = st.file_uploader(
    "네이버스토어 엑셀 파일 업로드 (여러 개 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="data_files"
)
if not uploaded_files:
    st.info("하나 이상의 네이버스토어 엑셀 파일을 업로드해주세요.")
    st.stop()

# 2) 컬럼 매핑 설정
column_mapping = {
    '주문번호': '주문번호',
    '정산완료일': '일자',
    '상품명': '판매품목',
    '옵션정보': '옵션명',
    '수량': '판매수량',
    '정산기준금액(A)': '판매금액',
    '판매수수료': '판매수수료',  # 원본에 해당 컬럼이 있을 경우
    '배송속성': '원본택배비',  # 입력 파일 택배비는 원본택배비로 남기고, 이후 덮어쓰기
    '주문상태': '배송상태',
    '정산상태': '정산현황',
    '클레임상태': '기타'
}

# 3) 여러 파일 읽어와서 필요한 컬럼으로 재구성
dfs = []
needed = [
    '주문번호', '일자', '판매품목', '옵션명',
    '판매수량', '판매금액', '판매수수료', '원본택배비',
    '배송상태', '정산현황', '기타'
]
for uploaded_file in uploaded_files:
    df = pd.read_excel(uploaded_file, engine='openpyxl', header=0)
    df.columns = [column_mapping.get(col, col) for col in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.reindex(columns=needed)
    dfs.append(df)

# 4) 데이터 병합 및 집계
combined = pd.concat(dfs, ignore_index=True)
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자': 'first',
    '판매품목': 'first',
    '옵션명': 'first',
    '판매수량': 'sum',
    '판매금액': 'sum',
    '판매수수료': 'sum',
    '원본택배비': 'sum',
    '배송상태': 'first',
    '정산현황': 'first',
    '기타': 'first'
})

# 5) 날짜 필터 옵션
if st.sidebar.checkbox("날짜 필터 사용", value=False):
    mn = merged['일자'].min().date()
    mx = merged['일자'].max().date()
    dr = st.sidebar.date_input(
        "날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx
    )
    if isinstance(dr, tuple) and len(dr) == 2:
        start, end = dr
        merged = merged[((merged['일자'].dt.date >= start) &
                         (merged['일자'].dt.date <= end)) |
                        merged['일자'].isna()]

# 6) 정상/문제 데이터 분류
df_ok = merged[(merged['판매수량'] > 0) & (merged['판매금액'] > 0) & merged['일자'].notna()]
df_err = merged[(merged['판매수수량'] == 0) | (merged['판매금액'] == 0) | merged['일자'].isna()]

# 7) 배송비 매핑 및 적용 (상품목록 기준, 수량 곱하기)
for df in (df_ok, df_err):
    df['택배비'] = df['판매품목'].map(shipping_map).fillna(0) * df['판매수량']

# 8) 결과 미리보기
cols_to_show = ['주문번호', '일자', '판매품목', '옵션명', '판매수량',
                '판매금액', '판매수수료', '택배비', '배송상태', '정산현황', '기타']
st.subheader("판매수량 및 판매금액 정상 데이터")
st.data_editor(df_ok[cols_to_show], num_rows="dynamic", key="ok_table")
st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
st.data_editor(df_err[cols_to_show], num_rows="dynamic", key="err_table")

# 9) 엑셀 다운로드 (정상/문제 시트 + 요약)
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def write_with_summary(df, sheet_name):
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        total_qty = df['판매수량'].sum()
        total_amount = df['판매금액'].sum()
        total_fee = df['판매수수료'].sum()
        total_delivery = -df['택배비'].sum()
        total_deposit = total_fee + total_delivery
        summary_row = ws.max_row + 2
        headers = list(df.columns)
        # 총판매량
        col_idx = headers.index('판매수량') + 1
        ws.cell(row=summary_row, column=col_idx, value='총판매량')
        ws.cell(row=summary_row, column=col_idx+1, value=total_qty)
        # 총금액
        col_idx = headers.index('판매금액') + 1
        ws.cell(row=summary_row+1, column=col_idx, value='총금액')
        ws.cell(row=summary_row+1, column=col_idx+1, value=total_amount)
        # 총수수료
        col_idx = headers.index('판매수수료') + 1
        ws.cell(row=summary_row+2, column=col_idx, value='총수수료')
        ws.cell(row=summary_row+2, column=col_idx+1, value=total_fee)
        # 총택배비
        col_idx = headers.index('택배비') + 1
        ws.cell(row=summary_row+3, column=col_idx, value='총택배비')
        ws.cell(row=summary_row+3, column=col_idx+1, value=total_delivery)
        # 총지출
        ws.cell(row=summary_row+4, column=col_idx, value='총지출')
        ws.cell(row=summary_row+4, column=col_idx+1, value=total_deposit)
        # 총이익
        col_idx = headers.index('판매금액') + 1
        ws.cell(row=summary_row+5, column=col_idx, value='총이익')
        ws.cell(row=summary_row+5, column=col_idx+1, value=total_amount + total_fee + total_delivery)

    write_with_summary(df_ok, '정상')
    write_with_summary(df_err, '문제')
buf.seek(0)

st.download_button(
    "결산 엑셀 다운로드", buf,
    file_name="네이버스토어_결산_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_button"
)
