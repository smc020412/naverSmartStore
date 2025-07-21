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
    '배송속성': '원본택배비',
    '주문상태': '배송상태',
    '정산상태': '정산현황',
    '클레임상태': '기타'
}
# 수수료 컬럼 리스트 (원본)
fee_columns = [
    '매출연동 수수료 합계(C)',
    '네이버페이 주문관리 수수료(B)',
    '무이자할부 수수료(D)'
]

# 3) 여러 파일 읽어와서 필요한 컬럼으로 재구성
dfs = []
for file in uploaded_files:
    df = pd.read_excel(file, engine='openpyxl', header=0)
    # 컬럼명 매핑
    df.rename(columns=column_mapping, inplace=True)
    # 수수료 합산
    existing_fees = [c for c in fee_columns if c in df.columns]
    if existing_fees:
        df[existing_fees] = df[existing_fees].apply(pd.to_numeric, errors='coerce')
        df['판매수수료'] = df[existing_fees].sum(axis=1)
        df.drop(columns=existing_fees, inplace=True)
    # 없으면 판매수수료 0 기본
    if '판매수수료' not in df.columns:
        df['판매수수료'] = 0
    # 필요한 컬럼 정의
    needed = [
        '주문번호', '일자', '판매품목', '옵션명',
        '판매수량', '판매금액', '판매수수료',
        '원본택배비', '배송상태', '정산현황', '기타'
    ]
    df = df.loc[:, df.columns.intersection(needed)]
    # 누락 컬럼 0 또는 빈 문자열 채우기
    for col in needed:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','판매수수료','원본택배비'] else ''
    df = df[needed]
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
    '원본택배비': 'first',
    '배송상태': lambda x: ', '.join(x.dropna().unique()),
    '정산현황': lambda x: ', '.join(x.dropna().unique()),
    '기타': lambda x: ', '.join(x.dropna().unique())
})

# 5) 날짜 필터 옵션
if st.sidebar.checkbox("날짜 필터 사용", value=False):
    valid_dates = merged['일자'].dropna().dt.date
    if not valid_dates.empty:
        mn, mx = valid_dates.min(), valid_dates.max()
        dr = st.sidebar.date_input(
            "날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx
        )
        if isinstance(dr, tuple) and len(dr) == 2:
            start, end = dr
            merged = merged[((merged['일자'].dt.date >= start) & (merged['일자'].dt.date <= end)) |
                             merged['일자'].isna()]

# 6) 정상/문제 데이터 분류
mask = (merged['판매수량'] > 0) & (merged['판매금액'] > 0) & merged['일자'].notna()
df_ok = merged[mask]
df_err = merged[~mask]

# 7) 배송비 적용 (상품목록 기준, 수량 곱하기)
for df in (df_ok, df_err):
    df['택배비'] = df['판매품목'].map(shipping_map).fillna(0) * df['판매수량']

# 8) 결과 미리보기
cols = [
    '주문번호','일자','판매품목','옵션명','판매수량',
    '판매금액','판매수수료','택배비','배송상태','정산현황','기타'
]
st.subheader("판매수량 및 판매금액 정상 데이터")
st.data_editor(df_ok[cols], num_rows="dynamic", key="ok_table")
st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
st.data_editor(df_err[cols], num_rows="dynamic", key="err_table")

# 9) 엑셀 다운로드 (정상/문제 시트 + 요약)
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def write_summary(df, name):
        df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        total_qty = df['판매수량'].sum()
        total_amt = df['판매금액'].sum()
        total_fee = df['판매수수료'].sum()
        total_ship = -df['택배비'].sum()
        total_dep = total_fee + total_ship
        start_row = ws.max_row + 2
        hdr = list(df.columns)
        def c(r, c, v): ws.cell(row=r, column=c, value=v)
        c(start_row, hdr.index('판매수량')+1, '총판매량'); c(start_row, hdr.index('판매수량')+2, total_qty)
        c(start_row+1, hdr.index('판매금액')+1, '총금액'); c(start_row+1, hdr.index('판매금액')+2, total_amt)
        c(start_row+2, hdr.index('판매수수료')+1, '총수수료'); c(start_row+2, hdr.index('판매수수료')+2, total_fee)
        c(start_row+3, hdr.index('택배비')+1, '총택배비'); c(start_row+3, hdr.index('택배비')+2, total_ship)
        c(start_row+4, hdr.index('택배비')+1, '총지출'); c(start_row+4, hdr.index('택배비')+2, total_dep)
        c(start_row+5, hdr.index('판매금액')+1, '총이익'); c(start_row+5, hdr.index('판매금액')+2, total_amt + total_dep)
    write_summary(df_ok, '정상')
    write_summary(df_err, '문제')
buf.seek(0)
st.download_button(
    "결산 엑셀 다운로드", buf,
    file_name="네이버스토어_결산_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
