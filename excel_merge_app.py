# excel_merge_app.py

import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 1) 아산율림_제품목록 파일 업로드 (선택) 및 배송비 매핑
shipping_map = {}
item_file = st.sidebar.file_uploader(
    "아산율림_제품목록 파일 업로드 (선택)",
    type=["xlsx"],
    accept_multiple_files=False,
    key="item_list"
)
if item_file:
    try:
        item_df = pd.read_excel(item_file, engine="openpyxl")
        shipping_map = dict(zip(item_df['상품명'], item_df['배송비']))
    except Exception as e:
        st.sidebar.error(f"상품목록 파일 처리 중 오류: {e}")
else:
    st.sidebar.info("상품목록 파일이 없으면 배송비는 0으로 처리됩니다.")

# 2) 네이버스토어 엑셀 파일 업로드 (여러 개 가능)
uploaded_files = st.file_uploader(
    "네이버스토어 엑셀 파일 업로드 (여러 개 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="data_files"
)
if not uploaded_files:
    st.info("하나 이상의 네이버스토어 엑셀 파일을 업로드해주세요.")
    st.stop()

# 3) 컬럼 매핑 및 수수료 컬럼 정의
column_mapping = {
    '주문번호': '주문번호',
    '정산완료일': '일자',
    '상품명': '판매품목',
    '옵션정보': '옵션명',
    '수량': '판매수량',
    '정산기준금액(A)': '판매금액',
    '주문상태': '배송상태',
    '정산상태': '정산현황',
    '클레임상태': '기타'
}
fee_columns = [
    '매출연동 수수료 합계(C)',
    '네이버페이 주문관리 수수료(B)',
    '무이자할부 수수료(D)'
]
needed_cols = [
    '주문번호','일자','판매품목','옵션명',
    '판매수량','판매금액','판매수수료',
    '배송상태','정산현황','기타'
]

# 4) 원본 데이터 로드 및 전처리
dfs = []
for f in uploaded_files:
    df = pd.read_excel(f, engine="openpyxl")
    df.rename(columns=column_mapping, inplace=True)
    # 수수료 합산
    existing = [c for c in fee_columns if c in df.columns]
    if existing:
        df[existing] = df[existing].apply(pd.to_numeric, errors='coerce')
        df['판매수수료'] = df[existing].sum(axis=1)
        df.drop(columns=existing, inplace=True)
    else:
        df['판매수수료'] = 0
    # 필요한 컬럼 채우기
    for col in needed_cols:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','판매수수료'] else ''
    df = df[needed_cols]
    dfs.append(df)

# 5) 데이터 결합 및 타입 변환
combined = pd.concat(dfs, ignore_index=True)
combined['일자'] = pd.to_datetime(combined['일자'], errors='coerce')
for c in ['판매수량','판매금액','판매수수료']:
    combined[c] = pd.to_numeric(combined[c], errors='coerce').fillna(0).astype(int)

# 6) 날짜 필터
st.sidebar.header("날짜 범위 필터")
valid_dates = combined['일자'].dt.date.dropna()
if not valid_dates.empty:
    mn, mx = valid_dates.min(), valid_dates.max()
    dr = st.sidebar.date_input("날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx)
    if isinstance(dr, tuple) and len(dr) == 2:
        start, end = dr
        combined = combined[((combined['일자'].dt.date >= start) & (combined['일자'].dt.date <= end)) |
                              combined['일자'].isna()]

# 7) 배송비 계산: 제품명 매칭 후 수량 곱하기 (없으면 0), 정수형, 음수표시
combined['택배비'] = combined['판매품목'].map(shipping_map).fillna(0) * combined['판매수량']
combined['택배비'] = -combined['택배비'].astype(int)

# 8) 주문 단위로 집계 및 순수익 계산 (merged)
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자': 'first',
    '판매품목': 'first',
    '옵션명': 'first',
    '판매수량': 'sum',
    '판매금액': 'sum',
    '판매수수료': 'sum',
    '택배비': 'sum',
    '배송상태': lambda x: ', '.join(x.dropna().unique()),
    '정산현황': lambda x: ', '.join(x.dropna().unique()),
    '기타': lambda x: ', '.join(x.dropna().unique())
})
# 순수익 계산: 판매금액 - 판매수수료 - (음수로 표시된 택배비)
merged['순수익'] = merged['판매금액'] - merged['판매수수료'] + merged['택배비']

# 9) 정상/문제 분류 및 미리보기
mask = (merged['판매수량'] > 0) & (merged['판매금액'] > 0) & merged['일자'].notna()
df_ok = merged[mask]
df_err = merged[~mask]
cols = ['주문번호','일자','판매품목','옵션명','판매수량','판매금액','판매수수료','택배비','순수익','배송상태','정산현황','기타']
st.subheader("판매수량 및 판매금액 정상 데이터")
st.data_editor(df_ok[cols], num_rows="dynamic", key="ok_table")
st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
st.data_editor(df_err[cols], num_rows="dynamic", key="err_table")

# 10) 엑셀 다운로드 (2개 시트 + 요약행 추가)
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def write_with_summary(df, sheet_name):
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        total_qty = df['판매수량'].sum()
        total_amount = df['판매금액'].sum()
        total_fee = df['판매수수료'].sum()
        total_delivery = df['택배비'].sum()  # 음수값 합산
        total_deposit = total_fee + total_delivery
        summary_row = ws.max_row + 2
        idx_amt = list(df.columns).index('판매금액') + 1
        # 총판매량
        ws.cell(row=summary_row, column=idx_amt, value='총판매량')
        ws.cell(row=summary_row, column=idx_amt+1, value=total_qty)
        # 총금액
        ws.cell(row=summary_row+1, column=idx_amt, value='총금액')
        ws.cell(row=summary_row+1, column=idx_amt+1, value=total_amount)
        # 총수수료
        ws.cell(row=summary_row+2, column=idx_amt+2, value='총수수료')
        ws.cell(row=summary_row+2, column=idx_amt+3, value=total_fee)
        # 총택배비
        ws.cell(row=summary_row+3, column=idx_amt+2, value='총택배비')
        ws.cell(row=summary_row+3, column=idx_amt+3, value=total_delivery)
        # 총지출
        ws.cell(row=summary_row+4, column=idx_amt+2, value='총지출')
        ws.cell(row=summary_row+4, column=idx_amt+3, value=total_deposit)
        # 총이익
        ws.cell(row=summary_row+5, column=idx_amt, value='총이익')
        ws.cell(row=summary_row+5, column=idx_amt+1, value=total_amount + total_deposit)
    write_with_summary(df_ok, '정상')
    write_with_summary(df_err, '문제')
buf.seek(0)
st.download_button(
    "결산 엑셀 다운로드", buf,
    file_name="네이버스토어_결산_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
