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
    st.sidebar.info("상품목록 파일이 없으면 모든 제품을 표시하고 배송비는 0으로 처리됩니다.")

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
    '옵션명': '옵션명',  # 옵션명 컬럼도 직접 매핑
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
    # 컬럼명 매핑
    df.rename(columns=column_mapping, inplace=True)
    # 옵션정보와 기존 옵션명 통합
    if '옵션정보' in df.columns:
        df['옵션명'] = df['옵션정보']
        df.drop(columns=['옵션정보'], inplace=True)
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

# 7) 제품 필터 (상품목록 기반)
if item_file:
    products = item_df['상품명'].dropna().unique().tolist()
    st.sidebar.header("제품 선택")
    select_all_products = st.sidebar.checkbox("전체 제품 선택", value=True)
    if select_all_products:
        selected_products = products
    else:
        selected_products = [prod for prod in products if st.sidebar.checkbox(prod, value=False)]
    # 제품으로만 필터
    combined = combined[combined['판매품목'].isin(selected_products)]

# 8) 배송비 계산 및 표시 (정수, 음수) 배송비 계산 및 표시 (정수, 음수) 및 표시 (정수, 음수)
combined['택배비'] = combined['판매품목'].map(shipping_map).fillna(0) * combined['판매수량']
combined['택배비'] = -combined['택배비'].astype(int)

# 9) 주문 단위 집계 및 순수익 계산
# 9) 주문 단위로 집계 및 순수익 계산
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자': 'first',
    '판매품목': 'first',
    '옵션명': 'first',  # 첫 옵션명 유지
    '판매수량': 'sum',
    '판매금액': 'sum',
    '판매수수료': 'sum',
    '택배비': 'sum',
    '배송상태': lambda x: next((v for v in x if pd.notna(v) and v!=''), ''),
    '정산현황': lambda x: next((v for v in x if pd.notna(v) and v!=''), ''),
    '기타': lambda x: ', '.join(x.dropna().unique())
})
# 순수익 계산
merged['순수익'] = merged['판매금액'] - merged['판매수수료'] + merged['택배비']

# 10) 미리보기
mask = (merged['판매수량'] > 0) & (merged['판매금액'] > 0) & merged['일자'].notna()
df_ok = merged[mask]
df_err = merged[~mask]
# 옵션명 포함하여 컬럼 순서 지정
preview_cols = ['주문번호','일자','판매품목','옵션명','판매수량','판매금액','판매수수료','택배비','순수익','배송상태','정산현황','기타']
st.subheader("판매수량 및 판매금액 정상 데이터")
st.data_editor(df_ok[preview_cols], num_rows="dynamic", key="ok_table")
st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
st.data_editor(df_err[preview_cols], num_rows="dynamic", key="err_table")

# 11) 엑셀 다운로드 및 요약행 추가 및 요약행 추가 및 요약행 추가 및 요약행 추가
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def write_with_summary(df, sheet_name):
        df_to_write = df[preview_cols]
        df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        total_qty = df_to_write['판매수량'].sum()
        total_amount = df_to_write['판매금액'].sum()
        total_fee = df_to_write['판매수수료'].sum()
        total_delivery = df_to_write['택배비'].sum()
        total_deposit = total_fee + total_delivery
        summary_row = ws.max_row + 2
        idx_amt = preview_cols.index('판매금액') + 1
        ws.cell(row=summary_row, column=idx_amt, value='총판매량')
        ws.cell(row=summary_row, column=idx_amt+1, value=total_qty)
        ws.cell(row=summary_row+1, column=idx_amt, value='총금액')
        ws.cell(row=summary_row+1, column=idx_amt+1, value=total_amount)
        ws.cell(row=summary_row+2, column=idx_amt+2, value='총수수료')
        ws.cell(row=summary_row+2, column=idx_amt+3, value=total_fee)
        ws.cell(row=summary_row+3, column=idx_amt+2, value='총택배비')
        ws.cell(row=summary_row+3, column=idx_amt+3, value=total_delivery)
        ws.cell(row=summary_row+4, column=idx_amt+2, value='총지출')
        ws.cell(row=summary_row+4, column=idx_amt+3, value=total_deposit)
        ws.cell(row=summary_row+5, column=idx_amt, value='총이익')
        ws.cell(row=summary_row+5, column=idx_amt+1, value=total_amount + total_deposit)
        statuses = ['정산완료','배송중','배송완료','구매확정']
        for i, status in enumerate(statuses):
            qty = df_to_write.loc[df_to_write['배송상태'] == status, '판매수량'].sum()
            ws.cell(row=summary_row+7+i, column=idx_amt, value=f'{status} 수량')
            ws.cell(row=summary_row+7+i, column=idx_amt+1, value=qty)
    write_with_summary(df_ok, '정상')
    write_with_summary(df_err, '문제')
buf.seek(0)
st.download_button(
    "결산 엑셀 다운로드", buf,
    file_name="네이버스토어_결산_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
