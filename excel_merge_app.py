import streamlit as st
import pandas as pd
from io import BytesIO
import msoffcrypto  # 암호화된 엑셀 처리용

# 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 1) 상품현황 배송비 파일 업로드 (선택) 및 상품번호 기반 배송비 매핑
shipping_map = {}
shipping_fee_file = st.sidebar.file_uploader(
    "상품현황 배송비 파일 업로드 (선택)",
    type=["xlsx"],
    accept_multiple_files=False,
    key="shipping_fee"
)
if shipping_fee_file:
    try:
        shipping_df = pd.read_excel(shipping_fee_file, engine="openpyxl")
        shipping_map = dict(zip(shipping_df['상품번호'], shipping_df['배송비']))
        st.sidebar.success(f"배송비 매핑 정보 {len(shipping_map)}건 로드됨")
    except Exception as e:
        st.sidebar.error(f"배송비 파일 처리 중 오류: {e}")
else:
    st.sidebar.info("배송비 파일이 없으면 택배비는 0으로 처리됩니다.")

# 2) 네이버스토어 엑셀 파일 업로드 및 비밀번호 입력
uploaded_files = st.sidebar.file_uploader(
    "네이버스토어 엑셀 파일 업로드 (여러 개 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="data_files"
)
password = st.sidebar.text_input(
    "암호화된 파일 비밀번호 입력 (없으면 비워두세요)",
    type="password",
    key="file_password"
)

# 3) 파일 로드 및 암호 해제 처리
file_dfs = []
for f in uploaded_files or []:
    df = None
    try:
        df = pd.read_excel(f, engine="openpyxl")
    except Exception:
        try:
            encrypted = msoffcrypto.OfficeFile(f)
            encrypted.load_key(password=password)
            decrypted = BytesIO()
            encrypted.decrypt(decrypted)
            df = pd.read_excel(decrypted, engine="openpyxl")
        except Exception as e2:
            st.sidebar.error(f"{f.name} 파일 열기 실패: {e2}")
            continue
    if df is not None:
        st.sidebar.write(f"{f.name} 로드됨: 행 {df.shape[0]}, 열 {df.shape[1]}")
        file_dfs.append(df)

# 4) 데이터프레임 병합 및 컬럼명 통일
if not file_dfs:
    st.error("업로드된 파일이 없습니다.")
    st.stop()

column_mapping = {
    '주문번호': '주문번호',
    '정산완료일': '일자',
    '상품명': '판매품목',
    '옵션정보': '옵션명',
    '수량': '판매수량',
    '정산기준금액(A)': '판매금액',
    '네이버페이 주문관리 수수료(B)': '판매수수료',
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

dfs = []
for df in file_dfs:
    df.rename(columns=column_mapping, inplace=True)
    if '옵션정보' in df.columns:
        df['옵션명'] = df['옵션정보']
        df.drop(columns=['옵션정보'], inplace=True)

    existing = [c for c in fee_columns if c in df.columns]
    if existing:
        df[existing] = df[existing].apply(pd.to_numeric, errors='coerce')
        df['판매수수료'] = df[existing].sum(axis=1)
        df.drop(columns=existing, inplace=True)
    else:
        df['판매수수료'] = 0

    for col in needed_cols:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','판매수수료'] else ''
    df = df[needed_cols]
    dfs.append(df)

if not dfs:
    st.error("유효한 데이터가 없습니다. 업로드한 파일과 비밀번호를 확인해주세요.")
    st.stop()

combined = pd.concat(dfs, ignore_index=True)
combined['일자'] = pd.to_datetime(combined['일자'], errors='coerce')
for col in ['판매수량','판매금액','판매수수료']:
    combined[col] = pd.to_numeric(combined[col], errors='coerce').fillna(0).astype(int)

# 5) 날짜 필터 UI
st.sidebar.header("날짜 범위 필터")
valid_dates = combined['일자'].dt.date.dropna()
if not valid_dates.empty:
    mn, mx = valid_dates.min(), valid_dates.max()
    dr = st.sidebar.date_input("날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx)
    if isinstance(dr, tuple) and len(dr) == 2:
        start, end = dr
        combined = combined[((combined['일자'].dt.date >= start) & (combined['일자'].dt.date <= end))|
                              combined['일자'].isna()]

# 6) 배송비 계산 및 표시 (상품번호 기준 매핑)
combined['택배비'] = combined['상품번호'].map(shipping_map).fillna(0) * combined['판매수량']
combined['택배비'] = -combined['택배비'].fillna(0).astype(int)

# 7) 주문 단위 집계 및 순수익 계산
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자': 'first',
    '판매품목': 'first',
    '옵션명': lambda x: x[x.notna() & (x!='')].iloc[0] if not x[x.notna() & (x!='')].empty else '',
    '판매수량': 'sum',
    '판매금액': 'sum',
    '판매수수료': 'sum',
    '택배비': 'sum',
    '배송상태': lambda x: next((v for v in x if pd.notna(v) and v!=''), ''),
    '정산현황': lambda x: next((v for v in x if pd.notna(v) and v!=''), ''),
    '기타': lambda x: ', '.join(x.dropna().unique())
})
merged['순수익'] = merged['판매금액'] - merged['판매수수료'] + merged['택배비']

# 8) 미리보기 (정상/문제 데이터 분리)
mask = (merged['판매수량'] > 0) & (merged['판매금액'] > 0) & merged['일자'].notna()
df_ok = merged[mask]
df_err = merged[~mask]
preview_cols = ['주문번호','일자','판매품목','옵션명','판매수량','판매금액','판매수수료','택배비','순수익','배송상태','정산현황','기타']
st.subheader("판매수량 및 판매금액 정상 데이터")
st.data_editor(df_ok[preview_cols], num_rows="dynamic", key="ok_table")
st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
st.data_editor(df_err[preview_cols], num_rows="dynamic", key="err_table")

# 9) 엑셀 다운로드 및 요약 (정상/문제 시트 분리)
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
        qty_fast = df_to_write.loc[df_to_write['정산현황'] == '빠른정산', '판매수량'].sum()
        ws.cell(row=summary_row+7, column=idx_amt, value='빠른정산 수량')
        ws.cell(row=summary_row+7, column=idx_amt+1, value=qty_fast)
        delivery_statuses = ['배송중','배송완료','구매확정']
        for j, status in enumerate(delivery_statuses):
            qty = df_to_write.loc[df_to_write['배송상태'] == status, '판매수량'].sum()
            ws.cell(row=summary_row+8+j, column=idx_amt, value=f'{status} 수량')
            ws.cell(row=summary_row+8+j, column=idx_amt+1, value=qty)
    write_with_summary(df_ok, '정상')
    write_with_summary(df_err, '문제')
buf.seek(0)
st.download_button("결산 엑셀 다운로드", buf, file_name="결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
