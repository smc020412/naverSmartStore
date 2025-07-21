import streamlit as st
import pandas as pd
from io import BytesIO
import msoffcrypto  # 암호화된 엑셀 처리용

# 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 1) 상품현황 배송비 파일 업로드 및 매핑
shipping_map = {}
shipping_fee_file = st.sidebar.file_uploader(
    "상품현황 배송비 파일 업로드 (선택)", type=["xlsx"], key="shipping_fee"
)
if shipping_fee_file:
    try:
        df_fee = pd.read_excel(shipping_fee_file, engine="openpyxl")
        shipping_map = dict(zip(df_fee['상품번호'], df_fee['배송비']))
        st.sidebar.success(f"배송비 매핑 {len(shipping_map)}건 로드됨")
    except Exception as e:
        st.sidebar.error(f"배송비 파일 오류: {e}")
else:
    st.sidebar.info("배송비 파일 없으면 0 처리")

# 2) 네이버스토어 엑셀 파일 업로드 & 비밀번호
upload_files = st.sidebar.file_uploader(
    "네이버스토어 엑셀 업로드 (다중)", type=["xlsx"], accept_multiple_files=True, key="data_files"
)
password = st.sidebar.text_input("암호 비밀번호", type="password", key="file_password")

# 3) 파일 로드 및 복호화
file_dfs = []
for f in upload_files or []:
    try:
        df = pd.read_excel(f, engine="openpyxl")
    except Exception:
        try:
            enc = msoffcrypto.OfficeFile(f)
            enc.load_key(password=password)
            buf = BytesIO()
            enc.decrypt(buf)
            df = pd.read_excel(buf, engine="openpyxl")
        except Exception as e2:
            st.sidebar.error(f"{f.name} 열기 실패: {e2}")
            continue
    st.sidebar.write(f"{f.name} 로드: {df.shape}")
    file_dfs.append(df)
if not file_dfs:
    st.error("업로드된 파일 없음")
    st.stop()

# 4) 병합 및 컬럼명 통일
mapping = {
    '주문번호':'주문번호','상품번호':'상품번호','정산완료일':'일자',
    '상품명':'판매품목','옵션정보':'옵션명','수량':'판매수량',
    '정산기준금액(A)':'판매금액','네이버페이 주문관리 수수료(B)':'판매수수료',
    '주문상태':'배송상태','정산상태':'정산현황','클레임상태':'기타'
}
needed = [
    '주문번호','상품번호','일자','판매품목','옵션명',
    '판매수량','판매금액','판매수수료','배송상태','정산현황','기타'
]
dfs = []
for df in file_dfs:
    df.rename(columns=mapping, inplace=True)
    if '옵션정보' in df.columns:
        df['옵션명'] = df['옵션정보']; df.drop(columns=['옵션정보'], inplace=True)
    for col in needed:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','판매수수료'] else ''
    dfs.append(df[needed])
combined = pd.concat(dfs, ignore_index=True)
combined['일자'] = pd.to_datetime(combined['일자'], errors='coerce')
for col in ['판매수량','판매금액','판매수수료']:
    combined[col] = pd.to_numeric(combined[col], errors='coerce').fillna(0).astype(int)

# 5) 날짜 필터
st.sidebar.header("날짜 범위")
dates = combined['일자'].dt.date.dropna()
if not dates.empty:
    mn, mx = dates.min(), dates.max()
    dr = st.sidebar.date_input("날짜 선택", value=(mn, mx), min_value=mn, max_value=mx)
    if isinstance(dr, tuple):
        start, end = dr
        combined = combined[((combined['일자'].dt.date >= start) & (combined['일자'].dt.date <= end)) | combined['일자'].isna()]

# 6) 제품 선택 필터
st.sidebar.header("제품 선택")
prod_map = combined[['상품번호','판매품목']].drop_duplicates().dropna().reset_index(drop=True)
select_all = st.sidebar.checkbox("전체 선택", value=True, key="sel_all")
if select_all:
    sel_nums = prod_map['상품번호'].tolist()
else:
    sel_nums = []
    for idx, row in prod_map.iterrows():
        if st.sidebar.checkbox(label=row['판매품목'], value=False, key=f"prod_cb_{idx}"):
            sel_nums.append(row['상품번호'])
if not select_all and not sel_nums:
    combined = combined.iloc[0:0]
else:
    combined = combined[combined['상품번호'].isin(sel_nums)]

# 7) 택배비 계산
combined['택배비'] = combined['상품번호'].map(shipping_map).fillna(0) * combined['판매수량']
combined['택배비'] = -combined['택배비'].astype(int)

# 8) 집계 및 순수익
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자':'first','판매품목':'first','옵션명':lambda x: x[x!=''].iloc[0] if any(x!='') else '',
    '판매수량':'sum','판매금액':'sum','판매수수료':'sum','택배비':'sum',
    '배송상태':lambda x: next((v for v in x if v), ''),
    '정산현황':lambda x: next((v for v in x if v), ''),'기타':lambda x: ','.join(x.dropna().unique())
})
merged['순수익'] = merged['판매금액'] - merged['판매수수료'] + merged['택배비']

# 9) 미리보기
mask = (merged['판매수량']>0)&(merged['판매금액']>0)&merged['일자'].notna()
df_ok = merged[mask]; df_err = merged[~mask]
cols = ['주문번호','일자','판매품목','옵션명','판매수량','판매금액','판매수수료','택배비','순수익','배송상태','정산현황','기타']
st.subheader("정상 데이터"); st.data_editor(df_ok[cols], num_rows="dynamic", key="ok")
st.subheader("문제 데이터"); st.data_editor(df_err[cols], num_rows="dynamic", key="err")

# 10) 엑셀 다운로드 및 요약행 추가
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def save(df, name):
        ws_df = df[cols]
        ws_df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        sums = {c: ws_df[c].sum() for c in ['판매수량','판매금액','판매수수료','택배비']}
        r = ws.max_row + 2
        j = cols.index('판매금액') + 1
        # 기본 요약
        ws.cell(row=r, column=j, value='총판매량');       ws.cell(row=r,   column=j+1, value=sums['판매수량'])
        ws.cell(row=r+1, column=j, value='총금액');         ws.cell(row=r+1, column=j+1, value=sums['판매금액'])
        ws.cell(row=r+2, column=j+2, value='총수수료');    ws.cell(row=r+2, column=j+3, value=sums['판매수수료'])
        ws.cell(row=r+3, column=j+2, value='총택배비');    ws.cell(row=r+3, column=j+3, value=sums['택배비'])
        ws.cell(row=r+4, column=j+2, value='총지출');     ws.cell(row=r+4, column=j+3, value=sums['판매수수료'] + sums['택배비'])
        ws.cell(row=r+5, column=j, value='총이익');        ws.cell(row=r+5, column=j+1, value=sums['판매금액'] - sums['판매수수료'] + sums['택배비'])
        # 빠른정산 수량
        qty_fast = ws_df.loc[ws_df['정산현황']=='빠른정산','판매수량'].sum()
        ws.cell(row=r+7, column=j, value='빠른정산 수량'); ws.cell(row=r+7, column=j+1, value=qty_fast)
        # 배송 상태별 수량
        for k, status in enumerate(['배송중','배송완료','구매확정']):
            qty = ws_df.loc[ws_df['배송상태']==status,'판매수량'].sum()
            ws.cell(row=r+8+k, column=j, value=f'{status} 수량'); ws.cell(row=r+8+k, column=j+1, value=qty)
    save(df_ok, '정상')
    save(df_err, '문제')
buf.seek(0)
st.download_button("결산 엑셀 다운로드", buf, file_name="결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

