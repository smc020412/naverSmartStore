import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import msoffcrypto

# 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# --- 1) 배송비 및 판매가격 파일 업로드 및 매핑 ---
shipping_map = {}
shipping_name_map = {}
price_map = {}
df_fee = pd.DataFrame()

shipping_fee_file = st.sidebar.file_uploader(
    "상품현황 배송비 + 가격 파일 업로드 (선택)", type=["xlsx"], key="shipping_fee"
)
if shipping_fee_file:
    try:
        df_fee = pd.read_excel(shipping_fee_file, engine="openpyxl")
        # 상품현황 컬럼 정리 및 매핑용 옵션 처리
        df_fee['상품번호'] = df_fee['상품번호'].astype(str).str.strip()
        df_fee['상품명'] = df_fee['상품명'].astype(str).str.strip()
        df_fee['옵션매칭'] = df_fee['옵션명'].fillna('').astype(str).apply(
            lambda x: x.split(':')[-1].strip() if ':' in x else x.strip()
        )
        # 배송비 매핑
        shipping_map = df_fee.set_index(['상품번호','옵션매칭'])['배송비'].to_dict()
        shipping_name_map = df_fee.set_index(['상품명','옵션매칭'])['배송비'].to_dict()
        # 가격 매핑
        price_map = df_fee.set_index(['상품명','옵션매칭'])['판매가격'].to_dict()
        # 제품 리스트
        df_products = df_fee[['상품번호','상품명']].drop_duplicates()
        st.sidebar.success("배송비 및 가격 매핑 완료")
    except Exception as e:
        st.sidebar.error(f"배송비/가격 파일 오류: {e}")
else:
    st.sidebar.info("배송비 및 가격 파일을 업로드해 주세요.")

# --- 2) 네이버스토어 엑셀 업로드 및 암호 입력 ---
upload_files = st.sidebar.file_uploader(
    "네이버스토어 엑셀 업로드 (다중)", type=["xlsx"], accept_multiple_files=True, key="data_files"
)
password = st.sidebar.text_input("암호 비밀번호", type="password", key="file_password")

# --- 3) 파일 로드 및 복호화 ---
file_dfs = []
for f in upload_files or []:
    try:
        df = pd.read_excel(f, engine="openpyxl")
    except Exception:
        try:
            enc = msoffcrypto.OfficeFile(f)
            enc.load_key(password=password)
            buf = BytesIO(); enc.decrypt(buf)
            df = pd.read_excel(buf, engine="openpyxl")
        except Exception as e2:
            st.sidebar.error(f"{f.name} 열기 실패: {e2}")
            continue
    file_dfs.append(df)
if not file_dfs:
    st.error("업로드된 파일이 없습니다.")
    st.stop()

# --- 4) 병합 및 컬럼명 통일 (정산수수료 병합 포함) ---
mapping = {
    '주문번호':'주문번호','상품번호':'상품번호','상품명':'판매품목',
    '옵션정보':'옵션명','수량':'판매수량','정산기준금액(A)':'판매금액',
    '네이버페이 주문관리 수수료(B)':'NP수수료','매출연동 수수료 합계(C)':'ML수수료',
    '주문상태':'배송상태','정산상태':'정산상태','클레임상태':'기타'
}
needed_cols = ['주문번호','상품번호','일자','판매품목','옵션명',
               '판매수량','판매금액','NP수수료','ML수수료','배송상태','정산상태','기타']
dfs = []
for df in file_dfs:
    df.rename(columns=mapping, inplace=True)
    df['상품번호'] = df.get('상품번호','').astype(str).str.strip()
    if '옵션정보' in df:
        df['옵션명'] = df['옵션정보'].astype(str).apply(
            lambda x: x.split(':')[-1].strip() if ':' in x else x.strip()
        )
        df.drop(columns=['옵션정보'], inplace=True)
    df['일자'] = pd.to_datetime(
        df.get('정산완료일', df.get('주문일시')), errors='coerce'
    )
    # 부족한 컬럼 채우기
    for col in needed_cols:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','NP수수료','ML수수료'] else ''
    dfs.append(df[needed_cols])
combined = pd.concat(dfs, ignore_index=True)

# 옵션매칭 & 타입 정리
combined['옵션매칭'] = combined['옵션명'].fillna('').astype(str).apply(
    lambda x: x.split(':')[-1].strip() if ':' in x else x.strip()
)
for col in ['판매수량','판매금액','NP수수료','ML수수료']:
    combined[col] = pd.to_numeric(combined[col], errors='coerce').fillna(0).astype(int)

# 판매수수료 합산
combined['판매수수료'] = combined['NP수수료'] + combined['ML수수료']
# 정산현황 설정
combined['정산현황'] = combined['정산상태']

# --- 5) 날짜 필터 ---
st.sidebar.header("날짜 범위")
mn, mx = combined['일자'].dt.date.min(), combined['일자'].dt.date.max()
start, end = st.sidebar.date_input("날짜 선택", value=(mn, mx), min_value=mn, max_value=mx)
combined = combined[(combined['일자'].dt.date.between(start, end)) | combined['일자'].isna()]

# --- 6) 제품 선택 체크박스 + 전체선택 ---
st.sidebar.header("제품 선택")
selected_codes = []
if not df_fee.empty:
    # 전체선택
    select_all = st.sidebar.checkbox("전체선택", value=True)
    df_products = df_fee[['상품번호','상품명']].drop_duplicates()
    for idx, row in df_products.iterrows():
        label = f"{row['상품번호']} - {row['상품명']}"
        checked = select_all or st.sidebar.checkbox(label, key=f"prod_{idx}")
        if checked:
            selected_codes.append(row['상품번호'])
    combined = combined[combined['상품번호'].isin(selected_codes)]
else:
    st.sidebar.info("상품현황 파일을 먼저 업로드하세요.")

# --- 7) 배송비 및 가격 보정 ---
combined['판매금액'] = combined.apply(
    lambda x: price_map.get((x['판매품목'],x['옵션매칭']),0) if x['판매금액']==0 else x['판매금액'],
    axis=1
).astype(int)
combined['택배비'] = -combined.apply(
    lambda x: shipping_map.get((x['상품번호'],x['옵션매칭']),0)*x['판매수량'],
    axis=1
).astype(int)
combined['총판매금액'] = combined['판매금액'] * combined['판매수량']

# --- 8) 집계 및 순수익 ---
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자':'first','판매품목':'first','옵션명':'first',
    '판매수량':'sum','판매금액':'sum','판매수수료':'sum','택배비':'sum',
    '배송상태':'first','정산현황':'first','기타':'first'
})
merged['순수익'] = merged['판매금액'] + merged['판매수수료'] + merged['택배비']

# --- 9) 미리보기 (순수익 순서 조정) ---
mask = (merged['판매수량']>0)&(merged['판매금액']>0)&merged['일자'].notna()
df_ok = merged[mask]
df_err = merged[~mask]
cols = ['주문번호','일자','판매품목','옵션명','판매수량','판매금액','판매수수료','택배비','순수익','배송상태','정산현황','기타']

st.subheader("정상 데이터")
st.data_editor(df_ok[cols], num_rows="dynamic", key="editor_ok")
st.subheader("진행중인 데이터")
st.data_editor(df_err[cols], num_rows="dynamic", key="editor_err")

# --- 10) 다운로드 ---
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    df_ok[cols].to_excel(writer, sheet_name='정상', index=False)
    df_err[cols].to_excel(writer, sheet_name='진행시트', index=False)
buf.seek(0)
st.download_button("결산 엑셀 다운로드", buf, file_name="결과.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
