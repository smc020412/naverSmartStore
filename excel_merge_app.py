# excel_merge_app.py

import streamlit as st
import pandas as pd
from io import BytesIO

# 0) 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 1) 아산율림_제품목록 파일 업로드 (선택)
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
    st.sidebar.info("상품목록 파일이 없으면 기본 위젯 택배비가 사용됩니다.")

# 2) 네이버스토어 엑셀 파일 업로드 (여러 개)
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
    '배송속성': '원본택배비',
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
    '판매수량','판매금액','판매수수료','원본택배비',
    '배송상태','정산현황','기타'
]

# 4) 여러 파일 읽어와서 전처리
dfs = []
for f in uploaded_files:
    df = pd.read_excel(f, engine="openpyxl")
    # 컬럼명 매핑
    df.rename(columns=column_mapping, inplace=True)
    # 수수료 합산
    existing = [c for c in fee_columns if c in df.columns]
    if existing:
        df[existing] = df[existing].apply(pd.to_numeric, errors='coerce')
        df['판매수수료'] = df[existing].sum(axis=1)
        df.drop(columns=existing, inplace=True)
    else:
        df['판매수수료'] = 0
    # 필요한 컬럼만 유지 및 누락 컬럼 채우기
    for col in needed_cols:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','판매수수료','원본택배비'] else ''
    df = df[needed_cols]
    dfs.append(df)

# 5) 데이터 병합 및 집계
combined = pd.concat(dfs, ignore_index=True)
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자':'first',
    '판매품목':'first',
    '옵션명':'first',
    '판매수량':'sum',
    '판매금액':'sum',
    '판매수수료':'sum',
    '원본택배비':'first',
    '배송상태': lambda x: ', '.join(x.dropna().unique()),
    '정산현황': lambda x: ', '.join(x.dropna().unique()),
    '기타': lambda x: ', '.join(x.dropna().unique())
})

# 타입 정리
date_cols = ['일자']
for c in date_cols:
    merged[c] = pd.to_datetime(merged[c], errors='coerce')
num_cols = ['판매수량','판매금액','판매수수료','원본택배비']
for c in num_cols:
    merged[c] = pd.to_numeric(merged[c], errors='coerce').fillna(0).astype(int)

# 6) 날짜 필터
st.sidebar.header("날짜 범위 필터")
valid = merged['일자'].dropna().dt.date
if not valid.empty:
    mn, mx = valid.min(), valid.max()
    dr = st.sidebar.date_input("날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx)
    if isinstance(dr, tuple) and len(dr)==2:
        start, end = dr
        merged = merged[((merged['일자'].dt.date>=start)&(merged['일자'].dt.date<=end))|
                         merged['일자'].isna()]

# 7) 기본 택배비 설정 위젯
st.sidebar.header("택배비 설정")
delivery_fee = st.sidebar.number_input("택배비 (정수)", min_value=0, value=0)

# 8) 택배비 계산: 제품목록 맵이 있으면 해당 배송비, 없으면 기본값, 모두 수량 곱하기
merged['택배비'] = merged.apply(
    lambda r: shipping_map.get(r['판매품목'], delivery_fee) * r['판매수량'],
    axis=1
)

# 9) 정상/문제 데이터 분류 및 미리보기
mask = (merged['판매수량']>0)&(merged['판매금액']>0)&merged['일자'].notna()
df_ok = merged[mask]
df_err = merged[~mask]
cols = ['주문번호','일자','판매품목','옵션명','판매수량',
        '판매금액','판매수수료','택배비','배송상태','정산현황','기타']
st.subheader("판매수량 및 판매금액 정상 데이터")
st.data_editor(df_ok[cols], num_rows="dynamic", key="ok_table")
st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
st.data_editor(df_err[cols], num_rows="dynamic", key="err_table")

# 10) 엑셀 다운로드 및 요약 생성
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def write_summary(df, name):
        df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        tot_qty = df['판매수량'].sum()
        tot_amt = df['판매금액'].sum()
        tot_fee = df['판매수수료'].sum()
        tot_ship = -df['택배비'].sum()
        tot_dep = tot_fee + tot_ship
        row0 = ws.max_row + 2
        hdr = list(df.columns)
        def c(r,c,v): ws.cell(row=r, column=c, value=v)
        c(row0,   hdr.index('판매수량')+1, '총판매량');    c(row0,   hdr.index('판매수량')+2, tot_qty)
        c(row0+1, hdr.index('판매금액')+1, '총금액');      c(row0+1, hdr.index('판매금액')+2, tot_amt)
        c(row0+2, hdr.index('판매수수료')+1, '총수수료');   c(row0+2, hdr.index('판매수수료')+2, tot_fee)
        c(row0+3, hdr.index('택배비')+1,     '총택배비');   c(row0+3, hdr.index('택배비')+2,     tot_ship)
        c(row0+4, hdr.index('택배비')+1,     '총지출');     c(row0+4, hdr.index('택배비')+2,     tot_dep)
        c(row0+5, hdr.index('판매금액')+1,  '총이익');     c(row0+5, hdr.index('판매금액')+2,  tot_amt + tot_dep)
    write_summary(df_ok, '정상')
    write_summary(df_err, '문제')
buf.seek(0)

st.download_button(
    "결산 엑셀 다운로드", buf,
    file_name="네이버스토어_결산_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_button"
)
