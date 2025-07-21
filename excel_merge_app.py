import streamlit as st
import pandas as pd
from io import BytesIO
import msoffcrypto  # 암호화된 엑셀 처리용

# 페이지 설정
st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 1) 아산율림_제품목록 파일 업로드 (선택) 및 배송비 매핑 (상품번호 기준)
shipping_map = {}
item_file = st.sidebar.file_uploader(
    "아산율림_제품목록 파일 업로드 (선택)",
    type=["xlsx"], accept_multiple_files=False, key="item_list"
)
if item_file:
    try:
        item_df = pd.read_excel(item_file, engine="openpyxl")
        item_df.columns = item_df.columns.str.strip()
        if '상품번호' in item_df.columns and '배송비' in item_df.columns:
            shipping_map = dict(zip(item_df['상품번호'], item_df['배송비']))
        else:
            st.sidebar.error("상품목록 파일에 '상품번호' 또는 '배송비' 컬럼이 없습니다.")
    except Exception as e:
        st.sidebar.error(f"상품목록 파일 처리 중 오류: {e}")
else:
    st.sidebar.info("상품목록 파일이 없으면 배송비를 0으로 처리합니다.")

# 2) 네이버스토어 엑셀 파일 업로드 (여러 개 가능) & 비밀번호 입력
uploaded_files = st.file_uploader(
    "네이버스토어 엑셀 파일 업로드 (여러 개 가능)",
    type=["xlsx"], accept_multiple_files=True, key="data_files"
)
password = st.sidebar.text_input(
    "암호화된 파일 비밀번호 입력 (없으면 비워두세요)",
    type="password", key="file_password"
)
if not uploaded_files:
    st.info("하나 이상의 네이버스토어 엑셀 파일을 업로드해주세요.")
    st.stop()

# 3) 컬럼 매핑 및 수수료 컬럼 정의
column_mapping = {
    '주문번호': '주문번호',
    '정산완료일': '일자',
    '상품번호': '상품번호',  # 상품번호 추가
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
    '주문번호','일자','상품번호','판매품목','옵션명',
    '판매수량','판매금액','판매수수료',
    '배송상태','정산현황','기타'
]

# 4) 데이터 로드 및 복호화
all_dfs = []
for f in uploaded_files:
    raw = f.read()
    buf = BytesIO(raw)
    try:
        df = pd.read_excel(buf, engine="openpyxl")
    except Exception:
        if password:
            buf.seek(0)
            office = msoffcrypto.OfficeFile(buf)
            office.load_key(password=password)
            dec = BytesIO()
            office.decrypt(dec)
            dec.seek(0)
            df = pd.read_excel(dec, engine="openpyxl")
        else:
            st.sidebar.error(f"{f.name} 암호화되어 있습니다. 비밀번호를 입력해주세요.")
            continue
    df.columns = df.columns.str.strip()
    df.rename(columns=column_mapping, inplace=True)
    # 옵션정보 처리
    if '옵션정보' in df.columns:
        df['옵션명'] = df['옵션정보']
        df.drop(columns=['옵션정보'], inplace=True)
    # 수수료 합산
    existing = [c for c in fee_columns if c in df.columns]
    if existing:
        df[existing] = df[existing].apply(pd.to_numeric, errors='coerce').fillna(0)
        df['판매수수료'] = df[existing].sum(axis=1)
        df.drop(columns=existing, inplace=True)
    else:
        df['판매수수료'] = 0
    # 부족한 컬럼 채우기
    for col in needed_cols:
        if col not in df.columns:
            df[col] = 0 if col in ['판매수량','판매금액','판매수수료'] else ''
    all_dfs.append(df[needed_cols])

if not all_dfs:
    st.error("유효한 데이터가 없습니다. 파일과 비밀번호를 확인해주세요.")
    st.stop()

# 5) 결합 및 타입 변환
combined = pd.concat(all_dfs, ignore_index=True)
combined['일자'] = pd.to_datetime(combined['일자'], errors='coerce')
for col in ['판매수량','판매금액','판매수수료']:
    combined[col] = pd.to_numeric(combined[col], errors='coerce').fillna(0).astype(int)

# 6) 날짜 필터
st.sidebar.header("날짜 범위 필터")
valid_dates = combined['일자'].dt.date.dropna()
if not valid_dates.empty:
    mn, mx = valid_dates.min(), valid_dates.max()
    dr = st.sidebar.date_input(
        "날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx
    )
    if isinstance(dr, tuple) and len(dr) == 2:
        combined = combined[
            (combined['일자'].dt.date >= dr[0]) & (combined['일자'].dt.date <= dr[1])
        ]

# 7) 배송비 계산 (상품번호 기준)
combined['택배비'] = combined['상품번호'].map(shipping_map).fillna(0) * combined['판매수량']
combined['택배비'] = -combined['택배비'].astype(int)

# 8) 주문 단위 집계 및 순수익 계산
merged = combined.groupby('주문번호', as_index=False).agg({
    '일자':'first',
    '상품번호':'first',
    '판매품목':'first',
    '옵션명': lambda x: next((v for v in x if v), ''),
    '판매수량':'sum',
    '판매금액':'sum',
    '판매수수료':'sum',
    '택배비':'sum',
    '배송상태': lambda x: next((v for v in x if v), ''),
    '정산현황': lambda x: next((v for v in x if v), ''),
    '기타': lambda x: ', '.join(x.dropna().unique())
})
merged['순수익'] = merged['판매금액'] - merged['판매수수료'] + merged['택배비']

# 9) 미리보기 및 다운로드
preview_cols = [
    '주문번호','일자','상품번호','판매품목','옵션명',
    '판매수량','판매금액','판매수수료','택배비','순수익',
    '배송상태','정산현황','기타'
]
st.data_editor(merged[preview_cols], num_rows="dynamic")

@st.cache_data
def to_excel(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()

st.download_button(
    "결산 엑셀 다운로드",
    data=to_excel(merged),
    file_name="결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
