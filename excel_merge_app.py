```python
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
        # 상품번호 → 배송비 매핑
        shipping_map = dict(zip(shipping_df['상품번호'], shipping_df['배송비']))
        st.sidebar.success(f"배송비 매핑 정보 {len(shipping_map)}건 로드됨")
    except Exception as e:
        st.sidebar.error(f"배송비 파일 처리 중 오류: {e}")
else:
    st.sidebar.info("배송비 파일이 없으면 택배비는 0으로 처리됩니다.")

# 2) 네이버스토어 엑셀 파일 업로드 (여러 개 가능) & 비밀번호 입력
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

# 컬럼 매핑 정의
column_mapping = {
    '주문번호': '주문번호',
    '상품번호': '상품번호',
    '정산완료일': '일자',
    '상품명': '판매품목',
    '옵션정보': '옵션명',
    '수량': '판매수량',
    '정산기준금액(A)': '판매금액',
    '네이버페이 주문관리 수수료(B)': '판매수수료',  # 추가 매핑
    '주문상태': '배송상태',
    '정산상태': '정산현황',
    '클레임상태': '기타'
}

# 필요한 컬럼
needed_cols = [
    '주문번호', '상품번호', '일자', '판매품목', '옵션명',
    '판매수량', '판매금액', '판매수수료',
    '배송상태', '정산현황', '기타'
]

# 데이터프레임 리스트를 하나로 합치고 컬럼명 일관성 확보
combined = pd.concat(file_dfs, ignore_index=True)
combined.rename(columns=column_mapping, inplace=True)

# 필요한 컬럼만 선택
combined = combined[needed_cols]

# 5) 택배비 계산 및 표시 (상품번호 기반 매핑)
combined['택배비'] = combined['상품번호'].map(shipping_map).fillna(0) * combined['판매수량']
combined['택배비'] = -combined['택배비'].astype(int)

# 6) 요약 및 결과 엑셀 생성
buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    def write_with_summary(df, sheet_name):
        df_to_write = df.copy()
        df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        # 요약 행 추가
        total_qty = df_to_write['판매수량'].sum()
        total_amount = df_to_write['판매금액'].sum()
        total_fee = df_to_write['판매수수료'].sum()
        total_delivery = df_to_write['택배비'].sum()
        total_deposit = total_fee + total_delivery
        summary_row = ws.max_row + 2
        idx_amt = list(df_to_write.columns).index('판매금액') + 1
        ws.cell(row=summary_row, column=idx_amt, value='총판매량')
        ws.cell(row=summary_row, column=idx_amt+1, value=total_qty)
        ws.cell(row=summary_row+1, column=idx_amt, value='총판매금액')
        ws.cell(row=summary_row+1, column=idx_amt+1, value=total_amount)
        ws.cell(row=summary_row+2, column=idx_amt, value='총수수료')
        ws.cell(row=summary_row+2, column=idx_amt+1, value=total_fee)
        ws.cell(row=summary_row+3, column=idx_amt, value='총택배비')
        ws.cell(row=summary_row+3, column=idx_amt+1, value=total_delivery)
        ws.cell(row=summary_row+4, column=idx_amt, value='최종입금액')
        ws.cell(row=summary_row+4, column=idx_amt+1, value=total_deposit)

    write_with_summary(combined, '결과')

buf.seek(0)
st.download_button(
    "결산 엑셀 다운로드", buf,
    file_name="네이버스토어_결산_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
```
