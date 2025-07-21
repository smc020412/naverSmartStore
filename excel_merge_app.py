# excel_merge_app.py

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="네이버스토어 엑셀 결산", layout="wide")
st.title("네이버스토어 엑셀 결산 앱")

# 1) 파일 업로드 (여러 개 가능)
uploaded_files = st.file_uploader(
    "네이버스토어 엑셀 파일 업로드 (여러 개 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="uploader"
)

if not uploaded_files:
    st.info("하나 이상의 네이버스토어 엑셀 파일을 업로드해주세요.")
else:
    # 2) 컬럼 매핑 설정
    column_mapping = {
        '주문번호': '주문번호',
        '정산완료일': '일자',
        '상품명': '판매품목',
        '옵션정보': '옵션명',
        '수량': '판매수량',
        '정산기준금액(A)': '판매금액',
        '배송속성': '택배비',  # 배송속성을 택배비로 대체
        '주문상태': '배송상태',
        '정산상태': '정산현황',
        '클레임상태': '기타'
    }
    fee_columns = [
        '매출연동 수수료 합계(C)',
        '네이버페이 주문관리 수수료(B)',
        '무이자할부 수수료(D)'
    ]

    # 3) 파일별 데이터프레임 생성
    dfs = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        df.rename(columns=column_mapping, inplace=True)
        # 수수료 합산
        fees = [c for c in fee_columns if c in df.columns]
        if fees:
            df[fees] = df[fees].apply(lambda col: pd.to_numeric(col, errors='coerce'))
            df['판매수수료'] = df[fees].sum(axis=1)
            df.drop(columns=fees, inplace=True)
        # 필요한 컬럼 유지
        needed = [
            '주문번호','일자','판매품목','옵션명','판매수량',
            '판매금액','판매수수료','택배비','배송상태','정산현황','기타'
        ]
        df = df.loc[:, ~df.columns.duplicated()]
        df = df.reindex(columns=needed)
        dfs.append(df)

    # 4) 병합 및 집계
    combined = pd.concat(dfs, ignore_index=True)
    merged = combined.groupby('주문번호', as_index=False).agg({
        '일자': 'first',
        '판매품목': 'first',
        '옵션명': 'first',
        '판매수량': 'sum',
        '판매금액': 'sum',
        '판매수수료': 'sum',
        '택배비': 'first',
        '배송상태': lambda x: ', '.join(x.dropna().unique()),
        '정산현황': lambda x: ', '.join(x.dropna().unique()),
        '기타': lambda x: ', '.join(x.dropna().unique())
    })

    # 5) 타입 변환
    merged['일자'] = pd.to_datetime(merged['일자'], errors='coerce')
    for col in ['판매수량','판매금액','판매수수료','택배비']:
        merged[col] = pd.to_numeric(merged[col], errors='coerce').fillna(0).astype(int)
    merged[['주문번호','판매품목','옵션명','배송상태','정산현황','기타']] = \
        merged[['주문번호','판매품목','옵션명','배송상태','정산현황','기타']].astype(str)

    # 6) 날짜 범위 필터
    st.sidebar.header("날짜 범위 필터")
    valid = merged['일자'].dropna().dt.date
    if not valid.empty:
        mn, mx = valid.min(), valid.max()
        dr = st.sidebar.date_input(
            "날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx,
            key="date_range"
        )
        if isinstance(dr, tuple) and len(dr)==2:
            start, end = dr
            merged = merged[(merged['일자'].dt.date>=start)&(merged['일자'].dt.date<=end)]

    # 7) 결과 미리보기
    st.subheader("결산 데이터 미리보기")
    st.data_editor(merged, num_rows="dynamic", key="main_table")

    # 8) 엑셀 다운로드 (요약행 추가)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        def write_with_summary(df, sheet_name):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            total_qty = df['판매수량'].sum()
            total_amount = df['판매금액'].sum()
            total_fee = df['판매수수료'].sum()
            total_delivery = df['택배비'].sum()
            # 요약 시작 행
            sr = ws.max_row + 2
            # 총판매량
            col_qty = list(df.columns).index('판매수량') + 1
            ws.cell(row=sr, column=col_qty, value='총판매량')
            ws.cell(row=sr, column=col_qty+1, value=total_qty)
            # 총금액
            col_amt = list(df.columns).index('판매금액') + 1
            ws.cell(row=sr+1, column=col_amt, value='총금액')
            ws.cell(row=sr+1, column=col_amt+1, value=total_amount)
            # 총수수료
            ws.cell(row=sr+2, column=col_amt, value='총수수료')
            ws.cell(row=sr+2, column=col_amt+1, value=total_fee)
            # 총택배비
            ws.cell(row=sr+3, column=col_amt, value='총택배비')
            ws.cell(row=sr+3, column=col_amt+1, value=total_delivery)
            # 총이익
            ws.cell(row=sr+4, column=col_amt, value='총이익')
            ws.cell(row=sr+4, column=col_amt+1, value=total_amount + total_fee - total_delivery)

        write_with_summary(merged, '결산')

    buf.seek(0)
    st.download_button(
        "결산 엑셀 다운로드", buf,
        file_name="네이버스토어_결산_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_button"
    )
