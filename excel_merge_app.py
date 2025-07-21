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
    accept_multiple_files=True
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

    # 3) 병합
    combined = pd.concat(dfs, ignore_index=True)
    merged = combined.groupby('주문번호', as_index=False).agg({
        '일자': 'first',
        '판매품목': 'first',
        '옵션명': 'first',
        '판매수량': 'sum',
        '판매금액': 'sum',
        '판매수수료': 'sum',
        '택배비': 'first',  # 입력된 택배비 유지
        '배송상태': lambda x: ', '.join(x.dropna().unique()),
        '정산현황': lambda x: ', '.join(x.dropna().unique()),
        '기타': lambda x: ', '.join(x.dropna().unique())
    })

    # 4) 타입 변환
    merged['일자'] = pd.to_datetime(merged['일자'], errors='coerce')
    for col in ['판매수량','판매금액','판매수수료','택배비']:
        merged[col] = pd.to_numeric(merged[col], errors='coerce').fillna(0).astype(int)
    for col in ['주문번호','판매품목','옵션명','배송상태','정산현황','기타']:
        merged[col] = merged[col].fillna('').astype(str)

    # 5) 날짜 범위 필터
    st.sidebar.header("날짜 범위 필터")
    valid = merged['일자'].dropna().dt.date
    if not valid.empty:
        mn, mx = valid.min(), valid.max()
        dr = st.sidebar.date_input(
            "날짜 범위 선택", value=(mn, mx), min_value=mn, max_value=mx
        )
        if isinstance(dr, tuple) and len(dr) == 2:
            start, end = dr
            merged = merged[((merged['일자'].dt.date >= start) & (merged['일자'].dt.date <= end)) | merged['일자'].isna()]

    # 5.1) 택배비 입력
    st.sidebar.header("택배비 설정")
    delivery_fee = st.sidebar.number_input("택배비 (정수)", min_value=0, value=0)

    # 6) 정상/문제 구분
    df_ok = merged[(merged['판매수량'] > 0) & (merged['판매금액'] > 0) & merged['일자'].notna()]
    df_err = merged[(merged['판매수량'] == 0) | (merged['판매금액'] == 0) | merged['일자'].isna()]

    # 6.1) 택배비 컬럼 업데이트 (모두 입력된 택배비로 설정)
    df_ok['택배비'] = delivery_fee
    df_err['택배비'] = delivery_fee

    # 7) 표시
    st.subheader("판매수량 및 판매금액 정상 데이터")
    st.data_editor(df_ok, num_rows="dynamic", key="ok_table")
    st.subheader("판매수량 또는 판매금액이 0이거나 일자가 없는 데이터")
    st.data_editor(df_err, num_rows="dynamic", key="err_table")

    # 8) 엑셀 다운로드 (2개의 시트 + 요약행 추가)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        def write_with_summary(df, sheet_name):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            total_qty = df['판매수량'].sum()
            total_amount = df['판매금액'].sum()
            total_fee = df['판매수수료'].sum()
            total_delivery = -(total_qty * delivery_fee)
            total_deposit = total_fee + total_delivery
            summary_row = ws.max_row + 2  # 한 칸 내려서 시작
            # 총판매량
            col_qty = list(df.columns).index('판매금액') + 1
            ws.cell(row=summary_row, column=col_qty, value='총판매량')
            ws.cell(row=summary_row, column=col_qty+1, value=total_qty)
            # 총금액
            col_amt = list(df.columns).index('판매금액') + 1
            ws.cell(row=summary_row+1, column=col_amt, value='총금액')
            ws.cell(row=summary_row+1, column=col_amt+1, value=total_amount)
            # 총수수료
            ws.cell(row=summary_row+2, column=col_amt+2, value='총수수료')
            ws.cell(row=summary_row+2, column=col_amt+3, value=total_fee)
            # 총택배비
            ws.cell(row=summary_row+3, column=col_amt+2, value='총택배비')
            ws.cell(row=summary_row+3, column=col_amt+3, value= total_delivery)

            # 총지출
            ws.cell(row=summary_row+4, column=col_amt+2, value='총지출')
            ws.cell(row=summary_row+4, column=col_amt+3, value= total_deposit)

            # 총이익
            ws.cell(row=summary_row+5, column=col_amt, value='총이익')
            ws.cell(row=summary_row+5, column=col_amt+1, value=total_amount + total_fee + total_delivery)
            

        write_with_summary(df_ok, '정상')
        write_with_summary(df_err, '문제')

    buf.seek(0)
    st.download_button(
        "결산 엑셀 다운로드", buf,
        file_name="네이버스토어_결산_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_button"
    )
