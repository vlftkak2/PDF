import re
import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from container.footer import footer
from modules.nav import sidebar
#import spacy
#from spacy import displacy

st.set_page_config(page_title="HKTVMALL 송장주문번호",
                page_icon="🌏",
                layout="wide")

vert_space = '<div style="padding: 20px 5px;"></div>'
sidebar()
footer()

#@st.cache_resource
#def load_model():
    #nlp = spacy.load("en_core_web_sm")
    #return nlp

#nlp = load_model()

curdate = datetime.now()
curdate = curdate.strftime("%Y-%m-%d")

curweek = datetime.now()
curweek = curweek.strftime("%w")

if curweek == "1":
    curweek = "월"
elif curweek == "2":
    curweek = "화"
elif curweek == "3":
    curweek = "수"
elif curweek == "4":
    curweek = "목"
elif curweek == "5":
    curweek = "금"
elif curweek =="6":
    curweek = "토"
elif curweek == "0":
    curweek ="일"
else:
    curweek = "에러"

st.title('🌏APR HKTVMALL')
st.subheader("📅날짜 : "+curdate+"("+curweek+")")

st.markdown(vert_space, unsafe_allow_html=True)

# Streamlit UI
st.write("PDF파일을 업로드 해주시면, 엑셀로 다운로드를 받으실 수 있습니다!")

# PDF 파일 업로드
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file:
    # PyMuPDF로 PDF 파일 읽기
    pdf_document = fitz.open("pdf", uploaded_file.read())

    # 주문번호 형식의 정규표현식
    order_number_pattern = r"Order No\.:?\s?(H\d{8}\s?\d{4})"

    # 모든 주문번호를 저장할 리스트
    order_numbers = []

    # 모든 페이지 순회
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]
        text = page.get_text("text")  # 텍스트 추출
        # 정규표현식을 사용해 주문번호 찾기
        matches = re.findall(order_number_pattern, text)
        # 공백 제거 후 리스트에 추가
        order_numbers.extend([match.replace(" ", "") for match in matches])

    pdf_document.close()

    # 중복 제거 및 정렬
    unique_order_numbers = list(set(order_numbers))
    unique_order_numbers.sort(key=lambda x: order_numbers.index(x))

    # DataFrame 생성 (순번 추가)
    df = pd.DataFrame({
        "Index": range(1, len(unique_order_numbers) + 1),
        "OrderNumbers": unique_order_numbers
    })

    # 엑셀 파일 변환
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Order Numbers")
    output.seek(0)

    # 현재 날짜와 시간을 파일 이름에 추가
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"HKTVMALL_OrderNumber_{timestamp}.xlsx"

    # 레이아웃 컨테이너 생성
    with st.container():
        # 다운로드 버튼 먼저 출력
        st.download_button(
            label="Download Order Numbers as Excel",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # 텍스트 결과 출력
        st.write(f"Total Unique Order Numbers Found: {len(unique_order_numbers)}")
        if len(unique_order_numbers) > 0:
            st.write("Extracted Order Numbers:")
            for order_number in unique_order_numbers:
                st.write(order_number)
