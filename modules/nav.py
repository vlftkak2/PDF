import streamlit as st

def sidebar():
    with st.sidebar:
        st.image("images\\APR.png")
        st.header('메뉴')
        st.page_link('pages/pdf.py', label='HKTVMALL 송장주문번호', icon='🌏')
        
