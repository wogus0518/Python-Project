# Python v3.9.7
import streamlit as st  # v0.89.0
import pandas_datareader as pdr  # v0.10.0

# 마크다운 문법이 먹힌다
st.write('''
# 삼성전자 주식 데이터
마감 가격과 거래량을 차트로 보여줍니다!
''')

# pdr.get_data_yahoo() 야후 주식 historical 데이터를 dataframe 형태로 반환한다.
df = pdr.get_data_yahoo('005930.KS', "2020-01-01", "2021-09-15")

st.line_chart(df.Close)
st.line_chart(df.Volume)
