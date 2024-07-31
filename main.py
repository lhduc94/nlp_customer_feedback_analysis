import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
from utils import compound_unicode,check
import openpyxl
st.set_page_config(layout="wide")
uploaded_feedback_file = st.file_uploader("Tải file feedback")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


if uploaded_feedback_file is not None:
    x=pd.ExcelFile(uploaded_feedback_file)
    sheet_names = x.sheet_names
    sheet_name = st.selectbox("Chọn Sheet name", sheet_names)
    df = x.parse(sheet_name)
    columns = df.columns
    id_column = st.selectbox("Chọn ID column", columns)
    feedback_column = st.selectbox("Chọn text column", [col for col in columns if col != id_column])
    outputs = df[[id_column,feedback_column]]
    outputs[feedback_column] = outputs[feedback_column].apply(lambda x: compound_unicode(x))
    st.markdown("## Xem trước data")
    st.write(outputs.head(5))
    uploaded_topic_keywords = st.file_uploader("Tải file keyword", key='topic')
    if uploaded_topic_keywords is not None:
        x=pd.ExcelFile(uploaded_topic_keywords)
        sheet_names = x.sheet_names
        sheet_name = st.selectbox("Chọn Sheet name", sheet_names)
        df = x.parse(sheet_name)
        columns = df.columns
        topic_column = st.selectbox("Chọn topic column", columns)
        keyword_column = st.selectbox("Chọn text column", [col for col in columns if col != topic_column])
        topic_df = df[[topic_column, keyword_column]]
        topic_df[keyword_column] = topic_df[keyword_column].str.lower()
        topic_df[keyword_column] = topic_df[keyword_column].apply(lambda x: compound_unicode(x))
        topic_df = topic_df.groupby(topic_column)[keyword_column].apply(lambda x: list(x)).reset_index()
        keywords = topic_df[keyword_column].values 
        topics  = topic_df[topic_column].values
        st.markdown("## Xem trước topic và keywords")
        st.write(topic_df)
        
        for i,(col,kw) in enumerate(zip(topics,keywords)):
            outputs[col] = outputs[feedback_column].apply(lambda x: check(kw, x))
        st.markdown("## Xem trước kết quả")
        st.write(outputs.head(10))

        st.markdown('## Tải xuống DataFrame dưới dạng Excel')

        # Tạo nút tải xuống
        st.download_button(
            label='Tải xuống Excel',
            data=to_excel(outputs),
            file_name='result.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )