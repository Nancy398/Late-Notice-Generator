import streamlit as st
from io import BytesIO
import pypandoc
from num2words import num2words  # 用于将数字转换为英文大写
from datetime import datetime
from docx import Document
from docx2pdf import convert

# 加载本地模板
def load_template(file_path="Late Rent Notice Template.docx"):
    return Document(file_path)
doc = load_template()
st.title("基于内置模板生成 PDF")

# 用户输入替换值
last_name = st.text_input("Last Name")
first_name = st.text_input("First Name")
address = st.text_area("Address")
postal = st.text_input("Postal Code")
title = st.selectbox("Title", ["Mr.", "Ms."])
amount = st.number_input("Amount", min_value=0.0, format="%.2f")
formatted_amount = "{:,.2f}".format(amount)

# 将金额转换为英文大写
def amount_to_words(amount):
    return num2words(amount, to='currency', lang='en', currency ='USD').title()
amount_words = amount_to_words(amount)

def get_current_date():
    now = datetime.now()
    return now.strftime("%B %d, %Y")
current_date = get_current_date()

if st.button("生成 PDF"):
    for paragraph in doc.paragraphs:
        if "{First Name}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{First Name}", first_name)
        if "{Last Name}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Last Name}", last_name)
        if "{Date}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Date}", current_date)
        if "{Address}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Address}", address)
        if "{Postal}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Postal}", postal)
        if "{gender}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{gender}", title)
        if "{Amount}" in paragraph.text:
            before_text = paragraph.text.split("{Amount}")[0]
            after_text = paragraph.text.split("{Amount}")[1]
            paragraph.clear()
            paragraph.add_run(before_text)
            run = paragraph.add_run(str(formatted_amount))
            run.bold = True
            run.underline = True
            paragraph.add_run(after_text)
        if "{Amount Words}" in paragraph.text:
            before_text = paragraph.text.split("{Amount Words}")[0]
            after_text = paragraph.text.split("{Amount Words}")[1]
            paragraph.clear()
            paragraph.add_run(before_text)
            run = paragraph.add_run(amount_words)
            run.bold = True
            run.underline = True
            paragraph.add_run(after_text)

buffer = BytesIO()
doc.save(buffer)
buffer.seek(0)
output_pdf = "output.pdf"
convert(buffer, output_pdf)

#     # 提供下载链接
st.success("PDF 已生成！")
with open(output_pdf, "rb") as f:
    st.download_button(
        label="下载 PDF 文件",
        data=f,
        file_name="Late Notice.pdf",
        mime="application/pdf"
    )




