import streamlit as st
from io import BytesIO
import pypandoc
from num2words import num2words  # 用于将数字转换为英文大写
from datetime import datetime
from docx import Document
import tempfile


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

    # 保存修改后的 Word 文档到内存
#     word_buffer = BytesIO()
#     doc.save(word_buffer)
#     word_buffer.seek(0)
    
#     pdf_output = BytesIO()
# # 通过 Pandoc 转换 DOCX 到 PDF
    
#     pypandoc.convert_file(word_buffer, to='pdf', format='docx', outputfile=pdf_output)
#     pdf_output.seek(0)


#     # 提供下载链接
#     st.success("PDF 已生成！")
#     with open(output_pdf, "rb") as f:
#         st.download_button(
#             label="下载 PDF 文件",
#             data=f,
#             file_name="Late Notice.pdf",
#             mime="application/pdf"
#         )

import os

# 使用 Pandoc 将 DOCX 转换为 PDF
def convert_docx_to_pdf(docx_path):
    # 创建临时 PDF 文件
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
        # 使用 Pandoc 转换文件
        pdf_output_path = temp_pdf.name
        pypandoc.convert_file("Late Rent Notice Template.docx", to='pdf', format='docx', outputfile=pdf_output_path)
        return pdf_output_path

pdf_output_path = convert_docx_to_pdf(doc)
        
        # 显示 PDF 文件下载链接
with open(pdf_output_path, "rb") as pdf_file:
    st.download_button(
        label="Download PDF",
        data=pdf_file,
        file_name="converted_document.pdf",
        mime="application/pdf"
            )

        # 删除临时文件
os.remove(docx_path)
os.remove(pdf_output_path)


