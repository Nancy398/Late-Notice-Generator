import streamlit as st
from io import BytesIO
import pypandoc
from io import BytesIO
from num2words import num2words  # 用于将数字转换为英文大写
from datetime import datetime
from docx import Document


# 加载本地模板
def load_template(file_path="Late Rent Notice Template.docx"):
    return Document(file_path)

st.title("基于内置模板生成 PDF")

# 用户输入替换值
last_name = st.text_input("Last Name")
first_name = st.text_input("First Name")
address = st.text_area("Address")
postal = st.text_input("Postal Code")
title = st.selectbox("Title", ["Mr.", "Ms."])
amount = st.number_input("Amount", min_value=0.0, format="%.2f")

# 将金额转换为英文大写
def amount_to_words(amount):
    return num2words(amount, to='currency', lang='en')
amount_words = amount_to_words(amount)

def get_current_date():
    now = datetime.now()
    return now.strftime("%B %d, %Y")
current_date = get_current_date()

if st.button("生成 PDF"):
    # 加载模板
    doc = load_template()

    # 替换占位符
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
        if "{Title}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Title}", title)
        if "{Amount}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Amount}", str(amount))
        if "{Amount_Words}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{Amount_Words}", amount_words)

    # 保存修改后的 Word 文档到内存
    word_buffer = BytesIO()
    doc.save(word_buffer)
    word_buffer.seek(0)

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
        temp_docx.write(word_buffer.getvalue())
        temp_docx.close()

    # 使用 Pandoc 转换为 PDF
    pdf_output_path = temp_docx.name.replace('.docx', '.pdf')
    pypandoc.convert_file(temp_docx.name, 'pdf', format='docx', outputfile=pdf_output_path)


    # 提供下载链接
    st.success("PDF 已生成！")
    st.download_button(
        label="下载生成的 PDF",
        data=word_buffer.getvalue(),
        file_name="generated_template.pdf",
        mime="application/pdf",
    )
