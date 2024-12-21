import streamlit as st
from docx import Document
from io import BytesIO
import pypandoc

# 加载本地模板
def load_template(file_path="template.docx"):
    return Document(file_path)

st.title("基于内置模板生成 PDF")

# 用户输入替换值
name = st.text_input("姓名", "示例姓名")
date = st.date_input("日期")
message = st.text_area("消息", "这是一个示例消息")

if st.button("生成 PDF"):
    # 加载模板
    doc = load_template()

    # 替换占位符
    for paragraph in doc.paragraphs:
        if "{name}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{name}", name)
        if "{date}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{date}", str(date))
        if "{message}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{message}", message)

    # 保存修改后的 Word 文档到内存
    word_buffer = BytesIO()
    doc.save(word_buffer)
    word_buffer.seek(0)

    # 转换为 PDF
    pdf_output = BytesIO()
    pypandoc.convert_file(
        word_buffer, "pdf", format="docx", outputfile=pdf_output
    )
    pdf_output.seek(0)

    # 提供下载链接
    st.success("PDF 已生成！")
    st.download_button(
        label="下载生成的 PDF",
        data=pdf_output,
        file_name="generated_template.pdf",
        mime="application/pdf",
    )
