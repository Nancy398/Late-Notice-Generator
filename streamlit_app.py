import streamlit as st
from io import BytesIO
import pypandoc
from num2words import num2words  # 用于将数字转换为英文大写
from datetime import datetime
from docx import Document
import pandoc
import fitz  # PyMuPDF
import os

# 加载本地模板
import streamlit as st
import fitz  # PyMuPDF
import os

# PDF 填充函数
def fill_pdf(toutput_path, data):
    pdf = fitz.open("Late Notice.pdf")
    
    for page_num in range(len(pdf)):
        page = pdf[page_num]
        for key, value in data.items():
            search_term = f"{{{{{ key }}}}}"  # 占位符格式 {{key}}
            matches = page.search_for(search_term)
            for match in matches:
                rect = match
                page.insert_text(
                    rect.tl, " " * len(search_term), fontsize=12, fontname="helv", color=(1, 1, 1)  # 使用白色替换占位符
                )
                page.insert_text(
                    rect.tl, value, fontsize=12, fontname="helv", color=(0, 0, 0)
                )
            

    pdf.save(output_path)
    pdf.close()

# Streamlit 界面
st.title("PDF 模板填充器")
last_name = st.text_input("Last Name")
first_name = st.text_input("First Name")
address = st.text_area("Address")
postal = st.text_input("Postal Code")
title = st.selectbox("Title", ["Mr.", "Ms."])
amount = st.number_input("Amount", min_value=0.0, format="%.2f")
formatted_amount = "{:,.2f}".format(amount)
def amount_to_words(amount):
    return num2words(amount, to='currency', lang='en', currency ='USD').title()
amount_words = amount_to_words(amount)

def get_current_date():
    now = datetime.now()
    return now.strftime("%B %d, %Y")
current_date = get_current_date()
    
data = {
    "First Name": first_name,
    "Last Name": last_name,
    "Address": address,
    "Postal": postal,
    "gender": title,
    "Amount": str(amount),
    "Amount Words":formatted_amount
}

# 生成 PDF
if st.button("生成 PDF"):
        output_path = "filled_template.pdf"
        fill_pdf(output_path, data)
        with open(output_path, "rb") as f:
            st.download_button(
                label="下载 PDF",
                data=f,
                file_name="filled_template.pdf",
                mime="application/pdf"
            )
        st.success("PDF 生成成功！")
#     else:
#         st.error("请先上传 PDF 模板！")


# # 用户输入替换值

# # 将金额转换为英文大写


# if st.button("生成 PDF"):
#     for paragraph in doc.paragraphs:
#         if "{First Name}" in paragraph.text:
#             paragraph.text = paragraph.text.replace("{First Name}", first_name)
#         if "{Last Name}" in paragraph.text:
#             paragraph.text = paragraph.text.replace("{Last Name}", last_name)
#         if "{Date}" in paragraph.text:
#             paragraph.text = paragraph.text.replace("{Date}", current_date)
#         if "{Address}" in paragraph.text:
#             paragraph.text = paragraph.text.replace("{Address}", address)
#         if "{Postal}" in paragraph.text:
#             paragraph.text = paragraph.text.replace("{Postal}", postal)
#         if "{gender}" in paragraph.text:
#             paragraph.text = paragraph.text.replace("{gender}", title)
#         if "{Amount}" in paragraph.text:
#             before_text = paragraph.text.split("{Amount}")[0]
#             after_text = paragraph.text.split("{Amount}")[1]
#             paragraph.clear()
#             paragraph.add_run(before_text)
#             run = paragraph.add_run(str(formatted_amount))
#             run.bold = True
#             run.underline = True
#             paragraph.add_run(after_text)
#         if "{Amount Words}" in paragraph.text:
#             before_text = paragraph.text.split("{Amount Words}")[0]
#             after_text = paragraph.text.split("{Amount Words}")[1]
#             paragraph.clear()
#             paragraph.add_run(before_text)
#             run = paragraph.add_run(amount_words)
#             run.bold = True
#             run.underline = True
#             paragraph.add_run(after_text)

# doc.save("Transfer.docx")
# #     # 提供下载链接
# st.success("Docx 已生成！")
# with open("Transfer.docx", "rb") as f:
#     st.download_button(
#         label="下载 docx 文件",
#         data=f,
#         file_name="Late Notice.docx",
#         mime="application/pdf"
#     )



# Title of the Streamlit app
# st.title("DOCX to PDF Converter")
# pypandoc.download_pandoc()
# # File uploader
# uploaded_file = st.file_uploader("Upload DOCX file", type="docx")

# if uploaded_file is not None:
#     # Save the uploaded DOCX file to BytesIO object
#     docx_file = BytesIO(uploaded_file.read())

#     # Convert the DOCX to PDF using pypandoc
#     try:
#         # We need to save the DOCX file to a temporary file first for pypandoc to work
#         temp_input_path = "/tmp/uploaded_file.docx"
#         with open(temp_input_path, 'wb') as f:
#             f.write(docx_file.getbuffer())
        
#         # Output PDF path
#         temp_output_path = "/tmp/output.pdf"
        
#         # Convert the DOCX file to PDF using pypandoc
#         pypandoc.convert_file(temp_input_path, to='pdf', format='docx', outputfile=temp_output_path,extra_args=['--pdf-engine=xelatex'])
        
#         # Provide download link for the converted PDF
#         with open(temp_output_path, "rb") as pdf_file:
#             st.download_button("Download PDF", pdf_file, file_name="converted_file.pdf", mime="application/pdf")
        
#     except Exception as e:
#         st.error(f"Error: {e}")



