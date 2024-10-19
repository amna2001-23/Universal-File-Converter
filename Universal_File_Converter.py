import streamlit as st
import pytesseract
from PIL import Image
from docx import Document
import io
import os

# Set Tesseract path for deployment on Streamlit Cloud
pytesseract.pytesseract.tesseract_cmd = '/app/.apt/usr/bin/tesseract'

# Conversion functions
def convert_images_to_word(images):
    doc = Document()
    
    for image in images:
        # Convert the image to text using Tesseract OCR
        text = pytesseract.image_to_string(Image.open(image))
        
        # Add the extracted text to the Word document
        doc.add_paragraph(text)
    
    # Save the Word document to a BytesIO object
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def convert_pdf_to_word(pdf_file):
    reader = PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    
    doc = Document()
    doc.add_paragraph(text)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def convert_word_to_pdf(word_file):
    temp_docx_path = 'temp.docx'
    with open(temp_docx_path, 'wb') as f:
        f.write(word_file.read())
    
    try:
        convert(temp_docx_path, "converted.pdf")
        with open("converted.pdf", "rb") as f:
            pdf_data = f.read()
        return pdf_data
    except Exception as e:
        st.error(f"Error converting Word to PDF: {e}")
        return None
    finally:
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)
        if os.path.exists("converted.pdf"):
            os.remove("converted.pdf")

# Streamlit app
st.sidebar.title('Navigation')
page = st.sidebar.radio('Go to', ['Introduction', 'Converter'])

if page == 'Introduction':
    st.title('Welcome to the Universal File Converter App')
    # Image slider
    with st.expander("Preview Images"):
        image_list = [
            'p2w.jpg',
            'p2w1.jpg',
            'w2p.png',
            'p2w2.jpg'
        ]
        image_index = st.slider('Slider', 0, len(image_list) - 1)
        st.image(image_list[image_index])
    st.subheader("Introducing Universal File Converter: The Ultimate File Converter")
    st.write("""
    This application allows you to convert files between Word and PDF formats.
    You can convert:
    - PDFs to Word.
    - Word documents to PDFs.
    - Images to Word documents.

    **Effortless Conversion**... (More details as before)""")
    
elif page == 'Converter':
    st.title('File Converter')

    conversion_type = st.selectbox("Select Conversion Type", ["PDF to Word", "Word to PDF", "Images to Word"])

    if conversion_type == "PDF to Word":
        pdf_file = st.file_uploader("Upload PDF file", type=['pdf'])
        if pdf_file is not None:
            word_data = convert_pdf_to_word(pdf_file)
            st.download_button(label="Download Word file", data=word_data, file_name="converted.docx")

    elif conversion_type == "Word to PDF":
        word_file = st.file_uploader("Upload Word file", type=['docx'])
        if word_file is not None:
            pdf_data = convert_word_to_pdf(word_file)
            if pdf_data:
                st.download_button(label="Download PDF file", data=pdf_data, file_name="converted.pdf")

    elif conversion_type == "Images to Word":
        images = st.file_uploader("Upload Images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        if images:
            word_data = convert_images_to_word(images)
            st.download_button(label="Download Word file", data=word_data, file_name="converted.docx")
