import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
from docx2pdf import convert
import io
import os
import pythoncom

# Conversion functions
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
    pythoncom.CoInitialize()  # Initialize the COM library
    temp_docx_path = 'temp.docx'
    with open(temp_docx_path, 'wb') as f:
        f.write(word_file.read())
    
    try:
        convert(temp_docx_path, "converted.pdf")
        with open("converted.pdf", "rb") as f:
            pdf_data = f.read()
        os.remove(temp_docx_path)
        os.remove("converted.pdf")
        return pdf_data
    except Exception as e:
        st.error(f"Error converting Word to PDF: {e}")
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)
        if os.path.exists("converted.pdf"):
            os.remove("converted.pdf")
        return None

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
    
    **Effortless Conversion**

    Universal File Converter is a versatile file converter that empowers users to effortlessly convert files between Word and PDF formats. With just a few clicks, you can transform your documents into the format that best suits your needs.

    **Convert PDFs to Word**

    This functionality is invaluable for users who need to extract text or data from PDF files or make edits to existing documents. Whether you're converting a PDF report into a Word document for further editing, Universal File Converter makes the process quick and hassle-free.

    **Create PDFs from Word Documents**

    This feature is ideal for users who need to share files in a secure and universally accessible format. Whether you're compiling a series of documents into a single PDF or converting an image into a professional-looking PDF, Universal File Converter ensures that your files are presented in the best possible way.

    **User-Friendly Interface**

    Universal File Converter boasts a user-friendly interface that makes file conversion a breeze. The intuitive design allows users of all experience levels to navigate the software effortlessly and complete conversions quickly. Whether you're a seasoned professional or a novice user, you'll find Universal File Converter to be a valuable tool for all your file conversion needs.

    **Secure and Reliable**

    Security is paramount when it comes to file conversion, and Universal File Converter takes this aspect seriously. The application ensures that your files are handled securely throughout the conversion process, giving you peace of mind knowing that your sensitive information is protected. Additionally, Universal File Converter delivers reliable performance, ensuring that your files are accurately converted every time.

    **Conclusion**

    In conclusion, Universal File Converter is the ultimate file converter for users who demand efficiency, versatility, and reliability. Whether you're converting PDFs to Word or creating PDFs from Word documents, Universal File Converter has you covered. With its user-friendly interface, secure handling of files, and seamless conversion capabilities, Universal File Converter is the go-to solution for all your file conversion needs.

    Ready to experience the power of Universal File Converter? Try it today and take your file conversion game to the next level!
    """)

elif page == 'Converter':
    st.title('File Converter')

    option = st.selectbox('Choose conversion type:', ['PDF to Word', 'Word to PDF'])
    uploaded_file = st.file_uploader("Upload a file", type=["pdf", "docx"])

    if uploaded_file is not None:
        if option == 'PDF to Word':
            word_data = convert_pdf_to_word(uploaded_file)
            st.download_button('Download Word file', word_data, file_name='converted.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

        elif option == 'Word to PDF':
            pdf_data = convert_word_to_pdf(uploaded_file)
            if pdf_data:
                st.download_button('Download PDF file', pdf_data, file_name='converted.pdf', mime='application/pdf')





#streamlit run "Universal File Converter.py"
