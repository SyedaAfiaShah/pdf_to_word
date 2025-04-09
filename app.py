import streamlit as st
import pdfplumber
from pdf2image import convert_from_bytes
import pytesseract
from docx import Document
import os

# Function to extract embedded text (for non-scanned PDFs)
def extract_text_pdfplumber(pdf_file):
    text_content = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_content += text + "\n"
    return text_content.strip()

# Function to extract text from scanned images using OCR
def extract_text_ocr(pdf_bytes):
    images = convert_from_bytes(pdf_bytes)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img) + "\n"
    return text.strip()

# Function to convert text to Word document
def convert_text_to_word(text, output_path):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_path)

# Streamlit app interface
st.set_page_config(page_title="Smart PDF to Word Converter", layout="centered")
st.title("📄 Smart PDF to Word Converter")
st.write("Upload any PDF (normal or scanned), and download it as a Word document.")

uploaded_file = st.file_uploader("Upload a PDF", type=["pdf"])

if uploaded_file:
    if st.button("Convert to Word"):
        with st.spinner("Processing..."):

            # First try embedded text
            embedded_text = extract_text_pdfplumber(uploaded_file)

            if embedded_text:
                final_text = embedded_text
                st.info("✅ Embedded text detected. Using direct text extraction.")
            else:
                final_text = extract_text_ocr(uploaded_file.read())
                st.info("🧠 No embedded text found. Using OCR to extract scanned text.")

            # Save to Word
            output_path = "converted_output.docx"
            convert_text_to_word(final_text, output_path)

            # Download button
            with open(output_path, "rb") as f:
                st.success("Conversion complete!")
                st.download_button(
                    label="📥 Download Word File",
                    data=f,
                    file_name="converted_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            os.remove(output_path)
