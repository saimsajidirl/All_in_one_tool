import streamlit as st
import os
import tempfile
from PIL import Image
from pdf2docx import Converter
from PyPDF2 import PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
import subprocess
from moviepy import VideoFileClip

st.set_page_config(
    page_title="All in one Tool",
    page_icon="logo.jpg",  # Path to your favicon file
    layout="centered"
)


def convert_png_to_ico(png_file):
    """Convert PNG to ICO file"""
    try:
        img = Image.open(png_file)

        original_filename = os.path.splitext(png_file.name)[0]
        with tempfile.NamedTemporaryFile(delete=False, suffix='.ico', prefix=f'{original_filename}_') as temp_ico:
            sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128)]
            img.save(temp_ico.name, format='ICO', sizes=sizes)

        return temp_ico.name, f"{original_filename}.ico"
    except Exception as e:
        st.error(f"Error converting PNG to ICO: {e}")
        return None, None


def convert_pdf_to_word(pdf_file):
    """Convert PDF to DOCX file"""
    try:
        original_filename = os.path.splitext(pdf_file.name)[0]

        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', prefix=f'{original_filename}_') as temp_pdf, \
                tempfile.NamedTemporaryFile(delete=False, suffix='.docx', prefix=f'{original_filename}_') as temp_docx:

            temp_pdf.write(pdf_file.getvalue())
            temp_pdf.close()

            cv = Converter(temp_pdf.name)
            cv.convert(temp_docx.name)
            cv.close()

        return temp_docx.name, f"{original_filename}.docx"
    except Exception as e:
        st.error(f"Error converting PDF to DOCX: {e}")
        return None, None


def merge_pdfs(pdf_files):
    """Merge multiple PDF files"""
    try:
        merger = PdfMerger()

        temp_pdf_paths = []
        for pdf_file in pdf_files:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                temp_pdf.write(pdf_file.getvalue())
                temp_pdf_paths.append(temp_pdf.name)

        for pdf_path in temp_pdf_paths:
            merger.append(pdf_path)

        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as merged_pdf:
            merger.write(merged_pdf.name)
            merger.close()

        for pdf_path in temp_pdf_paths:
            os.unlink(pdf_path)

        return merged_pdf.name, "merged_document.pdf"

    except Exception as e:
        st.error(f"Error merging PDFs: {e}")
        return None, None


def convert_images_to_pdf(image_files):
    """Convert multiple images to a single PDF"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            c = canvas.Canvas(temp_pdf.name, pagesize=letter)
            width, height = letter

            for img_file in image_files:
                img = Image.open(img_file)

                if img.mode != 'RGB':
                    img = img.convert('RGB')

                img_width, img_height = img.size
                aspect_ratio = img_width / img_height

                max_width = width - 2 * inch
                max_height = height - 2 * inch

                if img_width > max_width or img_height > max_height:
                    if aspect_ratio > 1:
                        draw_width = min(img_width, max_width)
                        draw_height = draw_width / aspect_ratio
                    else:
                        draw_height = min(img_height, max_height)
                        draw_width = draw_height * aspect_ratio
                else:
                    draw_width, draw_height = img_width, img_height

                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_img:
                    img.save(temp_img.name)

                x_centered = (width - draw_width) / 2
                y_centered = (height - draw_height) / 2

                c.drawImage(temp_img.name, x_centered, y_centered, width=draw_width, height=draw_height)

                c.showPage()

                os.unlink(temp_img.name)

            c.save()

        return temp_pdf.name, "converted_images.pdf"
    except Exception as e:
        st.error(f"Error converting images to PDF: {e}")
        return None, None


def convert_jpg_to_png(jpg_file):
    """Convert JPG to PNG file"""
    try:
        img = Image.open(jpg_file)

        original_filename = os.path.splitext(jpg_file.name)[0]
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png', prefix=f'{original_filename}_') as temp_png:
            if img.mode != 'RGB':
                img = img.convert('RGB')

            img.save(temp_png.name, format='PNG')

        return temp_png.name, f"{original_filename}.png"
    except Exception as e:
        st.error(f"Error converting JPG to PNG: {e}")
        return None, None


def convert_ppt_to_pdf(ppt_file):
    """Convert PPT to PDF using LibreOffice"""
    try:
        # Save the PPT file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_ppt:
            temp_ppt.write(ppt_file.getvalue())
            temp_ppt.close()

            # Use LibreOffice in headless mode to convert the PPTX file to PDF
            pdf_path = temp_ppt.name.replace('.pptx', '.pdf')
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', temp_ppt.name], check=True)

        return pdf_path, "converted_presentation.pdf"
    except Exception as e:
        st.error(f"Error converting PPT to PDF: {e}")
        return None, None


def convert_mp4_to_mp3(mp4_file):
    """Convert MP4 to MP3 file"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.mp4') as temp_mp4, \
                tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as temp_mp3:

            temp_mp4.write(mp4_file.getvalue())
            temp_mp4.close()

            video = VideoFileClip(temp_mp4.name)
            video.audio.write_audiofile(temp_mp3.name)
            video.close()

        original_filename = os.path.splitext(mp4_file.name)[0]
        return temp_mp3.name, f"{original_filename}.mp3"

    except Exception as e:
        st.error(f"Error converting MP4 to MP3: {e}")
        return None, None


def main():
    st.title("All in one Tool from Muhammad Saim")

    menu = st.sidebar.radio("Select Conversion", ["PNG to ICO", "PDF to Word", "Merge PDFs",
                                                  "Images to PDF", "JPG to PNG", "PPT to PDF",
                                                  "MP4 to MP3"])  # Added new conversion option

    if menu == "PNG to ICO":
        st.header("PNG to ICO Converter")
        png_file = st.file_uploader("Upload PNG file", type=['png'])
        if png_file is not None:
            if st.button("Convert to ICO"):
                ico_path, ico_filename = convert_png_to_ico(png_file)
                if ico_path:
                    with open(ico_path, 'rb') as f:
                        ico_bytes = f.read()
                    st.download_button(label="Download ICO", data=ico_bytes, file_name=ico_filename,
                                       mime="image/x-icon")
                    os.unlink(ico_path)

    elif menu == "PDF to Word":
        st.header("PDF to Word Converter")
        pdf_file = st.file_uploader("Upload PDF file", type=['pdf'])
        if pdf_file is not None:
            if st.button("Convert to DOCX"):
                docx_path, docx_filename = convert_pdf_to_word(pdf_file)
                if docx_path:
                    with open(docx_path, 'rb') as f:
                        docx_bytes = f.read()
                    st.download_button(label="Download DOCX", data=docx_bytes, file_name=docx_filename,
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    os.unlink(docx_path)

    elif menu == "Merge PDFs":
        st.header("Merge PDF Files")
        pdf_files = st.file_uploader("Upload PDF files", type=['pdf'], accept_multiple_files=True)
        if pdf_files:
            if st.button("Merge PDFs"):
                merged_pdf_path, merged_filename = merge_pdfs(pdf_files)
                if merged_pdf_path:
                    with open(merged_pdf_path, 'rb') as f:
                        merged_bytes = f.read()
                    st.download_button(label="Download Merged PDF", data=merged_bytes, file_name=merged_filename,
                                       mime="application/pdf")
                    os.unlink(merged_pdf_path)

    elif menu == "Images to PDF":
        st.header("Convert Images to PDF")
        image_files = st.file_uploader("Upload Image files", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
        if image_files:
            if st.button("Convert to PDF"):
                pdf_path, pdf_filename = convert_images_to_pdf(image_files)
                if pdf_path:
                    with open(pdf_path, 'rb') as f:
                        pdf_bytes = f.read()
                    st.download_button(label="Download PDF", data=pdf_bytes, file_name=pdf_filename,
                                       mime="application/pdf")
                    os.unlink(pdf_path)

    elif menu == "JPG to PNG":
        st.header("JPG to PNG Converter")
        jpg_file = st.file_uploader("Upload JPG file", type=['jpg', 'jpeg'])
        if jpg_file is not None:
            if st.button("Convert to PNG"):
                png_path, png_filename = convert_jpg_to_png(jpg_file)
                if png_path:
                    with open(png_path, 'rb') as f:
                        png_bytes = f.read()
                    st.download_button(label="Download PNG", data=png_bytes, file_name=png_filename, mime="image/png")
                    os.unlink(png_path)

    elif menu == "PPT to PDF":
        st.header("PPT to PDF Converter")
        ppt_file = st.file_uploader("Upload PPT file", type=['ppt', 'pptx'])
        if ppt_file is not None:
            if st.button("Convert to PDF"):
                pdf_path, pdf_filename = convert_ppt_to_pdf(ppt_file)
                if pdf_path:
                    with open(pdf_path, 'rb') as f:
                        pdf_bytes = f.read()
                    st.download_button(label="Download PDF", data=pdf_bytes, file_name=pdf_filename,
                                       mime="application/pdf")
                    os.unlink(pdf_path)

    elif menu == "MP4 to MP3":
        st.header("MP4 to MP3 Converter")
        mp4_file = st.file_uploader("Upload MP4 file", type=['mp4'])
        if mp4_file is not None:
            st.video(mp4_file)
            if st.button("Convert to MP3"):
                mp3_path, mp3_filename = convert_mp4_to_mp3(mp4_file)
                if mp3_path:
                    with open(mp3_path, 'rb') as f:
                        mp3_bytes = f.read()
                    st.download_button(label="Download MP3", data=mp3_bytes, file_name=mp3_filename, mime="audio/mpeg")
                    os.unlink(mp3_path)


if __name__ == "__main__":
    main()
