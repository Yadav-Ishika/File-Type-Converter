import os
import time
import fitz
import tempfile
import pandas as pd
from PIL import Image
import streamlit as st
from docx2pdf import convert
from pdf2docx import Converter
from reportlab.pdfgen import canvas
from pdf2image import convert_from_path
from reportlab.lib.pagesizes import letter

# Custom CSS to make buttons full width
st.markdown(
    """
    <style>
    .stButton > button {
        width: 100%;
    }
    </style>
    """,
    unsafe_allow_html=True
)


placeholder = st.empty()

# Title with full width
placeholder.markdown("""
    <style>
        .header-container {
            display: flex;
            flex-direction: column;
            gap: 0.5vw;
            align-items: center;
            padding: 20px;
            background-color: #f8f9fa; /* Light background for contrast */
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            margin-bottom: 10px;
        }
        .header-title {
            text-align: center;
            font-size: 4.2vw;
            margin: 0;
        }
        .header-subtitle {
            color: green;
            font-weight: bold;
        }
        .description {
            text-align: center;
            font-size: 1.2vw;
            color: #555;
            max-width: 600px; /* Limit width for readability */
        }
        hr {
            width: 80%; /* Horizontal rule width */
            border: 1px solid #ccc;
        }
    </style>
    <div class="header-container">
        <h1 class="header-title">"File Type Converter"<br /><span class="header-subtitle">Welcomes</span> You!</h1>
        <hr />
        <p class="description">
            Our service offers seamless conversions between various file formats, including:
            <ul>
                <li>CSV to PDF</li>
                <li>PDF to Image (JPEG, JPG, PNG)</li>
                <li>Image (JPEG, JPG, PNG) to PDF</li>
            </ul>
            Experience fast, reliable, and user-friendly file conversion to meet all your needs!
        </p>
    </div>
""", unsafe_allow_html=True)

# Initialize session state variables
if 'conversion_type' not in st.session_state:
    st.session_state.conversion_type = None
if 'docx_output_pdf_path' not in st.session_state:
    st.session_state.output_pdf_path = None
if 'pdf_output_docx_path' not in st.session_state:
    st.session_state.output_docx_path = None
if 'pdf_output_csv_path' not in st.session_state:
    st.session_state.output_csv_path = None
if 'csv_output_pdf_path' not in st.session_state:
    st.session_state.output_pdf_path_csv = None
if 'pdf_to_jpg_image_paths' not in st.session_state:
    st.session_state.jpg_image_paths = []
if 'pdf_to_png_image_paths' not in st.session_state:
    st.session_state.png_image_paths = []
if 'jpeg_image_paths' not in st.session_state:
    st.session_state.jpeg_image_paths = []
if 'jpg_image_paths' not in st.session_state:
    st.session_state.jpg_image_paths = []
if 'png_image_paths' not in st.session_state:
    st.session_state.png_image_paths = []


st.sidebar.markdown("<div style='position: relative; top: -30px;'><h1 style='text-align: center; color: #000; font-size: 3.5vw;'>Services</h1></div>", unsafe_allow_html=True)

# Sidebar with buttons
if st.button("Reload"):
    st.session_state.conversion_type = "reload_content"

if st.sidebar.button("DOCX to PDF Converter"):
    st.session_state.conversion_type = "docx_to_pdf"
    placeholder.empty()

if st.sidebar.button("PDF to DOCX Converter"):
    st.session_state.conversion_type = "pdf_to_docx"
    placeholder.empty()

if st.sidebar.button("PDF to CSV Converter"):
    st.session_state.conversion_type = "pdf_to_csv"
    placeholder.empty()

if st.sidebar.button("CSV to PDF Converter"):
    st.session_state.conversion_type = "csv_to_pdf"
    placeholder.empty()

if st.sidebar.button("PDF to JPEG Converter"):
    st.session_state.conversion_type = "pdf_to_jpeg"
    placeholder.empty()

if st.sidebar.button("PDF to JPG Converter"):
    st.session_state.conversion_type = "pdf_to_jpg"
    placeholder.empty()

if st.sidebar.button("PDF to PNG Converter"):
    st.session_state.conversion_type = "pdf_to_png"
    placeholder.empty()

if st.sidebar.button("JPEG to PDF Converter"):
    st.session_state.conversion_type = "jpeg_to_pdf"
    placeholder.empty()

if st.sidebar.button("JPG to PDF Converter"):
    st.session_state.conversion_type = "jpg_to_pdf"
    placeholder.empty()

if st.sidebar.button("PNG to PDF Converter"):
    st.session_state.conversion_type = "png_to_pdf"


if st.session_state.conversion_type == "reload_content":
    # Title with full width
    placeholder.markdown("""
        <style>
            .header-container {
                display: flex;
                flex-direction: column;
                gap: 0.5vw;
                align-items: center;
                padding: 20px;
                background-color: #f8f9fa; /* Light background for contrast */
                border-radius: 8px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                margin-bottom: 10px;
            }
            .header-title {
                text-align: center;
                font-size: 4.2vw;
                margin: 0;
            }
            .header-subtitle {
                color: green;
                font-weight: bold;
            }
            .description {
                text-align: center;
                font-size: 1.2vw;
                color: #555;
                max-width: 600px; /* Limit width for readability */
            }
            hr {
                width: 80%; /* Horizontal rule width */
                border: 1px solid #ccc;
            }
        </style>
        <div class="header-container">
            <h1 class="header-title">"File Type Converter"<br /><span class="header-subtitle">Welcomes</span> You!</h1>
            <hr />
            <p class="description">
                Our service offers seamless conversions between various file formats, including:
                <ul>
                    <li>CSV to PDF</li>
                    <li>PDF to Image (JPEG, JPG, PNG)</li>
                    <li>Image (JPEG, JPG, PNG) to PDF</li>
                </ul>
                Experience fast, reliable, and user-friendly file conversion to meet all your needs!
            </p>
        </div>
    """, unsafe_allow_html=True)
# Handle DOCX to PDF conversion
if st.session_state.conversion_type == "docx_to_pdf":
    st.title("DOCX to PDF Converter")

    # File Upload Section
    uploaded_file = st.file_uploader("Upload your DOCX file", type=["docx"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Conversion Logic
        docx_output_pdf_path = os.path.splitext(temp_file_path)[0] + ".pdf"  # Change file extension to .pdf

        try:
            # Convert DOCX to PDF
            convert(temp_file_path, docx_output_pdf_path)
            st.session_state.output_pdf_path = docx_output_pdf_path  # Store path in session state
            st.success(f"File converted successfully to {docx_output_pdf_path}")

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

    # Display download button for the output PDF if conversion is done
    if st.session_state.output_pdf_path:
        with open(st.session_state.output_pdf_path, "rb") as f:
            st.download_button(label="Download PDF", data=f, file_name=os.path.basename(st.session_state.output_pdf_path))

# Handle PDF to DOCX conversion
if st.session_state.conversion_type == "pdf_to_docx":
    st.title("PDF to DOCX Converter")

    # File Upload Section
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Conversion Logic
        pdf_output_docx_path = os.path.splitext(temp_file_path)[0] + ".docx"  # Change file extension to .docx

        try:
            # Convert PDF to DOCX
            cv = Converter(temp_file_path)
            cv.convert(pdf_output_docx_path)
            cv.close()
            st.session_state.output_docx_path = pdf_output_docx_path  # Store path in session state
            st.success(f"File converted successfully to {pdf_output_docx_path}")

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

    # Display download button for the output DOCX if conversion is done
    if st.session_state.output_docx_path:
        with open(st.session_state.output_docx_path, "rb") as f:
            st.download_button(label="Download DOCX", data=f, file_name=os.path.basename(st.session_state.output_docx_path))

# Handle PDF to CSV conversion
if st.session_state.conversion_type == "pdf_to_csv":
    # Function to extract text from PDF
    def extract_text_from_pdf(pdf_path):
        text = ""
        with fitz.open(pdf_path) as pdf_document:
            for page in pdf_document:
                text += page.get_text()
        return text

    # Function to convert text to CSV
    def text_to_csv(text, output_csv_path):
        # Split the text into lines
        lines = text.splitlines()
        # Split lines into columns (assuming comma-separated values)
        data = [line.split(',') for line in lines]
        # Create a DataFrame
        df = pd.DataFrame(data)
        # Save to CSV
        df.to_csv(output_csv_path, index=False)

    st.title("PDF to CSV Converter")

    # File Upload Section
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Conversion Logic
        output_csv_path = os.path.splitext(temp_file_path)[0] + ".csv"  # Change file extension to .csv

        try:
            # Extract text from PDF
            extracted_text = extract_text_from_pdf(temp_file_path)
            # Convert extracted text to CSV
            text_to_csv(extracted_text, output_csv_path)
            st.session_state.output_csv_path = output_csv_path  # Store path in session state
            st.success(f"File converted successfully to {output_csv_path}")

            # Display download button for the output CSV
            with open(st.session_state.output_csv_path, "rb") as f:
                st.download_button(label="Download CSV", data=f, file_name=os.path.basename(st.session_state.output_csv_path))

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

# Handle CSV to PDF conversion
if st.session_state.conversion_type == "csv_to_pdf":
    # Function to convert CSV to PDF
    def convert_csv_to_pdf(csv_path, output_pdf_path):
        # Read the CSV file into a DataFrame
        df = pd.read_csv(csv_path)

        # Create a PDF
        c = canvas.Canvas(output_pdf_path, pagesize=letter)
        width, height = letter

        # Set starting position for the text
        x = 50
        y = height - 50

        # Add a title
        c.setFont("Helvetica-Bold", 16)
        c.drawString(x, y, "CSV Data")
        y -= 30  # Move down for the data

        # Set font for the data
        c.setFont("Helvetica", 12)

        # Write the DataFrame to the PDF
        for index, row in df.iterrows():
            text = ', '.join(map(str, row.values))  # Convert row to string
            c.drawString(x, y, text)
            y -= 15  # Move down for the next row

            # Check if we need a new page
            if y < 50:  # If y position is too low, create a new page
                c.showPage()
                c.setFont("Helvetica", 12)
                y = height - 50  # Reset y position

        # Save the PDF
        c.save()

    st.title("CSV to PDF Converter")

    # File Upload Section
    uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Conversion Logic
        output_pdf_path = os.path.splitext(temp_file_path)[0] + ".pdf"  # Change file extension to .pdf

        try:
            # Convert CSV to PDF
            convert_csv_to_pdf(temp_file_path, output_pdf_path)

            st.session_state.output_pdf_path_csv = output_pdf_path  # Store path in session state
            st.success(f"File converted successfully to {output_pdf_path}")

            # Display download button for the output PDF
            with open(st.session_state.output_pdf_path_csv, "rb") as f:
                st.download_button(label="Download PDF", data=f, file_name=os.path.basename(st.session_state.output_pdf_path_csv))

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

# Handle PDF to JPEG conversion
if st.session_state.conversion_type == "pdf_to_jpeg":
    # Function to convert PDF to JPEG
    def convert_pdf_to_jpeg(pdf_path, output_folder):
        # Convert PDF pages to images
        images = convert_from_path(pdf_path)
        image_paths = []
        
        # Save each page as a JPEG image
        for i, img in enumerate(images):
            output_image_path = os.path.join(output_folder, f"page_{i + 1}.jpg")
            img.save(output_image_path, 'JPEG')
            image_paths.append(output_image_path)
        
        return image_paths
    
    st.title("PDF to JPEG Converter")
    
    # File Upload Section
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])
    
    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Create a temporary folder to save the images
        temp_output_folder = tempfile.mkdtemp()
    
        try:
            # Convert PDF to JPEG
            st.session_state.jpeg_image_paths = convert_pdf_to_jpeg(temp_file_path, temp_output_folder)  # Store image paths in session state
            
            st.success("File converted successfully!")

            # Display and provide download buttons for each image
            for i, image_path in enumerate(st.session_state.jpeg_image_paths):
                st.image(image_path, caption=f"Page {i + 1}", use_column_width=True)
    
                # Display download button for each JPEG image
                with open(image_path, "rb") as img_file:
                    st.download_button(label=f"Download Page {i + 1} as JPEG", data=img_file, file_name=os.path.basename(image_path))
    
        except Exception as e:
            st.error(f"Error during conversion: {e}")
        
        # Clean up temporary files
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

# Handle PDF to JPG conversion
if st.session_state.conversion_type == "pdf_to_jpg":
    # Function to convert PDF to JPG
    def convert_pdf_to_jpg(pdf_path, output_folder):
        # Convert PDF pages to images
        images = convert_from_path(pdf_path)
        image_paths = []

        # Save each page as a JPG image
        for i, img in enumerate(images):
            output_image_path = os.path.join(output_folder, f"page_{i + 1}.jpg")
            img.save(output_image_path, 'JPEG')  # Explicitly saving as JPG
            image_paths.append(output_image_path)

        return image_paths

    # Streamlit App Interface
    st.title("PDF to JPG Converter")

    # File Upload Section
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Create a temporary folder to save the images
        temp_output_folder = tempfile.mkdtemp()

        try:
            # Convert PDF to JPG
            st.session_state.jpg_image_paths = convert_pdf_to_jpg(temp_file_path, temp_output_folder)  # Store image paths in session state

            st.success("File converted successfully!")

            # Display and provide download buttons for each image
            for i, image_path in enumerate(st.session_state.jpg_image_paths):
                st.image(image_path, caption=f"Page {i + 1}", use_column_width=True)

                # Display download button for each JPG image
                with open(image_path, "rb") as img_file:
                    st.download_button(label=f"Download Page {i + 1} as JPG", data=img_file, file_name=os.path.basename(image_path))

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary files
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

# Handle PDF to PNG conversion
if st.session_state.conversion_type == "pdf_to_png":
    # Function to convert PDF to PNG
    def convert_pdf_to_png(pdf_path, output_folder):
        # Convert PDF pages to images
        images = convert_from_path(pdf_path)
        image_paths = []

        # Save each page as a PNG image
        for i, img in enumerate(images):
            output_image_path = os.path.join(output_folder, f"page_{i + 1}.png")
            img.save(output_image_path, 'PNG')  # Explicitly saving as PNG
            image_paths.append(output_image_path)

        return image_paths

    # Streamlit App Interface
    st.title("PDF to PNG Converter")

    # File Upload Section
    uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

    if uploaded_file:
        # Save the uploaded file to a temporary location
        temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Create a temporary folder to save the images
        temp_output_folder = tempfile.mkdtemp()

        try:
            # Convert PDF to PNG
            st.session_state.png_image_paths = convert_pdf_to_png(temp_file_path, temp_output_folder)  # Store image paths in session state

            st.success("File converted successfully!")

            # Display and provide download buttons for each image
            for i, image_path in enumerate(st.session_state.png_image_paths):
                st.image(image_path, caption=f"Page {i + 1}", use_column_width=True)

                # Display download button for each PNG image
                with open(image_path, "rb") as img_file:
                    st.download_button(label=f"Download Page {i + 1} as PNG", data=img_file, file_name=os.path.basename(image_path))

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary files
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

# Handle JPEG to PDF conversion
if st.session_state.conversion_type == "jpeg_to_pdf":
    # Function to convert JPEG images to PDF
    def convert_images_to_pdf(image_paths, output_pdf_path):
        # Open the first image
        image_list = [Image.open(img).convert("RGB") for img in image_paths]

        # Save all images as a PDF
        if image_list:
            image_list[0].save(output_pdf_path, save_all=True, append_images=image_list[1:])

    # Streamlit App Interface
    st.title("JPEG to PDF Converter")

    # File Upload Section
    uploaded_files = st.file_uploader("Upload your JPEG images", type=["jpg", "jpeg"], accept_multiple_files=True)

    if uploaded_files:
        # Save uploaded images to a temporary location
        st.session_state.jpeg_image_paths = []  # Reset session state for new upload
        for uploaded_file in uploaded_files:
            temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.session_state.jpeg_image_paths.append(temp_file_path)  # Store paths in session state

        # Define the output PDF path
        output_pdf_path = os.path.join(tempfile.gettempdir(), "output.pdf")

        try:
            # Convert JPEG to PDF
            convert_images_to_pdf(st.session_state.jpeg_image_paths, output_pdf_path)

            st.success("Images converted successfully to PDF!")

            # Provide download button for the generated PDF
            with open(output_pdf_path, "rb") as f:
                st.download_button(label="Download PDF", data=f, file_name="converted_images.pdf")

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary files
        for img_path in st.session_state.jpeg_image_paths:
            if os.path.exists(img_path):
                os.remove(img_path)

# Handle JPG to PDF conversion
if st.session_state.conversion_type == "jpg_to_pdf":
    # Function to convert JPG images to PDF
    def convert_images_to_pdf(image_paths, output_pdf_path):
        # Open the first image and convert to RGB
        image_list = [Image.open(img).convert("RGB") for img in image_paths]

        # Save all images as a PDF
        if image_list:
            image_list[0].save(output_pdf_path, save_all=True, append_images=image_list[1:])

    # Streamlit App Interface
    st.title("JPG to PDF Converter")

    # File Upload Section
    uploaded_files = st.file_uploader("Upload your JPG images", type=["jpg", "jpeg"], accept_multiple_files=True)

    if uploaded_files:
        # Save uploaded images to a temporary location
        st.session_state.jpg_image_paths = []  # Reset session state for new upload
        for uploaded_file in uploaded_files:
            temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.session_state.jpg_image_paths.append(temp_file_path)  # Store paths in session state

        # Define the output PDF path
        output_pdf_path = os.path.join(tempfile.gettempdir(), "output.pdf")

        try:
            # Convert JPG to PDF
            convert_images_to_pdf(st.session_state.jpg_image_paths, output_pdf_path)

            st.success("Images converted successfully to PDF!")

            # Provide download button for the generated PDF
            with open(output_pdf_path, "rb") as f:
                st.download_button(label="Download PDF", data=f, file_name="converted_images.pdf")

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary files
        for img_path in st.session_state.jpg_image_paths:
            if os.path.exists(img_path):
                os.remove(img_path)

# Handle PNG to PDF conversion
if st.session_state.conversion_type == "png_to_pdf":
    # Function to convert PNG images to PDF
    def convert_images_to_pdf(image_paths, output_pdf_path):
        # Open each image and convert to RGB (PDF doesn't support alpha channel)
        image_list = [Image.open(img).convert("RGB") for img in image_paths]

        # Save all images as a single PDF
        if image_list:
            image_list[0].save(output_pdf_path, save_all=True, append_images=image_list[1:])

    # Streamlit App Interface
    st.title("PNG to PDF Converter")

    # File Upload Section
    uploaded_files = st.file_uploader("Upload your PNG images", type=["png"], accept_multiple_files=True)

    if uploaded_files:
        # Save uploaded images to a temporary location
        st.session_state.png_image_paths = []  # Reset session state for new upload
        for uploaded_file in uploaded_files:
            temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.session_state.png_image_paths.append(temp_file_path)  # Store paths in session state

        # Define the output PDF path
        output_pdf_path = os.path.join(tempfile.gettempdir(), "output.pdf")

        try:
            # Convert PNG to PDF
            convert_images_to_pdf(st.session_state.png_image_paths, output_pdf_path)

            st.success("Images converted successfully to PDF!")

            # Provide download button for the generated PDF
            with open(output_pdf_path, "rb") as f:
                st.download_button(label="Download PDF", data=f, file_name="converted_images.pdf")

        except Exception as e:
            st.error(f"Error during conversion: {e}")

        # Clean up temporary files
        for img_path in st.session_state.png_image_paths:
            if os.path.exists(img_path):
                os.remove(img_path)

# Sidebar Content
st.sidebar.markdown("""
    <div style='text-align: center; display: flex; flex-direction: column; gap: -100px;'>
        <p style='font-size: 1vw;'>Developed by Rayyan Ashraf</p>
        <p style='font-size: 1vw;'>Copyright @ 2024</p>
        <p style='font-size: 1vw;'>Developed in Python 3.12.5</p>
    </div>
""", unsafe_allow_html=True)
