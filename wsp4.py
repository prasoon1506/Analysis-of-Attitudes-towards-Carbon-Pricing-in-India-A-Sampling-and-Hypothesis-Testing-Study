import os
import streamlit as st
import pandas as pd
import openpyxl
from matplotlib.path import Path
from matplotlib.patches import PathPatch
from io import BytesIO
import matplotlib.ticker as mticker
import base64
from pypdf import PdfReader, PdfWriter
from docx import Document
from reportlab.pdfgen.canvas import Canvas 
from pdf2docx import Converter
from docx2pdf import convert
import img2pdf
from PIL import Image
import PyPDF2
import fitz  # PyMuPDF
import tempfile
import shutil
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
from datetime import datetime
from streamlit_option_menu import option_menu
from matplotlib.patches import Rectangle
import matplotlib.backends.backend_pdf
from scipy import stats
from statsmodels.tsa.arima.model import ARIMA
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from streamlit_lottie import st_lottie
import time
import hashlib
import secrets
from streamlit_cookies_manager import EncryptedCookieManager
import json
import requests
from openpyxl.utils import get_column_letter
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
import math
import seaborn as sns
import xgboost as xgb
import plotly.graph_objs as go
import time
from collections import OrderedDict
import re
import plotly.graph_objects as go
import plotly.express as px
from concurrent.futures import ThreadPoolExecutor
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter, legal, landscape
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch, cm
from reportlab.platypus import Image as ReportLabImage
from reportlab.graphics.shapes import Line, Drawing
from reportlab.lib.colors import Color, HexColor
import io
import tempfile
def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
import statsmodels.api as sm
from statsmodels.stats.diagnostic import het_breuschpagan, acorr_ljungbox
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.tsa.stattools import adfuller
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, PolynomialFeatures
from sklearn.linear_model import LinearRegression, Ridge, Lasso
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor
from sklearn.svm import SVR
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import mean_squared_error, r2_score
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.stats import jarque_bera, kurtosis, skew
from statsmodels.stats.stattools import omni_normtest
def process_pdf(input_pdf, operations):
    from PyPDF2 import PdfReader, PdfWriter
    from io import BytesIO
    writer = PdfWriter()
    reader = PdfReader(input_pdf)
    if "extract" in operations:
        selected_pages = operations["extract"]["pages"]  # List of selected page numbers
        # Add only the selected pages
        for page_num in selected_pages:
            if 0 <= page_num - 1 < len(reader.pages):  # Convert to 0-based index and check bounds
                writer.add_page(reader.pages[page_num - 1])
    else:
        # Add all pages from input PDF
        for page in reader.pages:
            writer.add_page(page)
    if "merge" in operations and operations["merge"]["files"]:
        # Add pages from additional PDFs
        for additional_pdf in operations["merge"]["files"]:
            merge_reader = PdfReader(additional_pdf)
            for page in merge_reader.pages:
                writer.add_page(page)
    if len(writer.pages) == 0:
        return BytesIO(input_pdf.read())
    pdf_width = float(writer.pages[0].mediabox.width)
    pdf_height = float(writer.pages[0].mediabox.height)
    transformed_writer = PdfWriter()
    for i in range(len(writer.pages)):
        page = writer.pages[i]
        if "resize" in operations:
            scale = operations["resize"]["scale"] / 100
            page.scale(scale, scale)
        if "crop" in operations:
            left = operations["crop"]["left"] * pdf_width / 100
            bottom = operations["crop"]["bottom"] * pdf_height / 100
            right = operations["crop"]["right"] * pdf_width / 100
            top = operations["crop"]["top"] * pdf_height / 100
            page.cropbox.lower_left = (left, bottom)
            page.cropbox.upper_right = (right, top)  
        if "rotate" in operations:
            angle = operations["rotate"]["angle"]
            page.rotate(angle)
        transformed_writer.add_page(page)
    output = BytesIO()
    transformed_writer.write(output)
    return output
def add_watermark(pdf_writer, watermark_options):
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.colors import Color
    from pypdf import PdfReader
    from PIL import Image
    from io import BytesIO
    watermark_buffer = BytesIO()
    c = Canvas(watermark_buffer)
    first_page = pdf_writer.pages[0]
    page_width = float(first_page.mediabox.width)
    page_height = float(first_page.mediabox.height)  
    if watermark_options["type"] == "text":
        text = watermark_options["text"]
        color = watermark_options["color"]
        font_size = watermark_options["size"]
        opacity = watermark_options["opacity"]
        angle = watermark_options["angle"]
        position = watermark_options["position"]
        r = int(color[1:3], 16) / 255
        g = int(color[3:5], 16) / 255
        b = int(color[5:7], 16) / 255
        c.setFillColor(Color(r, g, b, alpha=opacity))
        c.setFont("Helvetica", font_size)
        if position == "center":
            x, y = page_width/2, page_height/2
        elif position == "top-left":
            x, y = 50, page_height-50
        elif position == "top-right":
            x, y = page_width-50, page_height-50
        elif position == "bottom-left":
            x, y = 50, 50
        elif position == "bottom-right":
            x, y = page_width-50, 50
        c.saveState()
        c.translate(x, y)
        c.rotate(angle)
        c.drawString(-len(text)*font_size/4, 0, text)
        c.restoreState()  
    else:
        image = Image.open(watermark_options["image"])
        opacity = watermark_options["opacity"]
        angle = watermark_options["angle"]
        position = watermark_options["position"]
        size = watermark_options["size"]  # percentage of page width
        img_width = page_width * size / 100
        img_height = img_width * image.height / image.width
        if position == "center":
            x, y = (page_width-img_width)/2, (page_height-img_height)/2
        elif position == "top-left":
            x, y = 0, page_height-img_height
        elif position == "top-right":
            x, y = page_width-img_width, page_height-img_height
        elif position == "bottom-left":
            x, y = 0, 0
        elif position == "bottom-right":
            x, y = page_width-img_width, 0
        img_buffer = BytesIO()
        if image.mode != 'RGBA':
            image = image.convert('RGBA')
        image.putalpha(int(opacity * 255))
        image.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        c.saveState()
        c.translate(x + img_width/2, y + img_height/2)
        c.rotate(angle)
        c.translate(-img_width/2, -img_height/2)
        c.drawImage(ImageReader(img_buffer), 0, 0, width=img_width, height=img_height)
        c.restoreState()
    c.save()
    watermark_buffer.seek(0)
    watermark_pdf = PdfReader(watermark_buffer)
    selected_pages = watermark_options.get("pages", "all")
    for i, page in enumerate(pdf_writer.pages):
        if selected_pages == "all" or (i+1) in selected_pages:
            page.merge_page(watermark_pdf.pages[0])
    return pdf_writer
def get_pdf_preview(pdf_file, page_num=0):
    import fitz
    from PIL import Image
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    page = doc[page_num]
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img
def get_image_size_metrics(original_image_bytes, processed_image_bytes):
    original_size = len(original_image_bytes) / 1024  # KB
    processed_size = len(processed_image_bytes) / 1024  # KB
    size_change = ((original_size - processed_size) / original_size) * 100  
    return {
        'original_size': original_size,
        'processed_size': processed_size,
        'size_change': size_change
    }
def process_image(image, operations):
    from PIL import Image, ImageEnhance
    if "resize" in operations:
        width = operations["resize"]["width"]
        height = operations["resize"]["height"]
        image = image.resize((width, height), Image.Resampling.LANCZOS)
    if "compress" in operations:
        quality = operations["compress"]["quality"]
        # Return image for JPEG saving with quality
        return image, quality
    if "crop" in operations:
        left = operations["crop"]["left"]
        top = operations["crop"]["top"]
        right = operations["crop"]["right"]
        bottom = operations["crop"]["bottom"]
        image = image.crop((left, top, right, bottom))
    if "rotate" in operations:
        angle = operations["rotate"]["angle"]
        image = image.rotate(angle, expand=True)
    if "brightness" in operations:
        factor = operations["brightness"]["factor"]
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(factor)
    if "contrast" in operations:
        factor = operations["contrast"]["factor"]
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(factor)
    return image, None
def convert_uploadedfile_to_image(uploaded_file):
    """Convert Streamlit UploadedFile to a temporary file path"""
    if uploaded_file is None:
        return None
    
    # Create a temporary file
    temp_dir = tempfile.mkdtemp()
    temp_file_path = os.path.join(temp_dir, uploaded_file.name)
    
    # Save the uploaded file to the temporary path
    with open(temp_file_path, 'wb') as f:
        f.write(uploaded_file.getvalue())
    
    return temp_file_path
# Professional Templates Dictionary
TEMPLATES = {
    "Classic Professional": {
        "background_type": "Color",
        "background_color": "#FFFFFF",
        "design_elements": ["Border"],
        "border_color": "#000000",
        "border_width": 1,
        "title_font": "Helvetica-Bold",
        "title_color": "#000000",
        "accent_color": "#4A4A4A",
        "layout": "centered"
    },
    "Modern Minimal": {
        "background_type": "Color",
        "background_color": "#FFFFFF",
        "design_elements": ["Accent Bar"],
        "accent_color": "#2C3E50",
        "title_font": "Helvetica",
        "title_color": "#2C3E50",
        "layout": "left-aligned"
    },
    "Corporate Blue": {
        "background_type": "Gradient",
        "gradient_start": "#E8F0FE",
        "gradient_end": "#FFFFFF",
        "design_elements": ["Corner Lines", "Header Bar"],
        "accent_color": "#1B4F72",
        "title_font": "Helvetica-Bold",
        "title_color": "#1B4F72",
        "layout": "centered"
    },
    "Creative Bold": {
        "background_type": "Color",
        "background_color": "#FFFFFF",
        "design_elements": ["Diagonal Lines", "Side Bar"],
        "accent_color": "#FF5733",
        "title_font": "Helvetica-Bold",
        "title_color": "#2C3E50",
        "layout": "asymmetric"
    },
    "Executive Elite": {
        "background_type": "Color",
        "background_color": "#F5F5F5",
        "design_elements": ["Gold Accents", "Double Border"],
        "accent_color": "#D4AF37",
        "border_color": "#000000",
        "title_font": "Times-Bold",
        "title_color": "#000000",
        "layout": "centered"
    }
}
def register_custom_fonts():
    """Register additional fonts for use in the PDF"""
    try:
        # Add more fonts as needed
        custom_fonts = [
            ("Roboto-Regular.ttf", "Roboto"),
            ("Montserrat-Regular.ttf", "Montserrat"),
            ("OpenSans-Regular.ttf", "OpenSans")
        ]
        for font_file, font_name in custom_fonts:
            if not font_name in pdfmetrics.getRegisteredFontNames():
                font_path = f"fonts/{font_file}"
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont(font_name, font_path))
    except Exception as e:
        st.warning(f"Some custom fonts couldn't be loaded: {str(e)}")
def draw_design_elements(c, options, width, height):
    margin = cm
    if "Border" in options["design_elements"]:
        c.setStrokeColor(options["border_color"])
        c.setLineWidth(options["border_width"])
        c.rect(margin, margin, width - 2*margin, height - 2*margin)
    if "Double Border" in options["design_elements"]:
        c.setStrokeColor(options["border_color"])
        c.setLineWidth(options["border_width"])
        # Outer border
        c.rect(margin, margin, width - 2*margin, height - 2*margin)
        # Inner border
        inner_margin = margin + 0.5*cm
        c.rect(inner_margin, inner_margin, width - 2*inner_margin, height - 2*inner_margin)
    if "Corner Lines" in options["design_elements"]:
        c.setStrokeColor(options["accent_color"])
        c.setLineWidth(2)
        corner_size = 3*cm
        # Draw sophisticated corner lines
        for x, y in [(margin, height-margin), (width-margin, height-margin),
                     (margin, margin), (width-margin, margin)]:
            c.saveState()
            c.translate(x, y)
            if y > height/2:
                c.rotate(180)
            c.lines([(0, 0, corner_size, 0), (0, 0, 0, -corner_size)])
            c.restoreState()
    if "Diagonal Lines" in options["design_elements"]:
        c.setStrokeColor(options["accent_color"])
        c.setLineWidth(1)
        spacing = 1*cm
        for i in range(int(height/(2*spacing))):
            y = i * 2*spacing
            c.line(0, y, 2*cm, y + 2*cm)
    if "Side Bar" in options["design_elements"]:
        c.setFillColor(options["accent_color"])
        c.rect(0, 0, 2*cm, height, fill=1)
    if "Header Bar" in options["design_elements"]:
        c.setFillColor(options["accent_color"])
        c.rect(0, height-3*cm, width, 3*cm, fill=1)
    if "Accent Bar" in options["design_elements"]:
        c.setFillColor(options["accent_color"])
        bar_width = 0.5*cm
        c.rect(margin, height/2, width - 2*margin, bar_width, fill=1)
    if "Gold Accents" in options["design_elements"]:
        c.setStrokeColor(options["accent_color"])
        c.setLineWidth(1)
        pattern_size = 1*cm
        for i in range(4):
            x = margin + i * pattern_size
            c.line(x, height-margin, x + pattern_size, height-margin-pattern_size)
            c.line(x, margin, x + pattern_size, margin+pattern_size)
def draw_watermark(c, options, width, height):
    if options.get("watermark_text"):
        c.saveState()
        c.translate(width/2, height/2)
        c.rotate(45)
        c.setFont(options["watermark_font"], options["watermark_size"])
        c.setFillColor(colors.Color(0, 0, 0, alpha=0.1))
        c.drawCentredString(0, 0, options["watermark_text"])
        c.restoreState()
def create_front_page(options):
    buffer = io.BytesIO()
    page_size = {
        "A4": A4,
        "A4 Landscape": landscape(A4),
        "Letter": letter,
        "Letter Landscape": landscape(letter),
        "Legal": legal
    }[options["page_size"]]
    c = canvas.Canvas(buffer, pagesize=page_size)
    width, height = page_size
    if options.get("template"):
        template = TEMPLATES[options["template"]]
        options = {**template, **options}  # Merge with user options, user options take precedence
    if options["background_type"] == "Color":
        c.setFillColor(options["background_color"])
        c.rect(0, 0, width, height, fill=True)
    elif options["background_type"] == "Gradient":
        steps = 100
        for i in range(steps):
            r = options["gradient_start"].red + (options["gradient_end"].red - options["gradient_start"].red) * i / steps
            g = options["gradient_start"].green + (options["gradient_end"].green - options["gradient_start"].green) * i / steps
            b = options["gradient_start"].blue + (options["gradient_end"].blue - options["gradient_start"].blue) * i / steps
            c.setFillColor((r, g, b))
            c.rect(0, height * i / steps, width, height / steps, fill=True)
    elif options["background_type"] == "Pattern":
        pattern_size = 1*cm
        c.setStrokeColor(colors.Color(0, 0, 0, alpha=0.1))
        for x in range(0, int(width), int(pattern_size)):
            for y in range(0, int(height), int(pattern_size)):
                if (x + y) % (2 * int(pattern_size)) == 0:
                    c.rect(x, y, pattern_size, pattern_size, fill=True)
    draw_design_elements(c, options, width, height)
    if options.get("watermark_text"):
        draw_watermark(c, options, width, height)
    if options.get("logo"):
        try:
            logo_path = convert_uploadedfile_to_image(options["logo"])
            if logo_path:
                logo_img = Image.open(logo_path)
                aspect = logo_img.height / logo_img.width
                logo_width = options["logo_width"]
                logo_height = logo_width * aspect
                if options["layout"] == "centered":
                    x = (width - logo_width) / 2
                    y = height - logo_height - 3*cm
                elif options["layout"] == "left-aligned":
                    x = 3*cm
                    y = height - logo_height - 3*cm
                elif options["layout"] == "asymmetric":
                    x = width - logo_width - 3*cm
                    y = height - logo_height - 3*cm
                c.drawImage(logo_path, x, y, width=logo_width, height=logo_height)
                os.unlink(logo_path)
                os.rmdir(os.path.dirname(logo_path))
        except Exception as e:
            st.error(f"Error processing logo: {str(e)}")
    c.setFont(options["title_font"], options["title_size"])
    c.setFillColor(options["title_color"])
    title_lines = options["title"].split('\n')
    if options["layout"] == "centered":
        title_height = (height + len(title_lines) * options["title_size"]) / 2
    elif options["layout"] == "left-aligned":
        title_height = height - 5*cm
    else:  # asymmetric
        title_height = (height + len(title_lines) * options["title_size"]) / 1.5
    for line in title_lines:
        title_width = c.stringWidth(line, options["title_font"], options["title_size"])
        if options["layout"] == "centered":
            x = (width - title_width) / 2
        elif options["layout"] == "left-aligned":
            x = 3*cm
        else:  # asymmetric
            x = width - title_width - 3*cm
        c.drawString(x, title_height, line)
        title_height -= options["title_size"] * 1.2
    if options["subtitle"]:
        c.setFont(options["subtitle_font"], options["subtitle_size"])
        c.setFillColor(options["subtitle_color"])
        subtitle_width = c.stringWidth(options["subtitle"], options["subtitle_font"], options["subtitle_size"])
        if options["layout"] == "centered":
            x = (width - subtitle_width) / 2
        elif options["layout"] == "left-aligned":
            x = 3*cm
        else:  # asymmetric
            x = width - subtitle_width - 3*cm
        c.drawString(x, title_height - cm, options["subtitle"])
    y_position = title_height - 4*cm
    for text_block in options["text_blocks"]:
        if text_block["text"]:
            c.setFont(text_block["font"], text_block["size"])
            c.setFillColor(text_block["color"])
            lines = text_block["text"].split('\n')
            for line in lines:
                text_width = c.stringWidth(line, text_block["font"], text_block["size"])
                if options["layout"] == "centered":
                    x = (width - text_width) / 2
                elif options["layout"] == "left-aligned":
                    x = 3*cm
                else:  # asymmetric
                    x = width - text_width - 3*cm
                c.drawString(x, y_position, line)
                y_position -= text_block["size"] * 1.5
    if options["show_date"] or options["footer_text"]:
        footer_font = options.get("footer_font", "Helvetica")
        footer_size = options.get("footer_size", 10)
        c.setFont(footer_font, footer_size)
        c.setFillColor(options.get("footer_color", colors.black))
        footer_elements = []
        if options["show_date"]:
            date_format = options.get("date_format", "%B %d, %Y")
            footer_elements.append(datetime.now().strftime(date_format))
        if options["footer_text"]:
            footer_elements.append(options["footer_text"])
        footer_text = " | ".join(footer_elements)
        footer_width = c.stringWidth(footer_text, footer_font, footer_size)
        if options["layout"] == "centered":
            x = (width - footer_width) / 2
        elif options["layout"] == "left-aligned":
            x = 3*cm
        else:  # asymmetric
            x = width - footer_width - 3*cm
        c.drawString(x, 2*cm, footer_text)
    c.save()
    buffer.seek(0)
    return buffer
def front_page_creator():
    st.header("📄 Professional Front Page Creator")
    st.subheader("Choose a Template")
    template = st.selectbox("Select a Template", 
        ["Custom"] + list(TEMPLATES.keys()),
        help="Choose a pre-designed template or create your own custom design")
    with st.expander("Preview Template", expanded=False):
        st.write("Template Preview would appear here")
    with st.container():
            # Basic Settings
            st.subheader("Basic Settings")
            col1, col2 = st.columns(2)
            with col1:
                page_size = st.selectbox(
                    "Page Size", 
                    ["A4", "A4 Landscape", "Letter", "Letter Landscape", "Legal"],
                    help="Choose the size and orientation of your front page")                
                title = st.text_area(
                    "Title",
                    placeholder="Enter title (can be multiple lines)\nUse new lines for multi-line titles",
                    help="Main title of your front page. Use new lines for multiple lines of text"
                )
                subtitle = st.text_input(
                    "Subtitle",
                    placeholder="Enter subtitle (optional)",
                    help="Optional subtitle that appears below the main title"
                )
            with col2:
                title_font = st.selectbox(
                    "Title Font",
                    ["Helvetica", "Helvetica-Bold", "Times-Roman", "Times-Bold", 
                     "Courier", "Courier-Bold", "Roboto", "Montserrat", "OpenSans"],
                    help="Choose the font for your title"
                )
                title_size = st.slider(
                    "Title Size",
                    20, 72, 48,
                    help="Adjust the size of your title text"
                )
                title_color = st.color_picker(
                    "Title Color",
                    "#000000",
                    help="Choose the color for your title"
                )
            st.subheader("Layout Settings")
            col3, col4 = st.columns(2)
            with col3:
                layout_style = st.selectbox(
                    "Layout Style",
                    ["centered", "left-aligned", "asymmetric"],
                    help="Choose how your content is aligned on the page")
                content_spacing = st.slider(
                    "Content Spacing",
                    1.0, 3.0, 1.5,
                    0.1,
                    help="Adjust the spacing between content elements")
            with col4:
                margins = st.slider(
                    "Page Margins (cm)",
                    1.0, 5.0, 2.5,
                    0.5,
                    help="Adjust the margins around your content")
            st.subheader("Background Settings")
            background_type = st.radio(
                "Background Type",
                ["Color", "Gradient", "Pattern", "None"],
                help="Choose the type of background for your front page")
            if background_type == "Color":
                background_color = st.color_picker(
                    "Background Color",
                    "#FFFFFF",
                    help="Choose a solid color for your background")
            elif background_type == "Gradient":
                col5, col6 = st.columns(2)
                with col5:
                    gradient_start = st.color_picker(
                        "Gradient Start Color",
                        "#FFFFFF",
                        help="Choose the starting color for your gradient")
                    gradient_direction = st.selectbox(
                        "Gradient Direction",
                        ["Top to Bottom", "Left to Right", "Diagonal"],
                        help="Choose the direction of your gradient")
                with col6:
                    gradient_end = st.color_picker(
                        "Gradient End Color",
                        "#E0E0E0",
                        help="Choose the ending color for your gradient")
            elif background_type == "Pattern":
                col7, col8 = st.columns(2)
                with col7:
                    pattern_type = st.selectbox(
                        "Pattern Type",
                        ["Dots", "Lines", "Grid", "Chevron"],
                        help="Choose the type of pattern")
                    pattern_color = st.color_picker(
                        "Pattern Color",
                        "#E0E0E0",
                        help="Choose the color for your pattern")
                with col8:
                    pattern_opacity = st.slider(
                        "Pattern Opacity",
                        0.0, 1.0, 0.1,
                        0.1,
                        help="Adjust the opacity of the pattern")
                    pattern_size = st.slider(
                        "Pattern Size",
                        0.5, 3.0, 1.0,
                        0.1,
                        help="Adjust the size of the pattern elements")
            st.subheader("Logo/Image Settings")
            logo = st.file_uploader(
                "Upload Logo/Image",
                type=["png", "jpg", "jpeg"],
                help="Upload your organization's logo or an image")
            if logo:
                col9, col10 = st.columns(2)
                with col9:
                    logo_width = st.slider(
                        "Logo Width",
                        50, 400, 200,
                        help="Adjust the width of your logo")
                    logo_opacity = st.slider(
                        "Logo Opacity",
                        0.1, 1.0, 1.0,
                        0.1,
                        help="Adjust the opacity of your logo")
                with col10:
                    logo_position = st.selectbox(
                        "Logo Position",
                        ["Top Center", "Top Left", "Top Right", "Bottom Center", "Bottom Left", "Bottom Right"],
                        help="Choose where to place your logo")
                    logo_padding = st.slider(
                        "Logo Padding (cm)",
                        0.5, 5.0, 2.0,
                        0.5,
                        help="Adjust the space around your logo")
            st.subheader("Design Elements")
            design_elements = st.multiselect(
                "Add Design Elements",
                ["Border", "Double Border", "Corner Lines", "Diagonal Lines", 
                 "Side Bar", "Header Bar", "Accent Bar", "Gold Accents"],
                default=["Border"],
                help="Choose decorative elements to enhance your design")
            if any(design_elements):
                col11, col12 = st.columns(2)
                with col11:
                    accent_color = st.color_picker(
                        "Accent Color",
                        "#000000",
                        help="Choose the color for decorative elements")
                if "Border" in design_elements or "Double Border" in design_elements:
                    with col12:
                        border_color = st.color_picker(
                            "Border Color",
                            "#000000",
                            help="Choose the color for the border")
                        border_width = st.slider(
                            "Border Width",
                            0.5, 5.0, 1.0,
                            0.5,
                            help="Adjust the thickness of the border")
            st.subheader("Additional Text Blocks")
            num_blocks = st.number_input("Number of Additional Text Blocks", 0, 5, 0)
            text_blocks = []
            for i in range(num_blocks):
             st.markdown(f"#### Text Block {i+1}")
             col13, col14 = st.columns(2)
             with col13:
                block_text = st.text_input(f"Text for Block {i+1}")
                block_font = st.selectbox(f"Font for Block {i+1}", ["Helvetica", "Times-Roman", "Courier"])
             with col14:
                block_size = st.slider(f"Size for Block {i+1}", 8, 36, 12)
                block_color = st.color_picker(f"Color for Block {i+1}", "#000000")
             text_blocks.append({
                "text": block_text,
                "font": block_font,
                "size": block_size,
                "color": block_color
            })
            st.subheader("Watermark Settings")
            add_watermark = st.checkbox(
                "Add Watermark",
                help="Add a watermark to your front page"
            )
            if add_watermark:
                col15, col16 = st.columns(2)
                with col15:
                    watermark_text = st.text_input(
                        "Watermark Text",
                        placeholder="Enter watermark text",
                        help="Text to use as watermark"
                    )
                    watermark_font = st.selectbox(
                        "Watermark Font",
                        ["Helvetica", "Times-Roman", "Courier"],
                        help="Choose the font for your watermark"
                    )
                with col16:
                    watermark_size = st.slider(
                        "Watermark Size",
                        20, 100, 60,
                        help="Adjust the size of your watermark"
                    )
                    watermark_opacity = st.slider(
                        "Watermark Opacity",
                        0.0, 1.0, 0.1,
                        0.1,
                        help="Adjust the opacity of your watermark"
                    )
            
            # Footer Options
            st.subheader("Footer Options")
            col15, col16 = st.columns(2)
            with col15:
                show_date = st.checkbox(
                    "Show Date",
                    help="Include the current date in the footer"
                )
                if show_date:
                    date_format = st.selectbox(
                        "Date Format",
                        ["%B %d, %Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"],
                        help="Choose how the date should be displayed"
                    )
            with col16:
                footer_text = st.text_input(
                    "Custom Footer Text",
                    placeholder="Enter custom footer text",
                    help="Add custom text to the footer"
                )
                footer_alignment = st.selectbox(
                    "Footer Alignment",
                    ["Center", "Left", "Right"],
                    help="Choose how the footer is aligned"
                )
            
            # Generate Button
            if st.button("Generate Front Page", help="Click to create your front page"):
                if not title:
                    st.error("Please enter a title for your front page.")
                    return
                    
                try:
                    options = {
                        "template": template if template != "Custom" else None,
                        "page_size": page_size,
                        "title": title,
                        "subtitle": subtitle,
                        "title_font": title_font,
                        "title_size": title_size,
                        "title_color": colors.HexColor(title_color),
                        "subtitle_font": title_font,
                        "subtitle_size": int(title_size * 0.6),
                        "subtitle_color": colors.HexColor(title_color),
                        "layout": layout_style,
                        "content_spacing": content_spacing,
                        "margins": margins,
                        "background_type": background_type,
                        "background_color": colors.HexColor(background_color) if background_type == "Color" else None,
                        "gradient_start": colors.HexColor(gradient_start) if background_type == "Gradient" else None,
                        "gradient_end": colors.HexColor(gradient_end) if background_type == "Gradient" else None,
                        "gradient_direction": gradient_direction if background_type == "Gradient" else None,
                        "pattern_type": pattern_type if background_type == "Pattern" else None,
                        "pattern_color": colors.HexColor(pattern_color) if background_type == "Pattern" else None,
                        "pattern_opacity": pattern_opacity if background_type == "Pattern" else None,
                        "pattern_size": pattern_size if background_type == "Pattern" else None,
                        "logo": logo,
                        "logo_width": logo_width if logo else None,
                        "logo_position": logo_position if logo else None,
                        "logo_opacity": logo_opacity if logo else None,
                        "logo_padding": logo_padding if logo else None,
                        "design_elements": design_elements,
                        "accent_color": colors.HexColor(accent_color) if any(design_elements) else None,
                        "border_color": colors.HexColor(border_color) if "Border" in design_elements or "Double Border" in design_elements else None,
                        "border_width": border_width if "Border" in design_elements or "Double Border" in design_elements else None,
                        "watermark_text": watermark_text if add_watermark else None,
                        "watermark_font": watermark_font if add_watermark else None,
                        "watermark_size": watermark_size if add_watermark else None,
                        "watermark_opacity": watermark_opacity if add_watermark else None,
                        "show_date": show_date,
                        "date_format": date_format if show_date else "%B %d, %Y",
                        "text_blocks": text_blocks,
                        "footer_text": footer_text,
                        "footer_alignment": footer_alignment,
                    }
                    
                    pdf_buffer = create_front_page(options)
                    
                    # Add download button with custom filename
                    filename = f"{title.split()[0].lower()}_front_page.pdf"
                    st.download_button(
                        label="📥 Download Front Page",
                        data=pdf_buffer,
                        file_name=filename,
                        mime="application/pdf",
                        help="Download your generated front page as a PDF"
                    )
                    
                    # Display success message with tips
                    st.success("✨ PDF generated successfully! Click the download button above to save your front page.")
                    
                    # Show preview tip
                    st.info("💡 Tip: After downloading, you may want to preview the PDF to ensure everything looks perfect.")
                    
                except Exception as e:
                    st.error(f"Error generating front page: {str(e)}")
                    st.info("Please try adjusting your settings or contact support if the problem persists.")
def excel_editor_and_analyzer():
    st.title("🧩 Advanced Excel Editor, File Converter and Data Analyzer")
    tab1, tab2, tab3, tab4 = st.tabs([
        "Excel Editor",
        "File Converter", 
        "Data Analyzer",
        "Front Page Creator"
    ])
    
    with tab1:
        excel_editor()
    with tab2:
        file_converter()
    with tab3:
        data_analyzer()
    with tab4:
        front_page_creator()
def file_converter():
    st.header("🔄 Universal File Converter")
    
    # Add custom CSS for better styling
    st.markdown("""
        <style>
        .converter-card {
            background-color: #f8f9fa;
            padding: 1.5rem;
            border-radius: 0.5rem;
            margin: 1rem 0;
            border: 1px solid #e9ecef;
        }
        .stButton>button {
            width: 100%;
            margin-top: 1rem;
        }
        .success-message {
            color: #28a745;
            padding: 0.75rem;
            border-radius: 0.25rem;
            margin: 1rem 0;
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
        }
        .info-message {
            color: #0c5460;
            padding: 0.75rem;
            border-radius: 0.25rem;
            margin: 1rem 0;
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
        }
        .conversion-stats {
            padding: 1rem;
            background-color: #fff;
            border-radius: 0.25rem;
            box-shadow: 0 0.125rem 0.25rem rgba(0,0,0,0.075);
            margin-top: 1rem;
        }
        </style>
    """, unsafe_allow_html=True)

    converter_type = st.selectbox(
        "Select Conversion Type",
        [
            "Excel ↔️ CSV Converter",
            "Word ↔️ PDF Converter",
            "Image to PDF Converter",
            "PDF Editor",
            "Image Editor"
        ]
    )

    # Excel ↔️ CSV Converter
    if converter_type == "Excel ↔️ CSV Converter":
        st.markdown("### Excel ↔️ CSV Converter")
        
        conversion_direction = st.radio(
            "Select conversion direction:",
            ["CSV to Excel", "Excel to CSV"],
            horizontal=True
        )

        if conversion_direction == "CSV to Excel":
            with st.container():
                st.markdown('<div class="converter-card">', unsafe_allow_html=True)
                
                uploaded_file = st.file_uploader("Upload CSV file", type="csv", key="csv_to_excel")
                
                if uploaded_file is not None:
                    try:
                        col1, col2 = st.columns(2)
                        with col1:
                            separator = st.selectbox(
                                "Select delimiter",
                                options=[",", ";", "|", "\t"],
                                index=0
                            )
                        with col2:
                            encoding = st.selectbox(
                                "Select encoding",
                                options=["utf-8", "iso-8859-1", "cp1252"],
                                index=0
                            )

                        df = pd.read_csv(uploaded_file, sep=separator, encoding=encoding)
                        
                        st.markdown("#### Preview")
                        st.dataframe(df.head(), use_container_width=True)
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Rows", df.shape[0])
                        with col2:
                            st.metric("Columns", df.shape[1])
                        with col3:
                            st.metric("Size", f"{uploaded_file.size / 1024:.2f} KB")

                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)
                        excel_data = output.getvalue()
                        
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_data,
                            file_name=f"{uploaded_file.name.split('.')[0]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    except Exception as e:
                        st.error(f"Error: {str(e)}")
                
                st.markdown('</div>', unsafe_allow_html=True)

        else:  # Excel to CSV
            with st.container():
                st.markdown('<div class="converter-card">', unsafe_allow_html=True)
                
                uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], key="excel_to_csv")
                
                if uploaded_file is not None:
                    try:
                        df = pd.read_excel(uploaded_file)
                        
                        st.markdown("#### Preview")
                        st.dataframe(df.head(), use_container_width=True)
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Rows", df.shape[0])
                        with col2:
                            st.metric("Columns", df.shape[1])
                        with col3:
                            st.metric("Size", f"{uploaded_file.size / 1024:.2f} KB")

                        csv_data = BytesIO()
                        df.to_csv(csv_data, index=False)
                        
                        st.download_button(
                            label="📥 Download CSV File",
                            data=csv_data.getvalue(),
                            file_name=f"{uploaded_file.name.split('.')[0]}.csv",
                            mime="text/csv"
                        )

                    except Exception as e:
                        st.error(f"Error: {str(e)}")
                
                st.markdown('</div>', unsafe_allow_html=True)

    # Word ↔️ PDF Converter
    elif converter_type == "Word ↔️ PDF Converter":
        st.markdown("### Word ↔️ PDF Converter")
        
        conversion_direction = st.radio(
            "Select conversion direction:",
            ["Word to PDF", "PDF to Word"],
            horizontal=True
        )

        with st.container():
            st.markdown('<div class="converter-card">', unsafe_allow_html=True)
            
            if conversion_direction == "Word to PDF":
                uploaded_file = st.file_uploader("Upload Word file", type=["docx", "doc"], key="word_to_pdf")
                
                if uploaded_file is not None:
                    try:
                        doc = Document(uploaded_file)
                        output = BytesIO()
                        
                        # Convert Word to PDF using ReportLab
                        pdf = SimpleDocTemplate(output, pagesize=letter)
                        story = []
                        for paragraph in doc.paragraphs:
                            story.append(Paragraph(paragraph.text))
                        pdf.build(story)
                        
                        st.download_button(
                            label="📥 Download PDF File",
                            data=output.getvalue(),
                            file_name=f"{uploaded_file.name.split('.')[0]}.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
                
            else:  # PDF to Word
                uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"], key="pdf_to_word")
                
                if uploaded_file is not None:
                    try:
                        pdf_reader = PdfReader(uploaded_file)
                        doc = Document()
                        
                        for page in pdf_reader.pages:
                            text = page.extract_text()
                            doc.add_paragraph(text)
                        
                        docx_output = BytesIO()
                        doc.save(docx_output)
                        
                        st.download_button(
                            label="📥 Download Word File",
                            data=docx_output.getvalue(),
                            file_name=f"{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)

    # Image to PDF Converter
    elif converter_type == "Image to PDF Converter":
        st.markdown("### Image to PDF Converter")
        
        with st.container():
            st.markdown('<div class="converter-card">', unsafe_allow_html=True)
            
            uploaded_files = st.file_uploader(
                "Upload images (you can select multiple files)",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key="image_to_pdf"
            )
            
            if uploaded_files:
                try:
                    # Preview uploaded images
                    if len(uploaded_files) > 0:
                        st.markdown("#### Preview")
                        cols = st.columns(min(3, len(uploaded_files)))
                        for idx, file in enumerate(uploaded_files[:3]):
                            cols[idx].image(file, use_column_width=True)
                        
                        if len(uploaded_files) > 3:
                            st.info(f"+ {len(uploaded_files) - 3} more images")
                    
                    # Convert images to PDF
                    output = BytesIO()
                    pdf = Canvas(output, pagesize=letter)
                    
                    for image_file in uploaded_files:
                        img = Image.open(image_file)
                        img_width, img_height = img.size
                        aspect = img_height / float(img_width)
                        
                        # Scale image to fit on page
                        if aspect > 1:
                            img_width = letter[0] - 40
                            img_height = img_width * aspect
                        else:
                            img_height = letter[1] - 40
                            img_width = img_height / aspect
                        
                        pdf.drawImage(ImageReader(img), 20, letter[1] - img_height - 20,
                                    width=img_width, height=img_height)
                        pdf.showPage()
                    
                    pdf.save()
                    
                    st.download_button(
                        label="📥 Download PDF File",
                        data=output.getvalue(),
                        file_name="converted_images.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"Error: {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)
    elif converter_type == "PDF Editor":
      st.markdown("### PDF Editor") 
      with st.container():
        st.markdown('<div class="converter-card">', unsafe_allow_html=True) 
        uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"], key="pdf_editor")
        if uploaded_file is not None:
            col1, col2 = st.columns(2)
            pdf_reader = PdfReader(uploaded_file)
            total_pages = len(pdf_reader.pages)
            with col1:
                st.markdown("#### Original PDF")
                preview_page = st.number_input("Preview page", 1, total_pages, 1) - 1
                uploaded_file.seek(0)
                original_preview = get_pdf_preview(uploaded_file, preview_page)
                st.image(original_preview, use_column_width=True)
                operations = st.multiselect(
                    "Select operations to perform",
                    ["Extract Pages", "Merge PDFs", "Rotate Pages", "Add Watermark", 
                     "Resize", "Crop"])
            try:
                pdf_operations = {}
            
                if "Extract Pages" in operations:
                    st.markdown("#### Extract Pages")
                    total_pages = len(pdf_reader.pages)
                    all_pages = list(range(1, total_pages + 1))
                    selected_pages = st.multiselect("Select pages to extract",options=all_pages,default=[1],help="You can select multiple non-consecutive pages")
                    if selected_pages:
                          selected_pages.sort()
                          pdf_operations["extract"] = {"pages": selected_pages}
                          st.info(f"Selected pages: {', '.join(map(str, selected_pages))}")
                if "Merge PDFs" in operations:
                    st.markdown("#### Merge PDFs")
                    additional_pdfs = st.file_uploader(
                        "Upload PDFs to merge",
                        type=["pdf"],
                        accept_multiple_files=True,
                        key="merge_pdfs"
                    )
                    if additional_pdfs:
                        pdf_operations["merge"] = {"files": additional_pdfs}
                if "Rotate Pages" in operations:
                    st.markdown("#### Rotate Pages")
                    rotation = st.selectbox("Rotation angle", [90, 180, 270])
                    pdf_operations["rotate"] = {"angle": rotation}
                if "Add Watermark" in operations:
                    st.markdown("#### Add Watermark")
                    watermark_type = st.radio("Watermark Type", ["Text", "Image"])
                    watermark_options = {
                        "type": watermark_type.lower(),
                        "position": st.selectbox(
                            "Position",
                            ["center", "top-left", "top-right", "bottom-left", "bottom-right"]
                        ),
                        "angle": st.slider("Rotation Angle", -180, 180, 45),
                        "opacity": st.slider("Opacity", 0.1, 1.0, 0.3)
                    }
                    
                    page_selection = st.radio("Apply watermark to", ["All Pages", "Selected Pages"])
                    if page_selection == "Selected Pages":
                        selected_pages = st.multiselect(
                            "Select pages",
                            range(1, total_pages + 1)
                        )
                        watermark_options["pages"] = selected_pages
                    else:
                        watermark_options["pages"] = "all"
                    
                    if watermark_type == "Text":
                        watermark_options.update({
                            "text": st.text_input("Watermark text"),
                            "color": st.color_picker("Color", "#000000"),
                            "size": st.slider("Size", 20, 100, 40)
                        })
                    else:
                        watermark_image = st.file_uploader(
                            "Upload watermark image",
                            type=["png", "jpg", "jpeg"]
                        )
                        if watermark_image:
                            watermark_options.update({
                                "image": watermark_image,
                                "size": st.slider("Size (% of page width)", 10, 100, 30)
                            })
                        
                    if (watermark_type == "Text" and watermark_options["text"]) or (watermark_type == "Image" and watermark_image):
                        pdf_operations["watermark"] = watermark_options
                if "Resize" in operations:
                    st.markdown("#### Resize PDF")
                    scale = st.slider("Scale percentage", 1, 200, 100,
                                    help="100% is original size")
                    pdf_operations["resize"] = {"scale": scale}
                
                if "Crop" in operations:
                    st.markdown("#### Crop PDF")
                    st.info("Values are in percentage of original size")
                    crop_col1, crop_col2 = st.columns(2)
                    with crop_col1:
                        left = st.number_input("Left", 0, 100, 0)
                        right = st.number_input("Right", 0, 100, 100)
                    with crop_col2:
                        top = st.number_input("Top", 0, 100, 100)
                        bottom = st.number_input("Bottom", 0, 100, 0)
                    pdf_operations["crop"] = {
                        "left": left,
                        "right": right,
                        "top": top,
                        "bottom": bottom
                    }
                # Process PDF if any operations are selected
                if pdf_operations:
                    output = BytesIO()
                    uploaded_file.seek(0)
                    
                    # Process other operations
                    output = process_pdf(uploaded_file, pdf_operations)
                    
                    # Handle watermark if selected
                    if "watermark" in pdf_operations:
                        output.seek(0)
                        pdf_writer = PdfWriter()
                        temp_reader = PdfReader(output)
                        for page in temp_reader.pages:
                            pdf_writer.add_page(page)
                        pdf_writer = add_watermark(pdf_writer, pdf_operations["watermark"])
                        final_output = BytesIO()
                        pdf_writer.write(final_output)
                        output = final_output
                    
                    # Show preview and metrics for all operations
                    with col2:
                        st.markdown("#### Processed PDF")
                        output.seek(0)
                        processed_preview = get_pdf_preview(output, preview_page)
                        st.image(processed_preview, use_column_width=True)
                    
                    # Display metrics
                    original_size = len(uploaded_file.getvalue()) / 1024  # KB
                    output.seek(0)
                    new_size = len(output.getvalue()) / 1024  # KB
                    
                    metric_col1, metric_col2, metric_col3 = st.columns(3)
                    with metric_col1:
                        st.metric("Original Size", f"{original_size:.1f} KB")
                    with metric_col2:
                        st.metric("New Size", f"{new_size:.1f} KB")
                    with metric_col3:
                        reduction = ((original_size - new_size) / original_size) * 100
                        st.metric("Size Change", f"{reduction:.1f}%")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Modified PDF",
                        data=output.getvalue(),
                        file_name=f"modified_{uploaded_file.name}",
                        mime="application/pdf"
                    )
                    
            except Exception as e:
                st.error(f"Error: {str(e)}")
            
      st.markdown('</div>', unsafe_allow_html=True)
    
    # Add new Image Editor section
    elif converter_type == "Image Editor":
        st.markdown("### Image Editor")
        
        with st.container():
            st.markdown('<div class="converter-card">', unsafe_allow_html=True)
            
            uploaded_file = st.file_uploader(
                "Upload image",
                type=["png", "jpg", "jpeg"],
                key="image_editor"
            )
            
            if uploaded_file is not None:
                try:
                    original_bytes = uploaded_file.getvalue()
                    image = Image.open(uploaded_file)
                    
                    # Show original image
                    st.markdown("#### Original Image")
                    st.image(image, use_column_width=True)
                    
                    # Image operations
                    operations = {}
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        if st.checkbox("Resize"):
                            st.markdown("##### Resize Settings")
                            orig_width, orig_height = image.size
                            width = st.number_input("Width", min_value=1, value=orig_width)
                            height = st.number_input("Height", min_value=1, value=orig_height)
                            operations["resize"] = {"width": width, "height": height}
                        
                        if st.checkbox("Crop"):
                            st.markdown("##### Crop Settings")
                            width, height = image.size
                            left = st.number_input("Left", 0, width-1, 0)
                            top = st.number_input("Top", 0, height-1, 0)
                            right = st.number_input("Right", left+1, width, width)
                            bottom = st.number_input("Bottom", top+1, height, height)
                            operations["crop"] = {"left": left, "top": top, "right": right, "bottom": bottom}
                    
                    with col2:
                        if st.checkbox("Rotate"):
                            angle = st.slider("Rotation Angle", -180, 180, 0)
                            operations["rotate"] = {"angle": angle}
                        
                        if st.checkbox("Adjust"):
                            brightness = st.slider("Brightness", 0.0, 2.0, 1.0)
                            contrast = st.slider("Contrast", 0.0, 2.0, 1.0)
                            operations["brightness"] = {"factor": brightness}
                            operations["contrast"] = {"factor": contrast}
                        
                        if st.checkbox("Compress"):
                            quality = st.slider("Quality", 1, 100, 85)
                            operations["compress"] = {"quality": quality}
                    
                    if operations:
                        # Process image
                        processed_image, quality = process_image(image, operations)
                        
                        # Show processed image
                        st.markdown("#### Processed Image")
                        st.image(processed_image, use_column_width=True)
                        
                        # Save processed image
                        output = BytesIO()
                        if quality is not None:
                            processed_image.save(output, format=image.format, quality=quality, optimize=True)
                        else:
                            processed_image.save(output, format=image.format, optimize=True)
                        processed_bytes = output.getvalue()
                        col1, col2, col3 = st.columns(3)
                        with col1:
                         st.metric("Width", f"{processed_image.width}px")
                        with col2:
                         st.metric("Height", f"{processed_image.height}px")
                        with col3:
                         st.metric("Size", f"{len(processed_bytes)/1024:.1f} KB")
                        st.markdown("#### Size Comparison")
                        metrics = get_image_size_metrics(original_bytes, processed_bytes)
                        col1, col2, col3 = st.columns(3)
                        with col1:
                         st.metric("Original Size", f"{metrics['original_size']:.1f} KB")
                        with col2:
                         st.metric("New Size", f"{metrics['processed_size']:.1f} KB")
                        with col3:
                         st.metric(
                            "Size Change", 
                            f"{metrics['size_change']:.1f}%",
                            delta_color="inverse")
                        st.download_button(
                        label="📥 Download Processed Image",
                        data=processed_bytes,
                        file_name=f"processed_{uploaded_file.name}",
                        mime=f"image/{image.format.lower()}")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)

    # Add helpful information
    with st.expander("ℹ️ Need Help?"):
        st.markdown("""
        ### Usage Instructions
        1. Select the type of conversion you want to perform
        2. Upload your file(s) in the supported format
        3. Configure any additional settings if available
        4. Click the download button to save your converted file
        
        ### Supported Formats
        - Excel: .xlsx, .xls
        - CSV: .csv
        - Word: .docx, .doc
        - PDF: .pdf
        - Images: .png, .jpg, .jpeg
        
        ### Common Issues
        - If you're having trouble with CSV encoding, try different encoding options
        - Large files may take longer to process
        """)
def excel_editor():
    st.header("Excel Editor")
    def create_excel_structure_html(sheet, max_rows=5):
        html = "<table class='excel-table'>"
        merged_cells = sheet.merged_cells.ranges

        for idx, row in enumerate(sheet.iter_rows(max_row=max_rows)):
            html += "<tr>"
            for cell in row:
                merged = False
                for merged_range in merged_cells:
                    if cell.coordinate in merged_range:
                        if cell.coordinate == merged_range.start_cell.coordinate:
                            rowspan = min(merged_range.max_row - merged_range.min_row + 1, max_rows - idx)
                            colspan = merged_range.max_col - merged_range.min_col + 1
                            html += f"<td rowspan='{rowspan}' colspan='{colspan}'>{cell.value}</td>"
                        merged = True
                        break
                if not merged:
                    html += f"<td>{cell.value}</td>"
            html += "</tr>"
        html += "</table>"
        return html

    # Function to get merged column groups
    def get_merged_column_groups(sheet):
        merged_groups = {}
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row == 1:  # Only consider merged cells in the first row (header)
                main_col = sheet.cell(1, merged_range.min_col).value
                merged_groups[main_col] = list(range(merged_range.min_col, merged_range.max_col + 1))
        return merged_groups

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        # Read Excel file
        excel_file = openpyxl.load_workbook(uploaded_file)
        sheet = excel_file.active

        # Display original Excel structure (first 5 rows)
        st.subheader("Original Excel Structure (First 5 Rows)")
        excel_html = create_excel_structure_html(sheet, max_rows=5)
        st.markdown(excel_html, unsafe_allow_html=True)

        # Get merged column groups
        merged_groups = get_merged_column_groups(sheet)

        # Create a list of column headers, considering merged cells
        column_headers = []
        column_indices = OrderedDict()  # To store the column indices for each header
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(1, col).value
            if cell_value is not None:
                column_headers.append(cell_value)
                if cell_value not in column_indices:
                    column_indices[cell_value] = []
                column_indices[cell_value].append(col - 1)  # pandas uses 0-based index
            else:
                # If the cell is empty, it's part of a merged cell, so use the previous header
                prev_header = column_headers[-1]
                column_headers.append(prev_header)
                column_indices[prev_header].append(col - 1)

        # Read as pandas DataFrame using the correct column headers
        df = pd.read_excel(uploaded_file, header=None, names=column_headers)
        df = df.iloc[1:]  # Remove the first row as it's now our header

        # Column selection for deletion
        st.subheader("Select columns to delete")
        all_columns = list(column_indices.keys())  # Use OrderedDict keys to maintain order
        cols_to_delete = st.multiselect("Choose columns to remove", all_columns)
        
        if cols_to_delete:
            columns_to_remove = []
            for col in cols_to_delete:
                columns_to_remove.extend(column_indices[col])
            
            df = df.drop(df.columns[columns_to_remove], axis=1)
            st.success(f"Deleted columns: {', '.join(cols_to_delete)}")

        # Row deletion
        st.subheader("Delete rows")
        num_rows = st.number_input("Enter the number of rows to delete from the start", min_value=0, max_value=len(df)-1, value=0)
        
        if num_rows > 0:
            df = df.iloc[num_rows:]
            st.success(f"Deleted first {num_rows} rows")
        
        # Display editable dataframe
        st.subheader("Edit Data")
        st.write("You can edit individual cell values directly in the table below:")
        
        # Replace NaN values with None and convert dataframe to a dictionary
        df_dict = df.where(pd.notnull(df), None).to_dict('records')
        
        # Use st.data_editor with the processed dictionary
        edited_data = st.data_editor(df_dict)
        
        # Convert edited data back to dataframe
        edited_df = pd.DataFrame(edited_data)
        st.subheader("Edited Data")
        st.dataframe(edited_df)
        
        # Download button
        def get_excel_download_link(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            excel_data = output.getvalue()
            b64 = base64.b64encode(excel_data).decode()
            return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="edited_file.xlsx">Download Edited Excel File</a>'
        
        st.markdown(get_excel_download_link(edited_df), unsafe_allow_html=True)

        # New button to upload edited file to Home
        if st.button("Upload Edited File to Home"):
            # Save the edited DataFrame to session state
            st.session_state.edited_df = edited_df
            st.session_state.edited_file_name = "edited_" + uploaded_file.name
            st.success("Edited file has been uploaded to Home. Please switch to the Home tab to see the uploaded file.")

    else:
        st.info("Please upload an Excel file to begin editing.")
def data_analyzer():
    st.header("Advanced Data Analyzer")

    uploaded_file = st.file_uploader("Choose an Excel file for analysis", type="xlsx", key="analyser")

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.write("Dataset Information:")
        st.write(f"Number of rows: {df.shape[0]}")
        st.write(f"Number of columns: {df.shape[1]}")
        
        numeric_columns = df.select_dtypes(include=['float64', 'int64']).columns
        categorical_columns = df.select_dtypes(include=['object']).columns
        
        analysis_type = st.selectbox("Select analysis type", ["Univariate Analysis", "Bivariate Analysis", "Regression Analysis", "Machine Learning Models", "Advanced Statistics"])
        
        if analysis_type == "Univariate Analysis":
            univariate_analysis(df, numeric_columns, categorical_columns)
        elif analysis_type == "Bivariate Analysis":
            bivariate_analysis(df, numeric_columns)
        elif analysis_type == "Regression Analysis":
            regression_analysis(df, numeric_columns, categorical_columns)
        elif analysis_type == "Machine Learning Models":
            machine_learning_models(df, numeric_columns, categorical_columns)
        elif analysis_type == "Advanced Statistics":
            advanced_statistics(df, numeric_columns)

def univariate_analysis(df, numeric_columns, categorical_columns):
    st.subheader("Univariate Analysis")
    
    column = st.selectbox("Select a column for analysis", numeric_columns.tolist() + categorical_columns.tolist())
    
    if column in numeric_columns:
        st.write(df[column].describe())
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig = go.Figure()
            fig.add_trace(go.Histogram(x=df[column], name="Histogram"))
            fig.update_layout(title=f"Histogram for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = go.Figure()
            fig.add_trace(go.Box(y=df[column], name="Box Plot"))
            fig.update_layout(title=f"Box Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            fig = go.Figure()
            fig.add_trace(go.Violin(y=df[column], box_visible=True, line_color='black', meanline_visible=True, fillcolor='lightseagreen', opacity=0.6, x0=column))
            fig.update_layout(title=f"Violin Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        with col4:
            fig = px.line(df, y=column, title=f"Line Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        # Additional statistics
        st.subheader("Additional Statistics")
        col5, col6, col7 = st.columns(3)
        with col5:
            st.metric("Skewness", f"{skew(df[column]):.4f}")
        with col6:
            st.metric("Kurtosis", f"{kurtosis(df[column]):.4f}")
        with col7:
            st.metric("Coefficient of Variation", f"{df[column].std() / df[column].mean():.4f}")
        
    else:
        st.write(df[column].value_counts())
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.bar(df[column].value_counts(), title=f"Bar Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = px.pie(df, names=column, title=f"Pie Chart for {column}")
            st.plotly_chart(fig, use_container_width=True)

def bivariate_analysis(df, numeric_columns):
    st.subheader("Bivariate Analysis")
    
    x_col = st.selectbox("Select X-axis variable", numeric_columns)
    y_col = st.selectbox("Select Y-axis variable", numeric_columns)
    
    chart_type = st.selectbox("Select chart type", ["Scatter", "Line", "Bar", "Box", "Violin", "3D Scatter", "Heatmap"])
    
    if chart_type == "Scatter":
        fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Line":
        fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Bar":
        fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Box":
        fig = px.box(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Violin":
        fig = px.violin(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "3D Scatter":
        z_col = st.selectbox("Select Z-axis variable", numeric_columns)
        fig = px.scatter_3d(df, x=x_col, y=y_col, z=z_col, title=f"3D Scatter Plot")
    elif chart_type == "Heatmap":
        corr_matrix = df[numeric_columns].corr()
        fig = px.imshow(corr_matrix, title="Correlation Heatmap")
    
    st.plotly_chart(fig, use_container_width=True)
    
    correlation = df[[x_col, y_col]].corr().iloc[0, 1]
    st.write(f"Correlation between {x_col} and {y_col}: {correlation:.4f}")
    
    # Add correlation interpretation
    st.subheader("Correlation Interpretation")
    st.write("""
    The correlation coefficient ranges from -1 to 1:
    - 1: Perfect positive correlation
    - 0: No correlation
    - -1: Perfect negative correlation
    
    Interpretation:
    - 0.00 to 0.19: Very weak correlation
    - 0.20 to 0.39: Weak correlation
    - 0.40 to 0.59: Moderate correlation
    - 0.60 to 0.79: Strong correlation
    - 0.80 to 1.00: Very strong correlation
    """)
    
    # Add correlation formula
    st.latex(r'''
    r = \frac{\sum_{i=1}^{n} (x_i - \bar{x})(y_i - \bar{y})}{\sqrt{\sum_{i=1}^{n} (x_i - \bar{x})^2} \sqrt{\sum_{i=1}^{n} (y_i - \bar{y})^2}}
    ''')
    st.write("Where:")
    st.write("- r is the correlation coefficient")
    st.write("- x_i and y_i are individual sample points")
    st.write("- x̄ and ȳ are the sample means")

def regression_analysis(df, numeric_columns, categorical_columns):
    st.subheader("Regression Analysis")
    
    regression_type = st.selectbox("Select regression type", ["Simple Linear", "Multiple Linear", "Polynomial", "Ridge", "Lasso"])
    
    y_col = st.selectbox("Select dependent variable", numeric_columns)
    x_cols = st.multiselect("Select independent variables", numeric_columns.tolist() + categorical_columns.tolist())
    
    if len(x_cols) == 0:
        st.warning("Please select at least one independent variable.")
        return
    
    X = df[x_cols]
    y = df[y_col]
    
    # Handle categorical variables
    X = pd.get_dummies(X, drop_first=True)
    
    if regression_type == "Polynomial":
        degree = st.slider("Select polynomial degree", 1, 5, 2)
        poly = PolynomialFeatures(degree=degree)
        X = poly.fit_transform(X)
    
    X = sm.add_constant(X)
    
    try:
        if regression_type == "Ridge":
            alpha = st.slider("Select alpha for Ridge regression", 0.0, 10.0, 1.0)
            model = sm.OLS(y, X).fit_regularized(alpha=alpha, L1_wt=0)
        elif regression_type == "Lasso":
            alpha = st.slider("Select alpha for Lasso regression", 0.0, 10.0, 1.0)
            model = sm.OLS(y, X).fit_regularized(alpha=alpha, L1_wt=1)
        else:
            model = sm.OLS(y, X).fit()
        
        st.write(model.summary())
        
        # Plot actual vs predicted values
        fig = px.scatter(x=y, y=model.predict(X), labels={'x': 'Actual', 'y': 'Predicted'}, title="Actual vs Predicted Values")
        fig.add_trace(go.Scatter(x=[y.min(), y.max()], y=[y.min(), y.max()], mode='lines', name='y=x'))
        st.plotly_chart(fig, use_container_width=True)
        
        # Residual plot
        residuals = model.resid
        fig = px.scatter(x=model.predict(X), y=residuals, labels={'x': 'Predicted', 'y': 'Residuals'}, title="Residual Plot")
        fig.add_hline(y=0, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)
        
        # Statistical tests
        st.subheader("Statistical Tests")
        
        # Normality test (Jarque-Bera)
        jb_statistic, jb_p_value = jarque_bera(residuals)
        st.write(f"Jarque-Bera Test for Normality: statistic = {jb_statistic:.4f}, p-value = {jb_p_value:.4f}")
        st.write(f"{'Reject' if jb_p_value < 0.05 else 'Fail to reject'} the null hypothesis of normality at 5% significance level.")
        
        # Heteroscedasticity test (Breusch-Pagan)
        _, bp_p_value, _, _ = het_breuschpagan(residuals, model.model.exog)
        st.write(f"Breusch-Pagan Test for Heteroscedasticity: p-value = {bp_p_value:.4f}")
        st.write(f"{'Reject' if bp_p_value < 0.05 else 'Fail to reject'} the null hypothesis of homoscedasticity at 5% significance level.")
        dw_statistic = durbin_watson(residuals)
        st.write(f"Durbin-Watson Test for Autocorrelation: {dw_statistic:.4f}")
        st.write("Values close to 2 suggest no autocorrelation, while values toward 0 or 4 suggest positive or negative autocorrelation.")
        
        # Multicollinearity (VIF)
        vif_data = pd.DataFrame()
        vif_data["Variable"] = X.columns
        vif_data["VIF"] = [variance_inflation_factor(X.values, i) for i in range(X.shape[1])]
        st.write("Variance Inflation Factors (VIF) for Multicollinearity:")
        st.write(vif_data)
        st.write("VIF > 5 suggests high multicollinearity.")
        
        # Add regression formulas and explanations
        st.subheader("Regression Formulas")
        if regression_type == "Simple Linear":
            st.latex(r'y = \beta_0 + \beta_1x + \epsilon')
        elif regression_type == "Multiple Linear":
            st.latex(r'y = \beta_0 + \beta_1x_1 + \beta_2x_2 + ... + \beta_nx_n + \epsilon')
        elif regression_type == "Polynomial":
            st.latex(r'y = \beta_0 + \beta_1x + \beta_2x^2 + ... + \beta_nx^n + \epsilon')
        elif regression_type == "Ridge":
            st.latex(r'\min_{\beta} \sum_{i=1}^n (y_i - \beta_0 - \sum_{j=1}^p \beta_jx_{ij})^2 + \lambda \sum_{j=1}^p \beta_j^2')
        elif regression_type == "Lasso":
            st.latex(r'\min_{\beta} \sum_{i=1}^n (y_i - \beta_0 - \sum_{j=1}^p \beta_jx_{ij})^2 + \lambda \sum_{j=1}^p |\beta_j|')
        
        st.write("Where:")
        st.write("- y is the dependent variable")
        st.write("- x, x_1, x_2, ..., x_n are independent variables")
        st.write("- β_0, β_1, β_2, ..., β_n are regression coefficients")
        st.write("- ε is the error term")
        st.write("- λ is the regularization parameter (for Ridge and Lasso)")
        
    except Exception as e:
        st.error(f"An error occurred during regression analysis: {str(e)}")
        st.write("This error might be due to multicollinearity, insufficient data, or other issues in the dataset.")
        st.write("Try selecting different variables or using a different regression type.")

def machine_learning_models(df, numeric_columns, categorical_columns):
    st.subheader("Machine Learning Models")
    
    model_type = st.selectbox("Select model type", ["Supervised", "Unsupervised"])
    
    if model_type == "Supervised":
        supervised_models(df, numeric_columns, categorical_columns)
    else:
        unsupervised_models(df, numeric_columns)

def supervised_models(df, numeric_columns, categorical_columns):
    st.write("Supervised Learning Models")
    
    y_col = st.selectbox("Select target variable", numeric_columns)
    x_cols = st.multiselect("Select features", numeric_columns.tolist() + categorical_columns.tolist())
    
    if len(x_cols) == 0:
        st.warning("Please select at least one feature.")
        return
    
    X = df[x_cols]
    y = df[y_col]
    
    # Handle categorical variables
    X = pd.get_dummies(X, drop_first=True)
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)
    
    models = {
        "Linear Regression": LinearRegression(),
        "Decision Tree": DecisionTreeRegressor(),
        "Random Forest": RandomForestRegressor(),
        "SVR": SVR()
    }
    
    selected_model = st.selectbox("Select a model", list(models.keys()))
    
    try:
        model = models[selected_model]
        model.fit(X_train_scaled, y_train)
        
        y_pred = model.predict(X_test_scaled)
        
        mse = mean_squared_error(y_test, y_pred)
        r2 = r2_score(y_test, y_pred)
        
        st.write(f"Mean Squared Error: {mse:.4f}")
        st.write(f"R-squared Score: {r2:.4f}")
        
        fig = px.scatter(x=y_test, y=y_pred, labels={'x': 'Actual', 'y': 'Predicted'}, title="Actual vs Predicted Values")
        fig.add_trace(go.Scatter(x=[y_test.min(), y_test.max()], y=[y_test.min(), y_test.max()], mode='lines', name='y=x'))
        st.plotly_chart(fig, use_container_width=True)
        
        # Feature importance (for tree-based models)
        if selected_model in ["Decision Tree", "Random Forest"]:
            feature_importance = pd.DataFrame({
                'feature': X.columns,
                'importance': model.feature_importances_
            }).sort_values('importance', ascending=False)
            
            st.write("Feature Importance:")
            fig = px.bar(feature_importance, x='feature', y='importance', title="Feature Importance")
            st.plotly_chart(fig, use_container_width=True)
        
        # Add model formulas and explanations
        st.subheader("Model Formulas and Explanations")
        if selected_model == "Linear Regression":
            st.latex(r'y = \beta_0 + \beta_1x_1 + \beta_2x_2 + ... + \beta_nx_n')
            st.write("Linear Regression finds the best-fitting linear relationship between the target variable and the features.")
        elif selected_model == "Decision Tree":
            st.write("Decision Trees make predictions by learning decision rules inferred from the data features.")
            st.image("https://scikit-learn.org/stable/_images/iris_dtc.png", caption="Example of a Decision Tree")
        elif selected_model == "Random Forest":
            st.write("Random Forest is an ensemble of Decision Trees, where each tree is trained on a random subset of the data and features.")
            st.image("https://scikit-learn.org/stable/_images/plot_forest_importances_faces_001.png", caption="Example of Random Forest Feature Importance")
        elif selected_model == "SVR":
            st.latex(r'\min_{w, b, \xi} \frac{1}{2} \|w\|^2 + C \sum_{i=1}^n \xi_i')
            st.write("Support Vector Regression (SVR) finds a function that deviates from y by a value no greater than ε for each training point x.")
    
    except Exception as e:
        st.error(f"An error occurred during model training: {str(e)}")
        st.write("This error might be due to insufficient data, incompatible data types, or other issues in the dataset.")
        st.write("Try selecting different variables or using a different model.")

def unsupervised_models(df, numeric_columns):
    st.write("Unsupervised Learning Models")
    
    x_cols = st.multiselect("Select features for clustering", numeric_columns)
    
    if len(x_cols) == 0:
        st.warning("Please select at least one feature.")
        return
    
    X = df[x_cols]
    
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    n_clusters = st.slider("Select number of clusters", 2, 10, 3)
    
    try:
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        cluster_labels = kmeans.fit_predict(X_scaled)
        
        df_clustered = df.copy()
        df_clustered['Cluster'] = cluster_labels
        
        if len(x_cols) >= 2:
            fig = px.scatter(df_clustered, x=x_cols[0], y=x_cols[1], color='Cluster', title="K-means Clustering")
            st.plotly_chart(fig, use_container_width=True)
        
        st.write("Cluster Centers:")
        cluster_centers = scaler.inverse_transform(kmeans.cluster_centers_)
        st.write(pd.DataFrame(cluster_centers, columns=x_cols))
        
        # Elbow method for optimal number of clusters
        inertias = []
        k_range = range(1, 11)
        for k in k_range:
            kmeans = KMeans(n_clusters=k, random_state=42)
            kmeans.fit(X_scaled)
            inertias.append(kmeans.inertia_)
        
        fig = px.line(x=k_range, y=inertias, title="Elbow Method for Optimal k",
                      labels={'x': 'Number of Clusters (k)', 'y': 'Inertia'})
        st.plotly_chart(fig, use_container_width=True)
        
        # PCA
        st.subheader("Principal Component Analysis (PCA)")
        n_components = st.slider("Select number of components", 2, min(len(x_cols), 10), 2)
        pca = PCA(n_components=n_components)
        pca_result = pca.fit_transform(X_scaled)
        
        df_pca = pd.DataFrame(data=pca_result, columns=[f'PC{i+1}' for i in range(n_components)])
        
        fig = px.scatter(df_pca, x='PC1', y='PC2', title="PCA Visualization")
        st.plotly_chart(fig, use_container_width=True)
        
        explained_variance_ratio = pca.explained_variance_ratio_
        cumulative_variance_ratio = np.cumsum(explained_variance_ratio)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=range(1, n_components+1), y=explained_variance_ratio, name='Individual'))
        fig.add_trace(go.Scatter(x=range(1, n_components+1), y=cumulative_variance_ratio, mode='lines+markers', name='Cumulative'))
        fig.update_layout(title='Explained Variance Ratio', xaxis_title='Principal Components', yaxis_title='Explained Variance Ratio')
        st.plotly_chart(fig, use_container_width=True)
        
        st.write("Explained Variance Ratio:")
        st.write(pd.DataFrame({'PC': range(1, n_components+1), 'Explained Variance Ratio': explained_variance_ratio, 'Cumulative Variance Ratio': cumulative_variance_ratio}))
        
        # Add formulas and explanations
        st.subheader("K-means Clustering Formula")
        st.latex(r'\min_{S} \sum_{i=1}^{k} \sum_{x \in S_i} \|x - \mu_i\|^2')
        st.write("Where:")
        st.write("- S is the set of clusters")
        st.write("- k is the number of clusters")
        st.write("- x is a data point")
        st.write("- μ_i is the mean of points in S_i")
        
        st.subheader("PCA Formula")
        st.latex(r'X = U\Sigma V^T')
        st.write("Where:")
        st.write("- X is the original data matrix")
        st.write("- U is the left singular vectors (eigenvectors of XX^T)")
        st.write("- Σ is a diagonal matrix of singular values")
        st.write("- V^T is the right singular vectors (eigenvectors of X^TX)")
    
    except Exception as e:
        st.error(f"An error occurred during unsupervised learning: {str(e)}")
        st.write("This error might be due to insufficient data, incompatible data types, or other issues in the dataset.")
        st.write("Try selecting different variables or adjusting the number of clusters/components.")

def advanced_statistics(df, numeric_columns):
    st.subheader("Advanced Statistics")
    
    column = st.selectbox("Select a column for advanced statistics", numeric_columns)
    
    st.write("Descriptive Statistics:")
    st.write(df[column].describe())
    
    st.subheader("Normality Tests")
    
    # Shapiro-Wilk Test
    shapiro_stat, shapiro_p = stats.shapiro(df[column])
    st.write(f"Shapiro-Wilk Test: statistic = {shapiro_stat:.4f}, p-value = {shapiro_p:.4f}")
    st.write(f"{'Reject' if shapiro_p < 0.05 else 'Fail to reject'} the null hypothesis of normality at 5% significance level.")
    
    # Anderson-Darling Test
    anderson_result = stats.anderson(df[column])
    st.write("Anderson-Darling Test:")
    st.write(f"Statistic: {anderson_result.statistic:.4f}")
    for i in range(len(anderson_result.critical_values)):
        sl, cv = anderson_result.significance_level[i], anderson_result.critical_values[i]
        st.write(f"At {sl}% significance level: critical value = {cv:.4f}")
        if anderson_result.statistic < cv:
            st.write(f"The null hypothesis of normality is not rejected at {sl}% significance level.")
        else:
            st.write(f"The null hypothesis of normality is rejected at {sl}% significance level.")
    
    # Jarque-Bera Test
    jb_stat, jb_p = stats.jarque_bera(df[column])
    st.write(f"Jarque-Bera Test: statistic = {jb_stat:.4f}, p-value = {jb_p:.4f}")
    st.write(f"{'Reject' if jb_p < 0.05 else 'Fail to reject'} the null hypothesis of normality at 5% significance level.")
    
    # Q-Q Plot
    fig, ax = plt.subplots()
    stats.probplot(df[column], dist="norm", plot=ax)
    ax.set_title("Q-Q Plot")
    st.pyplot(fig)
    
    st.subheader("Time Series Analysis")
    
    # Augmented Dickey-Fuller Test for Stationarity
    adf_result = adfuller(df[column])
    st.write("Augmented Dickey-Fuller Test:")
    st.write(f"ADF Statistic: {adf_result[0]:.4f}")
    st.write(f"p-value: {adf_result[1]:.4f}")
    for key, value in adf_result[4].items():
        st.write(f"Critical Value ({key}): {value:.4f}")
    st.write(f"{'Reject' if adf_result[1] < 0.05 else 'Fail to reject'} the null hypothesis of a unit root at 5% significance level.")
    
    # ACF and PACF plots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 10))
    plot_acf(df[column], ax=ax1)
    plot_pacf(df[column], ax=ax2)
    ax1.set_title("Autocorrelation Function (ACF)")
    ax2.set_title("Partial Autocorrelation Function (PACF)")
    st.pyplot(fig)
    
    st.subheader("Distribution Fitting")
    
    # Fit normal distribution
    mu, sigma = stats.norm.fit(df[column])
    x = np.linspace(df[column].min(), df[column].max(), 100)
    y = stats.norm.pdf(x, mu, sigma)
    
    fig, ax = plt.subplots()
    ax.hist(df[column], density=True, alpha=0.7, bins='auto')
    ax.plot(x, y, 'r-', lw=2, label='Normal fit')
    ax.set_title(f"Distribution Fitting for {column}")
    ax.legend()
    st.pyplot(fig)
    
    st.write(f"Fitted Normal Distribution: μ = {mu:.4f}, σ = {sigma:.4f}")
    
    # Kolmogorov-Smirnov Test
    ks_statistic, ks_p_value = stats.kstest(df[column], 'norm', args=(mu, sigma))
    st.write("Kolmogorov-Smirnov Test:")
    st.write(f"Statistic: {ks_statistic:.4f}")
    st.write(f"p-value: {ks_p_value:.4f}")
    st.write(f"{'Reject' if ks_p_value < 0.05 else 'Fail to reject'} the null hypothesis that the data comes from the fitted normal distribution at 5% significance level.")
def create_stats_pdf(stats_data, district):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    title = Paragraph(f"Descriptive Statistics for {district}", styles['Title'])
    elements.append(title)

    data = [['Brand', 'Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']]
    for brand, stats in stats_data.items():
        row = [brand]
        for stat in ['Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']:
            value = stats[stat]
            if isinstance(value, (int, float)):
                row.append(f"{value:.2f}")
            else:
                row.append(str(value))
        data.append(row)

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer
def create_prediction_pdf(prediction_data, district):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    title = Paragraph(f"Price Predictions for {district}", styles['Title'])
    elements.append(title)

    data = [['Brand', 'Predicted Price', 'Lower CI', 'Upper CI']]
    for brand, pred in prediction_data.items():
        row = [brand, f"{pred['forecast']:.2f}", f"{pred['lower_ci']:.2f}", f"{pred['upper_ci']:.2f}"]
        data.append(row)

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="WSP Analysis",page_icon="🔬", layout="wide")

# [Keep the existing custom CSS here]
# Custom CSS for the entire app
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .main .block-container {
        padding: 2rem;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 15px;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
    }
    h1 {
        color: #2c3e50;
        text-align: center;
        padding: 1.5rem;
        background: rgba(255, 255, 255, 0.95);
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stSelectbox, .stMultiSelect {
        background: white;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        background-color: #3498db;
        color: white;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2980b9;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stSlider > div > div > div {
        background-color: #3498db;
    }
    .stCheckbox > label {
        color: #2c3e50;
        font-weight: 500;
    }
    .stSubheader {
        color: #34495e;
        background: rgba(255, 255, 255, 0.9);
        padding: 0.8rem;
        border-radius: 8px;
        margin-top: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .uploadedFile {
        background-color: #e8f0fe;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .dataframe {
        font-size: 0.8em;
    }
    .dataframe thead tr th {
        background-color: #3498db;
        color: brown;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #f2f2f2;
    }
</style>
""", unsafe_allow_html=True)
# Global variables
if 'df' not in st.session_state:
    st.session_state.df = None
if 'week_names_input' not in st.session_state:
    st.session_state.week_names_input = []
if 'desired_diff_input' not in st.session_state:
    st.session_state.desired_diff_input = {}
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False
if 'diff_week' not in st.session_state:
    st.session_state.diff_week = 0

# [Keep the existing transform_data, plot_district_graph, process_file, and update_week_name functions]
def transform_data(df, week_names_input):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    
    # Region name replacements
    region_replacements = {
        '12_Madhya Pradesh(west)': 'Madhya Pradesh(West)',
        '20_Rajasthan': 'Rajasthan', '50_Rajasthan III': 'Rajasthan', '80_Rajasthan II': 'Rajasthan',
        '33_Chhattisgarh(2)': 'Chhattisgarh', '38_Chhattisgarh(3)': 'Chhattisgarh', '39_Chhattisgarh(1)': 'Chhattisgarh',
        '07_Haryana 1': 'Haryana', '07_Haryana 2': 'Haryana',
        '06_Gujarat 1': 'Gujarat', '66_Gujarat 2': 'Gujarat', '67_Gujarat 3': 'Gujarat', '68_Gujarat 4': 'Gujarat', '69_Gujarat 5': 'Gujarat',
        '13_Maharashtra': 'Maharashtra(West)',
        '24_Uttar Pradesh': 'Uttar Pradesh(West)',
        '35_Uttarakhand': 'Uttarakhand',
        '83_UP East Varanasi Region': 'Varanasi',
        '83_UP East Lucknow Region': 'Lucknow',
        '30_Delhi': 'Delhi',
        '19_Punjab': 'Punjab',
        '09_Jammu&Kashmir': 'Jammu&Kashmir',
        '08_Himachal Pradesh': 'Himachal Pradesh',
        '82_Maharashtra(East)': 'Maharashtra(East)',
        '81_Madhya Pradesh': 'Madhya Pradesh(East)',
        '34_Jharkhand': 'Jharkhand',
        '18_ODISHA': 'Odisha',
        '04_Bihar': 'Bihar',
        '27_Chandigarh': 'Chandigarh',
        '82_Maharashtra (East)': 'Maharashtra(East)',
        '25_West Bengal': 'West Bengal'
    }
    
    transformed_df['REGION'] = transformed_df['REGION'].replace(region_replacements)
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Delhi', 'Haryana', 'Punjab'], 'North-I')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Uttar Pradesh(West)','Uttarakhand'], 'North-II')
    
    zone_replacements = {
        'EZ_East Zone': 'East Zone',
        'CZ_Central Zone': 'Central Zone',
        'NZ_North Zone': 'North Zone',
        'UPEZ_UP East Zone': 'UP East Zone',
        'upWZ_up West Zone': 'UP West Zone',
        'WZ_West Zone': 'West Zone'
    }
    transformed_df['Zone'] = transformed_df['Zone'].replace(zone_replacements)
    
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    
    for i in range(num_weeks):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        week_name = week_names_input[i]
        week_data = week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True)
        
        # Use a unique suffix for each merge operation
        suffix = f'_{i}'
        transformed_df = pd.merge(transformed_df, week_data, left_index=True, right_index=True, suffixes=('', suffix))
    transformed_df = transformed_df.loc[:, ~transformed_df.columns.str.contains('_\d+$')]
    return transformed_df
def plot_district_graph(df, district_names, benchmark_brands_dict, desired_diff_dict, week_names, diff_week, download_pdf=False):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    if download_pdf:
        pdf = matplotlib.backends.backend_pdf.PdfPages("district_plots.pdf")
    
    for i, district_name in enumerate(district_names):
        fig,ax=plt.subplots(figsize=(10, 8))
        district_df = df[df["Dist Name"] == district_name]
        price_diffs = []
        for brand in brands:
            brand_prices = []
            for week_name in week_names:
                column_name = f"{brand} ({week_name})"
                if column_name in district_df.columns:
                    price = district_df[column_name].iloc[0]
                    brand_prices.append(price)
                else:
                    brand_prices.append(np.nan)
            valid_prices = [p for p in brand_prices if not np.isnan(p)]
            if len(valid_prices) > diff_week:
                price_diff = valid_prices[-1] - valid_prices[diff_week]
            else:
                price_diff = np.nan
            price_diff_label = price_diff
            if np.isnan(price_diff):
               price_diff = 'NA'
            label = f"{brand} ({price_diff if isinstance(price_diff, str) else f'{price_diff:.0f}'})"
            plt.plot(week_names, brand_prices, marker='o', linestyle='-', label=label)
            for week, price in zip(week_names, brand_prices):
                if not np.isnan(price):
                    plt.text(week, price, str(round(price)), fontsize=10)
        plt.grid(False)
        plt.xlabel('Month/Week', weight='bold')
        reference_week = week_names[diff_week]
        last_week = week_names[-1]
        
        explanation_text = f"***Numbers in brackets next to brand names show the price difference between {reference_week} and {last_week}.***"
        plt.annotate(explanation_text, 
                     xy=(0, -0.23), xycoords='axes fraction', 
                     ha='left', va='center', fontsize=8, style='italic', color='deeppink',
                     bbox=dict(facecolor="#f0f8ff", edgecolor='none', alpha=0.7, pad=3))
        
        region_name = district_df['REGION'].iloc[0]
        plt.ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        if i == 0:
            plt.text(0.5, 1.1, region_name, ha='center', va='center', transform=plt.gca().transAxes, weight='bold', fontsize=16)
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        else:
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()

        text_str = ''
        if district_name in benchmark_brands_dict:
            brand_texts = []
            max_left_length = 0
            for benchmark_brand in benchmark_brands_dict[district_name]:
                jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
                benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
                actual_diff = np.nan
                if jklc_prices and benchmark_prices:
                    for i in range(len(jklc_prices) - 1, -1, -1):
                        if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                            actual_diff = jklc_prices[i] - benchmark_prices[i]
                            break
                desired_diff_str = f" ({desired_diff_dict[district_name][benchmark_brand]:.0f} Rs.)" if district_name in desired_diff_dict and benchmark_brand in desired_diff_dict[district_name] else ""
                brand_text = [f"Benchmark Brand: {benchmark_brand}{desired_diff_str}", f"Actual Diff: {actual_diff:+.0f} Rs."]
                brand_texts.append(brand_text)
                max_left_length = max(max_left_length, len(brand_text[0]))
            num_brands = len(brand_texts)
            if num_brands == 1:
                text_str = "\n".join(brand_texts[0])
            elif num_brands > 1:
                half_num_brands = num_brands // 2
                left_side = brand_texts[:half_num_brands]
                right_side = brand_texts[half_num_brands:]
                lines = []
                for i in range(2):
                    left_text = left_side[0][i] if i < len(left_side[0]) else ""
                    right_text = right_side[0][i] if i < len(right_side[0]) else ""
                    lines.append(f"{left_text.ljust(max_left_length)} \u2502 {right_text.rjust(max_left_length)}")
                text_str = "\n".join(lines)
        plt.text(0.5, -0.3, text_str, weight='bold', ha='center', va='center', transform=plt.gca().transAxes, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
        plt.subplots_adjust(bottom=0.25)
        if download_pdf:
            pdf.savefig(fig, bbox_inches='tight')
        st.pyplot(fig)
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        b64_data = base64.b64encode(buf.getvalue()).decode()
        st.markdown(f'<a download="district_plot_{district_name}.png" href="data:image/png;base64,{b64_data}">Download Plot as PNG</a>', unsafe_allow_html=True)
        plt.close()
    
    if download_pdf:
        pdf.close()
        with open("district_plots.pdf", "rb") as f:
            pdf_data = f.read()
        b64_pdf = base64.b64encode(pdf_data).decode()
        st.markdown(f'<a download="{region_name}.pdf" href="data:application/pdf;base64,{b64_pdf}">Download All Plots as PDF</a>', unsafe_allow_html=True)
def update_week_name(index):
    def callback():
        if index < len(st.session_state.week_names_input):
            st.session_state.week_names_input[index] = st.session_state[f'week_{index}']
        else:
            st.warning(f"Attempted to update week {index + 1}, but only {len(st.session_state.week_names_input)} weeks are available.")
        st.session_state.all_weeks_filled = all(st.session_state.week_names_input)
    return callback


def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
    #-webkit-background-clip: text;
        #-webkit-text-fill-color: transparent;
def Home():
    # Custom CSS with more modern and professional styling
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
    
    body {
        font-family: 'Roboto', sans-serif;
        background-color: #f5f7fa;
        color: #333;
    }
    .title {
        font-size: 3.5rem;
        font-weight: 700;
        color: brown;
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 2rem;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        
    }
    .subtitle {
        font-size: 1.5rem;
        font-weight: 300;
        color: #34495e;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-box {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 2rem;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
    }
    .upload-section {
        background: linear-gradient(120deg, #a1c4fd 0%, #c2e9fb 100%);
        padding: 2rem;
        border-radius: 8px;
        margin-bottom: 2rem;
    }
    .btn-primary {
        background-color: #3498db;
        color: brown;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        border: none;
        cursor: pointer;
        transition: background-color 0.3s ease;
    }
    .btn-primary:hover {
        background-color: #2980b9;
    }
    </style>
    """, unsafe_allow_html=True)

    # Main title and subtitle
    st.markdown('<h1 class="title">Statistica</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Analyze, Visualize, Optimize.</p>', unsafe_allow_html=True)

    # Load and display Lottie animation
    lottie_url = "https://assets9.lottiefiles.com/packages/lf20_jcikwtux.json"
    lottie_json = load_lottie_url(lottie_url)

    col1, col2 = st.columns([1, 2])
    with col1:
        st_lottie(lottie_json, height=250, key="home_animation")
    with col2:
        st.markdown("""
        <div class="section-box">
        <h3>Welcome to Your Data Analysis Journey!</h3>
        <p>Our interactive dashboard empowers you to:</p>
        <ul>
            <li>Upload and process your WSP data effortlessly</li>
            <li>Visualize trends across different brands and regions</li>
            <li>Generate descriptive statistics and predictions</li>
            <li>Make data-driven decisions with confidence</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

    # How to use section
    st.markdown("""
    <div class="section-box">
    <h3>How to Use This Dashboard</h3>
    <ol>
        <li><strong>Upload Your Data:</strong> Start by uploading your Excel file containing the WSP data.</li>
        <li><strong>Enter Week Names:</strong> Provide names for each week column in your dataset.</li>
        <li><strong>Choose Your Analysis:</strong> Navigate to either the WSP Analysis Dashboard or Descriptive Statistics and Prediction sections.</li>
        <li><strong>Customize and Explore:</strong> Select your analysis parameters and generate valuable insights!</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)

    # File upload section
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.subheader("Upload Your Data")

    if 'file_processed' not in st.session_state:
        st.session_state.file_processed = False
    if 'file_ready' not in st.session_state:
        st.session_state.file_ready = False

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"],key="wsp_data")
    if 'edited_df' in st.session_state and 'edited_file_name' in st.session_state and not st.session_state.edited_df.empty:
        st.success(f"Edited file uploaded: {st.session_state.edited_file_name}")
        if st.button("Process Edited File", key="process_edited"):
            process_uploaded_file(st.session_state.edited_df)

    elif uploaded_file:
        st.success(f"File uploaded: {uploaded_file.name}")
        if st.button("Process Uploaded File", key="process_uploaded"):
            process_uploaded_file(uploaded_file)

    if st.session_state.file_ready:
        st.markdown("### Enter Week Names")
        num_weeks = st.session_state.num_weeks
        num_columns = min(4, num_weeks)  # Limit to 4 columns for better layout
        week_cols = st.columns(num_columns)

        for i in range(num_weeks):
            with week_cols[i % num_columns]:
                st.session_state.week_names_input[i] = st.text_input(
                    f'Week {i+1}', 
                    value=st.session_state.week_names_input[i],
                    key=f'week_{i}'
                )
        
        if st.button("Confirm Week Names", key="confirm_weeks"):
            if all(st.session_state.week_names_input):
                st.session_state.file_processed = True
                st.success("File processed successfully! You can now proceed to the analysis sections.")
            else:
                st.warning("Please fill in all week names before confirming.")

    if st.session_state.file_processed:
        st.success("File processed successfully! You can now proceed to the analysis sections.")
    else:
        st.info("Please upload a file and fill in all week names to proceed with the analysis.")

    st.markdown('</div>', unsafe_allow_html=True)

    # Help section
    st.markdown("""
    <div class="section-box">
    <h3>Need Assistance?</h3>
    <p>If you have any questions or need help using the dashboard, our support team is here for you. Don't hesitate to reach out!</p>
    <p>Email: prasoon.bajpai@lc.jkmail.com</p>
    <p>Phone: +91-9219393559</p>
    </div>
    """, unsafe_allow_html=True)

    # Footer
    st.markdown("""
    <div style="text-align: center; margin-top: 2rem; padding: 1rem; background-color: #34495e; color: #ecf0f1;">
    <p>© 2024 WSP Analysis Dashboard. All rights reserved.</p>
    </div>
    """, unsafe_allow_html=True)

def process_uploaded_file(uploaded_file):
    if (isinstance(uploaded_file, pd.DataFrame) or uploaded_file) and not st.session_state.file_processed:
        try:
            if isinstance(uploaded_file, pd.DataFrame):
                # Convert DataFrame to Excel file in memory
                buffer = BytesIO()
                uploaded_file.to_excel(buffer, index=False)
                buffer.seek(0)
                file_content = buffer.getvalue()
            else:
                file_content = uploaded_file.read()

            # Load workbook to check for hidden columns
            wb = openpyxl.load_workbook(BytesIO(file_content))
            ws = wb.active
            hidden_cols = [idx for idx, col in enumerate(ws.column_dimensions, 1) if ws.column_dimensions[col].hidden]

            # Read Excel file with header=2 for both cases
            df = pd.read_excel(BytesIO(file_content), header=2)
            df = df.dropna(axis=1, how='all')
            df = df.drop(columns=df.columns[hidden_cols], errors='ignore')

            if df.empty:
                st.error("The uploaded file resulted in an empty dataframe. Please check the file content.")
            else:
                st.session_state.df = df
                brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
                brand_columns = [col for col in st.session_state.df.columns if any(brand in str(col) for brand in brands)]
                num_weeks = len(brand_columns) // len(brands)
                
                if num_weeks > 0:
                    if 'week_names_input' not in st.session_state or len(st.session_state.week_names_input) != num_weeks:
                        st.session_state.week_names_input = [''] * num_weeks
                    
                    st.session_state.num_weeks = num_weeks
                    st.session_state.file_ready = True
                else:
                    st.warning("No weeks detected in the uploaded file. Please check the file content.")
                    st.session_state.week_names_input = []
                    st.session_state.file_processed = False
        except Exception as e:
            st.error(f"Error processing file: {e}")
            st.exception(e)
            st.session_state.file_processed = False
import streamlit as st
from streamlit_option_menu import option_menu
def wsp_analysis_dashboard():
    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: brown;
        text-align: center;
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 30px;
        font-family: 'Arial', sans-serif;
    }
    .title span {
        background: linear-gradient(45deg, #3366cc, #6699ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .section-box {
        background-color: #f9f9f9;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .stSelectbox, .stMultiSelect {
        background-color: white;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .stButton>button {
        border-radius: 20px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.05);
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title"><span>WSP Analysis Dashboard</span></div>', unsafe_allow_html=True)

    if not st.session_state.file_processed:
        st.warning("Please upload a file and fill in all week names in the Home section before using this dashboard.")
        return

    st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names_input)
    
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Analysis Settings")
    
    st.session_state.diff_week = st.slider("Select Week for Difference Calculation", 
                                           min_value=0, 
                                           max_value=len(st.session_state.week_names_input) - 1, 
                                           value=st.session_state.diff_week, 
                                           key="diff_week_slider") 
    download_pdf = st.checkbox("Download Plots as PDF", value=True)   
    col1, col2 = st.columns(2)
    with col1:
        zone_names = st.session_state.df["Zone"].unique().tolist()
        selected_zone = st.selectbox("Select Zone", zone_names, key="zone_select")
    with col2:
        filtered_df = st.session_state.df[st.session_state.df["Zone"] == selected_zone]
        region_names = filtered_df["REGION"].unique().tolist()
        selected_region = st.selectbox("Select Region", region_names, key="region_select")
        
    filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
    district_names = filtered_df["Dist Name"].unique().tolist()

    # Define recommended settings for each region
    region_recommendations = {
        "Gujarat": {
            "districts": ["Ahmadabad", "Mahesana", "Rajkot", "Vadodara", "Surat"],
            "benchmarks": ["UTCL", "Wonder"],
            "diffs": {"UTCL": -10.0, "Wonder": 0.0}
        },
        "Chhattisgarh": {
            "districts": ["Durg", "Raipur", "Bilaspur", "Raigarh", "Rajnandgaon"],
            "benchmarks": ["UTCL"],
            "diffs": {"UTCL": -10.0}
        },
        "Maharashtra(East)": {
            "districts": ["Nagpur", "Gondiya"],
            "benchmarks": ["UTCL"],
            "diffs": {"UTCL": -10.0}
        },
        "Odisha": {
            "districts": ["Cuttack", "Sambalpur", "Khorda"],
            "benchmarks": ["UTCL"],
            "diffs": {"UTCL": {"Sambalpur": -25.0, "Cuttack": -15.0, "Khorda": -15.0}}
        },
        "Rajasthan": {
            "districts": ["Alwar", "Jodhpur", "Udaipur", "Jaipur", "Kota", "Bikaner"],
            "benchmarks": [],
            "diffs": {}
        },
        "Madhya Pradesh(West)": {
            "districts": ["Indore", "Neemuch", "Ratlam", "Dhar"],
            "benchmarks": [],
            "diffs": {}
        },
        "Madhya Pradesh(East)": {
            "districts": ["Jabalpur", "Balaghat", "Chhindwara"],
            "benchmarks": [],
            "diffs": {}
        },
        "North-I": {
            "districts": ["East", "Gurugram", "Sonipat", "Hisar", "Yamunanagar", "Bathinda"],
            "benchmarks": [],
            "diffs": {}
        },
        "North-II": {
            "districts": ["Ghaziabad", "Meerut"],
            "benchmarks": [],
            "diffs": {}
        }
    }

    if selected_region in region_recommendations:
        recommended = region_recommendations[selected_region]
        suggested_districts = [d for d in recommended["districts"] if d in district_names]
        
        if suggested_districts:
            st.markdown(f"### Suggested Districts for {selected_region}")
            select_all = st.checkbox(f"Select all suggested districts for {selected_region}")
            
            if select_all:
                selected_districts = st.multiselect("Select District(s)", district_names, default=suggested_districts, key="district_select")
            else:
                selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    else:
        selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")

    st.markdown('</div>', unsafe_allow_html=True)

    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    benchmark_brands = [brand for brand in brands if brand != 'JKLC']
        
    benchmark_brands_dict = {}
    desired_diff_dict = {}
    if selected_districts:
        st.markdown("### Benchmark Settings")
        
        # First determine if recommendations are available
        has_recommendations = (
            selected_region in region_recommendations and 
            region_recommendations[selected_region]["benchmarks"]
        )
        
        if has_recommendations:
            use_recommended_benchmarks = st.checkbox(
                "Use recommended benchmarks and differences", 
                value=False
            )
        else:
            use_recommended_benchmarks = False
            
        if use_recommended_benchmarks:
            # Automatically use recommended settings without showing additional controls
            for district in selected_districts:
                benchmark_brands_dict[district] = region_recommendations[selected_region]["benchmarks"]
                desired_diff_dict[district] = {}
                
                if selected_region == "Odisha":
                    # Handle Odisha's district-specific differences
                    for brand in benchmark_brands_dict[district]:
                        if brand in region_recommendations[selected_region]["diffs"]:
                            desired_diff_dict[district][brand] = float(
                                region_recommendations[selected_region]["diffs"][brand].get(district, 0.0)
                            )
                else:
                    # Handle other regions' differences
                    for brand in benchmark_brands_dict[district]:
                        desired_diff_dict[district][brand] = float(
                            region_recommendations[selected_region]["diffs"].get(brand, 0.0)
                        )
                        
        else:
            # Show manual configuration options
            use_same_benchmarks = st.checkbox("Use same benchmarks for all districts", value=True)
            
            if use_same_benchmarks:
                selected_benchmarks = st.multiselect(
                    "Select Benchmark Brands for all districts", 
                    benchmark_brands, 
                    key="unified_benchmark_select"
                )
                
                for district in selected_districts:
                    benchmark_brands_dict[district] = selected_benchmarks
                    desired_diff_dict[district] = {}

                if selected_benchmarks:
                    st.markdown("#### Desired Differences")
                    num_cols = min(len(selected_benchmarks), 3)
                    diff_cols = st.columns(num_cols)
                    
                    for i, brand in enumerate(selected_benchmarks):
                        with diff_cols[i % num_cols]:
                            value = st.number_input(
                                f"{brand}",
                                min_value=-100.0,
                                value=0.0,
                                step=0.1,
                                format="%.1f",
                                key=f"unified_{brand}"
                            )
                            
                            for district in selected_districts:
                                desired_diff_dict[district][brand] = float(value)
                else:
                    st.warning("Please select at least one benchmark brand.")
            else:
                for district in selected_districts:
                    st.subheader(f"Settings for {district}")
                    
                    benchmark_brands_dict[district] = st.multiselect(
                        f"Select Benchmark Brands for {district}",
                        benchmark_brands,
                        key=f"benchmark_select_{district}"
                    )
                    
                    desired_diff_dict[district] = {}
                    
                    if benchmark_brands_dict[district]:
                        num_cols = min(len(benchmark_brands_dict[district]), 3)
                        diff_cols = st.columns(num_cols)
                        for i, brand in enumerate(benchmark_brands_dict[district]):
                            with diff_cols[i % num_cols]:
                                desired_diff_dict[district][brand] = st.number_input(
                                    f"{brand}",
                                    min_value=-100.0,
                                    value=0.0,
                                    step=0.1,
                                    format="%.1f",
                                    key=f"{district}_{brand}"
                                )
                    else:
                        st.warning(f"No benchmark brands selected for {district}.")

    st.markdown("### Generate Analysis")
    
    if st.button('Generate Plots', key='generate_plots', use_container_width=True):
        with st.spinner('Generating plots...'):
            plot_district_graph(filtered_df, selected_districts, benchmark_brands_dict, 
                              desired_diff_dict, 
                              st.session_state.week_names_input, 
                              st.session_state.diff_week, 
                              download_pdf)
            st.success('Plots generated successfully!')
    else:
        st.warning("Please upload a file in the Home section before using this dashboard.")
def descriptive_statistics_and_prediction():
    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: #3366cc;
        text-align: center;
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 30px;
        font-family: 'Arial', sans-serif;
    }
    .title span {
        background: linear-gradient(45deg, #3366cc, #6699ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .section-box {
        background-color: #f9f9f9;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .stSelectbox, .stMultiSelect {
        background-color: white;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .stButton>button {
        border-radius: 20px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.05);
    }
    .stats-box {
        background-color: #e6f3ff;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title"><span>Descriptive Statistics and Prediction</span></div>', unsafe_allow_html=True)

    if not st.session_state.file_processed:
        st.warning("Please upload a file in the Home section before using this feature.")
        return

    st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names_input)

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Analysis Settings")
    col1, col2 = st.columns(2)
    with col1:
        zone_names = st.session_state.df["Zone"].unique().tolist()
        selected_zone = st.selectbox("Select Zone", zone_names, key="stats_zone_select")
    with col2:
        filtered_df = st.session_state.df[st.session_state.df["Zone"] == selected_zone]
        region_names = filtered_df["REGION"].unique().tolist()
        selected_region = st.selectbox("Select Region", region_names, key="stats_region_select")
    

    filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
    district_names = filtered_df["Dist Name"].unique().tolist()
    
    if selected_region in ["Rajasthan", "Madhya Pradesh(West)","Madhya Pradesh(East)","Chhattisgarh","Maharashtra(East)","Odisha","North-I","North-II","Gujarat"]:
        suggested_districts = []
        
        if selected_region == "Rajasthan":
            rajasthan_districts = ["Alwar", "Jodhpur", "Udaipur", "Jaipur", "Kota", "Bikaner"]
            suggested_districts = [d for d in rajasthan_districts if d in district_names]
        elif selected_region == "Madhya Pradesh(West)":
            mp_west_districts = ["Indore", "Neemuch","Ratlam","Dhar"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Madhya Pradesh(East)":
            mp_west_districts = ["Jabalpur","Balaghat","Chhindwara"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Chhattisgarh":
            mp_west_districts = ["Durg","Raipur","Bilaspur","Raigarh","Rajnandgaon"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Maharashtra(East)":
            mp_west_districts = ["Nagpur","Gondiya"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Odisha":
            mp_west_districts = ["Cuttack","Sambalpur","Khorda"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "North-I":
            mp_west_districts = ["East","Gurugram","Sonipat","Hisar","Yamunanagar","Bathinda"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "North-II":
            mp_west_districts = ["Ghaziabad","Meerut"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Gujarat":
            mp_west_districts = ["Ahmadabad","Mahesana","Rajkot","Vadodara","Surat"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        
        
        
        if suggested_districts:
            st.markdown(f"### Suggested Districts for {selected_region}")
            select_all = st.checkbox(f"Select all suggested districts for {selected_region}")
            
            if select_all:
                selected_districts = st.multiselect("Select District(s)", district_names, default=suggested_districts, key="district_select")
            else:
                selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
        else:
            selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    else:
        selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    

    st.markdown('</div>', unsafe_allow_html=True)


    if selected_districts:
        # Add a button to download all stats and predictions in one PDF
        if len(selected_districts) > 1:
            if st.checkbox("Download All Stats and Predictions",value=True):
                all_stats_pdf = BytesIO()
                pdf = SimpleDocTemplate(all_stats_pdf, pagesize=letter)
                elements = []
                
                for district in selected_districts:
                    elements.append(Paragraph(f"Statistics and Predictions for {district}", getSampleStyleSheet()['Title']))
                    district_df = filtered_df[filtered_df["Dist Name"] == district]
                    
                    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
                    stats_data = {}
                    prediction_data = {}
                    
                    for brand in brands:
                        brand_data = district_df[[col for col in district_df.columns if brand in col]].values.flatten()
                        brand_data = brand_data[~np.isnan(brand_data)]
                        
                        if len(brand_data) > 0:
                            stats_data[brand] = pd.DataFrame({
                                'Mean': [np.mean(brand_data)],
                                'Median': [np.median(brand_data)],
                                'Std Dev': [np.std(brand_data)],
                                'Min': [np.min(brand_data)],
                                'Max': [np.max(brand_data)],
                                'Skewness': [stats.skew(brand_data)],
                                'Kurtosis': [stats.kurtosis(brand_data)],
                                'Range': [np.ptp(brand_data)],
                                'IQR': [np.percentile(brand_data, 75) - np.percentile(brand_data, 25)]
                            }).iloc[0]

                            if len(brand_data) > 2:
                                model = ARIMA(brand_data, order=(1,1,1))
                                model_fit = model.fit()
                                forecast = model_fit.forecast(steps=1)
                                confidence_interval = model_fit.get_forecast(steps=1).conf_int()
                                prediction_data[brand] = {
                                    'forecast': forecast[0],
                                    'lower_ci': confidence_interval[0, 0],
                                    'upper_ci': confidence_interval[0, 1]
                                }
                    
                    elements.append(Paragraph("Descriptive Statistics", getSampleStyleSheet()['Heading2']))
                    elements.append(create_stats_table(stats_data))
                    elements.append(Paragraph("Price Predictions", getSampleStyleSheet()['Heading2']))
                    elements.append(create_prediction_table(prediction_data))
                    elements.append(PageBreak())
                
                pdf.build(elements)
                st.download_button(
                    label="Download All Stats and Predictions PDF",
                    data=all_stats_pdf.getvalue(),
                    file_name=f"{selected_districts}stats_and_predictions.pdf",
                    mime="application/pdf"
                )

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.markdown("### Descriptive Statistics")
        
        for district in selected_districts:
            st.subheader(f"{district}")
            district_df = filtered_df[filtered_df["Dist Name"] == district]
            
            brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            stats_data = {}
            prediction_data = {}
            
            for brand in brands:
                st.markdown(f'<div class="stats-box">', unsafe_allow_html=True)
                st.markdown(f"#### {brand}")
                brand_data = district_df[[col for col in district_df.columns if brand in col]].values.flatten()
                brand_data = brand_data[~np.isnan(brand_data)]
                
                if len(brand_data) > 0:
                    basic_stats = pd.DataFrame({
                        'Mean': [np.mean(brand_data)],
                        'Median': [np.median(brand_data)],
                        'Std Dev': [np.std(brand_data)],
                        'Min': [np.min(brand_data)],
                        'Max': [np.max(brand_data)],
                        'Skewness': [stats.skew(brand_data)],
                        'Kurtosis': [stats.kurtosis(brand_data)],
                        'Range': [np.ptp(brand_data)],
                        'IQR': [np.percentile(brand_data, 75) - np.percentile(brand_data, 25)]
                    })
                    st.dataframe(basic_stats)
                    stats_data[brand] = basic_stats.iloc[0]

                    # ARIMA prediction for next week
                    if len(brand_data) > 2:  # Need at least 3 data points for ARIMA
                        model = ARIMA(brand_data, order=(1,1,1))
                        model_fit = model.fit()
                        forecast = model_fit.forecast(steps=1)
                        confidence_interval = model_fit.get_forecast(steps=1).conf_int()
                        st.markdown(f"Predicted price for next week: {forecast[0]:.2f}")
                        st.markdown(f"95% Confidence Interval: [{confidence_interval[0, 0]:.2f}, {confidence_interval[0, 1]:.2f}]")
                        prediction_data[brand] = {
                            'forecast': forecast[0],
                            'lower_ci': confidence_interval[0, 0],
                            'upper_ci': confidence_interval[0, 1]
                        }
                else:
                    st.warning(f"No data available for {brand} in this district.")
                st.markdown('</div>', unsafe_allow_html=True)

            # Create download buttons for stats and predictions
            stats_pdf = create_stats_pdf(stats_data, district)
            predictions_pdf = create_prediction_pdf(prediction_data, district)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download Statistics PDF",
                    data=stats_pdf,
                    file_name=f"{district}_statistics.pdf",
                    mime="application/pdf"
                )
            with col2:
                st.download_button(
                    label="Download Predictions PDF",
                    data=predictions_pdf,
                    file_name=f"{district}_predictions.pdf",
                    mime="application/pdf"
                )
        st.markdown('</div>', unsafe_allow_html=True)
def create_stats_table(stats_data):
    data = [['Brand', 'Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']]
    for brand, stats in stats_data.items():
        row = [brand]
        for stat in ['Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']:
            value = stats[stat]
            if isinstance(value, (int, float)):
                row.append(f"{value:.2f}")
            else:
                row.append(str(value))
        data.append(row)
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    return table

def create_prediction_table(prediction_data):
    data = [['Brand', 'Predicted Price', 'Lower CI', 'Upper CI']]
    for brand, pred in prediction_data.items():
        row = [brand, f"{pred['forecast']:.2f}", f"{pred['lower_ci']:.2f}", f"{pred['upper_ci']:.2f}"]
        data.append(row)
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    return table
from urllib.parse import quote
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df = df.fillna(0)
    regions = df['Zone'].unique().tolist()
    brands = df['Brand'].unique().tolist()
    return df, regions, brands
# Cache the model training
@st.cache_resource

def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
import plotly.subplots as sp
from scipy import stats
import matplotlib.image as mpimg
from PIL import Image
def create_visualization(region_data, region, brand, months):
    fig = plt.figure(figsize=(20, 34))
    gs = fig.add_gridspec(8, 3, height_ratios=[0.5,1, 1, 3, 2.25, 2, 2, 2])
    # Region and Brand Title
    ax_region = fig.add_subplot(gs[0, :])
    ax_region.axis('off')
    ax_region.text(0.5, 0.5, f'{region} ({brand})', fontsize=28, fontweight='bold', ha='center', va='center')
    
    # New layout for current month sales data
    ax_current = fig.add_subplot(gs[1, :])
    ax_current.axis('off')
    
    # Calculate values
    overall_oct = region_data['Monthly Achievement(Oct)'].iloc[-1]
    trade_oct = region_data['Trade Oct'].iloc[-1]
    non_trade_oct = overall_oct - trade_oct
    table_data_left = [
    ['AGS Target', f"{region_data['AGS Tgt (Oct)'].iloc[-1]:.0f}"],
    ['Plan', f"{region_data['Month Tgt (Oct)'].iloc[-1]:.0f}"],
    ['Trade Target', f"{region_data['Trade Tgt (Oct)'].iloc[-1]:.0f}"],
    ['Non-Trade Target', f"{region_data['Non-Trade Tgt (Oct)'].iloc[-1]:.0f}"]]
    table_data_right = [
    [f"{overall_oct:.0f}"],
    [f"{trade_oct:.0f}"],
    [f"{non_trade_oct:.0f}"]]
    ax_current.text(0.225, 0.9, 'Targets', fontsize=12, fontweight='bold', ha='center')
    ax_current.text(0.35, 0.9, 'Achievement', fontsize=12, fontweight='bold', ha='center')
    table_left = ax_current.table(
    cellText=table_data_left,
    cellLoc='center',
    loc='center',
    bbox=[0, 0.0, 0.3, 0.8]) 
    table_left.auto_set_font_size(False)
    table_left.set_fontsize(12)
    table_left.scale(1.2, 1.8)

    for i in range(len(table_data_left)):
      cell = table_left[i, 0]
      cell.set_facecolor('#ECF0F1')
      cell.set_text_props(fontweight='bold')
      cell = table_left[i, 1]
      cell.set_facecolor('#F7F9F9')
      cell.set_text_props(fontweight='bold')
    table_right = ax_current.table(
    cellText=table_data_right,
    cellLoc='center',
    loc='center',
    bbox=[0.3, 0.0, 0.1, 0.8])  
    cell = table_right.add_cell(0, 0,1, 2, 
                           text=f"{overall_oct:.0f}",
                           facecolor='#E8F6F3')
    cell.set_text_props(fontweight='bold')
    cell.set_text_props(fontweight='bold')
    cell = table_right.add_cell(1, 0, 1, 1,
                           text=f"{trade_oct:.0f}",
                           facecolor='#E8F6F3')
    cell.set_text_props(fontweight='bold')
    cell = table_right.add_cell(2, 0, 1, 1,
                           text=f"{non_trade_oct:.0f}",
                           facecolor='#E8F6F3')
    cell.set_text_props(fontweight='bold')
    table_right.auto_set_font_size(False)
    table_right.set_fontsize(13)
    table_right.scale(1.2, 1.8)
    ax_current.text(0.2, 1.0, 'October 2024 Performance Metrics', 
               fontsize=16, fontweight='bold', ha='center', va='bottom')
    # Modify the detailed metrics section
    detailed_metrics = [
    ('Trade', region_data['Trade Oct'].iloc[-1], region_data['Monthly Achievement(Oct)'].iloc[-1], 'Channel'),
    ('Green', region_data['Green Oct'].iloc[-1], region_data['Monthly Achievement(Oct)'].iloc[-1], 'Region'),
    ('Yellow', region_data['Yellow Oct'].iloc[-1], region_data['Monthly Achievement(Oct)'].iloc[-1], 'Region'),
    ('Red', region_data['Red Oct'].iloc[-1], region_data['Monthly Achievement(Oct)'].iloc[-1], 'Region'),
    ('Premium', region_data['Premium Oct'].iloc[-1], region_data['Monthly Achievement(Oct)'].iloc[-1], 'Product'),
    ('Blended', region_data['Blended Till Now Oct'].iloc[-1], region_data['Monthly Achievement(Oct)'].iloc[-1], 'Product')]
    colors = ['gold', 'green', 'yellow', 'red', 'darkmagenta', 'saddlebrown']
    
    # Add boxes for grouping metrics
    # Box 1 for Trade
    trade_box = patches.Rectangle((0.45, 0.74), 0.55, 0.125, 
                                facecolor='#F0F0F0', 
                                edgecolor='black',
                                alpha=1,
                                transform=ax_current.transAxes)
    ax_current.add_patch(trade_box)
    
    # Box 2 for Region types (Green, Yellow, Red)
    region_box = patches.Rectangle((0.45, 0.35), 0.55, 0.375,
                                 facecolor='#F0F0F0',
                                 edgecolor='black',
                                 alpha=1,
                                 transform=ax_current.transAxes)
    ax_current.add_patch(region_box)
    
    # Box 3 for Products (Premium, Blended)
    product_box = patches.Rectangle((0.45, 0.08), 0.55, 0.25,
                                  facecolor='#F0F0F0',
                                  edgecolor='black',
                                  alpha=1,
                                  transform=ax_current.transAxes)
    ax_current.add_patch(product_box)
    for i, (label, value, total, category) in enumerate(detailed_metrics):
     percentage = (value / total) * 100 if total != 0 else 0
     if i == 0:  
        y_pos = 0.77
     elif i <= 3:  
        y_pos = 0.63 - (i-1) * 0.11
     else:  
        y_pos = 0.24 - (i-4) * 0.11
     if category == 'Region' and value == 0:
        text = f'• {label} region not present'
     else:
        text = f'• {label} {category} has a share of {percentage:.1f}% in total sales, i.e., {value:.0f} MT.'
        ax_current.text(0.50, y_pos, text, fontsize=14, fontweight="bold", color=colors[i])
    ax_current.text(0.50, 0.90, 'Sales Breakown', fontsize=16, fontweight='bold', ha='center', va='bottom')
    ax_table = fig.add_subplot(gs[2, :])
    ax_table.axis('off')
    ax_table.set_title(f"Quarterly Requirement for November and Decemeber 2024", fontsize=18, fontweight='bold')
    table_data = [
                ['Overall\nRequirement', 'Trade Channel\nRequirement', 'Premium Product\nRequirement','Blended Product\nRequirement'],
                [f"{region_data['Q3 2023 Total'].iloc[-1]-region_data['Monthly Achievement(Oct)'].iloc[-1]:.0f}", f"{region_data['Q3 2023 Trade'].iloc[-1]-region_data['Trade Oct'].iloc[-1]:.0f}",f"{region_data['Q3 2023 Premium'].iloc[-1]-region_data['Premium Oct'].iloc[-1]:.0f}", 
                 f"{region_data['Q3 2023 Blended '].iloc[-1]-region_data['Blended Till Now Oct'].iloc[-1]:.0f}"],
            ]
    table = ax_table.table(cellText=table_data[1:], colLabels=table_data[0], cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(12)
    table.scale(1, 1.7)
    for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold', color='black')
                    cell.set_facecolor('goldenrod')
                cell.set_edgecolor('brown')
    # Main bar chart (same as before)
    ax1 = fig.add_subplot(gs[3, :])
    actual_ags = [region_data[f'AGS Tgt ({month})'].iloc[-1] for month in months]
    actual_achievements = [region_data[f'Monthly Achievement({month})'].iloc[-1] for month in months]
    actual_targets = [region_data[f'Month Tgt ({month})'].iloc[-1] for month in months]
    
    x = np.arange(len(months))
    width = 0.25
    rects1 = ax1.bar(x-width, actual_ags, width, label='AGS Target', color='brown', alpha=0.8)
    rects2 = ax1.bar(x, actual_targets, width, label='Plan', color='purple', alpha=0.8)
    rects3 = ax1.bar(x + width, actual_achievements, width, label='Achievement', color='yellow', alpha=0.8)
    
    ax1.set_ylabel('Targets and Achievement', fontsize=12, fontweight='bold')
    ax1.set_title(f"Monthly Targets and Achievements for FY 2025", fontsize=18, fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(months)
    ax1.legend()
    
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            ax1.annotate(f'{height:.0f}',
                        xy=(rect.get_x() + rect.get_width() / 3, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom', fontsize=11)
    
    autolabel(rects1)
    autolabel(rects2)
    autolabel(rects3)
    ax2 = fig.add_subplot(gs[4, :])
    percent_achievements_plan = [((ach / tgt) * 100) for ach, tgt in zip(actual_achievements, actual_targets)]
    percent_achievements_ags = [((ach / ags) * 100) for ach, ags in zip(actual_achievements, actual_ags)]
    
    # Plot both lines
    line1 = ax2.plot(x, percent_achievements_plan, marker='o', linestyle='-', color='purple', label='Achievement vs Plan')
    line2 = ax2.plot(x, percent_achievements_ags, marker='s', linestyle='-', color='brown', label='Achievement vs AGS')
    ax2.axhline(y=100, color='lightcoral', linestyle='--', alpha=0.7)
    
    ax2.set_xlabel('Month', fontsize=12, fontweight='bold')
    ax2.set_ylabel('% Achievement', fontsize=12, fontweight='bold')
    ax2.set_xticks(x)
    ax2.set_xticklabels(months)
    ax2.legend(loc='upper right')
    
    # Add annotations with dynamic positioning
    for i, (pct_plan, pct_ags) in enumerate(zip(percent_achievements_plan, percent_achievements_ags)):
        # Determine which value is higher
        if pct_plan >= pct_ags:
            # Plan is higher or equal, put Plan above and AGS below
            ax2.annotate(f'{pct_plan:.1f}%', 
                        (i, pct_plan), 
                        xytext=(0, 10), 
                        textcoords='offset points', 
                        ha='center', 
                        va='bottom', 
                        fontsize=12,
                        color='purple')
            
            ax2.annotate(f'{pct_ags:.1f}%', 
                        (i, pct_ags), 
                        xytext=(0, -15), 
                        textcoords='offset points', 
                        ha='center', 
                        va='top', 
                        fontsize=12,
                        color='brown')
        else:
            # AGS is higher, put AGS above and Plan below
            ax2.annotate(f'{pct_ags:.1f}%', 
                        (i, pct_ags), 
                        xytext=(0, 10), 
                        textcoords='offset points', 
                        ha='center', 
                        va='bottom', 
                        fontsize=12,
                        color='brown')
            
            ax2.annotate(f'{pct_plan:.1f}%', 
                        (i, pct_plan), 
                        xytext=(0, -15), 
                        textcoords='offset points', 
                        ha='center', 
                        va='top', 
                        fontsize=12,
                        color='purple')
    ax3 = fig.add_subplot(gs[5, :])
    ax3.axis('off')
    current_year = 2024
    last_year = 2023

    channel_data = [
        ('Trade', region_data['Trade Oct'].iloc[-1], region_data['Trade Oct 2023'].iloc[-1],'Channel'),
        ('Premium', region_data['Premium Oct'].iloc[-1], region_data['Premium Oct 2023'].iloc[-1],'Product'),
        ('Blended', region_data['Blended Till Now Oct'].iloc[-1], region_data['Blended Oct 2023'].iloc[-1],'Product')
    ]
    monthly_achievement_oct = region_data['Monthly Achievement(Oct)'].iloc[-1]
    total_oct_current = region_data['Monthly Achievement(Oct)'].iloc[-1]
    total_oct_last = region_data['Total Oct 2023'].iloc[-1]
    
    ax3.text(0.2, 1, f'October {current_year} Sales Comparison to October 2023:-', fontsize=16, fontweight='bold', ha='center', va='center')
    
    def get_arrow(value):
        return '↑' if value > 0 else '↓' if value < 0 else '→'
    def get_color(value):
        return 'green' if value > 0 else 'red' if value < 0 else 'black'

    total_change = ((total_oct_current - total_oct_last) / total_oct_last) * 100
    arrow = get_arrow(total_change)
    color = get_color(total_change)
    ax3.text(0.21, 0.9, f"October 2024: {total_oct_current:.0f}", fontsize=14, fontweight='bold', ha='center')
    ax3.text(0.22, 0.85, f"vs October 2023: {total_oct_last:.0f} ({total_change:.1f}% {arrow})", fontsize=12, color=color, ha='center')
    
    for i, (channel, value_current, value_last,x) in enumerate(channel_data):
        percentage = (value_current / monthly_achievement_oct) * 100
        percentage_last_year = (value_last / total_oct_last) * 100
        change = ((value_current - value_last) / value_last) * 100
        arrow = get_arrow(change)
        color = get_color(change)
        
        y_pos = 0.75 - i*0.25
        ax3.text(0.15, y_pos, f"{channel}:", fontsize=14, fontweight='bold')
        ax3.text(0.28, y_pos, f"{value_current:.0f}", fontsize=14)
        ax3.text(0.15, y_pos-0.05, f"vs Last Year: {value_last:.0f}", fontsize=12)
        ax3.text(0.28, y_pos-0.05, f"({change:.1f}% {arrow})", fontsize=12, color=color)
        # Add the share percentage comparison
        ax3.text(0.12, y_pos-0.1, 
                f"•{channel} {x} has share of {percentage_last_year:.1f}% in Oct. last year as compared to {percentage:.1f}% in Oct. this year.",
                fontsize=11, color='darkcyan')

    # Update the September comparison section similarly
    ax4 = fig.add_subplot(gs[5, 2])
    ax4.axis('off')
    
    channel_data1 = [
        ('Trade', region_data['Trade Oct'].iloc[-1], region_data['Trade Sep'].iloc[-1],'Channel'),
        ('Premium', region_data['Premium Oct'].iloc[-1], region_data['Premium Sep'].iloc[-1],'Product'),
        ('Blended', region_data['Blended Till Now Oct'].iloc[-1], region_data['Blended Sep'].iloc[-1],'Product')
    ]
    total_sep_current = region_data['Total Sep '].iloc[-1]
    
    ax4.text(0.35, 1, f'October {current_year} Sales Comparison to September 2024:-', fontsize=16, fontweight='bold', ha='center', va='center')
    
    total_change = ((total_oct_current - total_sep_current) / total_sep_current) * 100
    arrow = get_arrow(total_change)
    color = get_color(total_change)
    ax4.text(0.36, 0.9, f"October 2024: {total_oct_current:.0f}", fontsize=14, fontweight='bold', ha='center')
    ax4.text(0.37, 0.85, f"vs September 2024: {total_sep_current:.0f} ({total_change:.1f}% {arrow})", fontsize=12, color=color, ha='center')
    
    for i, (channel, value_current, value_last,t) in enumerate(channel_data1):
        percentage = (value_current / monthly_achievement_oct) * 100
        percentage_last_month = (value_last / total_sep_current) * 100
        change = ((value_current - value_last) / value_last) * 100
        arrow = get_arrow(change)
        color = get_color(change)
        
        y_pos = 0.75 - i*0.25
        ax4.text(0.10, y_pos, f"{channel}:", fontsize=14, fontweight='bold')
        ax4.text(0.65, y_pos, f"{value_current:.0f}", fontsize=14)
        ax4.text(0.10, y_pos-0.05, f"vs Last Month: {value_last:.0f}", fontsize=12)
        ax4.text(0.65, y_pos-0.05, f"({change:.1f}% {arrow})", fontsize=12, color=color)
        # Add the share percentage comparison
        ax4.text(0.00, y_pos-0.1, 
                f"•{channel} {t} has share of {percentage_last_month:.1f}% in Sept. as compared to {percentage:.1f}% in Oct.",
                fontsize=11, color='darkcyan')
    # Updated: August Region Type Breakdown with values
    def create_pie_data(data_values, labels, colors):
     non_zero_data = []
     non_zero_labels = []
     non_zero_colors = []
     for value, label, color in zip(data_values, labels, colors):
        if value > 0:
            non_zero_data.append(value)
            non_zero_labels.append(label)
            non_zero_colors.append(color)       
     return non_zero_data, non_zero_labels, non_zero_colors

    def make_autopct(values):
     def my_autopct(pct):
        total = sum(values)
        val = int(round(pct*total/100.0))
        return f'{pct:.0f}%\n({val:.0f})'
     return my_autopct
    ax5 = fig.add_subplot(gs[6, 0])
    region_type_data = [
    region_data['Green Oct'].iloc[-1],
    region_data['Yellow Oct'].iloc[-1],
    region_data['Red Oct'].iloc[-1],
    region_data['Unidentified Oct'].iloc[-1]]
    region_type_labels = ['G', 'Y', 'R', '']
    colors = ['green', 'yellow', 'red', 'gray']
    filtered_data, filtered_labels, filtered_colors = create_pie_data(
    region_type_data, region_type_labels, colors)
    explode = [0.05] * len(filtered_data)
    ax5.pie(filtered_data, labels=filtered_labels, colors=filtered_colors,
        autopct=make_autopct(filtered_data), startangle=90, explode=explode)
    ax5.set_title('October 2024 Region Type Breakdown:-', fontsize=16, fontweight='bold')
    ax6 = fig.add_subplot(gs[6, 1])
    region_type_data = [
    region_data['Green Oct 2023'].iloc[-1],
    region_data['Yellow Oct 2023'].iloc[-1],
    region_data['Red Oct 2023'].iloc[-1],
    region_data['Unidentified Oct 2023'].iloc[-1]]
    filtered_data, filtered_labels, filtered_colors = create_pie_data(
    region_type_data, region_type_labels, colors)
    explode = [0.05] * len(filtered_data)
    ax6.pie(filtered_data, labels=filtered_labels, colors=filtered_colors,
        autopct=make_autopct(filtered_data), startangle=90, explode=explode)
    ax6.set_title('October 2023 Region Type Breakdown:-', fontsize=16, fontweight='bold')
    ax7 = fig.add_subplot(gs[6, 2])
    region_type_data = [
    region_data['Green Sep'].iloc[-1],
    region_data['Yellow Sep'].iloc[-1],
    region_data['Red Sep'].iloc[-1],
    region_data['Unidentified Sep'].iloc[-1]]
    filtered_data, filtered_labels, filtered_colors = create_pie_data(
    region_type_data, region_type_labels, colors)
    explode = [0.05] * len(filtered_data)
    ax7.pie(filtered_data, labels=filtered_labels, colors=filtered_colors,
        autopct=make_autopct(filtered_data), startangle=90, explode=explode)
    ax7.set_title('September 2024 Region Type Breakdown:-', fontsize=16, fontweight='bold')
    ax_comparison = fig.add_subplot(gs[7, :])
    ax_comparison.axis('off')
    ax_comparison.set_title('Quarterly Performance Analysis (2023 vs 2024)', 
                          fontsize=20, fontweight='bold', pad=20)

    def create_modern_quarterly_box(ax, x, y, width, height, q_data, quarter):
        # Adjust the background rectangle
        rect = patches.Rectangle(
            (x, y), width, height,
            facecolor='#f8f9fa',
            edgecolor='#dee2e6',
            linewidth=2,
            alpha=0.9,
            zorder=1
        )
        ax.add_patch(rect)
        
        # Title bar
        title_height = height * 0.15
        title_bar = patches.Rectangle(
            (x, y + height - title_height),
            width,
            title_height,
            facecolor='#4a90e2',
            alpha=0.9,
            zorder=2
        )
        ax.add_patch(title_bar)
        
        # Quarter title
        ax.text(x + width/2, y + height - title_height/2,
                f"{quarter} Performance Overview",
                ha='center', va='center',
                fontsize=14, fontweight='bold',
                color='white',
                zorder=3)

        # Calculate metrics
        total_2023, total_2024 = q_data['total_2023'], q_data['total_2024']
        pct_change = ((total_2024 - total_2023) / total_2023) * 100
        trade_2023, trade_2024 = q_data['trade_2023'], q_data['trade_2024']
        trade_pct_change = ((trade_2024 - trade_2023) / trade_2023) * 100
        
        # Add metric comparisons
        y_offset = y + height - title_height - 0.1
        
        # Total Sales
        ax.text(x + 0.05, y_offset,
                "Total Sales Comparison:",
                fontsize=14, fontweight='bold',
                color='#2c3e50')
        
        y_offset -= 0.08
        ax.text(x + 0.05, y_offset,
                f"2023: {total_2023:,.0f}",
                fontsize=11)
        ax.text(x + width/2, y_offset,
                f"2024: {total_2024:,.0f}",
                fontsize=11)
        ax.text(x + 0.375*width, y_offset,
                f"{pct_change:+.1f}%",
                fontsize=11,
                color='green' if pct_change > 0 else 'red')
        
        # Trade Volume
        y_offset -= 0.12
        ax.text(x + 0.05, y_offset,
                "Trade Volume:",
                fontsize=14, fontweight='bold',
                color='#2c3e50')
        
        y_offset -= 0.08
        ax.text(x + 0.05, y_offset,
                f"2023: {trade_2023:,.0f}",
                fontsize=11)
        ax.text(x + width/2, y_offset,
                f"2024: {trade_2024:,.0f}",
                fontsize=11)
        ax.text(x + 0.375*width, y_offset,
                f"{trade_pct_change:+.1f}%",
                fontsize=11,
                color='green' if trade_pct_change > 0 else 'red')
        
        # Add trend arrow
        if pct_change > 0:
            arrow_style = 'fancy,head_length=4,head_width=6'
            arrow_color = 'green'
            start_x = x + 0.11
            end_x = x + width * 0.49
            start_y = y + 0.31
            end_y = y + 0.31
        else:
            arrow_style = 'fancy,head_length=4,head_width=6'
            arrow_color = 'red'
            start_x = x + 0.11
            end_x = x + width * 0.49
            start_y = y + 0.31
            end_y = y + 0.31
            
        arrow = patches.FancyArrowPatch(
            (start_x, start_y),
            (end_x, end_y),
            arrowstyle=arrow_style,
            color=arrow_color,
            linewidth=2,
            zorder=3
        )
        ax.add_patch(arrow)
        if trade_pct_change > 0:
            arrow_style = 'fancy,head_length=4,head_width=6'
            arrow_color = 'green'
            start_x = x + 0.11
            end_x = x + width * 0.49
            start_y = y + 0.11
            end_y = y + 0.11
        else:
            arrow_style = 'fancy,head_length=4,head_width=6'
            arrow_color = 'red'
            start_x = x + 0.11
            end_x = x + width * 0.49
            start_y = y + 0.11
            end_y = y + 0.11
            
        arrow = patches.FancyArrowPatch(
            (start_x, start_y),
            (end_x, end_y),
            arrowstyle=arrow_style,
            color=arrow_color,
            linewidth=2,
            zorder=3
        )
        ax.add_patch(arrow)
    # Create quarterly comparison boxes with adjusted positions
    # Make boxes taller and adjust vertical position
    box_height = 0.6  # Increased height
    box_y = 0.2      # Adjusted vertical position
    
    q1_data = {
        'total_2023': region_data['Q1 2023 Total'].iloc[-1],
        'total_2024': region_data['Q1 2024 Total'].iloc[-1],
        'trade_2023': region_data['Q1 2023 Trade'].iloc[-1],
        'trade_2024': region_data['Q1 2024 Trade'].iloc[-1]
    }

    q2_data = {
        'total_2023': region_data['Q2 2023 Total'].iloc[-1],
        'total_2024': region_data['Q2 2024 Total'].iloc[-1],
        'trade_2023': region_data['Q2 2023 Trade'].iloc[-1],
        'trade_2024': region_data['Q2 2024 Trade'].iloc[-1]
    }

    # Adjust box positioning
    create_modern_quarterly_box(ax_comparison, 0.1, box_y, 0.35, box_height, q1_data, "Q1")
    create_modern_quarterly_box(ax_comparison, 0.55, box_y, 0.35, box_height, q2_data, "Q2")

    # Set the axis limits explicitly
    ax_comparison.set_xlim(0, 1)
    ax_comparison.set_ylim(0, 1)

    plt.tight_layout()
    return fig


def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
def generate_full_report(df, regions):
    from matplotlib.backends.backend_pdf import PdfPages
    import matplotlib.pyplot as plt
    from io import BytesIO
    pdf_buffer = BytesIO()
    with PdfPages(pdf_buffer) as pdf:
        # Iterate through each region
        for region in regions:
            # Get unique brands for this region
            region_brands = df[df['Zone'] == region]['Brand'].unique().tolist()
            for brand in region_brands:
                # Filter data for current region and brand
                region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)]
                months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
                fig = create_visualization(region_data, region, brand, months)
                pdf.savefig(fig)
                plt.close(fig)
    pdf_buffer.seek(0)
    return pdf_buffer
def show_welcome_page():
        st.markdown("# 📈 Sales Review Report Generator")
        st.markdown("""
        ### Transform Your Sales Data into Actionable Insights
        
        This advanced analytics platform helps you:
        - 📊 Generate comprehensive sales review reports
        - 🎯 Track performance across regions and brands
        - 📈 Visualize key metrics and trends
        - 🔄 Compare historical data
        """)
        
        st.markdown("""
        <div class='reportBlock'>
        <h3>🚀 Getting Started</h3>
        <p>Upload your Excel file to begin analyzing your sales data:</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Choose your Excel file", type="xlsx", key="Sales_Prediction_uploader")
        
        if uploaded_file:
            with st.spinner("Processing your data..."):
                progress_bar = st.progress(0)
                for i in range(100):
                    time.sleep(0.01)
                    progress_bar.progress(i + 1)
                
                df, regions, brands = load_data(uploaded_file)
                st.session_state['df'] = df
                st.session_state['regions'] = regions
                st.session_state['brands'] = brands
                
                st.success("✅ File processed successfully!")

def show_report_generator():
    st.markdown("# 🎯 Report Generator")
    
    if st.session_state.get('df') is None:
        st.warning("⚠️ Please upload your data file on the Home page first.")
        return
    
    df = st.session_state['df']
    regions = st.session_state['regions']
    
    # Create tabs for different report types
    tab1, tab2 = st.tabs(["📑 Individual Report", "📚 Complete Report"])
    
    with tab1:
            st.markdown("""
            <div class='reportBlock'>
            <h3>Report Parameters</h3>
            </div>
            """, unsafe_allow_html=True)
            
            region = st.selectbox("Select Region", regions, key='region_select')
            region_brands = df[df['Zone'] == region]['Brand'].unique().tolist()
            brand = st.selectbox("Select Brand", region_brands, key='brand_select')
            
            if st.button("🔍 Generate Individual Report", key='individual_report'):
                with st.spinner("Creating your report..."):
                    region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)]
                    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
                    fig = create_visualization(region_data, region, brand, months)
                    
                    st.pyplot(fig)
                    
                    buf = BytesIO()
                    fig.savefig(buf, format="pdf")
                    buf.seek(0)
                    
                    st.download_button(
                        label="📥 Download Individual Report (PDF)",
                        data=buf,
                        file_name=f"sales_report_{region}_{brand}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
    
    with tab2:
        st.markdown("""
        <div class='reportBlock'>
        <h3>Complete Report Generation</h3>
        <p>Generate a comprehensive report covering all regions and brands in your dataset.</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("📊 Generate Complete Report", key='complete_report'):
            with st.spinner("Generating comprehensive report... This may take a few minutes."):
                progress_bar = st.progress(0)
                for i in range(100):
                    time.sleep(0.02)
                    progress_bar.progress(i + 1)
                
                pdf_buffer = generate_full_report(df, regions)
                
                st.success("✅ Report generated successfully!")
                
                st.download_button(
                    label="📥 Download Complete Report (PDF)",
                    data=pdf_buffer,
                    file_name=f"complete_sales_report_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf"
                )

def show_about_page():
        st.markdown("# ℹ️ About")
        st.markdown("""
        <div class='reportBlock'>
        <h2>Sales Review Report Generator Pro</h2>
        <p>Version 2.0 | Last Updated: October 2024</p>
        
        <h3>🎯 Purpose</h3>
        Our advanced analytics platform empowers sales teams to:
        - Generate detailed performance reports
        - Track KPIs across regions and brands
        - Identify trends and opportunities
        - Make data-driven decisions
        
        <h3>🛠️ Features</h3>
        - Automated report generation
        - Interactive visualizations
        - Multi-region analysis
        - Historical comparisons
        - PDF export capabilities
        
        <h3>📧 Support</h3>
        For technical support or feedback:
        - Email: prasoon.bajpai@lc.jkmail.com
        </div>
        """, unsafe_allow_html=True)

def sales_review_report_generator():
    # Sidebar navigation
    with st.sidebar:
        st.markdown("# 📊 Navigation")
        selected_page = st.radio(
            "",
            ["🏠 Home", "📈 Report Generator", "ℹ️ About"],
            key="navigation"
        )
        
        st.markdown("---")
        st.markdown("### 📅 Current Session")
        st.markdown(f"Date: {datetime.now().strftime('%B %d, %Y')}")
        if 'df' in st.session_state and st.session_state['df'] is not None:
            st.markdown("Status: ✅ Data Loaded")
        else:
            st.markdown("Status: ⚠️ Awaiting Data")
    
    # Main content
    if selected_page == "🏠 Home":
        show_welcome_page()
    elif selected_page == "📈 Report Generator":
        show_report_generator()
    else:
        show_about_page()
def load_lottie_url(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def get_online_editor_url(file_extension):
    extension_mapping = {
        '.xlsx': 'https://www.office.com/launch/excel?auth=2',
        '.xls': 'https://www.office.com/launch/excel?auth=2',
        '.doc': 'https://www.office.com/launch/word?auth=2',
        '.docx': 'https://www.office.com/launch/word?auth=2',
        '.ppt': 'https://www.office.com/launch/powerpoint?auth=2',
        '.pptx': 'https://www.office.com/launch/powerpoint?auth=2',
        '.pdf': 'https://documentcloud.adobe.com/link/home/'
    }
    return extension_mapping.get(file_extension.lower(), 'https://www.google.com/drive/')
def folder_menu():
    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: #3366cc;
        text-align: center;
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 30px;
        font-family: 'Arial', sans-serif;
    }
    .title span {
        background: linear-gradient(45deg, #3366cc, #6699ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .file-box {
        border: 1px solid #ddd;
        padding: 15px;
        margin: 15px 0;
        border-radius: 10px;
        background-color: #f9f9f9;
        transition: all 0.3s ease;
    }
    .file-box:hover {
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        transform: translateY(-5px);
    }
    .stButton>button {
        border-radius: 20px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.05);
    }
    .upload-section {
        background-color: #e6f3ff;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .todo-section {
        background-color: #f0f8ff;
        padding: 20px;
        border-radius: 10px;
        margin-top: 20px;
    }
    .todo-item {
        display: flex;
        align-items: center;
        margin-bottom: 10px;
    }
    .todo-text {
        margin-left: 10px;
    }
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="title"><span>📓 Advanced File Manager</span></div>', unsafe_allow_html=True)
    lottie_urls = [
        "https://assets9.lottiefiles.com/packages/lf20_3vbOcw.json",  # File manager animation
        "https://assets9.lottiefiles.com/packages/lf20_5lAtR7.json",  # Folder animation
        "https://assets1.lottiefiles.com/packages/lf20_4djadwfo.json",  # Document management
        "https://assets6.lottiefiles.com/packages/lf20_2a5yxpci.json"   # File transfer
    ]
    lottie_json = None
    for url in lottie_urls:
        lottie_json = load_lottie_url(url)
        if lottie_json:
            break
    col1, col2 = st.columns([1, 2])
    with col1:
        if lottie_json:
           st_lottie(lottie_json, height=200, key="file_animation")
        else:
           st.image("https://via.placeholder.com/200x200.png?text=File+Manager", use_column_width=True)
    with col2:
        st.markdown("""
        Welcome to the Advanced File Manager! 
        Here you can upload, download, and manage your files with ease. 
        Enjoy the smooth animations, user-friendly interface, and new features like file search and sorting.
        """)
    if not os.path.exists("uploaded_files"):
        os.makedirs("uploaded_files")
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload a file", type=["xlsx", "xls", "doc", "docx", "pdf", "ppt", "pptx", "txt", "csv"])
    if uploaded_file is not None:
        file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": uploaded_file.size}
        
        # Save the uploaded file
        with open(os.path.join("uploaded_files", uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"File {uploaded_file.name} saved successfully!")
    st.markdown('</div>', unsafe_allow_html=True)
    st.subheader("Your Files")
    search_query = st.text_input("Search files", "")
    sort_option = st.selectbox("Sort by", ["Name", "Size", "Date Modified"])
    if 'files_to_delete' not in st.session_state:
        st.session_state.files_to_delete = set()
    files = os.listdir("uploaded_files")
    if search_query:
        files = [f for f in files if search_query.lower() in f.lower()]
    
    # Apply sorting
    if sort_option == "Name":
        files.sort()
    elif sort_option == "Size":
        files.sort(key=lambda x: os.path.getsize(os.path.join("uploaded_files", x)), reverse=True)
    elif sort_option == "Date Modified":
        files.sort(key=lambda x: os.path.getmtime(os.path.join("uploaded_files", x)), reverse=True)

    for filename in files:
        file_path = os.path.join("uploaded_files", filename)
        file_stats = os.stat(file_path)
        
        st.markdown(f'<div class="file-box">', unsafe_allow_html=True)
        col1, col2, col3, col4= st.columns([3, 1, 1, 1])
        with col1:
            st.markdown(f"<h3>{filename}</h3>", unsafe_allow_html=True)
            st.text(f"Size: {file_stats.st_size / 1024:.2f} KB")
            st.text(f"Modified: {datetime.fromtimestamp(file_stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}")
        with col2:
            if st.button(f"📥 Download", key=f"download_{filename}"):
                with open(file_path, "rb") as file:
                    file_content = file.read()
                    b64 = base64.b64encode(file_content).decode()
                    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Click to download</a>'
                    st.markdown(href, unsafe_allow_html=True)
        with col3:
            if st.button(f"🗑️ Delete", key=f"delete_{filename}"):
                st.session_state.files_to_delete.add(filename)
        with col4:
            file_extension = os.path.splitext(filename)[1]
            editor_url = get_online_editor_url(file_extension)
            st.markdown(f"[🌐 Open Online]({editor_url})")
        st.markdown('</div>', unsafe_allow_html=True)

    # Process file deletion
    files_deleted = False
    for filename in st.session_state.files_to_delete:
        file_path = os.path.join("uploaded_files", filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            st.warning(f"{filename} has been deleted.")
            files_deleted = True
    
    # Clear the set of files to delete
    st.session_state.files_to_delete.clear()

    # Rerun the app if any files were deleted
    if files_deleted:
        st.rerun()

    st.info("Note: The 'Open Online' links will redirect you to the appropriate online editor. You may need to manually open your file once there.")

    # To-Do List / Diary Section
    st.markdown('<div class="todo-section">', unsafe_allow_html=True)
    st.subheader("📝 To-Do List / Diary")

    # Load existing to-do items
    if 'todo_items' not in st.session_state:
        st.session_state.todo_items = []

    # Add new to-do item
    new_item = st.text_input("Add a new to-do item or diary entry")
    if st.button("Add"):
        if new_item:
            st.session_state.todo_items.append({"text": new_item, "done": False, "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            st.success("Item added successfully!")

    # Display and manage to-do items
    for idx, item in enumerate(st.session_state.todo_items):
        col1, col2, col3 = st.columns([0.1, 3, 1])
        with col1:
            done = st.checkbox("", item["done"], key=f"todo_{idx}")
            if done != item["done"]:
                st.session_state.todo_items[idx]["done"] = done
        with col2:
            st.markdown(f"<div class='todo-text'>{'<s>' if item['done'] else ''}{item['text']}{'</s>' if item['done'] else ''}</div>", unsafe_allow_html=True)
        with col3:
            st.text(item["date"])
        if st.button("Delete", key=f"delete_todo_{idx}"):
            st.session_state.todo_items.pop(idx)
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # Add a fun fact section
    st.markdown("---")
    st.subheader("📚 Fun File Fact")
    fun_facts = [
        "The first computer virus was created in 1983 and was called the Elk Cloner.",
        "The most common file extension in the world is .dll (Dynamic Link Library).",
        "The largest file size theoretically possible in Windows is 16 exabytes minus 1 KB.",
        "The PDF file format was invented by Adobe in 1993.",
        "The first widely-used image format on the web was GIF, created in 1987.",
        "John McCarthy,an American computer scientist, coined the term Artificial Intelligence in 1956.",
        "About 90% of the World's Currency only exists on Computers.",
        "MyDoom is the most expensive computer virus in history.",
        "The original name of windows was Interface Manager.",
        "The first microprocessor created by Intel was the 4004."
    ]
    st.markdown(f"*{fun_facts[int(os.urandom(1)[0]) % len(fun_facts)]}*")
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

def sales_dashboard():
    
    st.title("Sales Dashboard")

    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: #3366cc;
        text-align: center;
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 30px;
        font-family: 'Arial', sans-serif;
    }
    .title span {
        background: linear-gradient(45deg, #3366cc, #6699ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .section-box {
        background-color: #f9f9f9;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .upload-section {
        background-color: #e6f3ff;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .stDataFrame {
        font-family: 'Arial', sans-serif;
    }
    </style>
    """, unsafe_allow_html=True)

    # Load Lottie animation
    lottie_url = "https://assets2.lottiefiles.com/packages/lf20_V9t630.json"  # New interesting animation
    lottie_json = load_lottie_url(lottie_url)
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st_lottie(lottie_json, height=200, key="home_animation")
    with col2:
        st.markdown("""
        Welcome to our interactive Sales Analysis Dashboard! 
        This powerful tool helps you analyze Sales data for JKLC and UCWL across different regions, districts and channels.
        Let's get started with your data analysis journey!
        """)

    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key="Sales_Dashboard_uploader")
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        df = process_dataframe(df)

        # Region selection
        regions = df['Region'].unique()
        selected_regions = st.multiselect('Select Regions', regions)

        # District selection
        districts = df[df['Region'].isin(selected_regions)]['Dist Name'].unique()
        selected_districts = st.multiselect('Select Districts', districts)

        # Channel selection
        channels = ['Overall', 'Trade', 'Non-Trade']
        selected_channels = st.multiselect('Select Channels', channels, default=channels)

        # Checkbox for whole region totals
        show_whole_region = st.checkbox('Show whole region totals')

        if st.button('Generate Report'):
            display_data(df, selected_regions, selected_districts, selected_channels, show_whole_region)

def process_dataframe(df):
    
    
    column_mapping = {
    pd.to_datetime('2024-09-23 00:00:00'): '23-Sep',
    pd.to_datetime('2024-08-23 00:00:00'): '23-Aug',
    pd.to_datetime('2024-07-23 00:00:00'): '23-Jul',
    pd.to_datetime('2024-06-23 00:00:00'): '23-Jun',
    pd.to_datetime('2024-05-23 00:00:00'): '23-May',
    pd.to_datetime('2024-04-23 00:00:00'): '23-Apr',
    pd.to_datetime('2024-08-24 00:00:00'): '24-Aug',
    pd.to_datetime('2024-07-24 00:00:00'): '24-Jul',
    pd.to_datetime('2024-06-24 00:00:00'): '24-Jun',
    pd.to_datetime('2024-05-24 00:00:00'): '24-May',
    pd.to_datetime('2024-04-24 00:00:00'): '24-Apr'
}

    df = df.rename(columns=column_mapping)
    df['FY 2024 till Aug'] = df['24-Apr'] + df['24-May'] + df['24-Jun'] + df['24-Jul'] + df['24-Aug']
    df['FY 2023 till Aug'] = df['23-Apr'] + df['23-May'] + df['23-Jun'] + df['23-Jul'] + df['23-Aug']
    df['Quarterly Requirement'] = df['23-Jul'] + df['23-Aug'] + df['23-Sep'] - df['24-Jul'] - df['24-Aug']
    df['Growth/Degrowth(MTD)'] = (df['24-Aug'] - df['23-Aug']) / df['23-Aug'] * 100
    df['Growth/Degrowth(YTD)'] = (df['FY 2024 till Aug'] - df['FY 2023 till Aug']) / df['FY 2023 till Aug'] * 100
    df['Q3 2023'] = df['23-Jul'] + df['23-Aug'] + df['23-Sep']
    df['Q3 2024 till August'] = df['24-Jul'] + df['24-Aug']

    # Non-Trade calculations
    for month in ['Sep', 'Aug', 'Jul', 'Jun', 'May', 'Apr']:
        df[f'23-{month} Non-Trade'] = df[f'23-{month}'] - df[f'23-{month} Trade']
        if month != 'Sep':
            df[f'24-{month} Non-Trade'] = df[f'24-{month}'] - df[f'24-{month} Trade']

    # Trade calculations
    df['FY 2024 till Aug Trade'] = df['24-Apr Trade'] + df['24-May Trade'] + df['24-Jun Trade'] + df['24-Jul Trade'] + df['24-Aug Trade']
    df['FY 2023 till Aug Trade'] = df['23-Apr Trade'] + df['23-May Trade'] + df['23-Jun Trade'] + df['23-Jul Trade'] + df['23-Aug Trade']
    df['Quarterly Requirement Trade'] = df['23-Jul Trade'] + df['23-Aug Trade'] + df['23-Sep Trade'] - df['24-Jul Trade'] - df['24-Aug Trade']
    df['Growth/Degrowth(MTD) Trade'] = (df['24-Aug Trade'] - df['23-Aug Trade']) / df['23-Aug Trade'] * 100
    df['Growth/Degrowth(YTD) Trade'] = (df['FY 2024 till Aug Trade'] - df['FY 2023 till Aug Trade']) / df['FY 2023 till Aug Trade'] * 100
    df['Q3 2023 Trade'] = df['23-Jul Trade'] + df['23-Aug Trade'] + df['23-Sep Trade']
    df['Q3 2024 till August Trade'] = df['24-Jul Trade'] + df['24-Aug Trade']

    # Non-Trade calculations
    df['FY 2024 till Aug Non-Trade'] = df['24-Apr Non-Trade'] + df['24-May Non-Trade'] + df['24-Jun Non-Trade'] + df['24-Jul Non-Trade'] + df['24-Aug Non-Trade']
    df['FY 2023 till Aug Non-Trade'] = df['23-Apr Non-Trade'] + df['23-May Non-Trade'] + df['23-Jun Non-Trade'] + df['23-Jul Non-Trade'] + df['23-Aug Non-Trade']
    df['Quarterly Requirement Non-Trade'] = df['23-Jul Non-Trade'] + df['23-Aug Non-Trade'] + df['23-Sep Non-Trade'] - df['24-Jul Non-Trade'] - df['24-Aug Non-Trade']
    df['Growth/Degrowth(MTD) Non-Trade'] = (df['24-Aug Non-Trade'] - df['23-Aug Non-Trade']) / df['23-Aug Non-Trade'] * 100
    df['Growth/Degrowth(YTD) Non-Trade'] = (df['FY 2024 till Aug Non-Trade'] - df['FY 2023 till Aug Non-Trade']) / df['FY 2023 till Aug Non-Trade'] * 100
    df['Q3 2023 Non-Trade'] = df['23-Jul Non-Trade'] + df['23-Aug Non-Trade'] + df['23-Sep Non-Trade']
    df['Q3 2024 till August Non-Trade'] = df['24-Jul Non-Trade'] + df['24-Aug Non-Trade']

    # Handle division by zero

    return df

    pass
def display_data(df, selected_regions, selected_districts, selected_channels, show_whole_region):
    def color_growth(val):
        try:
            value = float(val.strip('%'))
            color = 'green' if value > 0 else 'red' if value < 0 else 'black'
            return f'color: {color}'
        except:
            return 'color: black'

    if show_whole_region:
        filtered_data = df[df['Region'].isin(selected_regions)].copy()
        
        # Calculate sums for relevant columns first
        sum_columns = ['24-Apr','24-May','24-Jun','24-Jul','24-Aug','23-Apr','23-May','23-Jun','23-Jul', '23-Aug', 'FY 2024 till Aug', 'FY 2023 till Aug', 'Q3 2023', 'Q3 2024 till August','24-Apr Trade','24-May Trade','24-Jun Trade','24-Jul Trade', 
                        '24-Aug Trade','23-Apr Trade','23-May Trade','23-Jun Trade','23-Jul Trade', '23-Aug Trade', 'FY 2024 till Aug Trade', 'FY 2023 till Aug Trade', 
                        'Q3 2023 Trade', 'Q3 2024 till August Trade','24-Apr Non-Trade','24-May Non-Trade','24-Jun Non-Trade','24-Jul Non-Trade',
                        '24-Aug Non-Trade','23-Apr Non-Trade','23-May Non-Trade','23-Jun Non-Trade','23-Jul Non-Trade', '23-Aug Non-Trade', 'FY 2024 till Aug Non-Trade', 'FY 2023 till Aug Non-Trade', 
                        'Q3 2023 Non-Trade', 'Q3 2024 till August Non-Trade']
        grouped_data = filtered_data.groupby('Region')[sum_columns].sum().reset_index()

        # Then calculate Growth/Degrowth based on the summed values
        grouped_data['Growth/Degrowth(MTD)'] = (grouped_data['24-Aug'] - grouped_data['23-Aug']) / grouped_data['23-Aug'] * 100
        grouped_data['Growth/Degrowth(YTD)'] = (grouped_data['FY 2024 till Aug'] - grouped_data['FY 2023 till Aug']) / grouped_data['FY 2023 till Aug'] * 100
        grouped_data['Quarterly Requirement'] = grouped_data['Q3 2023'] - grouped_data['Q3 2024 till August']

        grouped_data['Growth/Degrowth(MTD) Trade'] = (grouped_data['24-Aug Trade'] - grouped_data['23-Aug Trade']) / grouped_data['23-Aug Trade'] * 100
        grouped_data['Growth/Degrowth(YTD) Trade'] = (grouped_data['FY 2024 till Aug Trade'] - grouped_data['FY 2023 till Aug Trade']) / grouped_data['FY 2023 till Aug Trade'] * 100
        grouped_data['Quarterly Requirement Trade'] = grouped_data['Q3 2023 Trade'] - grouped_data['Q3 2024 till August Trade']

        grouped_data['Growth/Degrowth(MTD) Non-Trade'] = (grouped_data['24-Aug Non-Trade'] - grouped_data['23-Aug Non-Trade']) / grouped_data['23-Aug Non-Trade'] * 100
        grouped_data['Growth/Degrowth(YTD) Non-Trade'] = (grouped_data['FY 2024 till Aug Non-Trade'] - grouped_data['FY 2023 till Aug Non-Trade']) / grouped_data['FY 2023 till Aug Non-Trade'] * 100
        grouped_data['Quarterly Requirement Non-Trade'] = grouped_data['Q3 2023 Non-Trade'] - grouped_data['Q3 2024 till August Non-Trade']
    else:
        if selected_districts:
            filtered_data = df[df['Dist Name'].isin(selected_districts)].copy()
        else:
            filtered_data = df[df['Region'].isin(selected_regions)].copy()
        grouped_data = filtered_data

    for selected_channel in selected_channels:
        if selected_channel == 'Trade':
            columns_to_display = ['24-Aug Trade','23-Aug Trade','Growth/Degrowth(MTD) Trade','FY 2024 till Aug Trade', 'FY 2023 till Aug Trade','Growth/Degrowth(YTD) Trade','Q3 2023 Trade','Q3 2024 till August Trade', 'Quarterly Requirement Trade']
            suffix = ' Trade'
        elif selected_channel == 'Non-Trade':
            columns_to_display = ['24-Aug Non-Trade','23-Aug Non-Trade','Growth/Degrowth(MTD) Non-Trade','FY 2024 till Aug Non-Trade', 'FY 2023 till Aug Non-Trade','Growth/Degrowth(YTD) Non-Trade','Q3 2023 Non-Trade','Q3 2024 till August Non-Trade', 'Quarterly Requirement Non-Trade']
            suffix = ' Non-Trade'
        else:  # Overall
            columns_to_display = ['24-Aug','23-Aug','Growth/Degrowth(MTD)','FY 2024 till Aug', 'FY 2023 till Aug','Growth/Degrowth(YTD)','Q3 2023','Q3 2024 till August', 'Quarterly Requirement']
            suffix = ''
        
        display_columns = ['Region' if show_whole_region else 'Dist Name'] + columns_to_display
        
        st.subheader(f"{selected_channel} Sales Data")
        
        # Create a copy of the dataframe with only the columns we want to display
        display_df = grouped_data[display_columns].copy()
        
        # Set the 'Region' or 'Dist Name' column as the index
        display_df.set_index('Region' if show_whole_region else 'Dist Name', inplace=True)
        
        # Style the dataframe
        styled_df = display_df.style.format({
            col: '{:,.0f}' if 'Growth' not in col else '{:.2f}%' for col in columns_to_display
        }).applymap(color_growth, subset=[col for col in columns_to_display if 'Growth' in col])
        
        st.dataframe(styled_df)

        # Add a bar chart for YTD comparison
        fig = go.Figure(data=[
            go.Bar(name='FY 2023', x=grouped_data['Region' if show_whole_region else 'Dist Name'], y=grouped_data[f'FY 2023 till Aug{suffix}']),
            go.Bar(name='FY 2024', x=grouped_data['Region' if show_whole_region else 'Dist Name'], y=grouped_data[f'FY 2024 till Aug{suffix}']),
        ])
        fig.update_layout(barmode='group', title=f'{selected_channel} YTD Comparison')
        st.plotly_chart(fig)

        # Add a line chart for monthly trends including September 2024
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug']
        fig_trend = go.Figure()
        for year in ['23', '24']:
            y_values = []
            for month in months:
                column_name = f'{year}-{month}{suffix}'
                if column_name in grouped_data.columns:
                    y_values.append(grouped_data[column_name].sum())
                else:
                    y_values.append(None)
            
            fig_trend.add_trace(go.Scatter(
                x=months, 
                y=y_values, 
                mode='lines+markers+text',
                name=f'FY 20{year}',
                text=[f'{y:,.0f}' if y is not None else '' for y in y_values],
                textposition='top center'
            ))
        
        fig_trend.update_layout(
            title=f'{selected_channel} Monthly Trends', 
            xaxis_title='Month', 
            yaxis_title='Sales',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_trend)



def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def normal():
 lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
 lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")
 with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Product-Mix Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )


 def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 18)
        header_text = f"Product Mix Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "Product Mix Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the Product Mix Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Normal
            [random.randint(1500, 2500) for _ in range(12)],  # Premium
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.blue
        lc.lines[2].strokeColor = colors.pink
        lc.lines[3].strokeColor = colors.brown

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Normal EBITDA'),
            (colors.blue, 'Premium EBITDA'),
            (colors.crimson, 'Overall EBITDA'),
            (colors.brown, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Normal and Premium EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If both (Trade,Non-Trade) are present: Premium +5%, Normal -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Normal and Premium Product for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Premium Product , which typically have higher EBITDA.",
            "Analyze factors contributing to higher EBITDA in Premium Channel,and apply insights to Normal.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all channels.",
            "Invest in product innovation to expand Premium Product offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for Normal and Premium Product ceteris paribus.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across products,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Normal EBITDA', 'Premium EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Normal'] + filtered_df['Premium']
                    filtered_df[overall_col] = (
                        (filtered_df['Normal'] * filtered_df['Normal EBITDA'] +
                         filtered_df['Premium'] * filtered_df['Premium EBITDA'])/ total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Normal Share'] = filtered_df['Normal'] / total_quantity
                    filtered_df['Average Premium Share'] = filtered_df['Premium'] / total_quantity
                    
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        normal = row['Average Normal Share']
                        premium = row['Average Premium Share']
                        
                        if normal == 1 or premium == 1 :
                            # If any share is 100%, don't change
                            return normal,premium
                        else:
                            premium = min(premium + 0.05, 1)
                            normal = max(normal - 0.05, 1 - premium)
                        
                        return normal,premium
                    filtered_df['Adjusted Normal Share'], filtered_df['Adjusted Premium Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Normal Share'] * filtered_df['Normal EBITDA'] +
                        filtered_df['Adjusted Premium Share'] * filtered_df['Premium EBITDA']
                    )

                    # Calculate differences
                    filtered_df['P-N Difference'] = filtered_df['Premium EBITDA'] - filtered_df['Normal EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Normal EBITDA'],
                                             mode='lines+markers', name='Normal EBITDA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Premium EBITDA'],
                                             mode='lines+markers', name='Premium EBITDA', line=dict(color='blue')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='crimson', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBITDA',
                                             line=dict(color='brown', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(P-N: {g_r:.0f})<br>(I-O: {g_y:.0f}))" 
                                for month, g_r, g_y in 
                                zip(filtered_df['Month'], 
                                    filtered_df['P-N Difference'],  
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}(Type:-{product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Normal','Premium']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [45,45,45] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                    
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Normal Share', 'Average Premium Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Normal Share': 'green', 'Average Premium Share': 'blue'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Normal', 'Premium']]
                    for _, row in filtered_df[['Month', 'Normal', 'Premium','Average Normal Share', 'Average Premium Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Normal']:.0f} ({row['Average Normal Share']:.2%})",
                            f"{row['Premium']:.0f} ({row['Average Premium Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer


 if selected == "Home":
    st.title("🔍 Advanced Product Mix Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key='NormalvsPremiumuploader')
        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success("File successfully uploaded! Please go to the Analysis page to view results.")

    with col2:
        if lottie_upload:
            st_lottie(lottie_upload, height=150, key="upload")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/4503/4503700.png", width=150)
    st.markdown("</div>", unsafe_allow_html=True)
 elif selected == "Product-Mix Analysis":
    st.title("📈 Product Mix Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                subset_df = df[(df['Region'] == region) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Product_Mix_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Product_Mix_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        brand = st.sidebar.selectbox("Select Brand", options=df[df['Region']==region]['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df[df['Region']==region]['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df[df['Region']==region]['Region subsets'].unique(), key="region_subset_select")

        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type
        premium_share = st.sidebar.slider("Adjust Premium Share (%)", 0, 100, 50)

        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Normal NSR', 'Premium NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Normal Contribution', 'Premium Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Normal EBITDA', 'Premium EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Normal'] * filtered_df[cols[0]] +
                                        filtered_df['Premium'] * filtered_df[cols[1]]) / (
                                            filtered_df['Normal'] + filtered_df['Premium'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - premium_share/100) * filtered_df[cols[0]] +
                                          (premium_share/100) * filtered_df[cols[1]])
            
            # Calculate difference between Premium and Normal
            filtered_df['Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            
            # Calculate difference between Imaginary and Overall
            filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
            
            # Create the plot
            fig = go.Figure()
            
            for col in cols:
                fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[col],
                                         mode='lines+markers', name=col))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                     mode='lines+markers', name=overall_col, line=dict(dash='dash')))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({premium_share}% Premium)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(P-N: {diff:.2f})<br>(I-O: {i_diff:.2f})" for month, diff, i_diff in 
                        zip(filtered_df['Month'], filtered_df['Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (P-N: Premium - Normal, I-O: Imaginary - Overall)',
                yaxis_title='Value',
                legend_title='Metrics',
                hovermode="x unified",
                xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Display descriptive statistics
            st.subheader("Descriptive Statistics")
            desc_stats = filtered_df[cols + [overall_col, imaginary_col]].describe()
            st.dataframe(desc_stats.style.format("{:.2f}"), use_container_width=True)
            
            # Display share of Normal and Premium Products
            st.subheader("Share of Normal and Premium Products")
            total_quantity = filtered_df['Normal'] + filtered_df['Premium']
            normal_share = (filtered_df['Normal'] / total_quantity * 100).round(2)
            premium_share = (filtered_df['Premium'] / total_quantity * 100).round(2)
            
            share_df = pd.DataFrame({
                'Month': filtered_df['Month'],
                'Premium Share (%)': premium_share,
                'Normal Share (%)': normal_share
            })
                  
            fig_pie = px.pie(share_df, values=[normal_share.mean(), premium_share.mean()], 
                                     names=['Normal', 'Premium'], title='Average Share Distribution',color=["N","P"],color_discrete_map={"N":"green","P":"blue"},hole=0.5)
            st.plotly_chart(fig_pie, use_container_width=True)
                    
            st.dataframe(share_df.set_index('Month').style.format("{:.2f}").background_gradient(cmap='RdYlGn'), use_container_width=True)
        
        
        else:
            st.warning("No data available for the selected combination.")
        
        st.markdown("</div>", unsafe_allow_html=True)

 elif selected == "About":
    st.title("About the Product Mix Analysis App")
    st.markdown("""
    This advanced data analysis application is designed to provide insightful visualizations and statistics for your Product(Normal and Premium) Mix data. 
    
    Key features include:
    - Interactive data filtering
    - Multiple analysis types (NSR, Contribution, EBITDA)
    - Dynamic visualizations with Plotly
    - Descriptive statistics and share analysis
    - Customizable Premium share adjustments
    """)
def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def trade():
 lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
 lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")
 with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Segment-Mix Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )


 def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 18)
        header_text = f"Segment Mix Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "Segment Mix Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the Segment Mix Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Trade
            [random.randint(1500, 2500) for _ in range(12)],  # Non-Trade
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.blue
        lc.lines[2].strokeColor = colors.pink
        lc.lines[3].strokeColor = colors.brown

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Trade EBITDA'),
            (colors.blue, 'Non-Trade EBITDA'),
            (colors.crimson, 'Overall EBITDA'),
            (colors.brown, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Trade and Non-Trade EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If both (Trade,Non-Trade) are present: Trade +5%, Non-Trade -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Trade and Non-Trade Channel for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Trade Channel specifically for PPC, which typically have higher EBIDTA.",
            "Analyze factors contributing to higher EBIDTA in Trade Channel,and apply insights to Non-Trade.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all channels.",
            "Invest in product innovation to expand Trade Channel offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for Trade and Non-Trade channel ceteris paribus.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across GYR,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Trade EBITDA', 'Non-Trade EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Trade'] + filtered_df['Non-Trade']
                    filtered_df[overall_col] = (
                        (filtered_df['Trade'] * filtered_df['Trade EBITDA'] +
                         filtered_df['Non-Trade'] * filtered_df['Non-Trade EBITDA'])/ total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Trade Share'] = filtered_df['Trade'] / total_quantity
                    filtered_df['Average Non-Trade Share'] = filtered_df['Non-Trade'] / total_quantity
                    
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        trade = row['Average Trade Share']
                        nontrade = row['Average Non-Trade Share']
                        
                        if trade == 1 or nontrade == 1 :
                            # If any share is 100%, don't change
                            return trade,nontrade
                        else:
                            trade = min(trade + 0.05, 1)
                            nontrade = min(nontrade - 0.05, 1 - trade)
                        
                        return trade,nontrade
                    filtered_df['Adjusted Trade Share'], filtered_df['Adjusted Non-Trade Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Trade Share'] * filtered_df['Trade EBITDA'] +
                        filtered_df['Adjusted Non-Trade Share'] * filtered_df['Non-Trade EBITDA']
                    )

                    # Calculate differences
                    filtered_df['T-NT Difference'] = filtered_df['Trade EBITDA'] - filtered_df['Non-Trade EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Trade EBITDA'],
                                             mode='lines+markers', name='Trade EBIDTA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Non-Trade EBITDA'],
                                             mode='lines+markers', name='Non-Trade EBIDTA', line=dict(color='blue')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='crimson', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBIDTA',
                                             line=dict(color='brown', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(T-NT: {g_r:.0f})<br>(I-O: {g_y:.0f}))" 
                                for month, g_r, g_y in 
                                zip(filtered_df['Month'], 
                                    filtered_df['T-NT Difference'],  
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}(Type:-{product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Trade','Non-Trade']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [45,45,45] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                    
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Trade Share', 'Average Non-Trade Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Trade Share': 'green', 'Average Non-Trade Share': 'blue'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Trade', 'Non-Trade']]
                    for _, row in filtered_df[['Month', 'Trade', 'Non-Trade','Average Trade Share', 'Average Non-Trade Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Trade']:.0f} ({row['Average Trade Share']:.2%})",
                            f"{row['Non-Trade']:.0f} ({row['Average Non-Trade Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer


 if selected == "Home":
    st.title("🔍 Advanced Segment Mix Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key='TradevsNontradeuploader')
        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success("File successfully uploaded! Please go to the Analysis page to view results.")

    with col2:
        if lottie_upload:
            st_lottie(lottie_upload, height=150, key="upload")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/4503/4503700.png", width=150)
    st.markdown("</div>", unsafe_allow_html=True)
 elif selected == "Segment-Mix Analysis":
    st.title("📈 Segment Mix Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                subset_df = df[(df['Region'] == region) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Segment_Mix_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        brand = st.sidebar.selectbox("Select Brand", options=df[df['Region']==region]['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df[df['Region']==region]['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df[df['Region']==region]['Region subsets'].unique(), key="region_subset_select")

        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type
        trade_share = st.sidebar.slider("Adjust Trade Share (%)", 0, 100, 50)

        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Trade NSR', 'Non-Trade NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Trade Contribution', 'Non-Trade Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Trade EBITDA', 'Non-Trade EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Trade'] * filtered_df[cols[0]] +
                                        filtered_df['Non-Trade'] * filtered_df[cols[1]]) / (
                                            filtered_df['Trade'] + filtered_df['Non-Trade'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - trade_share/100) * filtered_df[cols[1]] +
                                          (trade_share/100) * filtered_df[cols[0]])
            
            # Calculate difference between Premium and Normal
            filtered_df['Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            
            # Calculate difference between Imaginary and Overall
            filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
            
            # Create the plot
            fig = go.Figure()
            
            for col in cols:
                fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[col],
                                         mode='lines+markers', name=col))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                     mode='lines+markers', name=overall_col, line=dict(dash='dash')))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({trade_share}% Trade)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(T-NT: {diff:.2f})<br>(I-O: {i_diff:.2f})" for month, diff, i_diff in 
                        zip(filtered_df['Month'], filtered_df['Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (T-NT: Trade - Non-Trade, I-O: Imaginary - Overall)',
                yaxis_title='Value',
                legend_title='Metrics',
                hovermode="x unified",
                xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Display descriptive statistics
            st.subheader("Descriptive Statistics")
            desc_stats = filtered_df[cols + [overall_col, imaginary_col]].describe()
            st.dataframe(desc_stats.style.format("{:.2f}"), use_container_width=True)
            
            # Display share of Normal and Premium Products
            st.subheader("Share of Trade and Non-Trade Channel")
            total_quantity = filtered_df['Trade'] + filtered_df['Non-Trade']
            trade_share = (filtered_df['Trade'] / total_quantity * 100).round(2)
            nontrade_share = (filtered_df['Non-Trade'] / total_quantity * 100).round(2)
            
            share_df = pd.DataFrame({
                'Month': filtered_df['Month'],
                'Trade Share (%)': trade_share,
                'Non-Trade Share (%)': nontrade_share
            })
                  
            fig_pie = px.pie(share_df, values=[trade_share.mean(), nontrade_share.mean()], 
                                     names=['Trade', 'Non-Trade'], title='Average Share Distribution',color=["T","NT"],color_discrete_map={"T":"green","NT":"blue"},hole=0.5)
            st.plotly_chart(fig_pie, use_container_width=True)
                    
            st.dataframe(share_df.set_index('Month').style.format("{:.2f}").background_gradient(cmap='RdYlGn'), use_container_width=True)
        
        
        else:
            st.warning("No data available for the selected combination.")
        
        st.markdown("</div>", unsafe_allow_html=True)

 elif selected == "About":
    st.title("About the Segment Mix Analysis App")
    st.markdown("""
    This advanced data analysis application is designed to provide insightful visualizations and statistics for your Segment(Trade,Non-Trade) Mix data. 
    
    Key features include:
    - Interactive data filtering
    - Multiple analysis types (NSR, Contribution, EBITDA)
    - Dynamic visualizations with Plotly
    - Descriptive statistics and share analysis
    - Customizable Trade share adjustments
    """)
from plotly.subplots import make_subplots
import plotly.express as px
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER
from reportlab.graphics import renderPDF
import random
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from streamlit_option_menu import option_menu
from reportlab.platypus import Table, TableStyle
def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def green():
 with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Geo-Mix Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )
# Load Lottie animations
 lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
 lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")


 def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 18)
        header_text = f"GYR Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "GYR Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the GYR Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Green
            [random.randint(1500, 2500) for _ in range(12)],  # Yellow
            [random.randint(1000, 2000) for _ in range(12)],  # Red
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.yellow
        lc.lines[2].strokeColor = colors.red
        lc.lines[3].strokeColor = colors.blue
        lc.lines[4].strokeColor = colors.purple

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Green EBITDA'),
            (colors.yellow, 'Yellow EBITDA'),
            (colors.red, 'Red EBITDA'),
            (colors.blue, 'Overall EBITDA'),
            (colors.purple, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Green, Yellow, and Red EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If all three (Green, Yellow, Red) are present: Green +5%, Yellow +2.5%, Red -7.5%"),
            ("", "• If only two are present: Superior one (Green in GR or GY, Yellow in YR) +5%, other -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Green, Yellow, and Red products for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Green Region products, which typically have higher EBIDTA margins.",
            "Analyze factors contributing to higher EBIDTA in Green zone,and apply insights to Red zone.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all product categories.",
            "Invest in product innovation to expand Green and Yellow region offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for GYR keeping everything else constant.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across GYR,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Green EBITDA', 'Yellow EBITDA', 'Red EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red']
                    filtered_df[overall_col] = (
                        (filtered_df['Green'] * filtered_df['Green EBITDA'] +
                         filtered_df['Yellow'] * filtered_df['Yellow EBITDA'] + 
                         filtered_df['Red'] * filtered_df['Red EBITDA']) / total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Green Share'] = filtered_df['Green'] / total_quantity
                    filtered_df['Average Yellow Share'] = filtered_df['Yellow'] / total_quantity
                    filtered_df['Average Red Share'] = filtered_df['Red'] / total_quantity
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        green = row['Average Green Share']
                        yellow = row['Average Yellow Share']
                        red = row['Average Red Share']
                        
                        if green == 1 or yellow == 1 or red == 1:
                            # If any share is 100%, don't change
                            return green, yellow, red
                        elif red == 0:
                            green = min(green +0.05, 1)
                            yellow = max(1-green, 0)
                        elif green == 0 and yellow == 0:
                            # If both green and yellow are absent, don't change
                            return green, yellow, red
                        elif green == 0:
                            # If green is absent, increase yellow by 5% and decrease red by 5%
                            yellow = min(yellow + 0.05, 1)
                            red = max(1 - yellow, 0)
                        elif yellow == 0:
                            # If yellow is absent, increase green by 5% and decrease red by 5%
                            green = min(green + 0.05, 1)
                            red = max(1 - green, 0)
                        else:
                            # Normal case: increase green by 5%, yellow by 2.5%, decrease red by 7.5%
                            green = min(green + 0.05, 1)
                            yellow = min(yellow + 0.025, 1 - green)
                            red = max(1 - green - yellow, 0)
                        
                        return green, yellow, red
                    filtered_df['Adjusted Green Share'], filtered_df['Adjusted Yellow Share'], filtered_df['Adjusted Red Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Green Share'] * filtered_df['Green EBITDA'] +
                        filtered_df['Adjusted Yellow Share'] * filtered_df['Yellow EBITDA'] +
                        filtered_df['Adjusted Red Share'] * filtered_df['Red EBITDA']
                    )

                    # Calculate differences
                    filtered_df['G-R Difference'] = filtered_df['Green EBITDA'] - filtered_df['Red EBITDA']
                    filtered_df['G-Y Difference'] = filtered_df['Green EBITDA'] - filtered_df['Yellow EBITDA']
                    filtered_df['Y-R Difference'] = filtered_df['Yellow EBITDA'] - filtered_df['Red EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Green EBITDA'],
                                             mode='lines+markers', name='Green EBIDTA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Yellow EBITDA'],
                                             mode='lines+markers', name='Yellow EBIDTA', line=dict(color='yellow')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Red EBITDA'],
                                             mode='lines+markers', name='Red EBIDTA', line=dict(color='red')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='blue', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBIDTA',
                                             line=dict(color='purple', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(G-R: {g_r:.0f})<br>(G-Y: {g_y:.0f})<br>(Y-R: {y_r:.0f})" 
                                for month, g_r, g_y, y_r, i_o in 
                                zip(filtered_df['Month'], 
                                    filtered_df['G-R Difference'], 
                                    filtered_df['G-Y Difference'], 
                                    filtered_df['Y-R Difference'], 
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}({product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Green','Yellow','Red']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [40,40,40,40] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Green Share', 'Average Yellow Share', 'Average Red Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Green Share': 'green', 'Average Yellow Share': 'yellow', 'Average Red Share': 'red'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Green', 'Yellow', 'Red']]
                    for _, row in filtered_df[['Month', 'Green', 'Yellow', 'Red', 'Average Green Share', 'Average Yellow Share', 'Average Red Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Green']:.0f} ({row['Average Green Share']:.2%})",
                            f"{row['Yellow']:.0f} ({row['Average Yellow Share']:.2%})",
                            f"{row['Red']:.0f} ({row['Average Red Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer

 if selected == "Home":
    st.title("🔍 Advanced Geo Mix Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key="gyruploader")
        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success("File successfully uploaded! Please go to the Analysis page to view results.")

    with col2:
        if lottie_upload:
            st_lottie(lottie_upload, height=150, key="upload")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/4503/4503700.png", width=150)
    st.markdown("</div>", unsafe_allow_html=True)
 elif selected == "Geo-Mix Analysis":
    st.title("📈 Geo Mix Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                pdf_buffer = create_pdf_report(region, df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset)]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)

        # Add unique keys to each selectbox
        brand = st.sidebar.selectbox("Select Brand", options=df[df['Region']==region]['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df[df['Region']==region]['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df[df['Region']==region]['Region subsets'].unique(), key="region_subset_select")

        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type

        green_share = st.sidebar.slider("Adjust Green Share (%)", 0, 99, 50, key="green_share_slider")
        yellow_share = st.sidebar.slider("Adjust Yellow Share (%)", 0, 100-green_share, 0, key="yellow_share_slider")
        red_share = 100 - green_share - yellow_share
        st.sidebar.text(f"Red Share: {red_share}%")
        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Green NSR', 'Yellow NSR', 'Red NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Green Contribution', 'Yellow Contribution','Red Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Green EBITDA', 'Yellow EBITDA','Red EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Green'] * filtered_df[cols[0]] +
                                        filtered_df['Yellow'] * filtered_df[cols[1]] + filtered_df['Red']*filtered_df[cols[2]]) / (
                                            filtered_df['Green'] + filtered_df['Yellow']+filtered_df['Red'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - (green_share+yellow_share)/100) * filtered_df[cols[2]] +
                                          (green_share/100) * filtered_df[cols[0]] + (yellow_share/100) * filtered_df[cols[1]])
            
            # Calculate difference between Premium and Normal
            filtered_df['G-Y Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            filtered_df['G-R Difference'] = filtered_df[cols[0]] - filtered_df[cols[2]]
            filtered_df['Y-R Difference'] = filtered_df[cols[1]] - filtered_df[cols[2]]
            
            # Calculate difference between Imaginary and Overall
            filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
            
            # Create the plot
            fig = go.Figure()
            
            
            if cols[0] in cols:
                  fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[0]],
                                         mode='lines+markers', name=cols[0],line_color="green"))
            if cols[1] in cols:
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[1]],
                                         mode='lines+markers', name=cols[1],line_color="yellow"))
            if cols[2] in cols:
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[2]],
                                         mode='lines+markers', name=cols[2],line_color="red"))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                     mode='lines+markers', name=overall_col, line=dict(dash='dash')))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({green_share}% Green & {yellow_share}% Yellow)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(G-Y: {diff:.2f})<br>(G-R: {i_diff:.2f})<br>(Y-R: {j_diff:.2f})<br>(I-O: {k_diff:.2f})" for month, diff, i_diff, j_diff, k_diff in 
                        zip(filtered_df['Month'], filtered_df['G-Y Difference'], filtered_df['G-R Difference'], filtered_df['Y-R Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (G-Y: Green - Red,G-R: Green - Red,Y-R: Yellow - Red, I-O: Imaginary - Overall)',
                yaxis_title='Value',
                legend_title='Metrics',
                hovermode="x unified",
                xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            st.subheader("Descriptive Statistics")
            desc_stats = filtered_df[cols + [overall_col, imaginary_col]].describe()
            st.dataframe(desc_stats.style.format("{:.2f}").background_gradient(cmap='Blues'), use_container_width=True)
                    
                    # Display share of Green, Yellow, and Red Products
            st.subheader("Share of Green, Yellow, and Red Products")
            total_quantity = filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red']
            green_share = (filtered_df['Green'] / total_quantity * 100).round(2)
            yellow_share = (filtered_df['Yellow'] / total_quantity * 100).round(2)
            red_share = (filtered_df['Red'] / total_quantity * 100).round(2)
                    
            share_df = pd.DataFrame({
                        'Month': filtered_df['Month'],
                        'Green Share (%)': green_share,
                        'Yellow Share (%)': yellow_share,
                        'Red Share (%)': red_share
                    })
                    
            fig_pie = px.pie(share_df, values=[green_share.mean(), yellow_share.mean(), red_share.mean()], 
                                     names=['Green', 'Yellow', 'Red'], title='Average Share Distribution',color=["G","Y","R"],color_discrete_map={"G":"green","Y":"yellow","R":"red"},hole=0.5)
            st.plotly_chart(fig_pie, use_container_width=True)
                    
            st.dataframe(share_df.set_index('Month').style.format("{:.2f}").background_gradient(cmap='RdYlGn'), use_container_width=True)
        
        
        else:
            st.warning("No data available for the selected combination.")
        
        st.markdown("</div>", unsafe_allow_html=True)
 
 elif selected == "About":
    st.title("About the GYR Analysis App")
    st.markdown("""
    This advanced data analysis application is designed to provide insightful visualizations and statistics for your GYR (Green, Yellow, Red) data. 
    
    Key features include:
    - Interactive data filtering
    - Multiple analysis types (NSR, Contribution, EBITDA)
    - Dynamic visualizations with Plotly
    - Descriptive statistics and share analysis
    - Customizable Green and Yellow share adjustments
    """)
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Frame, Indenter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import time
# Set page config
def projection():
 def get_cookie_password():
    if 'cookie_password' not in st.session_state:
        st.session_state.cookie_password = secrets.token_hex(16)
    return st.session_state.cookie_password

# Initialize the encrypted cookie manager
 cookies = EncryptedCookieManager(
    prefix="sales_predictor_",
    password=get_cookie_password()
)

# Constants
 CORRECT_PASSWORD = "prasoonA1@"  # Replace with your desired password
 MAX_ATTEMPTS = 5
 LOCKOUT_DURATION = 3600  # 1 hour in seconds

 def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()
 def check_password():
    """Returns `True` if the user had the correct password."""
    if not cookies.ready():
        st.warning("Initializing cookies...")
        return False

    # Apply custom CSS
    st.markdown("""
    <style>
    .stTextInput > div > div > input {
        background-color: #f0f0f0;
        color: #333;
        border: 2px solid #4a69bd;
        border-radius: 5px;
        padding: 10px;
        font-size: 16px;
    }
    .stButton > button {
        background-color: #4a69bd;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 10px 20px;
        font-size: 16px;
        cursor: pointer;
        transition: background-color 0.3s;
    }
    .stButton > button:hover {
        background-color: #82ccdd;
    }
    .attempt-text {
        color: #ff4b4b;
        font-size: 14px;
        margin-top: 5px;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

    # Check if user is locked out
    lockout_time = cookies.get('lockout_time')
    if lockout_time is not None and time.time() < float(lockout_time):
        remaining_time = int(float(lockout_time) - time.time())
        st.error(f"Too many incorrect attempts. Please try again in {remaining_time // 60} minutes and {remaining_time % 60} seconds.")
        return False

    if 'login_attempts' not in st.session_state:
     login_attempts = cookies.get('login_attempts')
     st.session_state.login_attempts = int(login_attempts) if login_attempts is not None else 0
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if hash_password(st.session_state["password"]) == hash_password(CORRECT_PASSWORD):
            st.session_state["password_correct"] = True
            st.session_state.login_attempts = 0
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False
            st.session_state.login_attempts += 1
            if st.session_state.login_attempts >= MAX_ATTEMPTS:
                cookies['lockout_time'] = str(time.time() + LOCKOUT_DURATION)
        
        # Update login_attempts in cookies
        cookies['login_attempts'] = str(st.session_state.login_attempts)
        cookies.save()

    # First run, show input for password
    if "password_correct" not in st.session_state:
        st.markdown("<h1 style='text-align: center; color: #4a69bd;'>Sales Prediction Simulator</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center; color: #333;'>Please enter your password to access the application</h3>", unsafe_allow_html=True)
        st.text_input("Password", type="password", key="password")
        if st.button("Login"):
            password_entered()
        
        if st.session_state.login_attempts > 0:
            st.markdown(f"<p class='attempt-text'>Incorrect password. Attempt {st.session_state.login_attempts} of {MAX_ATTEMPTS}.</p>", unsafe_allow_html=True)
        
        return False
    
    # Password correct
    elif st.session_state.get("password_correct", False):
        return True
    
    # Password incorrect, show input box again
    else:
        st.markdown("<h1 style='text-align: center; color: #4a69bd;'>Sales Prediction Simulator</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center; color: #333;'>Please enter your password to access the application</h3>", unsafe_allow_html=True)
        st.text_input("Password", type="password", key="password")
        if st.button("Login"):
            password_entered()
        
        if st.session_state.login_attempts > 0:
            st.markdown(f"<p class='attempt-text'>Incorrect password. Attempt {st.session_state.login_attempts} of {MAX_ATTEMPTS}.</p>", unsafe_allow_html=True)
        
        return False
 if check_password():
  st.markdown("""
<style>
    body {
        background-color: #0e1117;
        color: #ffffff;
    }
    .stApp {
        background-image: linear-gradient(45deg, #1e3799, #0c2461);
    }
    .big-font {
        font-size: 48px !important;
        font-weight: bold;
        color: lime;
        text-align: center;
        text-shadow: 2px 2px 4px #000000;
    }
    .subheader {
        font-size: 24px;
        color: moccasin;
        text-align: center;
    }
    .stButton>button {
        background-color: #4a69bd;
        color: white;
        border-radius: 20px;
        border: 2px solid #82ccdd;
        padding: 10px 24px;
        font-size: 16px;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #82ccdd;
        color: #0c2461;
        transform: scale(1.05);
    }
    .stProgress > div > div > div > div {
        background-color: #4a69bd;
    }
    .stSelectbox {
        background-color: #1e3799;
    }
    .stDataFrame {
        background-color: #0c2461;
    }
    .metric-value {
        color: gold !important;
        font-size: 24px !important;
        font-weight: bold !important;
    }
    .metric-label {
        color: white !important;
    }
    h3 {
        color: #ff9f43 !important;
        font-size: 28px !important;
        font-weight: bold !important;
        text-shadow: 1px 1px 2px #000000;
    }
    /* Updated styles for file uploader */
    .stFileUploader {
        background-color: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
    }
    .custom-file-upload {
        display: inline-block;
        padding: 10px 20px;
        cursor: pointer;
        background-color: #4a69bd;
        color: #ffffff;
        border-radius: 5px;
        transition: all 0.3s;
    }
    .custom-file-upload:hover {
        background-color: #82ccdd;
        color: #0c2461;
    }
    .file-upload-text {
        font-size: 18px;
        color: fuchsia;
        font-weight: bold;
        margin-bottom: 10px;
    }
    /* Style for uploaded file name */
    .uploaded-filename {
        background-color: rgba(255, 255, 255, 0.2);
        color: cyan;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)
  def custom_file_uploader(label, type):
    st.markdown(f'<p class="file-upload-text">{label}</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Choose file", type=type, key="file_uploader", label_visibility="collapsed")
    return uploaded_file
  @st.cache_data
  def load_data(file):
    data = pd.read_excel(file)
    return data
  @st.cache_resource
  def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X_train, y_train)
    return model, X_test, y_test
  def create_monthly_performance_graph(data):
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
    colors = px.colors.qualitative.Pastel

    fig = go.Figure()

    for i, month in enumerate(months):
        if month != 'Oct':
            target = data[f'Month Tgt ({month})'].iloc[0]
            achievement = data[f'Monthly Achievement({month})'].iloc[0]
            percentage = (achievement / target * 100) if target != 0 else 0
            
            fig.add_trace(go.Bar(
                x=[f"{month} Tgt", f"{month} Ach"],
                y=[target, achievement],
                name=month,
                marker_color=colors[i],
                text=[f"{target:,.0f}", f"{achievement:,.0f}<br>{percentage:.1f}%"],
                textposition='auto'
            ))
        else:
            target = data['Month Tgt (Oct)'].iloc[0]
            projection = data['Predicted Oct 2024'].iloc[0]
            percentage = (projection / target * 100) if target != 0 else 0
            
            fig.add_trace(go.Bar(
                x=[f"{month} Tgt", f"{month} Proj"],
                y=[target, projection],
                name=month,
                marker_color=[colors[i], 'red'],
                text=[f"{target:,.0f}", f"{projection:,.0f}<br><span style='color:black'>{percentage:.1f}%</span>"],
                textposition='auto'
            ))

    fig.update_layout(
        title='Monthly Performance',
        plot_bgcolor='rgba(255,255,255,0.1)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='burlywood',
        title_font_color='burlywood',
        xaxis_title_font_color='burlywood',
        yaxis_title_font_color='burlywood',
        legend_font_color='burlywood',
        height=500,
        width=800,
        barmode='group'
    )
    fig.update_xaxes(tickfont_color='peru')
    fig.update_yaxes(title_text='Sales', tickfont_color='peru')
    fig.update_traces(textfont_color='black')
    
    return fig
  def create_target_vs_projected_graph(data):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=data['Zone'], y=data['Month Tgt (Oct)'], name='Month Target (Oct)', marker_color='#4a69bd'))
    fig.add_trace(go.Bar(x=data['Zone'], y=data['Predicted Oct 2024'], name='Projected Sales (Oct)', marker_color='#82ccdd'))
    
    fig.update_layout(
        title='October 2024: Target vs Projected Sales',
        barmode='group',
        plot_bgcolor='rgba(255,255,255,0.1)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='burlywood',
        title_font_color='burlywood',
        xaxis_title_font_color='burlywood',
        yaxis_title_font_color='burlywood',
        legend_font_color='burlywood',
        height=500
    )
    fig.update_xaxes(title_text='Zone', tickfont_color='peru')
    fig.update_yaxes(title_text='Sales', tickfont_color='peru')
    
    return fig

  def prepare_data_for_pdf(data):
    # Filter out specified zones
    excluded_zones = ['Bihar', 'J&K', 'North-I', 'Punjab,HP and J&K', 'U.P.+U.K.', 'Odisha+Jharkhand+Bihar']
    filtered_data = data[~data['Zone'].isin(excluded_zones)]

    # Further filter to include only LC and PHD brands
    filtered_data = filtered_data[filtered_data['Brand'].isin(['LC', 'PHD'])]

    # Calculate totals for LC, PHD, and LC+PHD
    lc_data = filtered_data[filtered_data['Brand'] == 'LC']
    phd_data = filtered_data[filtered_data['Brand'] == 'PHD']
    lc_phd_data = filtered_data

    totals = []
    for brand_data, brand_name in [(lc_data, 'LC'), (phd_data, 'PHD'), (lc_phd_data, 'LC+PHD')]:
        total_month_tgt_oct = brand_data['Month Tgt (Oct)'].sum()
        total_predicted_oct_2024 = brand_data['Predicted Oct 2024'].sum()
        total_oct_2023 = brand_data['Total Oct 2023'].sum()
        total_yoy_growth = (total_predicted_oct_2024 - total_oct_2023) / total_oct_2023 * 100

        totals.append({
            'Zone': 'All India Total',
            'Brand': brand_name,
            'Month Tgt (Oct)': total_month_tgt_oct,
            'Predicted Oct 2024': total_predicted_oct_2024,
            'Total Oct 2023': total_oct_2023,
            'YoY Growth': total_yoy_growth
        })

    # Concatenate filtered data with totals
    final_data = pd.concat([filtered_data, pd.DataFrame(totals)], ignore_index=True)

    # Round the values
    final_data['Month Tgt (Oct)'] = final_data['Month Tgt (Oct)'].round().astype(int)
    final_data['Predicted Oct 2024'] = final_data['Predicted Oct 2024'].round().astype(int)
    final_data['Total Oct 2023'] = final_data['Total Oct 2023'].round().astype(int)
    final_data['YoY Growth'] = final_data['YoY Growth'].round(2)

    return final_data

  def create_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            rightMargin=inch, leftMargin=inch,
                            topMargin=0.2*inch, bottomMargin=0.5*inch)
    elements = []

    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    title_style.alignment = 1
    title = Paragraph("Sales Predictions for October 2024", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.15*inch))
    elements.append(Paragraph("<br/><br/>", styles['Normal']))

    # Prepare data for PDF
    pdf_data = prepare_data_for_pdf(data)

    table_data = [['Zone', 'Brand', 'Month Tgt (Oct)', 'Predicted Oct 2024', 'Total Oct 2023', 'YoY Growth']]
    for _, row in pdf_data.iterrows():
        table_data.append([
            row['Zone'],
            row['Brand'],
            f"{row['Month Tgt (Oct)']:,}",
            f"{row['Predicted Oct 2024']:,}",
            f"{row['Total Oct 2023']:,}",
            f"{row['YoY Growth']:.2f}%"
        ])
    table_data[0][-1] = table_data[0][-1] + "*"  

    table = Table(table_data, colWidths=[1.25*inch, 0.80*inch, 1.5*inch, 1.75*inch, 1.5*inch, 1.20*inch], 
                  rowHeights=[0.60*inch] + [0.38*inch] * (len(table_data) - 1))
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A708B')),
        ('BACKGROUND', (0, len(table_data) - 3), (-1, len(table_data) - 1), colors.orange),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -4), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey)
    ])
    table.setStyle(style)
    elements.append(table)

    footnote_style = getSampleStyleSheet()['Normal']
    footnote_style.fontSize = 8
    footnote_style.leading = 10 
    footnote_style.alignment = 0
    footnote = Paragraph("*YoY Growth is calculated using October 2023 sales and predicted October 2024 sales.", footnote_style)
    indented_footnote = Indenter(left=-0.75*inch)
    elements.append(Spacer(1, 0.15*inch))
    elements.append(indented_footnote)
    elements.append(footnote)
    elements.append(Indenter(left=0.5*inch))

    doc.build(elements)
    buffer.seek(0)
    return buffer
  from reportlab.lib import colors

  def style_dataframe(df):
    styler = df.style

    for col in df.columns:
        if df[col].dtype in ['float64', 'int64']:
            styler.apply(lambda x: ['background-color: #f0f0f0'] * len(x), subset=[col])
        else:
            styler.apply(lambda x: ['background-color: #f0f0f0'] * len(x), subset=[col])

    numeric_format = {
        'October 2024 Target': '{:.0f}',
        'October Projection': '{:.2f}',
        'October 2023 Sales': '{:.0f}',
        'YoY Growth(Projected)': '{:.2f}%'
    }
    styler.format(numeric_format)
    return styler
  def main():
    st.markdown('<p class="big-font">Sales Prediction Simulator</p>', unsafe_allow_html=True)
    st.markdown('<p class="subheader">Upload your data and unlock the future of sales!</p>', unsafe_allow_html=True)
    uploaded_file = custom_file_uploader("Choose your sales data file (Excel format)", ["xlsx"])

    if uploaded_file is not None:
        st.markdown(f'<div class="uploaded-filename">Uploaded file: {uploaded_file.name}</div>', unsafe_allow_html=True)
        data = load_data(uploaded_file)

        features = ['Month Tgt (Oct)', 'Monthly Achievement(Sep)', 'Total Sep 2023', 'Total Oct 2023',
                    'Monthly Achievement(Apr)', 'Monthly Achievement(May)', 'Monthly Achievement(June)',
                    'Monthly Achievement(July)', 'Monthly Achievement(Aug)']

        X = data[features]
        y = data['Total Oct 2023']

        model, X_test, y_test = train_model(X, y)

        st.sidebar.header("Control Panel")
        
        # Initialize session state for filters if not already present
        if 'selected_brands' not in st.session_state:
            st.session_state.selected_brands = []
        if 'selected_zones' not in st.session_state:
            st.session_state.selected_zones = []

        # Brand filter
        st.sidebar.subheader("Select Brands")
        for brand in data['Brand'].unique():
            if st.sidebar.checkbox(brand, key=f"brand_{brand}"):
                if brand not in st.session_state.selected_brands:
                    st.session_state.selected_brands.append(brand)
            elif brand in st.session_state.selected_brands:
                st.session_state.selected_brands.remove(brand)

        # Zone filter
        st.sidebar.subheader("Select Zones")
        for zone in data['Zone'].unique():
            if st.sidebar.checkbox(zone, key=f"zone_{zone}"):
                if zone not in st.session_state.selected_zones:
                    st.session_state.selected_zones.append(zone)
            elif zone in st.session_state.selected_zones:
                st.session_state.selected_zones.remove(zone)

        # Apply filters
        if st.session_state.selected_brands and st.session_state.selected_zones:
            filtered_data = data[data['Brand'].isin(st.session_state.selected_brands) & 
                                 data['Zone'].isin(st.session_state.selected_zones)]
        else:
            filtered_data = data
        col1, col2 = st.columns(2)

        with col1:
            
            st.markdown("<h3>Model Performance Metrics</h3>", unsafe_allow_html=True)
            y_pred = model.predict(X_test)
            mse = mean_squared_error(y_test, y_pred)
            r2 = r2_score(y_test, y_pred)

            st.markdown(f'<div class="metric-label">Accuracy Score</div><div class="metric-value">{r2:.2f}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-label">Error Margin</div><div class="metric-value">{np.sqrt(mse):.2f}</div>', unsafe_allow_html=True)

            feature_importance = pd.DataFrame({
                'feature': features,
                'importance': model.feature_importances_
            }).sort_values('importance', ascending=False)

            fig_importance = px.bar(feature_importance, x='importance', y='feature', orientation='h',
                                    title='Feature Impact Analysis', labels={'importance': 'Impact', 'feature': 'Feature'})
            fig_importance.update_layout(
                plot_bgcolor='rgba(255,255,255,0.1)', 
                paper_bgcolor='rgba(0,0,0,0)', 
                font_color='burlywood',
                title_font_color='burlywood',
                xaxis_title_font_color='burlywood',
                yaxis_title_font_color='burlywood',
                legend_font_color='burlywood'
            )
            fig_importance.update_xaxes(tickfont_color='peru')
            fig_importance.update_yaxes(tickfont_color='peru')
            filtered_data['FY2025 Till Sep']= filtered_data['Monthly Achievement(Apr)']+filtered_data['Monthly Achievement(May)']+filtered_data['Monthly Achievement(June)']+filtered_data['Monthly Achievement(July)']+filtered_data['Monthly Achievement(Aug)']+filtered_data['Monthly Achievement(Sep)']
            fig_predictions1 = go.Figure()
            fig_predictions1.add_trace(go.Bar(x=filtered_data['Zone'], y=filtered_data['FY 2025 Till Sep'], name='Till SepSales', marker_color='#4a69bd'))
            fig_predictions1.update_layout(
                title='FY 2025 Till Sep', 
                barmode='group', 
                plot_bgcolor='rgba(255,255,255,0.1)', 
                paper_bgcolor='rgba(0,0,0,0)', 
                font_color='burlywood',
                xaxis_title_font_color='burlywood',
                yaxis_title_font_color='burlywood',
                title_font_color='burlywood',
                legend_font_color='burlywood'
            )
            fig_predictions1.update_xaxes(title_text='Zone', tickfont_color='peru')
            fig_predictions1.update_yaxes(title_text='Sales', tickfont_color='peru')
            st.plotly_chart(fig_importance, use_container_width=True)
            st.plotly_chart(fig_predictions1, use_container_width=True)

        with col2:
            st.markdown("<h3>Sales Forecast Visualization</h3>", unsafe_allow_html=True)
            X_2024 = filtered_data[features].copy()
            X_2024['Total Oct 2023'] = filtered_data['Total Oct 2023']
            predictions_2024 = model.predict(X_2024)
            filtered_data['Predicted Oct 2024'] = predictions_2024
            filtered_data['YoY Growth'] = (filtered_data['Predicted Oct 2024'] - filtered_data['Total Oct 2023']) / filtered_data['Total Oct 2023'] * 100

            fig_predictions = go.Figure()
            fig_predictions.add_trace(go.Bar(x=filtered_data['Zone'], y=filtered_data['Total Oct 2023'], name='Oct 2023 Sales', marker_color='#4a69bd'))
            fig_predictions.add_trace(go.Bar(x=filtered_data['Zone'], y=filtered_data['Predicted Oct 2024'], name='Predicted Oct 2024 Sales', marker_color='#82ccdd'))
            fig_predictions.update_layout(
                title='Sales Projection: 2023 vs 2024', 
                barmode='group', 
                plot_bgcolor='rgba(255,255,255,0.1)', 
                paper_bgcolor='rgba(0,0,0,0)', 
                font_color='burlywood',
                xaxis_title_font_color='burlywood',
                yaxis_title_font_color='burlywood',
                title_font_color='burlywood',
                legend_font_color='burlywood'
            )
            fig_predictions.update_xaxes(title_text='Zone', tickfont_color='peru')
            fig_predictions.update_yaxes(title_text='Sales', tickfont_color='peru')
            st.plotly_chart(fig_predictions, use_container_width=True)
            fig_target_vs_projected = create_target_vs_projected_graph(filtered_data)
            st.plotly_chart(fig_target_vs_projected, use_container_width=True)
        st.markdown("<h3>Monthly Performance by Zone and Brand</h3>", unsafe_allow_html=True)
        
        # Create dropdowns for zone and brand selection
        col_zone, col_brand = st.columns(2)
        with col_zone:
            selected_zone = st.selectbox("Select Zone", options=filtered_data['Zone'].unique())
        with col_brand:
            selected_brand = st.selectbox("Select Brand", options=filtered_data[filtered_data['Zone']==selected_zone]['Brand'].unique())
        
        # Filter data based on selection
        selected_data = filtered_data[(filtered_data['Zone'] == selected_zone) & (filtered_data['Brand']==selected_brand)]
        if not selected_data.empty:
            fig_monthly_performance = create_monthly_performance_graph(selected_data)
            
            # Update the graph with the selected data
            months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
            for i, month in enumerate(months):
                if month != 'Oct':
                    fig_monthly_performance.data[i].y = [
                        selected_data[f'Month Tgt ({month})'].iloc[0],
                        selected_data[f'Monthly Achievement({month})'].iloc[0]
                    ]
                else:
                    fig_monthly_performance.data[i].y = [
                        selected_data['Month Tgt (Oct)'].iloc[0],
                        selected_data['Predicted Oct 2024'].iloc[0]
                    ]
            
            st.plotly_chart(fig_monthly_performance, use_container_width=True)
        else:
            st.warning("No data available for the selected Zone and Brand combination.")
        st.markdown("<h3>Detailed Sales Forecast</h3>", unsafe_allow_html=True)
        
        share_df = pd.DataFrame({
           'Zone': filtered_data['Zone'],
            'Brand': filtered_data['Brand'],
             'October 2024 Target': filtered_data['Month Tgt (Oct)'],
           'October Projection': filtered_data['Predicted Oct 2024'],
           'October 2023 Sales': filtered_data['Total Oct 2023'],
          'YoY Growth(Projected)': filtered_data['YoY Growth']
             })
        styled_df = style_dataframe(share_df)
        st.dataframe(styled_df, use_container_width=True,hide_index=True)

        pdf_buffer = create_pdf(filtered_data)
        st.download_button(
            label="Download Forecast Report",
            data=pdf_buffer,
            file_name="sales_forecast_2024.pdf",
            mime="application/pdf"
        )
    else:
        st.info("Upload your sales data to begin the simulation!")

  if __name__ == "__main__":
    main()
 else:
    st.stop()
import distinctipy
from pathlib import Path
from collections import defaultdict
from matplotlib.backends.backend_pdf import PdfPages
def market_share():
    # Enhanced styling configuration
    THEME = {
        'PRIMARY': '#2563eb',
        'SECONDARY': '#64748b',
        'SUCCESS': '#10b981',
        'WARNING': '#f59e0b',
        'DANGER': '#ef4444',
        'BACKGROUND': '#ffffff',
        'SIDEBAR': '#f8fafc',
        'TEXT': '#1e293b',
        'HEADER': '#0f172a'
    }

    # Custom CSS with modern styling
    st.markdown("""
        <style>
        /* Global Styles */
        .stApp {
            background-color: #ffffff;
        }
        
        /* Main Content Area */
        .main {
            background-color: #f8fafc;
            padding: 2rem;
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
        }
        
        /* Headers */
        h1 {
            color: #0f172a;
            font-size: 2.25rem !important;
            font-weight: 700 !important;
            margin-bottom: 1.5rem !important;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #e2e8f0;
        }
        
        h2 {
            color: #1e293b;
            font-size: 1.875rem !important;
            font-weight: 600 !important;
            margin-top: 2rem !important;
        }
        
        h3 {
            color: #334155;
            font-size: 1.5rem !important;
            font-weight: 600 !important;
        }
        
        /* Sidebar */
        .css-1d391kg {
            background-color: #f8fafc;
            padding: 2rem 1.5rem;
            border-right: 1px solid #e2e8f0;
        }
        
        /* Cards */
        .stMetric {
            background-color: white;
            padding: 1rem;
            border-radius: 0.75rem;
            box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1);
            transition: transform 0.2s;
        }
        
        .stMetric:hover {
            transform: translateY(-2px);
        }
        
        /* Buttons */
        .stButton>button {
            background-color: #2563eb;
            color: white;
            border: none;
            padding: 0.5rem 1.25rem;
            border-radius: 0.5rem;
            font-weight: 500;
            transition: all 0.2s;
            box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
        }
        
        .stButton>button:hover {
            background-color: #1d4ed8;
            transform: translateY(-1px);
            box-shadow: 0 4px 6px rgba(37, 99, 235, 0.3);
        }
        
        /* Select Boxes */
        .stSelectbox>div>div {
            background-color: white;
            border-radius: 0.5rem;
            border: 1px solid #e2e8f0;
        }
        
        /* Expander */
        .streamlit-expanderHeader {
            background-color: white;
            border-radius: 0.5rem;
            border: 1px solid #e2e8f0;
            padding: 0.75rem 1rem;
        }
        
        /* Plots */
        .stPlot {
            background-color: white;
            padding: 1rem;
            border-radius: 0.75rem;
            box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1);
        }
        
        /* Loading Animation */
        .stSpinner {
            text-align: center;
            color: #2563eb;
        }
        
        /* Tooltips */
        .tooltip {
            position: relative;
            display: inline-block;
            border-bottom: 1px dotted #64748b;
        }
        
        .tooltip .tooltiptext {
            visibility: hidden;
            background-color: #1e293b;
            color: white;
            text-align: center;
            padding: 0.5rem 1rem;
            border-radius: 0.375rem;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            transform: translateX(-50%);
            opacity: 0;
            transition: opacity 0.2s;
        }
        
        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
        </style>
    """, unsafe_allow_html=True)
    # Global color mapping to maintain consistency across states and months
    COMPANY_COLORS = {}
    @st.cache_data
    def generate_distinct_color(existing_colors):
        """Generate a new distinct color that's visually different from existing ones"""
        if existing_colors:
            return distinctipy.get_colors(1, existing_colors)[0]
        return distinctipy.get_colors(1)[0]
    @st.cache_data
    def get_company_color(company):
     if 'company_colors' not in st.session_state:
        st.session_state.company_colors = {}
     if company not in st.session_state.company_colors:
        existing_colors = list(st.session_state.company_colors.values())
        st.session_state.company_colors[company] = generate_distinct_color(existing_colors)
     return st.session_state.company_colors[company]
    @st.cache_data
    def load_and_process_data(uploaded_file):
        """Load Excel file and return dict of dataframes and sheet names"""
        xl = pd.ExcelFile(uploaded_file)
        states = xl.sheet_names
        state_dfs = {state: pd.read_excel(uploaded_file, sheet_name=state) for state in states}
        
        # Initialize colors for all unique companies across all states
        all_companies = set()
        for df in state_dfs.values():
            all_companies.update(df['Company'].unique())
        
        # Assign consistent colors to all companies
        for company in all_companies:
            get_company_color(company)
            
        return state_dfs, states
    def get_available_months(df):
     share_cols = [col for col in df.columns if col.startswith('Share_')]
     months = [col.split('_')[1] for col in share_cols]
     month_order = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }
     sorted_months = sorted(months, key=lambda x: month_order[x])
     return sorted_months
    @st.cache_data
    def create_share_plot(df, month):
     from matplotlib.lines import Line2D
     def draw_curly_brace(ax, x, y1, y2):
        mid_y = (y1 + y2) / 2
        width = 0.03
        
        # Calculate points for the brace
        brace_points = [
            # Top horizontal line
            [x, y1],
            [x + width, y1],
            # Top curve
            [x + width, y1],
            [x + width, (y1 + mid_y)/2],
            # Middle point
            [x, mid_y],
            # Bottom curve
            [x + width, (mid_y + y2)/2],
            [x + width, y2],
            # Bottom horizontal line
            [x + width, y2],
            [x, y2]
        ]
        
        # Draw the brace using line segments
        for i in range(len(brace_points)-1):
            line = Line2D([brace_points[i][0], brace_points[i+1][0]],
                         [brace_points[i][1], brace_points[i+1][1]],
                         color='#2c3e50',
                         linewidth=1.5)
            ax.add_line(line)
        
        return mid_y
     def cascade_label_positions(positions, y_max, min_gap=12):
        if not positions:
            return [], {}
        
        # Group positions by x-coordinate (price range)
        x_groups = {}
        for vol, original_y, color, x_pos in positions:
            if x_pos not in x_groups:
                x_groups[x_pos] = []
            x_groups[x_pos].append((vol, original_y, color, x_pos))
        
        # Sort x_positions from right to left
        x_positions = sorted(x_groups.keys(), reverse=True)
        
        y_range = y_max * 0.9
        min_allowed_y = y_max * 0.05
        total_price_ranges = len(x_positions)
        height_per_range = y_range / total_price_ranges
        
        adjusted = []
        group_info = {}  # Store information about each group for brackets
        
        for i, x_pos in enumerate(x_positions):
            group = x_groups[x_pos]
            n_labels = len(group)
            
            top_y = y_range - (i * height_per_range)
            bottom_y = top_y - height_per_range
            group_gap = min(min_gap, height_per_range / (n_labels + 1))
            
            group = sorted(group, key=lambda x: x[1])
            group_volumes = []
            group_positions = []
            
            for j, (vol, original_y, color, x) in enumerate(group):
                label_y = top_y - ((j + 1) * group_gap)
                label_y = max(label_y, min_allowed_y)
                
                adjusted.append((vol, original_y, label_y, color, x_pos))
                group_volumes.append(vol)
                group_positions.append(label_y)
            
            if len(group) > 1:  # Only store group info if there are multiple companies
                group_info[x_pos] = {
                    'total_volume': sum(group_volumes),
                    'top_y': max(group_positions),
                    'bottom_y': min(group_positions)
                }
        return adjusted, group_info
     def adjust_label_positions(positions, y_max, min_gap=12):
        """
        Adjust label positions while respecting the plot's y-axis scale
        positions: list of (volume, original_y, color, x_pos) tuples
        y_max: maximum y value of the plot
        """
        if not positions:
            return positions
        
        # Sort positions by original y-coordinate
        positions = sorted(positions, key=lambda x: x[1])
        
        # Calculate the available space for labels
        y_range = y_max * 0.9  # Use 90% of the plot height for labels
        
        # Calculate optimal gap based on number of labels and available space
        n_labels = len(positions)
        optimal_gap = min(min_gap, y_range / (n_labels + 1))
        
        adjusted = []
        used_positions = set()
        min_allowed_y = y_max * 0.0005
        for vol, original_y, color, x_pos in positions:
            # Try to keep label close to original position if possible
            label_y = max(original_y, min_allowed_y)
            
            # Check for overlap with existing labels
            while any(abs(label_y - used_y) < optimal_gap for used_y in used_positions):
                label_y += optimal_gap
                
                # If we've gone too high, try positioning below
                if label_y > y_range:
                    label_y = min_allowed_y
                    while any(abs(label_y - used_y) < optimal_gap for used_y in used_positions):
                        label_y += optimal_gap
                        if label_y > y_range:
                         optimal_gap *= 0.8
                         label_y = max(original_y, min_allowed_y)
            used_positions.add(label_y)
            adjusted.append((vol, original_y, label_y, color, x_pos))
        
        return adjusted
     def check_overlap(y1, y2, height=10):  # height is the estimated text height in points
        return abs(y1 - y2) < height
     def adjust_positions(positions, min_gap=10):
        """
        Adjust y-positions of labels to prevent overlap
        positions: list of (volume, original_y, color, x_pos) tuples
        """
        if not positions:
            return positions
        
        # Sort positions by y-coordinate
        positions = sorted(positions, key=lambda x: x[1])
        
        adjusted_positions = [positions[0]]  # Keep first position as is
        
        # Adjust subsequent positions if they overlap
        for vol, y_pos, color, x_pos in positions[1:]:
            prev_y = adjusted_positions[-1][1]
            
            # If current position overlaps with previous
            if check_overlap(y_pos, prev_y):
                # Place new label above the previous one with minimum gap
                new_y = prev_y + min_gap
            else:
                new_y = y_pos
                
            adjusted_positions.append((vol, new_y, color, x_pos))
        
        return adjusted_positions

     plt.style.use('seaborn-v0_8-whitegrid')
     plt.rcParams.update({
        'font.family': 'sans-serif',
        'font.size': 10,
        'axes.labelweight': 'bold',
        'axes.titleweight': 'bold',
        'figure.facecolor': 'white',
        'axes.facecolor': '#f8f9fa',
        'grid.alpha': 0.2,
        'grid.color': '#b4b4b4',
        'figure.dpi': 120,
        'axes.spines.top': False,
        'axes.spines.right': False,
        'axes.linewidth': 1.5
    })
    
    # Data preparation (same as before)
     month_data = df[['Company', f'Share_{month}', f'WSP_{month}', f'Vol_{month}']].copy()
     month_data.columns = ['Company', 'Share', 'WSP', 'Volume']
    
     min_price = (month_data['WSP'].min() // 10) * 10
     max_price = (month_data['WSP'].max() // 10 + 1) * 10
     price_ranges = pd.interval_range(start=min_price, end=max_price, freq=10)
    
     month_data['Price_Range'] = pd.cut(month_data['WSP'], bins=price_ranges)
    
     pivot_df = pd.pivot_table(
        month_data,
        values=['Share', 'Volume'],
        index='Price_Range',
        columns='Company',
        aggfunc='sum',
        fill_value=0
    )
    
     share_df = pivot_df['Share']
     volume_df = pivot_df['Volume']
    
     share_df = share_df.loc[:, (share_df != 0).any(axis=0)]
     volume_df = volume_df.loc[:, (volume_df != 0).any(axis=0)]
    
     company_wsps = {company: month_data[month_data['Company'] == company]['WSP'].iloc[0]
                   for company in share_df.columns}
     sorted_companies = sorted(company_wsps.keys(), key=lambda x: company_wsps[x])
    
     share_df = share_df[sorted_companies]
     volume_df = volume_df[sorted_companies]
    
    # Create figure with more refined dimensions
     fig, ax1 = plt.subplots(figsize=(14, 9), dpi=120)
     ax2 = ax1.twinx()
    
    # Plot stacked bars with enhanced styling
     bottom = np.zeros(len(share_df))
     volume_positions = []
    
    # Calculate total share and volume for each price range
     total_shares = share_df.sum(axis=1)
     total_volumes = volume_df.sum(axis=1)
    
     for company in sorted_companies:
        values = share_df[company].values
        ax1.bar(range(len(share_df)), 
                values, 
                bottom=bottom,
                label=company,
                color=get_company_color(company),
                alpha=0.95,  # Slightly transparent bars
                edgecolor='white',  # White edges for contrast
                linewidth=0.5)
        
        # Add labels for individual company shares
        for i, v in enumerate(values):
            if v > 0:
                center = bottom[i] + v/2
                if v > 0.2:  # Only show percentage if > 1%
                    ax1.text(i, center, f'{v:.1f}%',
                            ha='center', va='center', 
                            fontsize=8,
                            color='white',
                            fontweight='bold')
                
                vol = volume_df.loc[share_df.index[i], company]
                if vol > 0:
                    volume_positions.append((vol, center, get_company_color(company), i))
        
        bottom += values
     max_total_share = total_shares.max()
     y_max = max_total_share * 1.15  # Add 15% padding
     ax1.set_ylim(0, y_max)
    
    # Add total share labels at the top of each stacked bar
     for i, total in enumerate(total_shares):
        ax1.text(i, total + (y_max * 0.02), f'Total: {total:.1f}%',
                ha='center', va='bottom',
                fontsize=12,
                fontweight='bold',
                color='#2c3e50')
     for i, total_vol in enumerate(total_volumes):
        ax1.text(i, -4, f'Vol: {total_vol:,.0f} MT',
                ha='center', va='top',
                fontsize=12,fontweight='bold',
                color='#34495e')
     adjusted_positions, group_info = cascade_label_positions(volume_positions, y_max)
    
    # Draw dashed lines and volume labels
     for vol, line_y, label_y, color, x_pos in adjusted_positions:
        # Draw connecting lines
        if abs(label_y - line_y) > 0.5:
            mid_x = x_pos + (len(share_df)-0.15 - x_pos) * 0.7
            ax1.plot([x_pos, mid_x, len(share_df)-0.15], 
                    [line_y, label_y, label_y],
                    color=color, linestyle='--', alpha=1, linewidth=1)
        else:
            ax1.plot([x_pos, len(share_df)-0.15], [line_y, line_y],
                    color=color, linestyle='--', alpha=1, linewidth=1)
        
        # Add individual volume labels
        label = f'{vol:,.0f} MT'
        ax2.text(0.98, label_y, label,
                transform=ax1.get_yaxis_transform(),
                va='center', ha='left',
                color=color,
                fontsize=11,
                fontweight='bold',
                bbox=dict(facecolor='white',
                         edgecolor='none',
                         alpha=1,
                         pad=1))
    
    # Draw braces and total volume labels for groups
     for x_pos, info in group_info.items():
        # Draw brace
        brace_x = 1.15  # Position after individual volume labels
        mid_y = draw_curly_brace(ax2, brace_x, info['top_y'], info['bottom_y'])
        
        # Add total volume label with nice formatting
        total_label = f'Total: {info["total_volume"]:,.0f} MT'
        ax2.text(brace_x + 0.005, mid_y, total_label,
                transform=ax1.get_yaxis_transform(),
                va='center', ha='left',
                color='#2c3e50',
                fontsize=11,
                fontweight='bold',
                bbox=dict(facecolor='white',
                         edgecolor='#bdc3c7',
                         boxstyle='round,pad=0.5',
                         alpha=0.9))
     plt.subplots_adjust(right=0.75)
     x_labels = [f'₹{interval.left:.0f}-{interval.right:.0f}'
                for interval in share_df.index]
     ax1.set_xticks(range(len(x_labels)))
     ax1.set_xticklabels(x_labels, ha='center')
    
    # Enhanced titles with better spacing and styling
     plt.suptitle('Market Share Distribution by Price Range',
                fontsize=16, y=1.05,
                color='#2c3e50',
                fontweight='bold')
     plt.title(f'{month.capitalize()}',
             fontsize=14,
             pad=15,
             color='#34495e',y=1.11)
    
    # Enhanced axis labels
     ax1.set_xlabel('WSP Price Range (₹)',
                  fontsize=11,
                  labelpad=15,
                  color='#2c3e50')
     ax1.set_ylabel('Market Share (%)',
                  fontsize=11,
                  labelpad=10,
                  color='#2c3e50')
    
    # Enhanced legend
     legend_labels = [f'{company} (WSP: ₹{company_wsps[company]:.0f})'
                    for company in sorted_companies]
     legend = ax1.legend(legend_labels,
                       bbox_to_anchor=(1.28, 0.8),
                       loc='upper left',
                       fontsize=9,
                       frameon=True,
                       facecolor='white',
                       edgecolor='brown',
                       title='Companies',
                       title_fontsize=10,
                       borderpad=1)
     legend.get_frame().set_alpha(1)
    
    # Clear right axis
     ax2.set_yticks([])
    
    # Enhanced total market size box
     total_market_size = volume_df.sum().sum()
     plt.figtext(0.45, 0.925,
                f'Total Market Size: {total_market_size:,.0f} MT',
                ha='center', va='center',
                bbox=dict(facecolor='#f8f9fa',
                         edgecolor='#bdc3c7',
                         boxstyle='round,pad=0.7',
                         alpha=1),
                fontsize=11,
                fontweight='bold',
                color='#2c3e50')
    
    # Adjusted layout
     plt.tight_layout()
     plt.subplots_adjust(right=0.82, bottom=0.2, top=0.88)
    
     return fig
    @st.cache_data
    def calculate_share_changes(shares, months):
        """Calculate sequential and total changes in share percentage"""
        sequential_changes = []
        for i in range(1, len(shares)):
            change = (shares[i] - shares[i-1])/shares[i]*100
            sequential_changes.append(change)
        
        total_change = (shares[-1] - shares[0])/shares[0]*100
        return sequential_changes, total_change
    @st.cache_data
    def create_trend_line_plot(_df, selected_companies, state_name):
     df = _df.copy()
     share_cols = [col for col in df.columns if col.startswith('Share_')]
     months = [col.split('_')[1] for col in share_cols]
     month_order = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }
     month_col_pairs = [(col, month_order[month]) 
                      for col, month in zip(share_cols, months)]
     sorted_pairs = sorted(month_col_pairs, key=lambda x: x[1])
    
     sorted_share_cols = [pair[0] for pair in sorted_pairs]
     sorted_months = [col.split('_')[1] for col in sorted_share_cols]
    
     fig, ax = plt.subplots(figsize=(14, 8))
    
     lines = []
     legend_labels = []
    
     for company in selected_companies:
        color = get_company_color(company)
        company_shares = df[df['Company'] == company][sorted_share_cols].iloc[0].values
        avg_share = company_shares.mean()
        
        line = ax.plot(range(len(sorted_months)), company_shares, 
                      marker='o', linewidth=2, color=color,
                      label=company)[0]
        lines.append(line)
        
        ax.axhline(y=avg_share, color=color, linestyle='--', alpha=0.3)
        
        for i, share in enumerate(company_shares):
            ax.annotate(f'{share:.1f}%', 
                      (i, share),
                      xytext=(0, 0),
                      textcoords='offset points',
                      ha='center',
                      va='bottom',
                      fontsize=8)
        
        sequential_changes, total_change = calculate_share_changes(company_shares, sorted_months)
        for i, change in enumerate(sequential_changes):
            mid_x = (i + 0.5)
            mid_y = (company_shares[i] + company_shares[i + 1]) / 2
            arrow_color = 'green' if change > 0 else 'red'
            arrow_symbol = '↑' if change > 0 else '↓'
            ax.annotate(f'{arrow_symbol}{abs(change):.1f}%',
                      (mid_x, mid_y),
                      xytext=(0, 0 if i % 2 == 0 else 0),
                      textcoords='offset points',
                      ha='center',
                      va='center',
                      color=arrow_color,
                      fontsize=8,
                      bbox=dict(facecolor='white', 
                              edgecolor='none',
                              alpha=0.7,
                              pad=0.5))
        
        change_symbol = '↑' if total_change > 0 else '↓'
        legend_labels.append(
            f"{company} (Avg: {avg_share:.1f}% | Total Change: {change_symbol}{abs(total_change):.1f}%)"
        )
    
     plt.title(f'Market Share Trends Over Time - {state_name}', 
             fontsize=20, 
             pad=20, 
             fontweight='bold')
     plt.xlabel('Months', fontsize=12, fontweight='bold')
     plt.ylabel('Market Share (%)', fontsize=12, fontweight='bold')
     plt.xticks(range(len(sorted_months)), sorted_months, rotation=45)
     plt.grid(True, linestyle='--', alpha=0.3)
     ax.legend(lines, legend_labels,
             bbox_to_anchor=(1.15, 1),
             loc='upper left',
             borderaxespad=0.,
             frameon=True,
             fontsize=10,
             title='Companies with Average Share & Total Change',
             title_fontsize=12,
             edgecolor='gray')
     ax.set_facecolor('#f8f9fa')
     fig.patch.set_facecolor('#ffffff')
     plt.tight_layout()
     return fig

    def create_title_page(state_name):
     fig, ax = plt.subplots(figsize=(11.7, 8.3))  # A4 size
     ax.axis('off')
     ax.text(0.5, 0.6, 'Market Share Analysis Report', 
            horizontalalignment='center',
            fontsize=24,
            fontweight='bold')
     ax.text(0.5, 0.5, f'State: {state_name}', 
            horizontalalignment='center',
            fontsize=20)
     current_date = datetime.now().strftime("%d %B %Y")
     ax.text(0.5, 0.4, f'Generated on: {current_date}', 
            horizontalalignment='center',
            fontsize=16)
     fig.patch.set_facecolor('#ffffff')
     return fig
    def create_dashboard_header():
        """Create an attractive dashboard header"""
        st.markdown("""
            <div style='padding: 1.5rem; background: linear-gradient(90deg, #2563eb 0%, #3b82f6 100%); 
                        border-radius: 1rem; margin-bottom: 2rem; color: white;'>
                <h1 style='color: brown; margin: 0; border: none;'>Market Share Analysis Dashboard</h1>
                <p style='margin: 0.5rem 0 0 0; opacity: 0.9;'>
                    Comprehensive market analysis and visualization tool
                </p>
            </div>
        """, unsafe_allow_html=True)

    def create_metric_card(title, value, delta=None, help_text=None):
            st.metric(
                label=title,
                value=value,
                delta=delta,
                help=help_text
            )
    def export_to_pdf(figs, filename):
        """Export multiple figures to a single PDF file"""
        with PdfPages(filename) as pdf:
            for fig in figs:
                pdf.savefig(fig, bbox_inches='tight')
                plt.close(fig)
    def main():
    # Initialize session state for storing computed figures
     if 'computed_figures' not in st.session_state:
        st.session_state.computed_figures = {}
    
     create_dashboard_header()
    
     col1, col2 = st.columns([1, 4])
    
     with col1:
        st.markdown("### 🎯 Analysis Controls")
        
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload your market share data file"
        )
        
        if uploaded_file:
            # Use cached data loading
            state_dfs, states = load_and_process_data(uploaded_file)
            
            st.markdown("### 🎯 Settings")
            selected_state = st.selectbox(
                "Select State",
                states,
                index=0,
                help="Choose the state for analysis"
            )
            
            available_months = get_available_months(state_dfs[selected_state])
            selected_months = st.multiselect(
                "Select Months",
                available_months,
                default=[available_months[0]],
                help="Choose months for comparison"
            )
            
            all_companies = state_dfs[selected_state]['Company'].unique()
            default_companies = [
                'JK Lakshmi', 'Ultratech', 'Ambuja',
                'Wonder', 'Shree', 'JK Cement (N)'
            ]
            available_defaults = [company for company in default_companies if company in all_companies]
            
            selected_companies = st.multiselect(
                "Select Companies for Trend Analysis",
                all_companies,
                default=available_defaults,
                help="Choose companies to show in the trend line graph"
            )
     with col2:
        if uploaded_file and selected_companies:
            st.markdown("### Market Share Trends")
            
            # Modified trend key to include data hash
            df_hash = hash(str(state_dfs[selected_state]))
            trend_key = f"trend_{selected_state}_{'-'.join(sorted(selected_companies))}_{df_hash}"
            
            if trend_key not in st.session_state.computed_figures:
                st.session_state.computed_figures[trend_key] = create_trend_line_plot(
                    state_dfs[selected_state], 
                    selected_companies,
                    selected_state
                )
            
            st.pyplot(st.session_state.computed_figures[trend_key])
            st.markdown("---")
        
        if uploaded_file and selected_months:
            st.markdown("### 📊 Key Metrics")
            metric_cols = st.columns(len(selected_months))
            
            for idx, month in enumerate(selected_months):
                df = state_dfs[selected_state]
                with metric_cols[idx]:
                    create_metric_card(
                        f"{month.capitalize()}",
                        f"{len(df[df[f'Share_{month}'] > 0])} Companies",
                        f"Avg WSP: ₹{df[f'WSP_{month}'].mean():.0f}",
                        "Number of active companies and average wholesale price"
                    )
            
            st.markdown("---")
            
            # Create share plots only for newly selected months
            for month in selected_months:
                plot_key = f"share_{selected_state}_{month}"
                
                # Only create plot if it hasn't been computed yet
                if plot_key not in st.session_state.computed_figures:
                    with st.spinner(f"📊 Generating visualization for {month.capitalize()}..."):
                        st.session_state.computed_figures[plot_key] = create_share_plot(state_dfs[selected_state],month)
                st.pyplot(st.session_state.computed_figures[plot_key])
                
                # Add download buttons
                col1, col2, col3 = st.columns([1, 1, 2])
                with col1:
                    buf = io.BytesIO()
                    st.session_state.computed_figures[plot_key].savefig(
                        buf,
                        format='png',
                        dpi=300,
                        bbox_inches='tight'
                    )
                    buf.seek(0)
                    st.download_button(
                        label="📥 Download PNG",
                        data=buf,
                        file_name=f'market_share_{selected_state}_{month}.png',
                        mime='image/png',
                        key=f"download_png_{month}"
                    )
                
                with col2:
                    pdf_buf = io.BytesIO()
                    with PdfPages(pdf_buf) as pdf:
                        pdf.savefig(st.session_state.computed_figures[plot_key], bbox_inches='tight')
                    pdf_buf.seek(0)
                    st.download_button(
                        label="📄 Download PDF",
                        data=pdf_buf,
                        file_name=f'market_share_{selected_state}_{month}.pdf',
                        mime='application/pdf',
                        key=f"download_pdf_{month}"
                    )
                
                st.markdown("---")
            if st.session_state.computed_figures:
             st.markdown("### 📑 Download Complete Report")
             all_pdf_buf = io.BytesIO()
             with PdfPages(all_pdf_buf) as pdf:
                # Add title page
                title_page = create_title_page(selected_state)
                pdf.savefig(title_page, bbox_inches='tight')
                plt.close(title_page)
                
                # Add trend plot if it exists
                df_hash = hash(str(state_dfs[selected_state]))
                trend_key = f"trend_{selected_state}_{'-'.join(sorted(selected_companies))}_{df_hash}"
                if trend_key in st.session_state.computed_figures:
                    pdf.savefig(st.session_state.computed_figures[trend_key], bbox_inches='tight')
                
                # Add all monthly plots
                for month in selected_months:
                    plot_key = f"share_{selected_state}_{month}"
                    if plot_key in st.session_state.computed_figures:
                        pdf.savefig(st.session_state.computed_figures[plot_key], bbox_inches='tight')
            
             all_pdf_buf.seek(0)
             st.download_button(
                label="📥 Download Complete Report (PDF)",
                data=all_pdf_buf,
                file_name=f'market_share_{selected_state}_complete_report.pdf',
                mime='application/pdf',
                key="download_complete_pdf"
            )

    if __name__ == "__main__":
        main()
def discount():
 import streamlit as st
 import pandas as pd
 import numpy as np
 import plotly.graph_objects as go
 from datetime import datetime
 import time
 import streamlit.components.v1 as components
 import io
 import warnings
 warnings.filterwarnings('ignore')
 st.markdown("""
<style>
    /* Global Styles */
    [data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    
    .stButton button {
        background-color: #3b82f6;
        color: white;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        border: none;
        transition: all 0.2s;
    }
    
    .stButton button:hover {
        background-color: #2563eb;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    /* Ticker Animation */
    @keyframes ticker {
        0% { transform: translateX(100%); }
        100% { transform: translateX(-100%); }
    }
    
    .ticker-container {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: white;
        padding: 16px;
        overflow: hidden;
        white-space: nowrap;
        position: relative;
        margin-bottom: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .ticker-content {
        display: inline-block;
        animation: ticker 2500s linear infinite;
        animation-delay: -1250s;
        padding-right: 100%;
        will-change: transform;
    }
    
    .ticker-content:hover {
        animation-play-state: paused;
    }
    
    .ticker-item {
        display: inline-block;
        margin-right: 80px;
        font-size: 16px;
        padding: 8px 16px;
        opacity: 1;
        transition: opacity 0.3s;
        background: rgba(255, 255, 255, 0.1);
        border-radius: 8px;
    }
    
    /* Enhanced Metrics */
    .state-name {
        color: #10B981;
        font-weight: 600;
    }
    
    .month-name {
        color: #60A5FA;
        font-weight: 600;
    }
    
    .discount-value {
        color: #FBBF24;
        font-weight: 600;
    }
    
    /* Card Styles */
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s;
        border: 1px solid #e2e8f0;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 600;
        color: #1e293b;
    }
    
    .metric-label {
        color: #64748b;
        font-size: 0.875rem;
        margin-top: 0.5rem;
    }
    
    /* Chart Container */
    .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin: 1rem 0;
        border: 1px solid #e2e8f0;
    }
    
    /* Selectbox Styling */
    .stSelectbox {
        background: white;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    
    /* Custom Header */
    .dashboard-header {
        padding: 1.5rem;
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: white;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .dashboard-title {
        font-size: 2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .dashboard-subtitle {
        color: #94a3b8;
        font-size: 1rem;
    }
 </style>
 """, unsafe_allow_html=True)
 @st.cache_data(ttl=3600)
 def process_excel_file(file_content, excluded_sheets):
    """Process Excel file and return processed data"""
    excel_data = io.BytesIO(file_content)
    excel_file = pd.ExcelFile(excel_data)
    processed_data = {}
    
    for sheet in excel_file.sheet_names:
        if not any(excluded_sheet in sheet for excluded_sheet in excluded_sheets):
            df = pd.read_excel(excel_data, sheet_name=sheet, usecols=range(22))
            
            # Find start index
            cash_discount_patterns = ['CASH DISCOUNT', 'Cash Discount', 'CD']
            start_idx = None
            
            for idx, value in enumerate(df.iloc[:, 0]):
                if isinstance(value, str):
                    if any(pattern.lower() in value.lower() for pattern in cash_discount_patterns):
                        start_idx = idx
                        break
            
            if start_idx is not None:
                df = df.iloc[start_idx:].reset_index(drop=True)
            
            # Trim at G. Total
            g_total_idx = None
            for idx, value in enumerate(df.iloc[:, 0]):
                if isinstance(value, str) and 'G. TOTAL' in value:
                    g_total_idx = idx
                    break
            
            if g_total_idx is not None:
                df = df.iloc[:g_total_idx].copy()
            
            processed_data[sheet] = df
            
    return processed_data

 class DiscountAnalytics:
    def __init__(self):
        self.excluded_discounts = [
            'Sub Total',
            'TOTAL OF DP PAYOUT',
            'TOTAL OF STS & RD',
            'Other (Please specify',
            'G. TOTAL'
        ]
        self.discount_mappings = {
            'group1': {
                'states': ['HP', 'JMU', 'PUN'],
                'discounts': ['CASH DISCOUNT', 'ADVANCE CD & NIL OS']
            },
            'group2': {
                'states': ['UP (W)'],
                'discounts': ['CD', 'Adv CD']
            }
        }
        
        self.combined_discount_name = 'CD and Advance CD'
        
        self.month_columns = {
            'April': {
                'quantity': 1,
                'approved': 2,
                'actual': 4
            },
            'May': {
                'quantity': 8,
                'approved': 9,
                'actual': 11
            },
            'June': {
                'quantity': 15,
                'approved': 16,
                'actual': 18
            }
        }
        self.total_patterns = ['G. TOTAL', 'G.TOTAL', 'G. Total', 'G.Total', 'GRAND TOTAL',"G. Total (STD + STS)"]
        self.excluded_states = ['MP (JK)', 'MP (U)','East']
    def create_ticker(self, data):
     """Create moving ticker with comprehensive discount information"""
     ticker_items = []
    
     # Get the last month (June in this case)
     last_month = "June"
     month_cols = self.month_columns[last_month]
    
     for state in data.keys():
        df = data[state]
        if not df.empty:
            state_text = f"<span class='state-name'>📍 {state}</span>"
            month_text = f"<span class='month-name'>📅 {last_month}</span>"
            
            # Get state group for combined discounts
            state_group = next(
                (group for group, config in self.discount_mappings.items()
                 if state in config['states']),
                None
            )
            
            discount_items = []
            
            if state_group:
                # Handle combined discounts
                relevant_discounts = self.discount_mappings[state_group]['discounts']
                combined_data = self.get_combined_data(df, month_cols, state)
                
                if combined_data:
                    actual = combined_data.get('actual', 0)
                    discount_items.append(
                        f"{self.combined_discount_name}: <span class='discount-value'>₹{actual:,.2f}</span>"
                    )
                
                # Add other non-combined discounts
                for discount in self.get_discount_types(df, state):
                    if discount != self.combined_discount_name:
                        mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == discount.strip()
                        filtered_df = df[mask]
                        if len(filtered_df) > 0:
                            actual = filtered_df.iloc[0, month_cols['actual']]
                            discount_items.append(
                                f"{discount}: <span class='discount-value'>₹{actual:,.2f}</span>"
                            )
            else:
                # Normal processing for states without combined discounts
                for discount in self.get_discount_types(df, state):
                    mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == discount.strip()
                    filtered_df = df[mask]
                    if len(filtered_df) > 0:
                        actual = filtered_df.iloc[0, month_cols['actual']]
                        discount_items.append(
                            f"{discount}: <span class='discount-value'>₹{actual:,.2f}</span>"
                        )
            
            if discount_items:
                full_text = f"{state_text} | {month_text} | {' | '.join(discount_items)}"
                ticker_items.append(f"<span class='ticker-item'>{full_text}</span>")
    
    # Repeat items 3 times for continuous flow
     ticker_items = ticker_items * 3
    
     ticker_html = f"""
     <div class="ticker-container">
        <div class="ticker-content">
            {' '.join(ticker_items)}
        </div>
     </div>
     """
     st.markdown(ticker_html, unsafe_allow_html=True)
    def create_summary_metrics(self, data):
        """Create summary metrics cards"""
        total_states = len(data)
        total_discounts = sum(len(self.get_discount_types(df)) for df in data.values())
        avg_discount = np.mean([df.iloc[0, 4] for df in data.values() if not df.empty])
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total States", total_states, "Active")
        with col2:
            st.metric("Total Discount Types", total_discounts, "Available")
        with col3:
            st.metric("Average Discount Rate", f"₹{avg_discount:,.2f}", "Per Bag")
    def create_monthly_metrics(self, data, selected_state, selected_discount):
        """Create monthly metrics based on selected discount type"""
        df = data[selected_state]
        
        if selected_discount == self.combined_discount_name:
            monthly_data = {
                month: self.get_combined_data(df, cols, selected_state)
                for month, cols in self.month_columns.items()
            }
        else:
            mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == selected_discount.strip()
            filtered_df = df[mask]
            if len(filtered_df) > 0:
                monthly_data = {
                    month: {
                        'actual': filtered_df.iloc[0, cols['actual']],
                        'approved': filtered_df.iloc[0, cols['approved']],
                        'quantity': filtered_df.iloc[0, cols['quantity']]
                    }
                    for month, cols in self.month_columns.items()
                }
        
        # Create three columns for each month
        for month, data in monthly_data.items():
            st.markdown(f"""
            <div style='text-align: center; margin-bottom: 10px;'>
                <h3 style='color: #1e293b; margin-bottom: 15px;'>{month}</h3>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                quantity = data.get('quantity', 0)
                st.metric(
                    "Quantity Sold",
                    f"{quantity:,.2f}",
                    delta=None,
                    help=f"Total quantity sold in {month}"
                )
            
            with col2:
                approved = data.get('approved', 0)
                st.metric(
                    "Approved Payout",
                    f"₹{approved:,.2f}",
                    delta=None,
                    help=f"Approved discount rate for {month}"
                )
            
            with col3:
                actual = data.get('actual', 0)
                difference = approved - actual
                delta_color = "normal" if difference >= 0 else "inverse"
                st.metric(
                    "Actual Payout",
                    f"₹{actual:,.2f}",
                    delta=f"₹{abs(difference):,.2f}" + (" under approved" if difference >= 0 else " over approved"),
                    delta_color=delta_color,
                    help=f"Actual discount rate for {month}"
                )
            
            st.markdown("---")
    def process_excel(self, uploaded_file):
        """Process uploaded Excel file using cached function"""
        return process_excel_file(uploaded_file.getvalue(), ['MP (U)', 'MP (JK)'])
    def create_trend_chart(self, data, selected_state, selected_discount):
        """Create trend chart using Plotly"""
        df = data[selected_state]
        
        if selected_discount == self.combined_discount_name:
            monthly_data = {
                month: self.get_combined_data(df, cols, selected_state)
                for month, cols in self.month_columns.items()
            }
        else:
            mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == selected_discount.strip()
            filtered_df = df[mask]
            if len(filtered_df) > 0:
                monthly_data = {
                    month: {
                        'actual': filtered_df.iloc[0, cols['actual']],
                        'approved': filtered_df.iloc[0, cols['approved']]
                    }
                    for month, cols in self.month_columns.items()
                }
        
        months = list(monthly_data.keys())
        actual_values = [data['actual'] for data in monthly_data.values()]
        approved_values = [data['approved'] for data in monthly_data.values()]
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=months,
            y=actual_values,
            name='Actual',
            line=dict(color='#10B981', width=3)
        ))
        
        fig.add_trace(go.Scatter(
            x=months,
            y=approved_values,
            name='Approved',
            line=dict(color='#3B82F6', width=3)
        ))
        
        fig.update_layout(
            title=f'Discount Trends - {selected_state}',
            xaxis_title='Month',
            yaxis_title='Discount Rate (₹/Bag)',
            template='plotly_white',
            height=400,
            margin=dict(t=50, b=50, l=50, r=50)
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Create and display the difference chart
        self.create_difference_chart(months, approved_values, actual_values, selected_state)

    def create_difference_chart(self, months, approved_values, actual_values, selected_state):
        """Create chart showing difference between approved and actual rates"""
        differences = [approved - actual for approved, actual in zip(approved_values, actual_values)]
        
        fig = go.Figure()
        
        # Add separate traces for positive and negative differences
        for i in range(len(months)):
            color = '#10B981' if differences[i] >= 0 else '#EF4444'  # Green for positive, red for negative
            fig.add_trace(go.Scatter(
                x=[months[i], months[i]],
                y=[0, differences[i]],
                mode='lines',
                line=dict(color=color, width=3),
                showlegend=False
            ))
        
        # Add markers at the difference points
        fig.add_trace(go.Scatter(
            x=months,
            y=differences,
            mode='markers',
            marker=dict(
                size=8,
                color=['#10B981' if d >= 0 else '#EF4444' for d in differences],
                line=dict(width=2, color='white')
            ),
            name='Difference'
        ))
        
        # Add a horizontal line at y=0
        fig.add_shape(
            type='line',
            x0=months[0],
            x1=months[-1],
            y0=0,
            y1=0,
            line=dict(color='gray', width=1, dash='dash')
        )
        
        fig.update_layout(
            title=f'Approved vs Actual Difference - {selected_state}',
            xaxis_title='Month',
            yaxis_title='Difference in Discount Rate (₹/Bag)',
            template='plotly_white',
            height=300,
            margin=dict(t=50, b=50, l=50, r=50)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    def get_discount_types(self, df, state=None):
     first_col = df.iloc[:, 0]
     valid_discounts = []
     if state:
        state_group = next(
            (group for group, config in self.discount_mappings.items()
             if state in config['states']),
            None
        )
        
        if state_group:
            # Get the relevant discounts for this state
            relevant_discounts = self.discount_mappings[state_group]['discounts']
            
            # Add the combined discount name if any of the discounts to combine exist
            if any(d in first_col.values for d in relevant_discounts):
                valid_discounts.append(self.combined_discount_name)
            
            # Add other discounts that aren't being combined
            for d in first_col.unique():
                if (isinstance(d, str) and 
                    d.strip() not in self.excluded_discounts and 
                    d.strip() not in relevant_discounts):
                    valid_discounts.append(d)
        else:
            # Normal processing for other states
            valid_discounts = [
                d for d in first_col.unique() 
                if isinstance(d, str) and d.strip() not in self.excluded_discounts
            ]
     else:
        # When no state is provided (for ticker), return all unique discounts
        valid_discounts = [
            d for d in first_col.unique() 
            if isinstance(d, str) and d.strip() not in self.excluded_discounts
        ]
    
     return sorted(valid_discounts)
    def get_combined_data(self, df, month_cols, state):
     combined_data = {
        'actual': np.nan, 
        'approved': np.nan,
        'quantity': np.nan
    }
    
     state_group = next(
        (group for group, config in self.discount_mappings.items()
         if state in config['states']),
        None
    )
    
     if state_group:
        relevant_discounts = self.discount_mappings[state_group]['discounts']
        mask = df.iloc[:, 0].fillna('').astype(str).str.strip().isin(relevant_discounts)
        filtered_df = df[mask]
        
        if len(filtered_df) > 0:
            # Sum up the values for all relevant discounts
            combined_data['approved'] = filtered_df.iloc[:, month_cols['approved']].sum()
            combined_data['actual'] = filtered_df.iloc[:, month_cols['actual']].sum()
            
            # Calculate total quantity and divide by 2 for CD and Advance CD
            total_quantity = filtered_df.iloc[:, month_cols['quantity']].sum()
            combined_data['quantity'] = total_quantity / 2  # Divide summed quantity by 2
    
     return combined_data
 def main():
    processor = DiscountAnalytics()
    
    # Enhanced Sidebar
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 1rem;'>
            <h2 style='color: #1e293b;'>Dashboard Controls</h2>
        </div>
        """, unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    # Enhanced Header
    st.markdown("""
    <div class='dashboard-header'>
        <div class='dashboard-title'>Discount Analytics Dashboard</div>
        <div class='dashboard-subtitle'>Monitor and analyze discount performance across states</div>
    </div>
    """, unsafe_allow_html=True)
    
    if uploaded_file is not None:
        with st.spinner('Processing data...'):
            data = processor.process_excel(uploaded_file)
            processor.create_ticker(data)
        
        # Enhanced Metrics Layout
        st.markdown("""
        <div style='margin: 2rem 0;'>
            <h3 style='color: #1e293b; margin-bottom: 1rem;'>Key Performance Indicators</h3>
        </div>
        """, unsafe_allow_html=True)
        
        processor.create_summary_metrics(data)
        
        # Enhanced Selection Controls
        st.markdown("""
        <div style='margin: 2rem 0;'>
            <h3 style='color: #1e293b; margin-bottom: 1rem;'>Detailed Analysis</h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            selected_state = st.selectbox("Select State", list(data.keys()))
        
        if selected_state:
            with col2:
                discount_types = processor.get_discount_types(data[selected_state], selected_state)
                selected_discount = st.selectbox("Select Discount Type", discount_types)
            
            # Wrap charts in custom containers
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            processor.create_monthly_metrics(data, selected_state, selected_discount)
            st.markdown("</div>", unsafe_allow_html=True)
            
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            processor.create_trend_chart(data, selected_state, selected_discount)
            st.markdown("</div>", unsafe_allow_html=True)
    
    else:
        st.markdown("""
        <div style='text-align: center; padding: 3rem; background: white; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);'>
            <h2 style='color: #1e293b; margin-bottom: 1rem;'>Welcome to Discount Analytics</h2>
            <p style='color: #64748b; margin-bottom: 2rem;'>Please upload an Excel file to begin your analysis.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Enhanced placeholder metrics
        st.markdown("<div style='margin-top: 2rem;'>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total States", "0", "Waiting")
        with col2:
            st.metric("Total Discount Types", "0", "Waiting")
        with col3:
            st.metric("Average Discount Rate", "₹0.00", "Waiting")
        st.markdown("</div>", unsafe_allow_html=True)

 if __name__ == "__main__":
    main()
def load_visit_data():
    try:
        with open('visit_data.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {'total_visits': 0, 'daily_visits': {}}
def save_visit_data(data):
    with open('visit_data.json', 'w') as f:
        json.dump(data, f)
def update_visit_count():
    visit_data = load_visit_data()
    today = datetime.now().strftime('%Y-%m-%d')
    visit_data['total_visits'] += 1
    visit_data['daily_visits'][today] = visit_data['daily_visits'].get(today, 0) + 1
    save_visit_data(visit_data)
    return visit_data['total_visits'], visit_data['daily_visits'][today]
def load_visit_data():
    try:
        with open('visit_data.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {'total_visits': 0, 'daily_visits': {}}
def save_visit_data(data):
    with open('visit_data.json', 'w') as f:
        json.dump(data, f)
def update_visit_count():
    visit_data = load_visit_data()
    today = datetime.now().strftime('%Y-%m-%d')
    visit_data['total_visits'] += 1
    visit_data['daily_visits'][today] = visit_data['daily_visits'].get(today, 0) + 1
    save_visit_data(visit_data)
    return visit_data['total_visits'], visit_data['daily_visits'][today]
def main():
    # Custom CSS for the sidebar and main content
    st.markdown("""
    <style>
    .sidebar .sidebar-content {
        background-image: linear-gradient(180deg, #2e7bcf 25%, #4527A0 100%);
        color: white;
    }
    .sidebar-text {
        color: white !important;
    }
    .stButton>button {
        width: 100%;
        border-radius: 20px;
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }
    .stProgress .st-bo {
        background-color: #4CAF50;
    }
    .stProgress .st-bp {
        background-color: #E0E0E0;
    }
    .settings-container {
        background-color: rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        padding: 20px;
        border-radius: 10px;
        margin-top: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .visit-counter {
        background-color: rgba(255, 228, 225, 0.7);
        border-radius: 10px;
        padding: 15px;
        margin-top: 20px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .visit-counter h3 {
        color: #FFD700;
        font-size: 18px;
        margin-bottom: 10px;
    }
    .visit-counter p {
        color: #8B4513;
        font-size: 14px;
        margin: 5px 0;
    }
    .user-info {
        background-color: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 10px;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)
    st.sidebar.title("Analytics Dashboard")
    if 'username' not in st.session_state:
        st.session_state.username = "Guest"
    st.sidebar.markdown(f"""
    <div class="user-info">
        <i class="fas fa-user"></i> Logged in as: {st.session_state.username}
        <br>
        <small>Last login: {datetime.now().strftime('%Y-%m-%d %H:%M')}</small>
    </div>
    """, unsafe_allow_html=True)

    # Main menu with icons and hover effects
    with st.sidebar:
        selected = option_menu(
            menu_title="Main Menu",
            options=[
                "Home", 
                "Data Management", 
                "Analysis Dashboards", 
                "Predictions", 
                "Settings"
            ],
            icons=[
                "house-fill", 
                "database-fill-gear", 
                "graph-up-arrow", 
                "lightbulb-fill", 
                "gear-fill"
            ],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent"},
                "icon": {"color": "orange", "font-size": "20px"}, 
                "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#eee"},
                "nav-link-selected": {"background-color": "rgba(255, 255, 255, 0.2)"},
            }
        )
    # Submenu based on main selection
    if selected == "Home":
        Home()
    elif selected == "Data Management":
        data_management_menu = option_menu(
            menu_title="Data Management",
            options=["Editor", "File Manager"],
            icons=["pencil-square", "folder"],
            orientation="horizontal",
        )
        if data_management_menu == "Editor":
            excel_editor_and_analyzer()
        elif data_management_menu == "File Manager":
            folder_menu()
    elif selected == "Analysis Dashboards":
        analysis_menu = option_menu(
            menu_title="Analysis Dashboards",
            options=["WSP Analysis", "Sales Dashboard","Sales Review Report","Market Share Analysis","Discount Analysis", "Product-Mix", "Segment-Mix","Geo-Mix"],
            icons=["clipboard-data", "cash","bar-chart", "arrow-up-right", "shuffle", "globe"],
            orientation="horizontal",
        )
        if analysis_menu == "WSP Analysis":
            wsp_analysis_dashboard()
        elif analysis_menu == "Sales Dashboard":
            sales_dashboard()
        elif analysis_menu == "Sales Review Report":
            sales_review_report_generator()
        elif analysis_menu == "Product-Mix":
            normal()
        elif analysis_menu == "Segment-Mix":
            trade()
        elif analysis_menu == "Market Share Analysis":
            market_share()
        elif analysis_menu == "Geo-Mix":
            green()
        elif analysis_menu == "Discount Analysis":
            discount()
    elif selected == "Predictions":
        prediction_menu = option_menu(
            menu_title="Predictions",
            options=["WSP Projection","Sales Projection"],
            icons=["bar-chart", "graph-up-arrow"],
            orientation="horizontal",
        )
        if prediction_menu == "WSP Projection":
            descriptive_statistics_and_prediction()
        elif prediction_menu == "Sales Projection":
            projection()
    elif selected == "Settings":
        st.title("Settings")
        st.markdown('<div class="settings-container">', unsafe_allow_html=True)
        st.subheader("User Settings")
        username = st.text_input("Username", value=st.session_state.username)
        email = st.text_input("Email", value="johndoe@example.com")
        if st.button("Update Profile"):
            st.session_state.username = username
            st.success("Profile updated successfully!")
        st.subheader("Appearance")
        theme = st.selectbox("Theme", ["Light", "Dark", "System Default"])
        chart_color = st.color_picker("Default Chart Color", "#2e7bcf")
        st.subheader("Notifications")
        email_notifications = st.checkbox("Receive Email Notifications", value=True)
        notification_frequency = st.select_slider("Notification Frequency", options=["Daily", "Weekly", "Monthly"])
        # Save Settings Button
        if st.button("Save Settings"):
            st.success("Settings saved successfully!")
        st.markdown('</div>', unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.subheader("📢 Feedback")
    feedback = st.sidebar.text_area("Share your thoughts:")
    if st.sidebar.button("Submit Feedback", key="submit_feedback"):
        # Here you would typically send this feedback to a database or email
        st.sidebar.success("Thank you for your valuable feedback!")
    # Display visit counter with animations
    total_visits, daily_visits = update_visit_count()
    st.sidebar.markdown(f"""
    <div class="visit-counter">
        <h3>📊 Visit Statistics</h3>
        <p>Total Visits: <span class="count">{total_visits}</span></p>
        <p>Visits Today: <span class="count">{daily_visits}</span></p>
    </div>
    <script>
        const countElements = document.querySelectorAll('.count');
        countElements.forEach(element => {{
            const target = parseInt(element.innerText);
            let count = 0;
            const timer = setInterval(() => {{
                element.innerText = count;
                if (count === target) {{
                    clearInterval(timer);
                }}
                count++;
            }}, 20);
        }});
    </script>
    """, unsafe_allow_html=True)
if __name__ == "__main__":
    main()
