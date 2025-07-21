from flask import Flask, render_template, request, send_file, redirect, url_for, session, after_this_request
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from datetime import datetime
from layout_manager import get_layout_positions

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Needed for session

# --- CONFIG ---
EXCEL_PATH = r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\Excel Data\Master Product List.xlsx"
FIXED_SLIDE_IMAGES = [
    r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\Photos\2nd page.jpg",
    r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\Photos\3rd Page.jpg"
]
LAST_SLIDE_IMAGE = [ r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\Photos\Last Page.jpg" ]
IMAGE_FOLDER = r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\ERP Image Export 14-05-2025"
LOGO_PATH = r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\Photos\logo.png"
OUTPUT_FOLDER = r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\TRALALA\SPec Presentation Generator\Presentations"

TEMPLATES_PER_PRODUCT_COUNT = {
    1: 4,
    2: 3,
    3: 3,
    4: 2,
    5: 1,
    6: 1,
    7: 1,
}

COUNTER_FILE = os.path.join(OUTPUT_FOLDER, "ppt_counter.txt")
START_FROM = 300

# --- HELPER FUNCTIONS (Title slide, product slides, etc.) ---
# Paste your add_title_slide, add_fixed_slide_from_image, add_product_slide functions here

def load_excel_data():
    return pd.read_excel(EXCEL_PATH)

def get_product_data(df, codes):
    return df[df["Code"].isin(codes)]

def get_next_counter(counter_file, start_from):
    if os.path.exists(counter_file):
        with open(counter_file, 'r') as f:
            try:
                num = int(f.read().strip())
            except:
                num = start_from
    else:
        num = start_from
    return num

def increment_counter(counter_file, num):
    with open(counter_file, 'w') as f:
        f.write(str(num + 1))

def add_title_slide(
    prs, 
    main_title="PROPOSED SELECTIONS", 
    subtitle="SANITARYWARE |", 
    date=" APRIL 15,2025",
    project="MAF - MOE", 
    logo_path=None
):
    # Use a blank slide for full control
    blank_layout = None
    for layout in prs.slide_layouts:
        if layout.name.lower() == "blank":
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank_layout)
    SLIDE_WIDTH = prs.slide_width
    SLIDE_HEIGHT = prs.slide_height
    # Main Title
    title_box = slide.shapes.add_textbox(SLIDE_WIDTH - Inches(11), Inches(0.8), Inches(10), Inches(3.5))
    tf = title_box.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = main_title
    p.font.size = Pt(100)
    p.font.name = 'Belleza'
    p.font.color.rgb = RGBColor(225, 25, 25)
    p.alignment = PP_ALIGN.RIGHT
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(SLIDE_WIDTH - Inches(11), Inches(4.5), Inches(10), Inches(0.4))
    tf2 = subtitle_box.text_frame
    tf2.clear()
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(30)
    p2.font.name = 'Poppins Light'
    p2.font.color.rgb = RGBColor(225, 25, 25)
    p2.alignment = PP_ALIGN.RIGHT
    # Date
    date_box = slide.shapes.add_textbox(SLIDE_WIDTH - Inches(11), Inches(5), Inches(10), Inches(0.5))
    tf2 = date_box.text_frame
    tf2.clear()
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = date
    p2.font.size = Pt(30)
    p2.font.name = 'Poppins Light'
    p2.font.color.rgb = RGBColor(225, 25, 25)
    p2.alignment = PP_ALIGN.RIGHT
    # Project
    proj_box = slide.shapes.add_textbox(SLIDE_WIDTH - Inches(7), Inches(8.5), Inches(6), Inches(1))
    tf3 = proj_box.text_frame
    tf3.clear()
    p3 = tf3.paragraphs[0]
    p3.text = project
    p3.font.size = Pt(50)
    p3.font.name = 'Belleza'
    p3.font.color.rgb = RGBColor(237, 137, 39)
    p3.alignment = PP_ALIGN.RIGHT
    # Logo
    LOGO_WIDTH = Inches(5)
    LOGO_HEIGHT = Inches(2)
    margin_right = Inches(0.3)
    margin_bottom = Inches(0.1)
    left = SLIDE_WIDTH - LOGO_WIDTH - margin_right
    top = SLIDE_HEIGHT - LOGO_HEIGHT - margin_bottom
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, left, top, width=LOGO_WIDTH, height=LOGO_HEIGHT)
    else:
        print("⚠️ Logo image not found for title slide!")

def add_fixed_slide_from_image(prs, image_path):
    blank_layout = None
    for layout in prs.slide_layouts:
        if layout.name.lower() == "blank":
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank_layout)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    slide.shapes.add_picture(image_path, 0, 0, width=slide_width, height=slide_height)

from layout_manager import get_layout_positions

def add_product_slide(prs, products, title, template_number):
    # Use a safe blank layout
    blank_layout = None
    for layout in prs.slide_layouts:
        if layout.name.lower() == "blank":
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank_layout)

    # Add page title manually at the top
    # title_box = slide.shapes.add_textbox(Inches(3.335), Inches(0.3), Inches(20), Inches(1)) #LEFT TOP WIDDTH HEIGHT
    slide.shapes.add_picture( r"C:\Users\Divyansh\OneDrive - Casa Mia LLC\Desktop\SPec Presentation Generator\Photos\Heading Background.png", Inches(9.835), Inches(0), Inches(7), Inches(1.25))
    title_box = slide.shapes.add_textbox(Inches(9.835), Inches(0.25), Inches(7), Inches(1.25))

    tf_title = title_box.text_frame
    tf_title.clear()
    p = tf_title.paragraphs[0]
    p.text = title
    p.font.size = Pt(36)
    p.font.name = 'Poppins Light'
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

    product_count = len(products)
    positions, image_size, text_height = get_layout_positions(product_count, template_number)

    for i, (_, product) in enumerate(products.iterrows()):
        if i >= len(positions):
            break

        left, top = positions[i]
        img_filename = f"{str(product['Code'])}.jpeg"  # or .png if you use png
        img_path = os.path.join(IMAGE_FOLDER, img_filename)
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path, left, top, width=image_size, height=image_size)
        else:
            print(f"⚠️ Image not found for ERP Code: {product['Code']}")

        # Add textbox under image
        textbox = slide.shapes.add_textbox(left, top + image_size + Inches(0.3), image_size + Inches(0.5), text_height)
        tf = textbox.text_frame
        tf.clear()
        tf.word_wrap = True
        # text = f"{product['Code']} - {product['Description']}\n"
        # for line in text.strip().split('\n'):
        #     p = tf.add_paragraph()
        #     p.text = line
        #     p.font.size = Pt(18)
        #     p.font.name = 'Poppins Light'
        #     p.font.color.rgb = RGBColor(0, 0, 0)
        #     p.alignment = PP_ALIGN.JUSTIFY
        
        code = product['Code']
        description = product['Description']

        # First line: Code (bold) + Description (normal)
        p = tf.add_paragraph()
        run_code = p.add_run()
        run_code.text = f"{code} - "
        run_code.font.bold = True  # Make the code bold
        run_code.font.size = Pt(18)
        run_code.font.name = 'Poppins Light'
        run_code.font.color.rgb = RGBColor(0, 0, 0)

        run_desc = p.add_run()
        run_desc.text = description
        run_desc.font.bold = False
        run_desc.font.size = Pt(18)
        run_desc.font.name = 'Poppins Light'
        run_desc.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.JUSTIFY

    # --- Add logo at bottom right ---
    LOGO_WIDTH = Inches(2.25)
    LOGO_HEIGHT = Inches(0.9)
    SLIDE_WIDTH = prs.slide_width
    SLIDE_HEIGHT = prs.slide_height
    margin_right = Inches(0.35)
    margin_bottom = Inches(0.175)
    left = SLIDE_WIDTH - LOGO_WIDTH - margin_right
    top = SLIDE_HEIGHT - LOGO_HEIGHT - margin_bottom
    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, left, top, width=LOGO_WIDTH, height=LOGO_HEIGHT)
    else:
        print("⚠️ Logo image not found for product slide!")

# def add_product_slide(prs, products, title, template_number):
#     # slide = prs.slides.add_slide(prs.slide_layouts[5])
#     blank_layout = prs.slide_layouts[14]  # This is usually the blank layout
#     slide = prs.slides.add_slide(blank_layout)
#     # title_shape = slide.shapes.title
#     # title_shape.text = title
    
#     # Add page title manually at the top
#     # title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
#     # Centered textbox: width = 20", left = (26.67 - 20) / 2 = 3.335
#     title_box = slide.shapes.add_textbox(Inches(3.335), Inches(0.3), Inches(20), Inches(1))
#     tf_title = title_box.text_frame
#     tf_title.clear()
#     p = tf_title.paragraphs[0]
#     p.text = title
#     p.font.size = Pt(48)
#     p.font.name = 'Poppins Light'  # Change this to any valid font name installed on your system
#     p.font.bold = True
#     p.font.color.rgb = RGBColor(0, 0, 0)  # Black
#     p.alignment = PP_ALIGN.CENTER



#     product_count = len(products)
#     positions, image_size, text_height = get_layout_positions(product_count, template_number)

#     for i, (_, product) in enumerate(products.iterrows()):
#         if i >= len(positions):
#             break

#         left, top = positions[i]
#         # img_path = os.path.join(IMAGE_FOLDER, str(product["Image Path"]))
#         img_filename = f"{str(product['Code'])}.jpeg"  # or .png, based on your image type
#         img_path = os.path.join(IMAGE_FOLDER, img_filename)


#         # if os.path.exists(img_path):
#         #     slide.shapes.add_picture(img_path, left, top, width=image_size, height=image_size)
        
#         if os.path.exists(img_path):
#             slide.shapes.add_picture(img_path, left, top, width=image_size, height=image_size)
#         else:
#             print(f"⚠️ Image not found for ERP Code: {product['Code']}")


#         # Add textbox under image
#         textbox = slide.shapes.add_textbox(left, top + image_size + Inches(0.3), image_size + Inches(0.5), text_height)
#         tf = textbox.text_frame
#         tf.clear()  # Clear default paragraph
#         tf.word_wrap = True 
#         # --- Basic product info ---
#         text = f"{product['Code']} - {product['Description']}\n"

#         # {product['Product Name']}\n
#         # --- Optional fields (e.g., Color, Material) if they exist ---
#         # Add these fields in Excel, and they will automatically appear here
        
#         # for field in ["Colour", "Category", "Thickness", "Finish", "Collection"]:
#         #     if field in product and pd.notna(product[field]):
#         #         text += f"\n{field}: {product[field]}"

#         # tf.text = text
        
#         for line in text.strip().split('\n'):
#             p = tf.add_paragraph()
#             p.text = line
#             p.font.size = Pt(18)
#             p.font.name = 'Poppins Light'  # Change this to any valid font name installed on your system
#             p.font.color.rgb = RGBColor(0, 0, 0)  # Black
#             p.alignment = PP_ALIGN.JUSTIFY  # <-- This is the justification setting


    # slide = prs.slides.add_slide(prs.slide_layouts[5])
    # title_shape = slide.shapes.title
    # title_shape.text = title


    # for i, (_, product) in enumerate(products.iterrows()):
    #     pic_path = os.path.join(IMAGE_FOLDER, str(product["Image Path"]))
    #     product_box_left = left + Inches(3) * (i % 3)
    #     product_box_top = top + Inches(3.5) * (i // 3)

    #     if os.path.exists(pic_path):
    #         slide.shapes.add_picture(pic_path, product_box_left, product_box_top, width, height)

    #     textbox = slide.shapes.add_textbox(product_box_left, product_box_top + Inches(3.1), width, Inches(1))
    #     tf = textbox.text_frame
    #     tf.text = f"{product['ERP ID']} - {product['Product Name']}\n{product['Description']}\nPrice: {product['Price']}"

# --- FLASK ROUTES ---
@app.route('/', methods=['GET', 'POST'])
def index():
    # Step 1: Get presentation meta info
    if request.method == 'POST':
        session['meta'] = {
            'main_title': request.form.get('main_title', 'PROPOSED SELECTIONS'),
            'subtitle': request.form.get('subtitle', 'SANITARYWARE |'),
            'date': request.form.get('date') or datetime.today().strftime("%B %d, %Y"),
            'project': request.form.get('project', ''),
        }
        session['slides'] = []
        return redirect(url_for('add_slide'))
    return render_template('meta.html')

@app.route('/add_slide', methods=['GET', 'POST'])
def add_slide():
    # Track which slide we are editing
    if 'slides' not in session:
        session['slides'] = []
    if 'slide_index' not in session:
        session['slide_index'] = 0

    slides = session['slides']
    slide_index = session['slide_index']

    # Handle form submission
    if request.method == 'POST':
        product_count = int(request.form['product_count'])
        template_number = int(request.form['template'])
        page_title = request.form['page_title']
        codes = [request.form.get(f'code{i}').strip() for i in range(1, product_count + 1) if request.form.get(f'code{i}').strip()]

        # If we're editing an existing slide, update it; else append new
        slide_data = {
            'product_count': product_count,
            'template': template_number,
            'page_title': page_title,
            'codes': codes,
        }
        if slide_index < len(slides):
            slides[slide_index] = slide_data
        else:
            slides.append(slide_data)
        session['slides'] = slides  # Save

        # Navigation logic
        if request.form.get('action') == 'add_more':
            session['slide_index'] = slide_index + 1
            return redirect(url_for('add_slide'))
        elif request.form.get('action') == 'back':
            if slide_index == 0:
                # Go back to meta page
                session.pop('slide_index', None)
                return redirect(url_for('index'))
            else:
                session['slide_index'] = slide_index - 1
                return redirect(url_for('add_slide'))
        elif request.form.get('action') == 'finish':
            # Reset index before download
            session.pop('slide_index', None)
            return redirect(url_for('download'))

    # GET request
    # If we're editing, pre-fill the form with existing slide data
    templates_per_count = TEMPLATES_PER_PRODUCT_COUNT
    slide_data = slides[slide_index] if slide_index < len(slides) else None
    return render_template('slide.html',
                           templates_per_count=templates_per_count,
                           slide_data=slide_data,
                           slide_index=slide_index,
                           total_slides=len(slides))


@app.route('/download')
def download():
    # Step 3: Generate PPT and send to user
    meta = session.get('meta')
    slides = session.get('slides', [])
    df = load_excel_data()
    prs = Presentation()
    prs.slide_width = Inches(26.67)
    prs.slide_height = Inches(15)

    # Title slide
    add_title_slide(
        prs,
        main_title=meta['main_title'],
        subtitle=meta['subtitle'],
        date=meta['date'],
        project=meta['project'],
        logo_path=LOGO_PATH
    )
    # Fixed slides
    for img_path in FIXED_SLIDE_IMAGES:
        add_fixed_slide_from_image(prs, img_path)
    # Product slides
    for s in slides:
        selected_products = get_product_data(df, s['codes'])
        if selected_products.empty:
            continue  # skip
        add_product_slide(prs, selected_products, s['page_title'], s['template'])
    # Last slide
    for img_path in LAST_SLIDE_IMAGE:
        add_fixed_slide_from_image(prs, img_path)

    # Output path
    today = datetime.today()
    year_str = f"YR{today.strftime('%y')}"
    prefix = "PP"
    current_number = get_next_counter(COUNTER_FILE, START_FROM)
    file_number_str = f"{current_number:04d}"
    output_path = os.path.join(
        OUTPUT_FOLDER,
        f"{prefix}-{year_str}-{file_number_str}.pptx"
    )
    prs.save(output_path)
    increment_counter(COUNTER_FILE, current_number)

    # Clean up session for new run
    session.pop('meta', None)
    session.pop('slides', None)
    session.pop('slide_index', None)    
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
