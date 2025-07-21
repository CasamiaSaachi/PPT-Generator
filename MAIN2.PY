from flask import Flask, render_template, request, send_file, redirect, url_for, session, jsonify
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from datetime import datetime
from layout_manager import get_layout_positions

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

# --- RELATIVE PATHS ---
BASE = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE, "Excel Data", "Master Product List.xlsx")
FIXED_SLIDE_IMAGES = [os.path.join(BASE, "Photos", "2nd page.jpg"), os.path.join(BASE, "Photos", "3rd Page.jpg")]
LAST_SLIDE_IMAGE = [os.path.join(BASE, "Photos", "Last Page.jpg")]
IMAGE_FOLDER = os.path.join(BASE, "ERP Image Export 14-05-2025")
LOGO_PATH = os.path.join(BASE, "Photos", "logo.png")
OUTPUT_FOLDER = os.path.join(BASE, "Presentations")
TEMPLATE_INFO_PATH = os.path.join("static", "template-info.docx")
HEADING_BG_PATH = os.path.join(BASE, "Photos", "Heading Background.png")
TEMPLATES_PER_PRODUCT_COUNT = {1: 3, 2: 3, 3: 3, 4: 3, 5: 2, 6: 1, 7: 1}
COUNTER_FILE = os.path.join(OUTPUT_FOLDER, "ppt_counter.txt")
LOG_FILE = os.path.join(BASE, "download_logs.csv")


START_FROM = 300 # THUS IS BACK UP, WHEN TRY FAILS EXCEPT WILL TAKE FROM HERE


# --- HELPERS ---
# Below helps in reading the excel file
def load_excel_data():
    return pd.read_excel(EXCEL_PATH)
# check for the code in excel file
def get_product_data(df, codes):
    matched_dfs = [df[df["Code"] == code] for code in codes if code in df["Code"].values]
    if matched_dfs:
        return pd.concat(matched_dfs, ignore_index=True)
    else:
        return pd.DataFrame(columns=df.columns)

# chceck for what will be tje next counter eg: 301,302 etc
def get_next_counter(counter_file, start_from):
    if os.path.exists(counter_file):
        with open(counter_file, 'r') as f:
            try: num = int(f.read().strip())
            except: num = start_from
    else: num = start_from
    return num
# updates the counter file after the no is used
def increment_counter(counter_file, num):
    with open(counter_file, 'w') as f:
        f.write(str(num + 1))
# this will update teh log file whenever teh app is used
def log_download(filename, n_slides, n_products):
    with open(LOG_FILE, 'a') as f:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        f.write(f"{now},{filename},{n_slides},{n_products}\n")
# this will create the title slide, the date and the project aren't used but this needs to be there for the code to work
def add_title_slide(prs, main_title="PROPOSED SELECTIONS", subtitle="SANITARYWARE |", date="APRIL 15,2025", project="MAF - MOE", logo_path=None, file_name=""):
    blank_layout = next((l for l in prs.slide_layouts if l.name.lower() == "blank"), prs.slide_layouts[-1])
    slide = prs.slides.add_slide(blank_layout)
    SLIDE_WIDTH, SLIDE_HEIGHT = prs.slide_width, prs.slide_height
    # Use a blank slide for full control
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
    # Filename (below date)
    filename_box = slide.shapes.add_textbox(SLIDE_WIDTH - Inches(11), Inches(5.6), Inches(10), Inches(0.4))
    tf_file = filename_box.text_frame
    tf_file.clear()
    tf_file.word_wrap = True
    p_file = tf_file.paragraphs[0]
    p_file.text = file_name  # We'll pass this as a parameter
    p_file.font.size = Pt(19)
    p_file.font.name = 'Helvetica'
    p_file.font.color.rgb = RGBColor(225, 25, 25)
    p_file.alignment = PP_ALIGN.RIGHT

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
# this adds the fixed slides which are same in all ppts
def add_fixed_slide_from_image(prs, image_path):
    blank_layout = next((l for l in prs.slide_layouts if l.name.lower() == "blank"), prs.slide_layouts[-1])
    slide = prs.slides.add_slide(blank_layout)
    slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

EXCEL_DF = load_excel_data()
from layout_manager import get_layout_positions
# this add products in main slide
def add_product_slide(prs, df, codes_list, page_title, template_number, product_count):
    """
    NEW signature:

      prs             : Presentation object
      df              : DataFrame loaded from Excel (EXCEL_DF)
      codes_list      : list of strings, length == product_count
      page_title      : str
      template_number : int (1..N available for that count)
      product_count   : int (the exact number of slots the user asked for)
    """

    # 1) Use a truly blank layout
    blank_layout = None
    for layout in prs.slide_layouts:
        if layout.name.lower() == "blank":
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[-1]

    slide = prs.slides.add_slide(blank_layout)

    # 2) Draw the page_title at the top (exactly as before)
    slide.shapes.add_picture(
        HEADING_BG_PATH,
        Inches(9.835), Inches(0),
        Inches(7), Inches(1.25)
    )
    title_box = slide.shapes.add_textbox(
        Inches(9.835), Inches(0.25),
        Inches(7), Inches(1.25)
    )
    tf_title = title_box.text_frame
    tf_title.clear()
    p = tf_title.paragraphs[0]
    p.text       = page_title
    p.font.size  = Pt(36)
    p.font.name  = 'Poppins Light'
    p.font.bold  = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment  = PP_ALIGN.CENTER

    # 3) Always ask get_layout_positions for exactly product_count slots
    positions, image_size, text_height = get_layout_positions(product_count, template_number)

    # 4) Now loop over each index i in [0 .. product_count-1]:
    for i in range(product_count):
        code = None
        if i < len(codes_list):
            code = codes_list[i].strip()
        else:
            code = ""   # (should never happen if codes_list was built correctly)

        left, top = positions[i]

        # a) Look up Excel row, if any
        row = df[df['Code'] == code]
        desc = ""
        if not row.empty:
            desc = str(row.iloc[0].get('Description', "")).strip()

        # b) Try to add picture (if the file exists). If not, leave blank.
        img_filename = f"{code}.jpeg"
        img_path = os.path.join(IMAGE_FOLDER, img_filename)

        if code and os.path.exists(img_path):
            slide.shapes.add_picture(img_path, left, top,
                                     width=image_size, height=image_size)
        else:
            # ⚠️ image is missing OR code was blank ⇒ leave that slot empty
            print(f"⚠️ Image not found (or missing code) for slot #{i+1}: '{code}'")

        # c) Add a textbox under that slot, but only if we have at least the 'code'.
        #    If the code itself is blank, you may choose to skip drawing even the box.
        if code:
            textbox = slide.shapes.add_textbox(
                left,
                top + image_size + Inches(0.3),
                image_size + Inches(0.5),
                text_height
            )
            tf = textbox.text_frame
            tf.clear()
            tf.word_wrap = True

            # Put “CODE – Description” (desc may be empty if no match in Excel)
            p = tf.add_paragraph()
            run_code = p.add_run()
            run_code.text = f"{code} – "
            run_code.font.bold = True
            run_code.font.size = Pt(18)
            run_code.font.name = 'Poppins Light'
            run_code.font.color.rgb = RGBColor(0, 0, 0)

            run_desc = p.add_run()
            run_desc.text = desc
            run_desc.font.bold = False
            run_desc.font.size = Pt(18)
            run_desc.font.name = 'Poppins Light'
            run_desc.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.JUSTIFY

    # 5) Finally, draw the bottom-right logo exactly as before
    LOGO_WIDTH  = Inches(2.25)
    LOGO_HEIGHT = Inches(0.9)
    SLIDE_WIDTH = prs.slide_width
    SLIDE_HEIGHT= prs.slide_height
    margin_right  = Inches(0.35)
    margin_bottom = Inches(0.175)
    left_logo = SLIDE_WIDTH  - LOGO_WIDTH  - margin_right
    top_logo  = SLIDE_HEIGHT - LOGO_HEIGHT - margin_bottom

    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH,
                                 left_logo,
                                 top_logo,
                                 width=LOGO_WIDTH,
                                 height=LOGO_HEIGHT)
    else:
        print("⚠️ Logo image not found for product slide!")
# --- PRODUCT VALIDATION AJAX ENDPOINT ---

# checks for the product description and image is present or not
@app.route('/validate_products', methods=['POST'])
def validate_products():
    data = request.json
    codes = data.get('codes', [])
    result = []
    df = EXCEL_DF #USING CACHED MEMORY
    for code in codes:
        img_path = os.path.join(IMAGE_FOLDER, f"{code}.jpeg")
        desc = None
        row = df[df['Code'] == code]
        if not row.empty:
            desc = str(row.iloc[0].get('Description', '')).strip()
        has_img = os.path.exists(img_path)
        has_desc = bool(desc)
        result.append({
            'code': code,
            'has_image': has_img,
            'has_desc': has_desc,
        })
    return jsonify(result)

# --- FLASK ROUTES (same as before, but no logic change) ---

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session['meta'] = {
            'main_title': request.form.get('main_title', 'PROPOSED SELECTIONS'),
            'subtitle': request.form.get('subtitle', 'SANITARYWARE |'),
            'date': request.form.get('date') or datetime.today().strftime("%B %d, %Y"),
            'project': request.form.get('project', ''),
        }
        session['slides'] = []
        session['slide_index'] = 0
        return redirect(url_for('add_slide'))
    session.pop('slide_index', None)
        # 👇 make sure this clears everything
    session.pop('meta', None)
    session.pop('slides', None)
    session.pop('last_file_path', None)
    return render_template('meta.html')

@app.route('/add_slide', methods=['GET', 'POST'])
def add_slide():
    if 'slides' not in session:
        session['slides'] = []
    if 'slide_index' not in session:
        session['slide_index'] = 0
    slides = session['slides']
    slide_index = session['slide_index']
    if request.method == 'POST':
        product_count = int(request.form['product_count'])
        template_number = int(request.form['template'])
        page_title = request.form['page_title']
        codes = [request.form.get(f'code{i}').strip() for i in range(1, product_count + 1) if request.form.get(f'code{i}').strip()]
        slide_data = {'product_count': product_count, 'template': template_number, 'page_title': page_title, 'codes': codes}
        if slide_index < len(slides): slides[slide_index] = slide_data
        else: slides.append(slide_data)
        session['slides'] = slides
        if request.form.get('action') == 'add_more':
            session['slide_index'] = slide_index + 1
            return redirect(url_for('add_slide'))
        # elif request.form.get('action') == 'back':
        #     if slide_index == 0:
        #         session.pop('slide_index', None)
        #         return redirect(url_for('index'))
        #     else:
        #         session['slide_index'] = slide_index - 1
        #         return redirect(url_for('add_slide'))
        elif request.form.get('action') == 'back':
            if slide_index > 0:
                session['slide_index'] = slide_index - 1
            # Even if it's 0, we stay on the first slide, not reset the session
            return redirect(url_for('add_slide'))
        elif request.form.get('action') == 'finish' or request.form.get('action') == 'force_finish':
            session.pop('slide_index', None)
            return redirect(url_for('download'))
    templates_per_count = TEMPLATES_PER_PRODUCT_COUNT
    slide_data = slides[slide_index] if slide_index < len(slides) else None
    return render_template('slide.html', templates_per_count=templates_per_count, slide_data=slide_data, slide_index=slide_index, total_slides=len(slides), template_info_url=url_for('static', filename='template-info.docx'))

@app.route('/download')
def download():
    meta = session.get('meta')
    slides = session.get('slides', [])
    df = EXCEL_DF #USING CACHED MEMORY
    prs = Presentation()
    prs.slide_width = Inches(26.67)
    prs.slide_height = Inches(15)
    today = datetime.today()
    year_str = f"YR{today.strftime('%y')}"
    prefix = "PP"
    current_number = get_next_counter(COUNTER_FILE, START_FROM)
    file_number_str = f"{current_number:04d}"
    file_name_str = f"{prefix}-{year_str}-{file_number_str}"
    add_title_slide(prs, main_title=meta['main_title'], subtitle=meta['subtitle'], date=meta['date'], project=meta['project'], logo_path=LOGO_PATH, file_name=file_name_str)
    for img_path in FIXED_SLIDE_IMAGES:
        add_fixed_slide_from_image(prs, img_path)
    total_products = 0
    for s in slides:
        product_count   = s['product_count']
        template_number = s['template']
        codes_list      = s['codes']          # e.g. ['ERP100', 'ERP120', 'ERP200', 'ERP400']

        # Count how many valid products/images we actually find, if you still want to log that:
        actual_rows = get_product_data(df, codes_list)  # just to count for your "total_products"
        total_products += len(actual_rows)

        # → Pass product_count and codes_list into add_product_slide, instead of a filtered DataFrame
        add_product_slide(prs,df,codes_list,s['page_title'],template_number,product_count)

    for img_path in LAST_SLIDE_IMAGE:
        add_fixed_slide_from_image(prs, img_path)
    # today = datetime.today()
    # year_str = f"YR{today.strftime('%y')}"
    # prefix = "PP"
    # current_number = get_next_counter(COUNTER_FILE, START_FROM)
    # file_number_str = f"{current_number:04d}"
    output_path = os.path.join(OUTPUT_FOLDER, f"{prefix}-{year_str}-{file_number_str}.pptx")
    prs.save(output_path)
    increment_counter(COUNTER_FILE, current_number)
    log_download(os.path.basename(output_path), len(slides), total_products)
    # session.pop('meta', None)
    # session.pop('slides', None)
    session['last_file_path'] = output_path
    return send_file(output_path, as_attachment=True)

@app.errorhandler(Exception)
def handle_exception(e):
    import traceback
    tb = traceback.format_exc()
    return render_template('error.html', error=tb), 500

if __name__ == '__main__':
    # app.run(debug=True)
    app.run(host='0.0.0.0', port=5000, debug=True) 