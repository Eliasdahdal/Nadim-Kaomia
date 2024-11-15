import streamlit as st
import random
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from io import BytesIO
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import zipfile

# Function to convert numbers to Arabic numerals
def convert_to_arabic_numerals(number):
    arabic_numbers = str(number).translate(str.maketrans('0123456789', '٠١٢٣٤٥٦٧٨٩'))
    return arabic_numbers

# Function to set text direction to RTL and justify alignment for paragraphs
def set_rtl_and_justify(paragraph):
    paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # Justify alignment
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

# Function to set the entire document to RTL
def set_document_rtl(doc):
    for section in doc.sections:
        sectPr = section._sectPr
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        sectPr.append(bidi)

# Function to set the table to RTL
def set_table_rtl(table):
    tblPr = table._tblPr
    tblPr.append(parse_xml(r'<w:bidiVisual w:val="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT

# Function to set cell properties for RTL
def set_cell_rtl(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(r'<w:rtl w:val="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # Justify alignment
        pPr = paragraph._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)

# Function to create a Word document for a single model
def create_word_file(header, footer, questions, true_false_statements):
    doc = Document()
    set_document_rtl(doc)  # Set the entire document to RTL

    # Set default paragraph style to RTL and justify alignment
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    style.font.size = Pt(12)

    # Apply RTL to style
    pPr = style.element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

    # Add header
    if header:
        header_paragraph = doc.add_paragraph(header, style='Heading 1')
        set_rtl_and_justify(header_paragraph)

    # Add True/False section
    if true_false_statements:
        tf_intro = doc.add_paragraph("أولا: ضع كلمة صح أو غلط أمام العبارات التالية")
        set_rtl_and_justify(tf_intro)
        for statement in true_false_statements:
            tf_paragraph = doc.add_paragraph(f"{statement} (       )")
            set_rtl_and_justify(tf_paragraph)

    # Add a separator
    separator = doc.add_paragraph("-" * 80)
    set_rtl_and_justify(separator)

    # Add multiple-choice questions
    mcq_intro = doc.add_paragraph("ثانيا : اختر الإجابة الصحيحة مما يلي")
    set_rtl_and_justify(mcq_intro)

    for idx, q in enumerate(questions, 1):
        # Convert question number to Arabic numerals and add space after the number
        arabic_number = convert_to_arabic_numerals(idx)
        question_paragraph = doc.add_paragraph(f"{arabic_number}- {q['question']}")
        set_rtl_and_justify(question_paragraph)

        # Add options in table format
        table = doc.add_table(rows=2, cols=2)
        set_table_rtl(table)  # Set the table to RTL
        table.style = 'Table Grid'

        options_labels = ['أ', 'ب', 'ج', 'د']
        for i, option in enumerate(q['options']):
            row_index = i // 2
            col_index = i % 2
            cell = table.cell(row_index, col_index)
            set_cell_rtl(cell)  # Set cell to RTL
            cell.text = f"{options_labels[i]}- {option}"

        # Adjust columns for RTL by swapping cells in each row
        for row in table.rows:
            cells = row.cells
            cells[0].text, cells[1].text = cells[1].text, cells[0].text

        # Add space between questions
        spacer = doc.add_paragraph()
        set_rtl_and_justify(spacer)

    # Add footer
    if footer:
        footer_paragraph = doc.add_paragraph(footer, style='Heading 2')
        set_rtl_and_justify(footer_paragraph)

    # Save the file to BytesIO
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit app
st.title("إنشاء نماذج أسئلة")

# Header and footer inputs
header = st.text_area("أدخل العنوان الرئيسي (Header):", "أ.نديم_دحدل")
footer = st.text_area("أدخل التذييل (Footer):", "انتهت الأسئلة")

# True/False statements
st.subheader("إضافة عبارات صح وخطأ")
true_false_statements_input = st.text_area(
    "أدخل عبارات صح وخطأ (افصل كل عبارة بسطر جديد):",
    "ألا يعاقب المسؤول أمام مرؤوسيه من مهارات القيادة\n"
    "يقتضي القيام بعملية التحليل السياسي لتحديد كيفية التعامل مع الظاهرة\n"
    "المؤامرة في التحليل السياسي شرط لازم وكاف\n"
    "التحليل السياسي ترف فكري يمارسه البعض بقصد الدعاية"
)
true_false_statements = [s for s in true_false_statements_input.split('\n') if s.strip()]

# Questions and options
st.subheader("إضافة أسئلة اختيار من متعدد")
num_questions = st.number_input("عدد الأسئلة:", min_value=1, value=5, step=1)

questions = []
for i in range(num_questions):
    with st.expander(f"السؤال {i + 1}"):
        question = st.text_input(f"السؤال {i + 1}:", f"سؤال {i + 1}")
        options = []
        col1, col2 = st.columns(2)
        with col1:
            options.append(st.text_input(f"الخيار الأول للسؤال {i + 1}:", f"خيار 1"))
            options.append(st.text_input(f"الخيار الثالث للسؤال {i + 1}:", f"خيار 3"))
        with col2:
            options.append(st.text_input(f"الخيار الثاني للسؤال {i + 1}:", f"خيار 2"))
            options.append(st.text_input(f"الخيار الرابع للسؤال {i + 1}:", f"خيار 4"))
        questions.append({"question": question, "options": options})

# Number of models and file name
num_models = st.number_input("عدد النماذج:", min_value=1, value=3, step=1)
file_name = st.text_input("اسم الملف (بدون الامتداد):", "quiz_model")

# Upload Word file for headers and footers
uploaded_file = st.file_uploader("تحميل ملف Word لاستخدام العنوان والتذييل:", type="docx")
if uploaded_file:
    doc = Document(uploaded_file)
    header = doc.paragraphs[0].text if doc.paragraphs else header
    footer = doc.paragraphs[-1].text if doc.paragraphs else footer

# Generate models and create ZIP file
if st.button("إنشاء النماذج"):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for i in range(1, num_models + 1):
            shuffled_questions = random.sample(questions, len(questions))
            shuffled_statements = random.sample(true_false_statements, len(true_false_statements))
            buffer = create_word_file(
                header,
                footer,
                shuffled_questions,
                shuffled_statements
            )
            zip_file.writestr(f"{file_name}_{i}.docx", buffer.getvalue())

    zip_buffer.seek(0)
    st.download_button(
        label="تحميل جميع النماذج كملف مضغوط",
        data=zip_buffer,
        file_name=f"{file_name}_models.zip",
        mime="application/zip"
    )