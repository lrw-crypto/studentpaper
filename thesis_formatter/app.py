import os
import re
from flask import Flask, render_template, request, send_file, after_this_request
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from io import BytesIO

app = Flask(__name__)


# --- 核心排版功能區 ---

def add_page_number(run):
    """
    在 python-docx 中插入頁碼需要操作底層 XML
    """
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = " PAGE "

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)


def set_chinese_font(run, size_pt=12):
    """
    強制設定中英文雙字型：
    中文：新細明體 (PMingLiU)
    英文：Times New Roman
    """
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '新細明體')
    run.font.size = Pt(size_pt)
    run.font.color.rgb = None  # 確保是黑色(預設)


def format_thesis(doc_stream, paper_title):
    doc = Document(doc_stream)

    # 1. 版面設定 (Page Layout)
    # 規定：A4, 上下左右邊界皆為 2公分
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    # 2. 設定預設樣式 (Normal Style)
    style = doc.styles['Normal']
    style.paragraph_format.line_spacing = 1.0  # 單行間距
    set_chinese_font(style, 12)  # 預設全域字體設定

    # 3. 頁首與頁尾 (Header & Footer)
    # 頁首：小論文篇名 (置中, 10pt)
    header = section.header
    header.is_linked_to_previous = False
    # 清除預設段落並建立新段落
    if header.paragraphs:
        header_para = header.paragraphs[0]
        header_para.clear()
    else:
        header_para = header.add_paragraph()

    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_run = header_para.add_run(paper_title)
    set_chinese_font(h_run, 10)  # 標題 10級字

    # 頁尾：頁碼 (置中, 10pt)
    footer = section.footer
    footer.is_linked_to_previous = False
    if footer.paragraphs:
        footer_para = footer.paragraphs[0]
        footer_para.clear()
    else:
        footer_para = footer.add_paragraph()

    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f_run = footer_para.add_run()
    add_page_number(f_run)  # 插入動態頁碼
    set_chinese_font(f_run, 10)

    # 4. 內容遍歷與智慧格式化
    is_reference_section = False

    # 用於偵測標題的 Regex
    # 層級一：壹、 (Chinese Number + comma)
    regex_h1 = re.compile(r'^[壹貳參肆伍陸]、')
    # 層級二：一、 (Chinese One + comma)
    regex_h2 = re.compile(r'^[一二三四五六七八九十]+、')
    # 層級三：(一) (Parenthesis + Chinese One)
    regex_h3 = re.compile(r'^\（[一二三四五六七八九十]+\）')
    # 層級四：１、 (Fullwidth digit + comma)
    regex_h4 = re.compile(r'^[０-９]+、')

    for para in doc.paragraphs:
        text = para.text.strip()

        # 跳過空白段落，但保留格式
        if not text:
            continue

        # 4.1 偵測是否進入「參考文獻」區域
        # 進入此區後，後續段落都需要懸掛縮排 (Hanging Indent)
        if "參考文獻" in text and (text.startswith("陸") or text.startswith("六") or len(text) < 10):
            is_reference_section = True
            # 確保標題本身格式正確
            para.paragraph_format.first_line_indent = Pt(0)
            para.paragraph_format.left_indent = Pt(0)
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # 強制重設字體
            for run in para.runs:
                set_chinese_font(run, 12)
                run.bold = False  # 規定標題不粗體
            continue

        # 4.2 處理「參考文獻」內容的 APA 格式 (懸掛縮排)
        if is_reference_section:
            # 懸掛縮排：第一行靠左，第二行開始縮排
            # 實作：左縮排 2個字元 (24pt)，第一行縮排 -2個字元 (-24pt)
            para.paragraph_format.left_indent = Pt(24)
            para.paragraph_format.first_line_indent = Pt(-24)
            para.paragraph_format.line_spacing = 1.0

            # 內容字體重設
            for run in para.runs:
                # 參考文獻中的書名/期刊名依照 APA 標準可能需要斜體
                # 這裡保留使用者原本的斜體設定，但強制設定字型與大小
                is_italic = run.italic
                set_chinese_font(run, 12)
                run.italic = is_italic
                run.bold = False  # 參考文獻內文通常不粗體
            continue

        # 4.3 一般內文標題層級處理
        # 預設段落縮排 (2個字元)
        para.paragraph_format.first_line_indent = Pt(24)
        para.paragraph_format.left_indent = Pt(0)

        # 檢查是否為標題，如果是，取消首行縮排(或依習慣調整)，並確保單獨成行
        is_heading = False

        if regex_h1.match(text):  # 壹、前言
            is_heading = True
            para.paragraph_format.first_line_indent = Pt(0)  # 大標靠左
        elif regex_h2.match(text):  # 一、研究動機
            is_heading = True
            para.paragraph_format.first_line_indent = Pt(0)  # 次標靠左
        elif regex_h3.match(text):  # (一)
            is_heading = True
            para.paragraph_format.first_line_indent = Pt(24)  # 第三層通常縮排
        elif regex_h4.match(text):  # １、
            is_heading = True
            para.paragraph_format.first_line_indent = Pt(24)

        # 統一處理字體 (移除粗體、斜體，除了特定需求)
        for run in para.runs:
            set_chinese_font(run, 12)
            # 根據比賽規定：字體限黑色，不可使用粗體、斜體、底線
            run.bold = False
            run.italic = False
            run.underline = False

    # 存入緩衝區
    output_stream = BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream


# --- 路由區 ---

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return '沒有上傳檔案', 400

    file = request.files['file']
    title = request.form.get('title', '小論文')

    if file.filename == '':
        return '未選擇檔案', 400

    if file and file.filename.endswith('.docx'):
        # 執行排版
        formatted_file = format_thesis(file, title)

        return send_file(
            formatted_file,
            as_attachment=True,
            download_name=f'已排版_{file.filename}',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        return '請上傳 .docx 格式的 Word 檔案', 400


if __name__ == '__main__':
    # 確保 template 資料夾存在
    if not os.path.exists('templates'):
        os.makedirs('templates')
    app.run(debug=True, port=5000)