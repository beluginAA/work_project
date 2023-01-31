#!/usr/bin/python
# -*- coding: cp1251 -*-
import docx
import docx2txt
import pypandoc
import os
import time
import numpy as np
from PIL import Image
from pathlib import Path
from docx.oxml import OxmlElement, ns
from docx2python import docx2python
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT

start_time = time.time()

def new_document(pt = 14, line_spacing = 1.15, font_name = 'Times New Roman', left_margin = 30, right_margin = 15, top_margin = 20, bottom_margin = 20):

    def adding_margins(name): 
        text, foot_flag, num_footnotes, header_text, table_word = [], False, [], [], []
        document = docx2python(Path("Documents") / str(name))
        for obj in document.body:
            if len(obj) > 1 and len(obj[0][0]) == 1:
                numbers = [[obj[_][number][0] for number in range(len(obj[0]))] for _ in range(len(obj))]
                text.append(numbers)
                for main in numbers:
                    for letter in main:
                        table_word.append(letter)
                header_text.append(numbers)
                num_footnotes.append(0)
            elif len(obj) == 1 and len(obj[0][0]) >= 1:
                for line in range(len(obj[0][0])):
                    if obj[0][0][line] != '':
                        if obj[0][0][line].find('.png') != -1:
                            text.append(obj[0][0][line][obj[0][0][line].find('media') + 6 :][: obj[0][0][line][obj[0][0][line].find('media') + 6 :].find('----')])
                            header_text.append(obj[0][0][line][obj[0][0][line].find('media') + 6 :][: obj[0][0][line][obj[0][0][line].find('media') + 6 :].find('----')])
                            num_footnotes.append(0)
                            image_flag = True
                        elif obj[0][0][line].find('footnote') != -1:
                            phrase = list(obj[0][0][line][: obj[0][0][line].find('----')])
                            while True:
                                if phrase[0].isalpha() : break
                                if phrase[0] == ' ': del phrase[0]
                            for _ in range(12): phrase = np.insert(phrase, 0, ' ')
                            phrase = ''.join(phrase)
                            text.append(phrase)
                            header_text.append(phrase)
                            num_footnotes.append(1)
                            foot_flag = True
                        else:
                            header_text.append(obj[0][0][line])
                            phrase = list(obj[0][0][line])
                            if obj[0][0][line].find('\t') != -1: 
                                text.append(obj[0][0][line][obj[0][0][line].find('--\t') + 1 :])
                                header_text.append(obj[0][0][line][obj[0][0][line].find('--\t') + 3 :])
                                num_footnotes.append(0)
                                continue
                            if len(phrase) == 0: 
                                text.append(obj[0][0][line])
                                header_text.append(obj[0][0][line])
                                num_footnotes.append(0)
                                continue
                            if len(obj[0][0][line].split()) <= 10: 
                                text.append(obj[0][0][line].capitalize())
                                header_text.append(obj[0][0][line].capitalize())
                                num_footnotes.append(0)
                                continue
                            while True:
                                if phrase[0].isalpha() : break
                                if phrase[0] == ' ': del phrase[0]
                            for _ in range(12): phrase = np.insert(phrase, 0, ' ')
                            phrase = ''.join(phrase)
                            text.append(phrase)
                            num_footnotes.append(0)
        return text, num_footnotes, foot_flag, header_text, table_word

    def find_words(words):
        if len(words) > 0:
            for word in words:
                if word.lower() == 'из':
                    pass

    def adding_list(name):
        if text[0][name].find('-\t') != -1: 
            find_words(text[0][name][text[0][name].find('\t') + 1 :].split())
            paragrahp = new_doc.add_paragraph(text[0][name][text[0][name].find('\t') + 1 :], style = 'List Bullet')
        else: paragrahp = new_doc.add_paragraph(text[0][name][text[name].find('\t') + 1 :], style = 'List Number')
        p_fmt = paragrahp.paragraph_format
        p_fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_fmt.line_spacing = line_spacing
        p_fmt.space_before = Pt(line_spacing)
        p_fmt.space_after = Pt(line_spacing)
        return paragrahp

    def adding_heading(name):
        if text[0][name] == '' and headings[name - 1] == 0 and name + 1 <= len(headings): 
            try: headings[name + 1] == 0
            except IndexError: return 
            else: return
        if headings[name] and text[0][name - 1] != '' and name != 0:
                paragrahp = new_doc.add_paragraph('')
        paragrahp = new_doc.add_paragraph('')
        find_words(text[0][name].split())
        run = paragrahp.add_run(text[0][name])
        run.bold = True
        run.font.size = Pt(pt+2)
        p_fmt = paragrahp.paragraph_format
        p_fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fmt.line_spacing = 1.5
        p_fmt.space_before = Pt(1.5)
        p_fmt.space_after = Pt(1.5)

    def adding_paragraph(name):
        if name != 0:
            if tables[name + 1] == 1:
                new_doc.add_paragraph(' ')
                find_words(text[0][name].split())
                paragrahp = new_doc.add_paragraph('')
                run = paragrahp.add_run(text[0][name])
                run.italic = True
                run.font.size = Pt(pt - 2)
                p_fmt = paragrahp.paragraph_format
                p_fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif pictures[name - 1] == 1 and text[0][name] != ' ':
                find_words(text[0][name].split())
                paragrahp = new_doc.add_paragraph('')
                run = paragrahp.add_run(text[0][name])
                run.italic = True
                run.font.size = Pt(pt - 2)
                p_fmt = paragrahp.paragraph_format
                p_fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                find_words(text[0][name].split())
                paragrahp = new_doc.add_paragraph(text[0][name])
                p_fmt = paragrahp.paragraph_format
                p_fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_fmt.line_spacing = line_spacing
            p_fmt.space_before = Pt(line_spacing)
            p_fmt.space_after = Pt(line_spacing)
        return paragrahp

    def adding_table(txt, name, pt):
        doc = docx.Document(Path("Documents") / str(filename))
        unique, merged, max_row, max_col, added, cells_text, all_cells_text = [], [], np.array([]), np.array([]), [], [], []
        for row in doc.tables[table_num].rows:
            info = []
            for cell in row.cells:
                tc = cell._tc
                max_row, max_col = np.append(tc.bottom, max_row), np.append(tc.right, max_col)
                cell_loc = (tc.top, tc.bottom, tc.left, tc.right)
                if tc.bottom - tc.top > 1:
                    if cell_loc not in merged: 
                        info.append(cell.text)
                        all_cells_text.append(cell.text)
                    else: 
                        info.append(' ')
                        all_cells_text.append(' ')
                elif tc.right - tc.left > 1:
                    if cell_loc not in merged: 
                        info.append(cell.text)
                        all_cells_text.append(cell.text)
                    else: 
                        info.append(' ')
                        all_cells_text.append(' ')
                else: 
                    info.append(cell.text)
                    all_cells_text.append(cell.text)
                if tc.bottom - tc.top > 1 or tc.right - tc.left > 1 and cell_loc not in merged: merged.append(cell_loc)
                else: unique.append(cell_loc)
            cells_text.append(info)
        table = new_doc.add_table(rows = 0, cols = int(np.amax(max_col)), style='Table Grid')
        for row in range(int(np.amax(max_row))): 
            table.add_row().cells  
            for col in range(int(np.amax(max_col))):
                cell_paragraphs = [paragraph for paragraph in table.cell(row ,col).paragraphs]
                for paragraph in cell_paragraphs:
                    p = paragraph._element
                    p.getparent().remove(p)
                    paragraph._p = paragraph._element = None
                table.cell(row ,col).vertical_alignment = WD_ALIGN_VERTICAL.CENTER                     
                if row == 0: 
                    p = table.cell(row, col).add_paragraph('')
                    if cells_text[row][col] != ' ':
                        find_words(cells_text[row][col])
                        run = p.add_run(cells_text[row][col])
                        if len(txt[name][0]) == 1:
                            run.font.size = Pt(pt)
                            p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        else:
                            run.bold = True
                            run.font.size = Pt(pt+2)
                            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                else:
                    if cells_text[row][col] != '':
                        find_words(cells_text[row][col])
                        p = table.cell(row, col).add_paragraph(cells_text[row][col])
                        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        if tables[name + 1] == 1: new_doc.add_paragraph(' ')
        for row in new_doc.tables[table_num].rows:
            for cell in row.cells:
                tc = cell._tc
                loc = (tc.top, tc.bottom, tc.left, tc.right)
                if loc not in unique: added.append(loc)
        for merge in merged:
            for cell_1 in range(len(added) - 1):
                for cell_2 in range(cell_1 + 1, len(added)):
                    if (added[cell_1][0] == merge[0] and added[cell_1][1] == merge[1] and added[cell_1][2] == merge[2] and added[cell_2][3] == merge[3] and added[cell_2][0] == merge[0] and added[cell_2][1] == merge[1]) or (
                        added[cell_2][0] == merge[0] and added[cell_2][1] == merge[1] and added[cell_2][2] == merge[2] and added[cell_1][3] == merge[3] and added[cell_1][0] == merge[0] and added[cell_1][1] == merge[1]) or (
                        added[cell_1][1] == merge[1] and added[cell_1][2] == merge[2] and added[cell_1][3] == merge[3] and added[cell_2][0] == merge[0] and added[cell_2][2] == merge[2] and added[cell_2][3] == merge[3]) or (
                        added[cell_2][1] == merge[1] and added[cell_2][2] == merge[2] and added[cell_2][3] == merge[3] and added[cell_1][0] == merge[0] and added[cell_1][2] == merge[2] and added[cell_1][3] == merge[3]):
                        table.cell(min([added[cell_1][0], added[cell_1][1]]), min([added[cell_1][2], added[cell_1][3]])).merge(table.cell(min([added[cell_2][0], added[cell_2][1]]), min([added[cell_2][2], added[cell_2][3]])))               
        for row in range(int(np.amax(max_row))):  
            for col in range(int(np.amax(max_col))):
                cell_paragraphs = [paragraph for paragraph in table.cell(row ,col).paragraphs]
                for paragraph in cell_paragraphs:
                    if ' ' in paragraph.text: 
                        p = paragraph._element
                        p.getparent().remove(p)
                        paragraph._p = paragraph._element = None
        return all_cells_text

    def adding_picture(name):
        with Image.open(Path("Documents") / str(text[0][name])) as img:
            width, height = img.size
            img.save(Path("Documents") / str(text[0][name]) , dpi=(800, 800), optimize=True, quality=100)
        if width > 700:
            new_doc.add_picture(os.path.join("Documents", str(text[0][name])), width = Mm(width/4.25), height = Mm(height/4.25))
        else:
            new_doc.add_picture(os.path.join("Documents", str(text[0][name])), width = Mm(width/3.5), height = Mm(height/3.5))
        new_doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        os.remove(Path("Documents") / str(text[0][name]))

    def adding_footnotes(name):
        with open(os.path.join("Documents", str(name[:name.find('.docx')]) + '.txt' ), 'r', encoding = 'utf8') as file:
            f = file.readlines()
            footnotes, summa = {}, 0
            for word in range(len(f)):
                if f[word].find('footnote') != -1:
                    summa += 1
                    footnotes[summa] = f[word][f[word].find('footnote') + 10 :][: f[word][f[word].find('footnote') + 10 :].find(']')]
                    find_words(f[word][f[word].find('footnote') + 10 :][: f[word][f[word].find('footnote') + 10 :].find(']')].split())
        os.remove(os.path.join("Documents", str(name[:name.find('.docx')]) + '.txt' ))
        return footnotes
    
    def adding_headers_and_footers(up, down):
        if len(up) > 0:
            header = new_doc.sections[0].header
            header_para = header.add_paragraph('')
            run = header_para.add_run(str(up[0]))
            run.italic = True
            run.font.size = Pt(pt)
            header_para.alignment=WD_ALIGN_PARAGRAPH.CENTER
        if len(down) > 0:
            footer = new_doc.sections[0].footer
            footer_para = footer.paragraphs[0]
            footer_para.text = str(down[0]) + '\n\n'
            footer_para.runs[0].italic = True
            footer_para.runs[0].font.size = Pt(pt)
    
    def add_page_number(paragraph):

        def create_attribute(element, name, value):
            element.set(ns.qn(name), value)
        
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        page_num_run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        create_attribute(fldChar1, 'w:fldCharType', 'begin')
        instrText = OxmlElement('w:instrText')
        create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        create_attribute(fldChar2, 'w:fldCharType', 'end')
        page_num_run._r.append(fldChar1)
        page_num_run._r.append(instrText)
        page_num_run._r.append(fldChar2)

    text = adding_margins(str(filename))
    if text[2]: 
        docxFilename = os.path.join("Documents", str(filename) ) 
        pypandoc.convert_file(docxFilename, to = 'asciidoc', outputfile = os.path.join("Documents", str(filename[:filename.find('.docx')]) + '.txt'))
        # Codecs: asciidoc, asciidoctor, beamer, biblatex, bibtex, commonmark, commonmark_x, context, csljson, docbook, 
        # docbook4, docbook5, docx, dokuwiki, dzslides, epub, epub2, epub3, fb2, gfm, haddock, html, html4, html5, icml, ipynb, 
        # jats, jats_archiving, jats_articleauthoring, jats_publishing, jira, json, latex, man, markdown, markdown_github, 
        # markdown_mmd, markdown_phpextra, markdown_strict, markua, mediawiki, ms, muse, native, odt, opendocument, opml, 
        # org, pdf, plain, pptx, revealjs, rst, rtf, s5, slideous, slidy, tei, texinfo, textile, xwiki, zimwiki
        footnotes = adding_footnotes(filename)
    my_doc = docx2txt.process(Path("Documents") / str(filename),directory).split('\n')
    new_doc = docx.Document()
    new_doc.sections[0].orientation = WD_ORIENT.PORTRAIT
    new_doc.sections[0].left_margin = Mm(left_margin)
    new_doc.sections[0].right_margin = Mm(right_margin)
    new_doc.sections[0].top_margin = Mm(top_margin)
    new_doc.sections[0].bottom_margin = Mm(bottom_margin)
    new_doc.styles['Normal'].font.name = font_name
    new_doc.styles['Normal'].font.size = Pt(pt)
    headings = [1 if type(text[0][j]) is not list and len(text[0][j].split()) <= 10 and text[0][j] != '' and text[0][j].find('\t') == -1  else 0 for j in range(len(text[0]))]
    pictures, summa, tables_text, table_num = [1 if type(text[0][phrase]) is not list and text[0][phrase].find('.png') != -1 else 0 for phrase in range(len(text[0]))], 0, [], 0
    tables = [1 if type(text[0][j]) is list else 0 for j in range(len(text[0]))]
    tables.append(0)
    for phrase in range(len(text[0])):
        if type(text[0][phrase]) is list: 
            tables_text.extend(adding_table(text[0], phrase, pt))
            table_num += 1
        elif text[0][phrase].find('.png') != -1: adding_picture(phrase)
        elif text[0][phrase].find('\t') != -1: 
            paragrahp = adding_list(phrase)
            if text[1][phrase]:
                summa += 1
                #fghjk
                run = paragrahp.add_run(' [' + str(footnotes[summa]) + ']')
                run.bold = True
                run.font.size = Pt(10)
        elif len(text[0][phrase].split()) <= 10: adding_heading(phrase)
        else: 
            paragrahp = adding_paragraph(phrase)
            if text[1][phrase]:
                summa += 1
                run = paragrahp.add_run(' [' + str(footnotes[summa]) + ']')
                run.bold = True
                run.font.size = Pt(10)
    upper_header_list, lower_header_list, header_flag = [], [], False
    for split in my_doc:
        if split != '':
            if split == text[0][0]:break
            else: upper_header_list.append(split)
    for split in my_doc:
        if split != '':
            if split not in text[-2] and split not in text[-1] and split not in upper_header_list and split not in tables_text: 
                lower_header_list.append(split)
    if len(upper_header_list) > 0 or len(lower_header_list) > 0: header_flag = True
    if header_flag: adding_headers_and_footers(upper_header_list, lower_header_list)
    add_page_number(new_doc.sections[0].footer.paragraphs[0])
    save_name = Path("New Documents") / str(filename)
    new_doc.save(save_name)

def delete_files():
    directory = Path("New Documents")
    if os.path.isdir("New Documents"):
        for filename in os.listdir(directory):
            f = os.path.join(directory, filename)
            if os.path.isfile(f) and filename.endswith('.docx'): os.remove(f)

directory, files = Path("Documents"), []
delete_files()
if not os.path.isdir("New Documents"): os.mkdir("New Documents")
assert len(os.listdir(directory)) > 0, 'Download files in the folder'
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    if os.path.isfile(f) and filename.endswith('.docx'): 
        files.append(filename[: filename.find('.docx')])
        new_document(14, 1.15, 'Times New Roman', 30, 15, 20, 20)


end_time = time.time()
print(end_time - start_time)

