import os
import re
import random
import copy
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """Đặt màu nền cho ô trong bảng docx"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Đặt lề bên trong ô"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def clean_tex_tags(text):
    """Dọn dẹp một số thẻ TeX cơ bản để hiển thị dễ đọc hơn nếu cần"""
    if not text:
        return ""
    text = re.sub(r'\\textbf\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\textit\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\underline\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\text\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\quad', ' ', text)
    text = re.sub(r'\\qquad', '  ', text)
    text = re.sub(r'\\\\', '\n', text)
    return text.strip()

class ExamParser:
    """Bóc tách đề thi từ file TeX hoặc Word/Text theo cấu trúc chuẩn Bộ GD&ĐT 2025"""
    
    @staticmethod
    def parse_tex(tex_content):
        """
        Bóc tách file TeX định dạng ex-test:
        - \\begin{ex} ... \\choice ... \\loigiai{...} \\end{ex}
        - \\choiceTF
        - \\shortans{...}
        """
        exam_data = {
            "title": "ĐỀ KIỂM TRA ĐỊNH KỲ",
            "subject": "HÓA HỌC",
            "time": "45 phút",
            "part1": [], # Nhiều lựa chọn
            "part2": [], # Đúng / Sai
            "part3": []  # Trả lời ngắn
        }
        
        # Tìm tiêu đề, môn
        title_match = re.search(r'\\textbf\{\\LARGE\s+([^}]+)\}', tex_content)
        if title_match:
            exam_data["title"] = clean_tex_tags(title_match.group(1))
            
        subj_match = re.search(r'MÔN:\s*([^\\}]+)', tex_content)
        if subj_match:
            exam_data["subject"] = clean_tex_tags(subj_match.group(1).replace('--', '-').strip())
            
        time_match = re.search(r'Thời gian làm bài:\s*([^,\\}]+)', tex_content)
        if time_match:
            exam_data["time"] = time_match.group(1).strip()

        # Bóc tách từng khối \begin{ex} ... \end{ex}
        ex_blocks = re.findall(r'\\begin\{ex\}(.*?)\\end\{ex\}', tex_content, re.DOTALL)
        
        for block in ex_blocks:
            # Loại bỏ \loigiai nếu có nhưng lưu lại lời giải
            explanation = ""
            loigiai_match = re.search(r'\\loigiai\{(.*)\}', block, re.DOTALL)
            if loigiai_match:
                explanation = clean_tex_tags(loigiai_match.group(1).strip())
                block_content = block[:loigiai_match.start()]
            else:
                block_content = block
                
            # 1. Kiểm tra phần III: Trả lời ngắn \shortans{...}
            shortans_match = re.search(r'\\shortans\{([^}]+)\}', block_content)
            if shortans_match:
                answer = shortans_match.group(1).strip()
                q_text = block_content[:shortans_match.start()].strip()
                exam_data["part3"].append({
                    "question": clean_tex_tags(q_text),
                    "answer": answer,
                    "explanation": explanation
                })
                continue

            # 2. Kiểm tra phần II: Trắc nghiệm Đúng / Sai (\choiceTF hoặc có các ý a, b, c, d)
            # Một số đề TeX dùng \choiceTF {..} {..} {..} {..} hoặc dạng \item a) ...
            if "\\choiceTF" in block_content or ("a)" in block_content and "b)" in block_content and "c)" in block_content and "d)" in block_content):
                # Xử lý choiceTF
                q_text = block_content
                items = []
                if "\\choiceTF" in block_content:
                    parts = block_content.split("\\choiceTF")
                    q_text = clean_tex_tags(parts[0].strip())
                    tf_body = parts[1]
                    raw_items = re.findall(r'\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', tf_body)
                    for item_str in raw_items:
                        is_true = "\\True" in item_str
                        cleaned_str = clean_tex_tags(item_str.replace("\\True", "").strip())
                        items.append({
                            "text": cleaned_str,
                            "is_true": is_true
                        })
                else:
                    # Tách theo a), b), c), d)
                    q_parts = re.split(r'(?=[a-d]\))', block_content)
                    q_text = clean_tex_tags(q_parts[0].strip())
                    for ip in q_parts[1:]:
                        ip_clean = ip.strip()
                        if len(ip_clean) >= 2 and ip_clean[1] == ')':
                            sub_text = ip_clean[2:].strip()
                            is_true = "\\True" in sub_text
                            items.append({
                                "text": clean_tex_tags(sub_text.replace("\\True", "").strip()),
                                "is_true": is_true
                            })
                
                # Bổ sung nhãn a, b, c, d
                sub_labels = ['a', 'b', 'c', 'd']
                labeled_items = []
                for i, it in enumerate(items[:4]):
                    labeled_items.append({
                        "id": sub_labels[i],
                        "text": it["text"],
                        "is_true": it.get("is_true", False)
                    })
                    
                exam_data["part2"].append({
                    "question": q_text,
                    "items": labeled_items,
                    "explanation": explanation
                })
                continue

            # 3. Phần I: Trắc nghiệm 4 lựa chọn (\choice)
            if "\\choice" in block_content:
                parts = block_content.split("\\choice")
                q_text = clean_tex_tags(parts[0].strip())
                choice_body = parts[1]
                # Tách 4 lựa chọn trong các ngoặc {}
                raw_choices = re.findall(r'\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', choice_body)
                choices = []
                correct_idx = 0
                for idx, c_str in enumerate(raw_choices[:4]):
                    is_correct = "\\True" in c_str
                    if is_correct:
                        correct_idx = idx
                    choices.append(clean_tex_tags(c_str.replace("\\True", "").strip()))
                
                if len(choices) == 4:
                    exam_data["part1"].append({
                        "question": q_text,
                        "options": choices,
                        "correct_idx": correct_idx,
                        "explanation": explanation
                    })
                continue
                
        return exam_data

    @staticmethod
    def _extract_rich_text(element):
        """Trích xuất cả văn bản thường, oMath và công thức MathType (o:title) sang dạng LaTeX"""
        import xml.etree.ElementTree as ET
        try:
            root = ET.fromstring(element.xml)
            parts = []
            for elem in root.iter():
                if elem.tag.endswith('}t') and elem.text:
                    parts.append(elem.text)
                elif elem.attrib.get('{urn:schemas-microsoft-com:office:office}title'):
                    formula = elem.attrib['{urn:schemas-microsoft-com:office:office}title'].strip()
                    parts.append(f" ${formula}$ ")
            return "".join(parts).strip()
        except:
            return element.text.strip() if hasattr(element, 'text') else ""

    @staticmethod
    def parse_docx(doc_path):
        """Bóc tách đề thi từ file docx chuẩn 2025 kể cả khi các phương án nằm trong bảng hoặc chứa MathType"""
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        
        doc = Document(doc_path)
        lines = []
        for child in doc.element.body:
            if child.tag.endswith('p'):
                txt = ExamParser._extract_rich_text(child)
                if txt:
                    lines.append(txt)
            elif child.tag.endswith('tbl'):
                t = Table(child, doc)
                row_items = []
                for row in t.rows:
                    cells = [ExamParser._extract_rich_text(c._tc) for c in row.cells]
                    cells = [c for c in cells if c]
                    if cells:
                        row_items.append("   ".join(cells))
                if row_items:
                    lines.append("\n".join(row_items))

        full_text = "\n".join(lines)
        return ExamParser.parse_raw_text(full_text)

    @staticmethod
    def parse_raw_text(text):
        """Bóc tách từ văn bản thuần với các phân đoạn PHẦN I, PHẦN II, PHẦN III"""
        exam_data = {
            "title": "ĐỀ KIỂM TRA ĐỊNH KỲ",
            "subject": "HÓA HỌC",
            "time": "45 phút",
            "part1": [],
            "part2": [],
            "part3": []
        }

        # Tìm tiêu đề, môn học nếu có
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        for line in lines[:10]:
            if "MÔN:" in line.upper():
                exam_data["subject"] = line.split(":")[-1].strip()
            elif "KIỂM TRA" in line.upper() or "HỌC KỲ" in line.upper():
                exam_data["title"] = line
            elif "THỜI GIAN" in line.upper():
                m = re.search(r'(\d+\s*phút)', line, re.IGNORECASE)
                if m:
                    exam_data["time"] = m.group(1)

        # Phân chia 3 phần dựa trên tiêu đề PHẦN I, PHẦN II, PHẦN III
        p1_idx = text.find("PHẦN I")
        p2_idx = text.find("PHẦN II")
        p3_idx = text.find("PHẦN III")

        p1_text = ""
        p2_text = ""
        p3_text = ""

        if p1_idx != -1:
            end_p1 = p2_idx if p2_idx != -1 else (p3_idx if p3_idx != -1 else len(text))
            p1_text = text[p1_idx:end_p1]

        if p2_idx != -1:
            end_p2 = p3_idx if p3_idx != -1 else len(text)
            p2_text = text[p2_idx:end_p2]

        if p3_idx != -1:
            p3_text = text[p3_idx:]

        # Nếu không có chữ PHẦN I rõ ràng, thử bóc toàn bộ văn bản vào Phần 1
        if not p1_text and not p2_text and not p3_text:
            p1_text = text

        # 1. Bóc tách PHẦN I (Trắc nghiệm nhiều lựa chọn: Câu X: ... A. ... B. ... C. ... D. ...)
        if p1_text:
            # Tách các câu hỏi bắt đầu bằng "Câu 1:", "Câu 2:" hoặc "Câu 1."
            q_chunks = re.split(r'\n(?=Câu\s+\d+[:.])', p1_text)
            for chunk in q_chunks:
                if not re.match(r'^Câu\s+\d+[:.]', chunk.strip()):
                    continue
                # Tách nội dung câu và các phương án A, B, C, D
                # Regex tìm vị trí của A., B., C., D.
                opt_matches = list(re.finditer(r'(?:^|\n|\s+)([A-D])[\.:\)]\s+', chunk))
                if len(opt_matches) >= 4:
                    q_content = chunk[:opt_matches[0].start()].strip()
                    # Bỏ tiền tố "Câu X:"
                    q_content = re.sub(r'^Câu\s+\d+[:.]\s*', '', q_content)
                    
                    options = []
                    for i in range(4):
                        start_pos = opt_matches[i].end()
                        end_pos = opt_matches[i+1].start() if i < 3 else len(chunk)
                        opt_text = chunk[start_pos:end_pos].strip()
                        options.append(opt_text)
                        
                    # Mặc định correct_idx = 0 nếu không đánh dấu sao/đáp án
                    exam_data["part1"].append({
                        "question": q_content,
                        "options": options,
                        "correct_idx": 0,
                        "explanation": ""
                    })

        # 2. Bóc tách PHẦN II (Trắc nghiệm đúng sai: a), b), c), d))
        if p2_text:
            q_chunks = re.split(r'\n(?=Câu\s+\d+[:.])', p2_text)
            for chunk in q_chunks:
                if not re.match(r'^Câu\s+\d+[:.]', chunk.strip()):
                    continue
                sub_matches = list(re.finditer(r'(?:^|\n|\s+)([a-d])[\)\.:]\s+', chunk))
                if len(sub_matches) >= 4:
                    q_content = chunk[:sub_matches[0].start()].strip()
                    q_content = re.sub(r'^Câu\s+\d+[:.]\s*', '', q_content)
                    items = []
                    labels = ['a', 'b', 'c', 'd']
                    for i in range(4):
                        start_pos = sub_matches[i].end()
                        end_pos = sub_matches[i+1].start() if i < 3 else len(chunk)
                        item_text = chunk[start_pos:end_pos].strip()
                        # Kiểm tra xem có ghi chú [Đ] hoặc [S] không
                        is_true = True
                        if "[S]" in item_text or "(Sai)" in item_text:
                            is_true = False
                        items.append({
                            "id": labels[i],
                            "text": item_text,
                            "is_true": is_true
                        })
                    exam_data["part2"].append({
                        "question": q_content,
                        "items": items,
                        "explanation": ""
                    })

        # 3. Bóc tách PHẦN III (Trả lời ngắn)
        if p3_text:
            q_chunks = re.split(r'\n(?=Câu\s+\d+[:.])', p3_text)
            for chunk in q_chunks:
                if not re.match(r'^Câu\s+\d+[:.]', chunk.strip()):
                    continue
                q_content = re.sub(r'^Câu\s+\d+[:.]\s*', '', chunk.strip())
                # Tìm xem có đáp án mẫu không
                ans = ""
                ans_match = re.search(r'(?:Đáp án|Đáp số|Đ/A)[:=]\s*([^\n]+)', q_content, re.IGNORECASE)
                if ans_match:
                    ans = ans_match.group(1).strip()
                    q_content = q_content[:ans_match.start()].strip()
                    
                exam_data["part3"].append({
                    "question": q_content,
                    "answer": ans,
                    "explanation": ""
                })

        return exam_data


class ExamShuffler:
    """Xáo đề thi tạo các mã đề khác nhau kèm ma trận đáp án"""

    @staticmethod
    def shuffle(original_exam, exam_codes=["101", "102", "103", "104"], shuffle_part1_opts=True, shuffle_part2_items=True):
        """
        Xáo đề và tạo danh sách các mã đề
        """
        shuffled_exams = {}
        
        for code in exam_codes:
            # Dùng mã đề làm seed để nếu chạy lại thì kết quả sinh ra cố định
            rng = random.Random(int(code) if code.isdigit() else hash(code))
            
            new_exam = {
                "code": code,
                "title": original_exam.get("title", "ĐỀ KIỂM TRA ĐỊNH KỲ"),
                "subject": original_exam.get("subject", "HÓA HỌC"),
                "time": original_exam.get("time", "45 phút"),
                "part1": [],
                "part2": [],
                "part3": []
            }

            # --- XÁO PHẦN I ---
            # Xáo thứ tự câu hỏi
            p1_questions = copy.deepcopy(original_exam.get("part1", []))
            rng.shuffle(p1_questions)
            
            for q_idx, q in enumerate(p1_questions):
                opts = list(enumerate(q["options"])) # [(0, opt0), (1, opt1), (2, opt2), (3, opt3)]
                correct_orig_idx = q.get("correct_idx", 0)
                
                if shuffle_part1_opts:
                    rng.shuffle(opts)
                    
                new_opts = [o[1] for o in opts]
                # Tìm vị trí mới của đáp án đúng ban đầu
                new_correct_idx = 0
                for new_i, (orig_i, _) in enumerate(opts):
                    if orig_i == correct_orig_idx:
                        new_correct_idx = new_i
                        break
                        
                new_exam["part1"].append({
                    "number": q_idx + 1,
                    "question": q["question"],
                    "options": new_opts,
                    "correct_idx": new_correct_idx,
                    "correct_label": ["A", "B", "C", "D"][new_correct_idx],
                    "explanation": q.get("explanation", "")
                })

            # --- XÁO PHẦN II ---
            p2_questions = copy.deepcopy(original_exam.get("part2", []))
            rng.shuffle(p2_questions)
            
            for q_idx, q in enumerate(p2_questions):
                items = copy.deepcopy(q["items"])
                if shuffle_part2_items:
                    rng.shuffle(items)
                
                # Gán lại id a, b, c, d
                labels = ['a', 'b', 'c', 'd']
                labeled_items = []
                for i, it in enumerate(items[:4]):
                    it["id"] = labels[i]
                    labeled_items.append(it)
                    
                new_exam["part2"].append({
                    "number": q_idx + 1,
                    "question": q["question"],
                    "items": labeled_items,
                    "explanation": q.get("explanation", "")
                })

            # --- XÁO PHẦN III ---
            p3_questions = copy.deepcopy(original_exam.get("part3", []))
            rng.shuffle(p3_questions)
            for q_idx, q in enumerate(p3_questions):
                new_exam["part3"].append({
                    "number": q_idx + 1,
                    "question": q["question"],
                    "answer": q.get("answer", ""),
                    "explanation": q.get("explanation", "")
                })

            shuffled_exams[code] = new_exam
            
        return shuffled_exams


class ExamWordExporter:
    """Xuất đề thi ra file Word (.docx) chuẩn mẫu in ấn của Bộ GD&ĐT"""

    @staticmethod
    def export_single_exam(exam_data, output_path, school_name="TRƯỜNG THPT CHUYÊN ...", exam_name=None):
        doc = Document()

        # Thiết lập lề trang A4 chuẩn (Top: 2cm, Bottom: 2cm, Left: 2.5cm, Right: 2cm)
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.79)
            section.bottom_margin = Inches(0.79)
            section.left_margin = Inches(0.98)
            section.right_margin = Inches(0.79)

        # Style mặc định Times New Roman 12pt
        normal_style = doc.styles['Normal']
        normal_style.font.name = 'Times New Roman'
        normal_style.font.size = Pt(12)
        normal_style.font.color.rgb = RGBColor(0, 0, 0)
        normal_style.paragraph_format.line_spacing = 1.15
        normal_style.paragraph_format.space_after = Pt(3)

        # 1. BẢNG TIÊU ĐỀ TRƯỜNG & MÃ ĐỀ
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.columns[0].width = Inches(4.0)
        table.columns[1].width = Inches(2.7)

        # Cột trái: Tên trường & Kỳ thi
        cell_left = table.cell(0, 0)
        set_cell_margins(cell_left, top=60, bottom=60, left=60, right=60)
        p_l1 = cell_left.paragraphs[0]
        p_l1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p_l1.add_run(f"{school_name.upper()}\n")
        r1.font.bold = True
        r1.font.size = Pt(11)
        
        exam_title = exam_name if exam_name else exam_data.get("title", "ĐỀ KIỂM TRA ĐỊNH KỲ")
        r2 = p_l1.add_run(f"{exam_title.upper()}\n")
        r2.font.bold = True
        r2.font.size = Pt(12)
        
        r3 = p_l1.add_run(f"MÔN: {exam_data.get('subject', 'HÓA HỌC').upper()}\n")
        r3.font.bold = True
        r3.font.size = Pt(11)
        
        r4 = p_l1.add_run(f"Thời gian làm bài: {exam_data.get('time', '45 phút')}\n(Không kể thời gian phát đề)")
        r4.font.italic = True
        r4.font.size = Pt(10)

        # Cột phải: Khung Mã Đề & Thông tin thí sinh
        cell_right = table.cell(0, 1)
        set_cell_margins(cell_right, top=60, bottom=60, left=60, right=60)
        p_r1 = cell_right.paragraphs[0]
        p_r1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        code_run = p_r1.add_run(f"MÃ ĐỀ THI: {exam_data.get('code', '101')}\n")
        code_run.font.bold = True
        code_run.font.size = Pt(14)
        code_run.font.color.rgb = RGBColor(180, 0, 0)

        p_info = cell_right.add_paragraph()
        p_info.paragraph_format.line_spacing = 1.15
        p_info.paragraph_format.space_after = Pt(2)
        p_info.add_run("Họ, tên thí sinh: .......................................\n")
        p_info.add_run("Số báo danh: .............................................\n")
        p_info.add_run("Lớp: ............................................................")

        doc.add_paragraph() # Dòng trống phân cách

        # 2. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
        p1_list = exam_data.get("part1", [])
        if p1_list:
            h1 = doc.add_paragraph()
            r_h1 = h1.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. ")
            r_h1.font.bold = True
            r_h1.font.size = Pt(12)
            
            num_q = len(p1_list)
            score = round(num_q * 0.25, 2)
            r_score = h1.add_run(f"({score} điểm)\n")
            r_score.font.bold = True
            
            r_inst = h1.add_run(f"Thí sinh trả lời từ câu 1 đến câu {num_q}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
            r_inst.font.italic = True
            
            for q in p1_list:
                p_q = doc.add_paragraph()
                r_qnum = p_q.add_run(f"Câu {q['number']}: ")
                r_qnum.font.bold = True
                p_q.add_run(q["question"])
                
                # 4 phương án A, B, C, D
                opts = q["options"]
                labels = ["A.", "B.", "C.", "D."]
                
                # Trình bày dạng 2 cột hoặc 1 dòng tùy độ dài
                p_opts = doc.add_paragraph()
                p_opts.paragraph_format.left_indent = Inches(0.2)
                p_opts.paragraph_format.space_after = Pt(4)
                
                # Kiểm tra nếu đáp án ngắn thì trên 1 dòng hoặc 2 dòng
                total_len = sum(len(o) for o in opts)
                if total_len < 80:
                    for i in range(4):
                        r_lbl = p_opts.add_run(f"{labels[i]} ")
                        r_lbl.font.bold = True
                        p_opts.add_run(f"{opts[i]}\t\t")
                else:
                    for i in range(4):
                        if i > 0:
                            p_opts = doc.add_paragraph()
                            p_opts.paragraph_format.left_indent = Inches(0.2)
                            p_opts.paragraph_format.space_after = Pt(2)
                        r_lbl = p_opts.add_run(f"{labels[i]} ")
                        r_lbl.font.bold = True
                        p_opts.add_run(opts[i])

        # 3. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI
        p2_list = exam_data.get("part2", [])
        if p2_list:
            doc.add_paragraph()
            h2 = doc.add_paragraph()
            r_h2 = h2.add_run("PHẦN II. Câu trắc nghiệm đúng sai. ")
            r_h2.font.bold = True
            r_h2.font.size = Pt(12)
            
            num_q2 = len(p2_list)
            r_inst2 = h2.add_run(f"Thí sinh trả lời từ câu 1 đến câu {num_q2}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
            r_inst2.font.italic = True
            
            for q in p2_list:
                p_q = doc.add_paragraph()
                r_qnum = p_q.add_run(f"Câu {q['number']}: ")
                r_qnum.font.bold = True
                p_q.add_run(q["question"])
                
                for item in q["items"]:
                    p_it = doc.add_paragraph()
                    p_it.paragraph_format.left_indent = Inches(0.25)
                    p_it.paragraph_format.space_after = Pt(2)
                    r_lbl = p_it.add_run(f"{item['id']}) ")
                    r_lbl.font.bold = True
                    p_it.add_run(item["text"])

        # 4. PHẦN III: TRẢ LỜI NGẮN
        p3_list = exam_data.get("part3", [])
        if p3_list:
            doc.add_paragraph()
            h3 = doc.add_paragraph()
            r_h3 = h3.add_run("PHẦN III. Câu trắc nghiệm trả lời ngắn. ")
            r_h3.font.bold = True
            r_h3.font.size = Pt(12)
            
            num_q3 = len(p3_list)
            r_inst3 = h3.add_run(f"Thí sinh trả lời từ câu 1 đến câu {num_q3}. Thí sinh ghi đáp số/kết quả vào khung tương ứng.")
            r_inst3.font.italic = True
            
            for q in p3_list:
                p_q = doc.add_paragraph()
                r_qnum = p_q.add_run(f"Câu {q['number']}: ")
                r_qnum.font.bold = True
                p_q.add_run(q["question"])
                
                p_ans_box = doc.add_paragraph()
                p_ans_box.paragraph_format.left_indent = Inches(0.25)
                p_ans_box.paragraph_format.space_after = Pt(4)
                r_box = p_ans_box.add_run("Đáp số: [ ............................................. ]")
                r_box.font.italic = True

        # Cuối trang: Dòng kết thúc
        p_end = doc.add_paragraph()
        p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_end.paragraph_format.space_before = Pt(15)
        r_end = p_end.add_run("-------------------- HẾT --------------------\n(Cán bộ coi thi không giải thích gì thêm)")
        r_end.font.bold = True
        r_end.font.size = Pt(11)

        doc.save(output_path)
        return output_path

    @staticmethod
    def export_answer_keys(shuffled_exams, output_path, exam_name=None):
        """Xuất file bảng đáp án tổng hợp đối chiếu tất cả các mã đề"""
        doc = Document()
        normal_style = doc.styles['Normal']
        normal_style.font.name = 'Times New Roman'
        normal_style.font.size = Pt(11)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t_run = p_title.add_run(f"BẢNG ĐÁP ÁN VÀ MA TRẬN CHẤM ĐIỂM TỔNG HỢP\n")
        t_run.font.bold = True
        t_run.font.size = Pt(14)
        t_run.font.color.rgb = RGBColor(180, 0, 0)
        
        if exam_name:
            sub_run = p_title.add_run(f"{exam_name.upper()}\n")
            sub_run.font.bold = True
            sub_run.font.size = Pt(12)

        codes = list(shuffled_exams.keys())

        # --- BẢNG ĐÁP ÁN PHẦN I ---
        h1 = doc.add_paragraph()
        r1 = h1.add_run("1. ĐÁP ÁN PHẦN I (Trắc nghiệm nhiều lựa chọn)")
        r1.font.bold = True
        r1.font.size = Pt(12)

        first_exam = shuffled_exams[codes[0]]
        num_p1 = len(first_exam.get("part1", []))

        if num_p1 > 0:
            table = doc.add_table(rows=num_p1 + 1, cols=len(codes) + 1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = True

            # Tiêu đề cột
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = "Câu"
            set_cell_background(hdr_cells[0], "E0E0E0")
            hdr_cells[0].paragraphs[0].runs[0].font.bold = True
            hdr_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for col_i, code in enumerate(codes):
                hdr_cells[col_i + 1].text = f"Mã {code}"
                set_cell_background(hdr_cells[col_i + 1], "D9EDF7")
                hdr_cells[col_i + 1].paragraphs[0].runs[0].font.bold = True
                hdr_cells[col_i + 1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Dữ liệu đáp án từng câu
            for q_i in range(num_p1):
                row_cells = table.rows[q_i + 1].cells
                row_cells[0].text = str(q_i + 1)
                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                row_cells[0].paragraphs[0].runs[0].font.bold = True
                
                for col_i, code in enumerate(codes):
                    ex = shuffled_exams[code]
                    ans = ex["part1"][q_i]["correct_label"]
                    row_cells[col_i + 1].text = ans
                    row_cells[col_i + 1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[col_i + 1].paragraphs[0].runs[0].font.bold = True

        # --- BẢNG ĐÁP ÁN PHẦN II ---
        num_p2 = len(first_exam.get("part2", []))
        if num_p2 > 0:
            doc.add_paragraph()
            h2 = doc.add_paragraph()
            r2 = h2.add_run("2. ĐÁP ÁN PHẦN II (Trắc nghiệm đúng sai)")
            r2.font.bold = True
            r2.font.size = Pt(12)

            for code in codes:
                ex = shuffled_exams[code]
                p_code = doc.add_paragraph()
                r_c = p_code.add_run(f"• Mã đề {code}:")
                r_c.font.bold = True
                
                for q in ex.get("part2", []):
                    ans_str = ", ".join([f"{it['id']}: {'Đ' if it.get('is_true', True) else 'S'}" for it in q["items"]])
                    p_ans = doc.add_paragraph()
                    p_ans.paragraph_format.left_indent = Inches(0.3)
                    p_ans.add_run(f"Câu {q['number']}: ").font.bold = True
                    p_ans.add_run(ans_str)

        # --- BẢNG ĐÁP ÁN PHẦN III ---
        num_p3 = len(first_exam.get("part3", []))
        if num_p3 > 0:
            doc.add_paragraph()
            h3 = doc.add_paragraph()
            r3 = h3.add_run("3. ĐÁP ÁN PHẦN III (Trả lời ngắn)")
            r3.font.bold = True
            r3.font.size = Pt(12)

            for code in codes:
                ex = shuffled_exams[code]
                p_code = doc.add_paragraph()
                r_c = p_code.add_run(f"• Mã đề {code}:")
                r_c.font.bold = True
                
                for q in ex.get("part3", []):
                    p_ans = doc.add_paragraph()
                    p_ans.paragraph_format.left_indent = Inches(0.3)
                    p_ans.add_run(f"Câu {q['number']}: ").font.bold = True
                    p_ans.add_run(f"{q.get('answer', '---')}")

        doc.save(output_path)
        return output_path
