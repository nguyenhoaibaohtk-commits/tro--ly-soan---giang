import os
import io
import zipfile
import json
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename
from exam_processor import ExamParser, ExamShuffler, ExamWordExporter
from khbd_processor import KHBDProcessor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024 # 32MB

# Bộ nhớ tạm lưu trữ đề hiện tại
CURRENT_EXAM_STORE = {
    "parsed": None,
    "shuffled": None
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/list_local_files', methods=['GET'])
def list_local_files():
    """Liệt kê các file tài liệu có sẵn trong thư mục làm việc"""
    try:
        supported_exts = ('.tex', '.docx', '.txt')
        files = []
        for f in os.listdir(BASE_DIR):
            if f.lower().endswith(supported_exts):
                f_path = os.path.join(BASE_DIR, f)
                size_kb = round(os.path.getsize(f_path) / 1024, 1)
                files.append({
                    "name": f,
                    "size": f"{size_kb} KB",
                    "ext": os.path.splitext(f)[1].lower()
                })
        return jsonify({"status": "success", "files": files})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/load_file', methods=['POST'])
def load_file():
    """Nạp và phân tích file có sẵn trong máy"""
    try:
        data = request.json or {}
        filename = data.get("filename")
        if not filename:
            return jsonify({"status": "error", "message": "Chưa chọn file"}), 400
            
        file_path = os.path.join(BASE_DIR, filename)
        if not os.path.exists(file_path):
            return jsonify({"status": "error", "message": "File không tồn tại"}), 404
            
        ext = os.path.splitext(filename)[1].lower()
        if ext == '.tex':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            parsed = ExamParser.parse_tex(content)
            raw_preview = content[:2000]
        elif ext == '.docx':
            parsed = ExamParser.parse_docx(file_path)
            raw_preview = f"Đã đọc file Word: {filename}"
        else: # .txt
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            parsed = ExamParser.parse_raw_text(content)
            raw_preview = content[:2000]

        CURRENT_EXAM_STORE["parsed"] = parsed
        return jsonify({
            "status": "success",
            "filename": filename,
            "parsed": parsed,
            "counts": {
                "part1": len(parsed.get("part1", [])),
                "part2": len(parsed.get("part2", [])),
                "part3": len(parsed.get("part3", []))
            },
            "preview": raw_preview
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/upload_file', methods=['POST'])
def upload_file():
    """Tải file từ máy người dùng lên để xử lý"""
    try:
        if 'file' not in request.files:
            return jsonify({"status": "error", "message": "Không tìm thấy file tải lên"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"status": "error", "message": "Chưa chọn file"}), 400

        filename = secure_filename(file.filename)
        save_path = os.path.join(UPLOAD_DIR, filename)
        file.save(save_path)

        ext = os.path.splitext(filename)[1].lower()
        if ext == '.tex':
            with open(save_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            parsed = ExamParser.parse_tex(content)
        elif ext == '.docx':
            parsed = ExamParser.parse_docx(save_path)
        else:
            with open(save_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            parsed = ExamParser.parse_raw_text(content)

        CURRENT_EXAM_STORE["parsed"] = parsed
        return jsonify({
            "status": "success",
            "filename": filename,
            "parsed": parsed,
            "counts": {
                "part1": len(parsed.get("part1", [])),
                "part2": len(parsed.get("part2", [])),
                "part3": len(parsed.get("part3", []))
            }
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/shuffle_exam', methods=['POST'])
def shuffle_exam():
    """Xáo đề thi thành các mã đề mong muốn và tạo file Word"""
    try:
        req_data = request.json or {}
        exam_data = req_data.get("exam_data") or CURRENT_EXAM_STORE.get("parsed")
        
        if not exam_data:
            return jsonify({"status": "error", "message": "Chưa có dữ liệu đề thi để xáo"}), 400

        exam_codes = req_data.get("codes", ["101", "102", "103", "104"])
        school_name = req_data.get("school_name", "TRƯỜNG THPT .....................")
        exam_name = req_data.get("exam_name", exam_data.get("title", "ĐỀ KIỂM TRA ĐỊNH KỲ"))
        shuffle_opts = req_data.get("shuffle_opts", True)
        shuffle_tf = req_data.get("shuffle_tf", True)

        # Xóa các file cũ trong thư mục output
        for f in os.listdir(OUTPUT_DIR):
            try:
                os.remove(os.path.join(OUTPUT_DIR, f))
            except:
                pass

        # Thực hiện xáo đề
        shuffled = ExamShuffler.shuffle(
            exam_data,
            exam_codes=exam_codes,
            shuffle_part1_opts=shuffle_opts,
            shuffle_part2_items=shuffle_tf
        )
        CURRENT_EXAM_STORE["shuffled"] = shuffled

        # Xuất file Word từng mã đề
        exported_files = []
        for code, ex_content in shuffled.items():
            fname = f"De_Kiem_Tra_Ma_{code}.docx"
            out_path = os.path.join(OUTPUT_DIR, fname)
            ExamWordExporter.export_single_exam(ex_content, out_path, school_name=school_name, exam_name=exam_name)
            exported_files.append({
                "code": code,
                "filename": fname,
                "url": f"/api/download/{fname}"
            })

        # Xuất file Bảng đáp án tổng hợp
        ans_fname = "Bang_Dap_An_Va_Ma_Tran_Tat_Ca_Ma_De.docx"
        ans_path = os.path.join(OUTPUT_DIR, ans_fname)
        ExamWordExporter.export_answer_keys(shuffled, ans_path, exam_name=exam_name)
        exported_files.append({
            "code": "DAP_AN",
            "filename": ans_fname,
            "url": f"/api/download/{ans_fname}"
        })

        return jsonify({
            "status": "success",
            "shuffled": shuffled,
            "files": exported_files,
            "zip_url": "/api/download_all_zip"
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    """Tải 1 file Word cụ thể"""
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

@app.route('/api/download_all_zip', methods=['GET'])
def download_all_zip():
    """Tải toàn bộ đề thi và đáp án dạng file nén .zip"""
    try:
        mem_zip = io.BytesIO()
        with zipfile.ZipFile(mem_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
            for f in os.listdir(OUTPUT_DIR):
                f_path = os.path.join(OUTPUT_DIR, f)
                if os.path.isfile(f_path):
                    zf.write(f_path, arcname=f)
        mem_zip.seek(0)
        return send_file(
            mem_zip,
            mimetype='application/zip',
            as_attachment=True,
            download_name='Bo_De_Kiem_Tra_Va_Dap_An_2025.zip'
        )
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# ----------------- PHẦN KẾ HOẠCH BÀI DẠY (KHBD) -----------------

@app.route('/api/khbd/list_lessons', methods=['GET'])
def list_khbd_lessons():
    """Lấy danh mục các bài học mẫu chuẩn 5512"""
    try:
        lessons = KHBDProcessor.get_lesson_list()
        return jsonify({"status": "success", "lessons": lessons})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/khbd/get_lesson/<lesson_id>', methods=['GET'])
def get_khbd_lesson(lesson_id):
    """Lấy dữ liệu chi tiết bài học theo ID"""
    try:
        data = KHBDProcessor.get_lesson(lesson_id)
        return jsonify({"status": "success", "data": data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/khbd/generate_smart', methods=['POST'])
def generate_smart_khbd():
    """Tự động sinh bộ khung giáo án chuẩn 5512 theo tên bài nhập vào"""
    try:
        req = request.json or {}
        title = req.get("lesson_title", "BÀI HỌC MỚI")
        grade = req.get("grade", "10")
        subject = req.get("subject", "HÓA HỌC")
        data = KHBDProcessor.generate_smart_outline(title, grade, subject)
        return jsonify({"status": "success", "data": data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/khbd/get_sample', methods=['GET'])
def get_khbd_sample():
    """Lấy dữ liệu mẫu Kế hoạch bài dạy chuẩn 5512"""
    try:
        sample = KHBDProcessor.get_sample_data()
        return jsonify({"status": "success", "data": sample})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/khbd/export', methods=['POST'])
def export_khbd():
    """Xuất Kế hoạch bài dạy ra file Word"""
    try:
        data = request.json or {}
        if not data:
            return jsonify({"status": "error", "message": "Không có dữ liệu KHBD"}), 400

        lesson_title = data.get("lesson_title", "KHBD_5512")
        clean_title = "".join(c for c in lesson_title if c.isalnum() or c in (' ', '_', '-')).strip()
        filename = f"KHBD_{clean_title[:30]}.docx"
        out_path = os.path.join(OUTPUT_DIR, filename)

        KHBDProcessor.export_word(data, out_path)
        return jsonify({
            "status": "success",
            "filename": filename,
            "url": f"/api/download/{filename}"
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    import webbrowser
    import threading
    import time

    def open_browser():
        time.sleep(1.2)
        try:
            webbrowser.open("http://127.0.0.1:5000")
        except:
            pass

    threading.Thread(target=open_browser, daemon=True).start()

    print("="*60)
    print(" ỨNG DỤNG TRỢ LÝ GIÁO VIÊN (EDUKIT 2025)")
    print(" Soạn Kế Hoạch Bài Dạy CV 5512 & Trộn Đề Thi 2025")
    print(" Máy của bạn:       http://localhost:5000")
    print(" Cho máy khác dùng: http://192.168.1.13:5000 (Cùng mạng Wi-Fi)")
    print(" Vui lòng không đóng cửa sổ này trong khi dùng app.")
    print("="*60)
    app.run(host='0.0.0.0', port=5000, debug=False)
