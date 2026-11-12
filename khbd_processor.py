import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

class KHBDProcessor:
    """
    Xử lý tạo và xuất Kế hoạch bài dạy theo Công văn 5512/BGDĐT-GDTrH
    Tích hợp Năng lực số (CV 3456/BGDĐT-GDPT) và An toàn mạng toàn diện (QĐ 6206/BGDĐT-GDPT)
    """

    LESSON_DATABASE = {
        # ==================== LỚP 10 ====================
        "hoa10_bai3": {
            "id": "hoa10_bai3",
            "grade": "10",
            "lesson_title": "BÀI 3: CẤU TRÚC LỚP VỎ ELECTRON CỦA NGUYÊN TỬ",
            "duration": "4 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Trình bày được mô hình của Rutherford - Bohr và mô hình hiện đại về orbital nguyên tử (AO).\n- Nêu được khái niệm về orbital nguyên tử (AO s, AO p) và số lượng electron trong một AO.\n- Trình bày được cấu trúc lớp vỏ electron: lớp, phân lớp electron và số lượng electron tối đa trong một phân lớp, một lớp.\n- Viết được cấu hình electron nguyên tử của 20 nguyên tố đầu tiên theo ô orbital và dạng số.",
            "objective_competence_general": "- Tự chủ và tự học: Chủ động đọc SGK, quan sát video mô phỏng và hoàn thành phiếu học tập cá nhân.\n- Giao tiếp và hợp tác: Tích cực thảo luận nhóm, phân công nhiệm vụ rõ ràng và chia sẻ kết quả.\n- Giải quyết vấn đề và sáng tạo: Vận dụng quy tắc viết cấu hình electron để dự đoán tính chất cơ bản của nguyên tố.",
            "objective_competence_subject": "- Nhận thức hóa học: Phân biệt được mô hình hành tinh nguyên tử cổ điển và mô hình hiện đại về đám mây electron; hiểu được ý nghĩa các lớp và phân lớp electron.\n- Tìm hiểu thế giới tự nhiên: Sử dụng mô hình 3D để hình dung không gian của AO s (hình cầu) và AO p (hình số 8 nổi).\n- Vận dụng kiến thức: Dự đoán được nguyên tố kim loại, phi kim hay khí hiếm dựa vào số electron lớp ngoài cùng.",
            "objective_competence_digital_safety": "- Năng lực số (Mã 5.2.NC1a, 3.1.NC1a): Sử dụng công cụ mô phỏng 3D/thí nghiệm ảo (PhET Interactive Simulations) để quan sát hình dạng đám mây electron, orbital nguyên tử s, p; sử dụng bảng tương tác số/phần mềm học tập để luyện tập viết cấu hình electron nguyên tử.\n- An toàn mạng ([ATM4, ATM6]): Tuân thủ quy định an toàn thiết bị khi truy cập các nền tảng thí nghiệm ảo; tôn trọng bản quyền, trích dẫn nguồn gốc học liệu số và hình ảnh mô phỏng 3D khi hoàn thành báo cáo học tập.",
            "objective_qualities": "- Chăm chỉ: Nghiêm túc hoàn thành các nhiệm vụ học tập trên lớp và bài tập về nhà.\n- Trung thực: Khách quan, trung thực trong quá trình báo cáo kết quả thảo luận nhóm.\n- Trách nhiệm: Có ý thức hợp tác trách nhiệm với các thành viên trong nhóm.",
            "equipment_teacher": "- Máy tính, máy chiếu, bài giảng PowerPoint có tích hợp liên kết thí nghiệm ảo PhET và video mô phỏng 3D orbital s, p.\n- Phiếu học tập số 1, 2 và bảng tương tác số/bảng phụ.",
            "equipment_student": "- Sách giáo khoa Hóa học 10, vở ghi chép.\n- Điện thoại thông minh/máy tính bảng có kết nối Internet để truy cập mô phỏng PhET.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (7 phút)",
                    "goal": "Tạo mâu thuẫn nhận thức về sự phân bố và chuyển động của electron quanh hạt nhân nguyên tử.",
                    "content": "GV chiếu hình ảnh so sánh mô hình hành tinh nguyên tử của Rutherford - Bohr với mô hình đám mây xác suất hiện đại. HS trả lời câu hỏi: Electron chuyển động theo quỹ đạo cố định hay trong một vùng không gian xác suất?",
                    "product": "Câu trả lời của HS: nhận thức được sự giới hạn của mô hình cổ điển và tính ưu việt của mô hình hiện đại.",
                    "implementation_gv": "GV trình chiếu video ngắn (1 phút) mô phỏng chuyển động hỗn loạn của electron quanh hạt nhân. Nêu câu hỏi gợi mở.",
                    "implementation_hs": "HS quan sát, trao đổi cặp đôi và phát biểu dự đoán."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (30 phút)",
                    "goal": "Khám phá hình dạng AO s, p qua mô phỏng 3D; quy tắc viết cấu hình e.",
                    "content": "HS làm việc nhóm: truy cập mô phỏng PhET trên máy tính bảng/điện thoại, xoay các góc nhìn 3D của AO s và AO p. Điền Phiếu học tập số 1 về số lượng e tối đa và quy tắc điền electron.",
                    "product": "Phiếu học tập số 1 hoàn chỉnh; viết đúng cấu hình e của 20 nguyên tố đầu tiên.",
                    "implementation_gv": "1. Chuyển giao: GV hướng dẫn HS quét mã QR truy cập trang mô phỏng PhET an toàn, không bấm vào liên kết lạ ([ATM4]). Giao Phiếu học tập số 1.\n2. Thực hiện: Các nhóm thao tác trên mô phỏng 3D, thảo luận hoàn thành phiếu.\n3. Báo cáo: Đại diện nhóm trình chiếu kết quả trên màn hình lớp.\n4. Kết luận: GV chuẩn hóa kiến thức về AO và cấu hình electron.",
                    "implementation_hs": "HS sử dụng công cụ số an toàn, quan sát trực quan AO 3D, ghi nhận kết quả vào phiếu học tập."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (5 phút)",
                    "goal": "Rèn luyện kĩ năng viết cấu hình electron và xác định tính chất nguyên tố.",
                    "content": "Làm bài tập tương tác trên Quizizz/bảng số: viết cấu hình electron cho Na (Z=11), Cl (Z=17), Fe (Z=26).",
                    "product": "Kết quả làm bài tập trực tuyến của học sinh.",
                    "implementation_gv": "GV mở phòng luyện tập trực tuyến, nhắc nhở HS sử dụng tên tài khoản nghiêm túc ([ATM2]).",
                    "implementation_hs": "HS tham gia làm bài, đối chiếu kết quả với bạn."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG VÀ MỞ RỘNG (3 phút)",
                    "goal": "Tìm hiểu ứng dụng của cấu hình electron bền vững của khí hiếm.",
                    "content": "Tìm kiếm thông tin trực tuyến có dẫn nguồn về ứng dụng của khí hiếm Argon trong hàn kim loại và đèn Neon quảng cáo.",
                    "product": "Bản báo cáo ngắn gọn có ghi rõ nguồn tài liệu số tham khảo ([ATM6]).",
                    "implementation_gv": "Giao nhiệm vụ về nhà, hướng dẫn chọn nguồn tin cậy từ các trang khoa học chính thống.",
                    "implementation_hs": "Ghi nhận nhiệm vụ và thực hiện tại nhà."
                }
            ],
            "appendix": "PHIẾU HỌC TẬP SỐ 1:\n1. Quan sát mô phỏng PhET: Nêu hình dạng và sự định hướng trong không gian của AO s và các AO px, py, pz.\n2. Lớp thứ n có bao nhiêu phân lớp? Số electron tối đa ở lớp n là bao nhiêu?\n3. Viết cấu hình electron của Al (Z=13) và S (Z=16)."
        },

        "hoa10_bai5": {
            "id": "hoa10_bai5",
            "grade": "10",
            "lesson_title": "BÀI 5: CẤU TẠO CỦA BẢNG TUẦN HOÀN CÁC NGUYÊN TỐ HÓA HỌC",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Nêu được lịch sử phát minh bảng tuần hoàn của Mendeleev và nguyên tắc sắp xếp các nguyên tố.\n- Mô tả được cấu tạo bảng tuần hoàn gồm: ô nguyên tố, chu kì (chu kì nhỏ, chu kì lớn) và nhóm (nhóm A, nhóm B).\n- Xác định được mối liên hệ giữa cấu hình electron nguyên tử với vị trí (ô, chu kì, nhóm) trong bảng tuần hoàn.",
            "objective_competence_general": "- Tự chủ và tự học: Tự tra cứu bảng tuần hoàn để xác định vị trí các nguyên tố.\n- Giao tiếp và hợp tác: Làm việc nhóm để phân tích đặc điểm cấu tạo của các chu kì và nhóm nguyên tố.\n- Giải quyết vấn đề: Suy luận vị trí của nguyên tố khi biết số hiệu nguyên tử Z hoặc cấu hình electron.",
            "objective_competence_subject": "- Nhận thức hóa học: Hiểu được nguyên tắc số thứ tự chu kì bằng số lớp electron; số thứ tự nhóm A bằng số electron hóa trị.\n- Tìm hiểu thế giới tự nhiên: Khám phá trật tự hài hòa của các nguyên tố trong vũ trụ thông qua bảng tuần hoàn.",
            "objective_competence_digital_safety": "- Năng lực số (Mã 1.1.NC1a, 5.2.NC1a): Sử dụng ứng dụng Bảng tuần hoàn tương tác (Periodic Table App / ptable.com) để tra cứu nhanh chóng và chính xác số hiệu nguyên tử, bán kính nguyên tử, độ âm điện, cấu hình electron và phân loại nguyên tố (s, p, d, f; kim loại, phi kim, khí hiếm).\n- An toàn mạng ([ATM1, ATM4]): Nhận diện và chọn lọc nguồn thông tin/ứng dụng học tập số uy tín, an toàn; cảnh giác với các liên kết quảng cáo hoặc mã độc khi cài đặt phần mềm từ kho ứng dụng.",
            "objective_qualities": "- Chăm chỉ: Chủ động thực hiện các bài tập tra cứu bảng tuần hoàn số.\n- Trách nhiệm: Tham gia tích cực vào các hoạt động nhóm trực tuyến.",
            "equipment_teacher": "- Máy chiếu, máy tính kết nối Internet truy cập ptable.com hoặc ứng dụng BTH tương tác.\n- Phiếu học tập số 1, 2.",
            "equipment_student": "- Thiết bị thông minh có cài ứng dụng Bảng tuần hoàn, SGK Hóa học 10.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (5 phút)",
                    "goal": "Tạo nhu cầu phân loại và sắp xếp các nguyên tố hóa học một cách khoa học.",
                    "content": "GV đưa ra danh sách 20 nguyên tố hóa học lộn xộn và đặt thử thách: Hãy sắp xếp các nguyên tố này sao cho những nguyên tố có tính chất tương tự nhau nằm cùng một cột.",
                    "product": "Các phương án sắp xếp sơ khởi của học sinh.",
                    "implementation_gv": "GV nêu câu hỏi thử thách, kích thích nhu cầu tìm hiểu bảng tuần hoàn của Mendeleev.",
                    "implementation_hs": "HS làm việc cá nhân 2 phút rồi thảo luận nhanh với bạn bên cạnh."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (30 phút)",
                    "goal": "Tìm hiểu 3 nguyên tắc sắp xếp và cấu tạo ô, chu kì, nhóm qua ứng dụng số.",
                    "content": "HS sử dụng ứng dụng Bảng tuần hoàn số (ptable.com) để tra cứu thông tin trong ô nguyên tố, tìm hiểu sự thay đổi số lớp e qua các chu kì và số e hóa trị trong các nhóm A.",
                    "product": "Bản báo cáo phiếu học tập đã hoàn thiện của các nhóm.",
                    "implementation_gv": "1. Chuyển giao: GV chiếu giao diện ptable.com, hướng dẫn tính năng lọc nguyên tố theo khối s, p, d, f.\n2. Thực hiện: HS thực hành tra cứu, đối chiếu cấu hình electron với vị trí ô, chu kì, nhóm.\n3. Báo cáo: Đại diện nhóm chia sẻ màn hình kết quả tra cứu.\n4. Kết luận: GV chốt 3 nguyên tắc sắp xếp và quy tắc xác định vị trí.",
                    "implementation_hs": "HS tương tác trực tiếp với ứng dụng số, ghi nhận kết quả vào phiếu học tập."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Xác định vị trí nguyên tố trong bảng tuần hoàn từ số hiệu Z.",
                    "content": "Xác định vị trí (ô, chu kì, nhóm) của các nguyên tố có Z = 12, 16, 19 trong bảng tuần hoàn.",
                    "product": "Lời giải của học sinh trên bảng và trong vở.",
                    "implementation_gv": "GV gọi 3 học sinh lên bảng làm 3 câu, hướng dẫn nhận xét.",
                    "implementation_hs": "HS làm bài tập vào vở và nhận xét bài của bạn trên bảng."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Tìm hiểu câu chuyện phát minh bảng tuần hoàn.",
                    "content": "Sưu tầm bài viết hoặc video ngắn từ nguồn tin cậy về giấc mơ thấy bảng tuần hoàn của Mendeleev.",
                    "product": "Bài chia sẻ ngắn kèm đường dẫn uy tín ([ATM1]).",
                    "implementation_gv": "Giao nhiệm vụ về nhà.",
                    "implementation_hs": "Ghi nhận bài tập về nhà."
                }
            ],
            "appendix": "PHIẾU HỌC TẬP:\n1. Số thứ tự của ô nguyên tố bằng đại lượng nào?\n2. Vì sao chu kì 1 chỉ có 2 nguyên tố, chu kì 2 và 3 có 8 nguyên tố?\n3. Cho nguyên tử X có cấu hình: 1s2 2s2 2p6 3s2 3p4. Xác định ô, chu kì, nhóm của X."
        },

        "hoa10_bai12": {
            "id": "hoa10_bai12",
            "grade": "10",
            "lesson_title": "BÀI 12: LIÊN KẾT CỘNG HÓA TRỊ",
            "duration": "3 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Trình bày được khái niệm liên kết cộng hóa trị (đơn, đôi, ba), viết được công thức electron và công thức Lewis.\n- Phân biệt được liên kết cộng hóa trị không phân cực và phân cực dựa vào hiệu độ âm điện.\n- Giải thích được sự hình thành liên kết sigma (σ) và liên kết pi (π) qua sự xen phủ orbital nguyên tử.",
            "objective_competence_general": "- Tự chủ, hợp tác và giải quyết vấn đề trong nhóm học tập.",
            "objective_competence_subject": "- Nhận thức hóa học: Vẽ đúng công thức Lewis và hình dung cấu trúc không gian của các phân tử HCl, O2, N2, CO2, H2O, NH4+.",
            "objective_competence_digital_safety": "- Năng lực số (Mã 3.1.NC1a, 3.2.NC1a): Sử dụng phần mềm vẽ cấu trúc hóa học chuyên dụng (KingDraw / ChemSketch / MolView) để thiết kế công thức Lewis, công thức cấu tạo và quan sát mô hình cấu trúc phân tử dạng 3D (mô hình rỗng, mô hình đặc).\n- An toàn mạng ([ATM6]): Tôn trọng quyền sở hữu trí tuệ, ghi rõ nguồn xuất xứ của các mô hình phân tử 3D và không sử dụng các bản bẻ khóa (crack) phần mềm vẽ hóa học.",
            "objective_qualities": "- Cẩn thận, chính xác trong biểu diễn công thức hóa học.",
            "equipment_teacher": "- Máy tính có cài đặt phần mềm KingDraw hoặc trình duyệt mở trang web MolView (molview.org).",
            "equipment_student": "- SGK, thiết bị di động cài đặt ứng dụng KingDraw.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Kích thích hứng thú về cách các nguyên tử phi kim liên kết với nhau.",
                    "content": "Vì sao hai nguyên tử Chlorine (Cl) đều thiếu 1 electron nhưng lại có thể kết hợp với nhau thành phân tử Cl2 bền vững?",
                    "product": "Dự đoán của HS: dùng chung cặp electron.",
                    "implementation_gv": "Nêu tình huống xuất phát từ quy tắc octet.",
                    "implementation_hs": "Thảo luận và trả lời."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Viết công thức Lewis và dựng mô hình 3D trên KingDraw/MolView.",
                    "content": "HS sử dụng ứng dụng KingDraw hoặc MolView để vẽ cấu trúc Lewis và chuyển sang chế độ xem không gian 3D của các phân tử H2O, NH3, CH4.",
                    "product": "Ảnh chụp màn hình cấu trúc 3D của các phân tử do HS dựng.",
                    "implementation_gv": "GV hướng dẫn các thao tác cơ bản trên thanh công cụ MolView: chọn nguyên tố, nối liên kết đơn/đôi, bấm nút 3D.",
                    "implementation_hs": "HS tự tay thiết kế phân tử, xoay mô hình 3D để quan sát góc liên kết."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Phân loại liên kết dựa vào hiệu độ âm điện.",
                    "content": "Tính hiệu độ âm điện và xác định loại liên kết trong các phân tử: Na2O, AlCl3, HCl, N2.",
                    "product": "Bảng phân loại liên kết hoàn chỉnh.",
                    "implementation_gv": "GV nhận xét và chuẩn hóa.",
                    "implementation_hs": "Làm bài tập vào vở."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Xuất tệp hình ảnh công thức hóa học chuẩn.",
                    "content": "Xuất hình ảnh công thức cấu tạo của phân tử CO2 và chèn vào tệp tài liệu bài tập cá nhân kèm nguồn phần mềm ([ATM6]).",
                    "product": "Tệp bài làm có chèn hình ảnh công thức số hóa.",
                    "implementation_gv": "Hướng dẫn kĩ thuật lưu ảnh PNG trong suốt.",
                    "implementation_hs": "Thực hiện ở nhà."
                }
            ],
            "appendix": "BẢNG HIỆU ĐỘ ÂM ĐIỆN VÀ LOẠI LIÊN KẾT:\n- 0 <= Delta Chi < 0.4: LK CHT không phân cực\n- 0.4 <= Delta Chi < 1.7: LK CHT phân cực\n- Delta Chi >= 1.7: Liên kết ion"
        },

        "hoa10_bai15": {
            "id": "hoa10_bai15",
            "grade": "10",
            "lesson_title": "BÀI 15: PHẢN ỨNG OXI HÓA – KHỬ",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Nêu được khái niệm số oxi hóa và quy tắc xác định số oxi hóa của nguyên tố trong hợp chất và ion.\n- Trình bày được định nghĩa chất khử, chất oxi hóa, sự khử, sự oxi hóa và phản ứng oxi hóa - khử.\n- Cân bằng được phương pháp thăng bằng electron cho các phản ứng oxi hóa - khử đơn giản và phức tạp.",
            "objective_competence_general": "- Hợp tác nhóm số, giao tiếp văn minh trên môi trường trực tuyến.",
            "objective_competence_subject": "- Nhận thức hóa học: Nhận diện vai trò các chất trong phản ứng oxi hóa - khử đời sống (cháy, gỉ sét, quang hợp, pin).",
            "objective_competence_digital_safety": "- Năng lực số (Mã 3.1.NC1a, 2.3.NC1a): Tổ chức làm việc nhóm trực tuyến: sử dụng phần mềm thiết kế đồ họa số (Canva / PowerPoint) để tạo sản phẩm số Infographic tóm tắt 4 bước cân bằng phản ứng oxi hóa - khử bằng phương pháp thăng bằng electron và sưu tầm hình ảnh ứng dụng thực tế.\n- An toàn mạng ([ATM2, ATM6, ATM8]): Ứng xử văn minh, tôn trọng ý kiến đồng đội khi thảo luận nhóm trên nhóm Zalo/nền tảng học tập số; ghi rõ nguồn bản quyền hình ảnh thực tế sưu tầm; hỗ trợ bạn bè xử lý sự cố kĩ thuật khi nộp bài số.",
            "objective_qualities": "- Chăm chỉ, có tinh thần trách nhiệm tập thể.",
            "equipment_teacher": "- Máy tính, máy chiếu, biểu mẫu giao nhiệm vụ dự án nhóm trên Canva.",
            "equipment_student": "- Điện thoại/máy tính có kết nối Internet.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Nhận biết các hiện tượng oxi hóa - khử quanh ta.",
                    "content": "Quan sát video đinh sắt bị gỉ, sự cháy của than củi, sự đổi màu quả táo sau khi gọt. Điểm chung của các quá trình này là gì?",
                    "product": "Phát hiện của HS: đều có sự tương tác với oxygen và thay đổi bản chất chất.",
                    "implementation_gv": "Chiếu video trực quan.",
                    "implementation_hs": "Quan sát và nêu ý kiến."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Nắm vững quy tắc xác định số oxi hóa và 4 bước cân bằng thăng bằng e.",
                    "content": "Nghiên cứu SGK, thảo luận nhóm thống nhất 4 bước cân bằng: Xác định số oxi hóa -> Viết quá trình nhường/nhận e -> Tìm hệ số -> Điền hệ số và kiểm tra.",
                    "product": "Bản nháp 4 bước cân bằng phản ứng Cu + HNO3 loãng.",
                    "implementation_gv": "GV làm mẫu 1 phương trình, sau đó chia nhóm thiết kế Infographic tóm tắt quy trình.",
                    "implementation_hs": "HS thảo luận nhóm, phân công tạo sản phẩm Infographic trên Canva."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Luyện tập cân bằng nhanh phản ứng oxi hóa khử.",
                    "content": "Cân bằng các phản ứng: Fe + H2SO4 đặc nóng; Al + HNO3 tạo N2O.",
                    "product": "Lời giải chi tiết của học sinh.",
                    "implementation_gv": "Theo dõi, sửa lỗi xác định sai số oxi hóa.",
                    "implementation_hs": "Thực hiện bài tập cá nhân."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Nộp sản phẩm Infographic nhóm lên không gian lưu trữ chung an toàn.",
                    "content": "Xuất file Infographic định dạng PDF/PNG, nộp bài vào thư mục Google Drive của lớp với phân quyền an toàn ([ATM4]).",
                    "product": "Liên kết sản phẩm số Infographic đã chia sẻ quyền xem hợp lệ.",
                    "implementation_gv": "Nhắc nhở HS không chia sẻ công khai dữ liệu cá nhân (SĐT, địa chỉ) trên bài tập nộp.",
                    "implementation_hs": "Nộp bài đúng hạn."
                }
            ],
            "appendix": "4 BƯỚC CÂN BẰNG PHẢN ỨNG OXI HÓA - KHỬ:\n1. Xác định số oxi hóa của các nguyên tố thay đổi.\n2. Viết quá trình oxi hóa và quá trình khử.\n3. Tìm hệ số thích hợp sao cho tổng e nhường = tổng e nhận.\n4. Đặt hệ số vào phương trình và đếm kiểm tra theo thứ tự: Kim loại -> Phi kim -> Hydro -> Oxygen."
        },

        # ==================== LỚP 11 ====================
        "hoa11_bai1": {
            "id": "hoa11_bai1",
            "grade": "11",
            "lesson_title": "BÀI 1: KHÁI NIỆM VỀ CÂN BẰNG HÓA HỌC",
            "duration": "4 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Trình bày được khái niệm phản ứng thuận nghịch và trạng thái cân bằng hóa học.\n- Viết được biểu thức hằng số cân bằng Kc của phản ứng thuận nghịch.\n- Vận dụng nguyên lí chuyển dịch cân bằng Le Chatelier để dự đoán chiều chuyển dịch cân bằng khi thay đổi nhiệt độ, nồng độ, áp suất.",
            "objective_competence_general": "- Năng lực tự học, hợp tác và giải quyết vấn đề bằng công nghệ.",
            "objective_competence_subject": "- Nhận thức hóa học: Hiểu sâu sắc bản chất cân bằng động (vt = vn > 0).",
            "objective_competence_digital_safety": "- Năng lực số (Mã 5.2.NC1a, 3.4.NC1a): Sử dụng mô phỏng số (PhET / Thí nghiệm ảo) quan sát sự chuyển dịch cân bằng của hệ khí N2O4 (không màu) ⇌ 2NO2 (nâu đỏ); ứng dụng phần mềm bảng tính MS Excel/Google Sheets lập công thức tự động tính hằng số cân bằng KC và nồng độ cân bằng các chất.\n- An toàn mạng ([ATM4, ATM6]): Tuân thủ bảo mật tài khoản học tập trực tuyến khi chia sẻ bảng tính tính toán KC; lưu trữ và phân quyền tệp dữ liệu bảng tính chỉ trong nội bộ nhóm học tập; tôn trọng bản quyền phần mềm bảng tính.",
            "objective_qualities": "- Chăm chỉ, cẩn thận, yêu thích khám phá khoa học.",
            "equipment_teacher": "- Máy tính kết nối Internet mở mô phỏng PhET Reversible Reactions.\n- Tệp bảng tính Excel mẫu tính hằng số KC.",
            "equipment_student": "- SGK Hóa học 11, máy tính xách tay/điện thoại thông minh có ứng dụng bảng tính.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Phân biệt phản ứng một chiều và phản ứng thuận nghịch.",
                    "content": "Quan sát hiện tượng đốt cháy củi (phản ứng 1 chiều) và quá trình tạo thạch nhũ trong hang động (thuận nghịch).",
                    "product": "Phát biểu ban đầu của học sinh về phản ứng xảy ra theo 2 chiều trái ngược.",
                    "implementation_gv": "Chiếu video hang động Phong Nha - Kẻ Bàng.",
                    "implementation_hs": "Quan sát và trả lời câu hỏi."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Tìm hiểu đặc điểm cân bằng động, hằng số Kc và nguyên lí Le Chatelier qua mô phỏng PhET.",
                    "content": "HS thao tác trên mô phỏng PhET: thay đổi nhiệt độ bình chứa NO2/N2O4, quan sát màu sắc thay đổi. Mở tệp Excel nhập số liệu nồng độ ban đầu và nồng độ cân bằng để tự động tính KC.",
                    "product": "Bảng tính Excel chứa công thức tính KC và đồ thị biểu diễn tốc độ phản ứng vt, vn theo thời gian.",
                    "implementation_gv": "GV hướng dẫn HS thiết lập công thức trong ô Excel: =C2/(A2*B2^3). Nhắc nhở quy tắc phân quyền bảng tính chia sẻ ([ATM4]).",
                    "implementation_hs": "HS thực hành nhập công thức bảng tính, đối chiếu kết quả với bạn."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Dự đoán chiều chuyển dịch cân bằng trong tổng hợp ammonia (quá trình Haber - Bosch).",
                    "content": "N2(g) + 3H2(g) <=> 2NH3(g) (Delta H < 0). Cần điều chỉnh áp suất và nhiệt độ như thế nào để tăng hiệu suất tạo NH3?",
                    "product": "Kết luận: Cần tăng áp suất và hạ nhiệt độ phù hợp.",
                    "implementation_gv": "GV hướng dẫn HS lập luận theo nguyên lí Le Chatelier.",
                    "implementation_hs": "HS suy luận và ghi chép."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Giải thích vai trò của cân bằng hóa học trong sản xuất công nghiệp và đời sống.",
                    "content": "Tại sao khi mở nắp chai nước ngọt có ga lại thấy bọt khí sủi lên?",
                    "product": "Giải thích cân bằng khí CO2 tan trong nước ngọt phụ thuộc áp suất.",
                    "implementation_gv": "GV đặt câu hỏi mở rộng.",
                    "implementation_hs": "Ghi chép câu hỏi và hoàn thành ở nhà."
                }
            ],
            "appendix": "PHIẾU BÀI TẬP: Cho cân bằng: CaCO3(s) <=> CaO(s) + CO2(g) (Delta H > 0). Cân bằng chuyển dịch theo chiều nào khi: a) Thêm CaCO3? b) Tăng nhiệt độ? c) Tăng áp suất?"
        },

        # ==================== LỚP 12 ====================
        "hoa12_bai1": {
            "id": "hoa12_bai1",
            "grade": "12",
            "lesson_title": "BÀI 1: ESTER VÀ LIPID",
            "duration": "3 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Nêu được khái niệm, đặc điểm cấu tạo phân tử của ester và chất béo (triglyceride).\n- Trình bày được tính chất vật lí (mùi thơm hoa quả, tính tan) và tính chất hóa học đặc trưng (thủy phân trong môi trường acid và base - xà phòng hóa).\n- Nêu được ứng dụng của ester và chất béo trong công nghiệp thực phẩm, mỹ phẩm và đời sống.",
            "objective_competence_general": "- Tự chủ, giao tiếp, hợp tác và ứng dụng kiến thức vào thực tế đời sống.",
            "objective_competence_subject": "- Nhận thức hóa học: Viết đúng phương trình phản ứng este hóa và xà phòng hóa chất béo.\n- Vận dụng kiến thức: Giải thích quy trình sản xuất xà phòng và sử dụng dầu mỡ hợp lí.",
            "objective_competence_digital_safety": "- Năng lực số (Mã 1.1.NC1a, 1.2.NC1a): Tra cứu thông tin từ các nguồn tài nguyên số uy tín (trang web của Bộ Y tế, Viện Dinh dưỡng, PubMed, Wikipedia tiếng Việt/Anh có kiểm chứng) về thành phần dinh dưỡng, vai trò của chất béo omega-3, omega-6 đối với sức khỏe và ứng dụng của ester trong đời sống.\n- An toàn mạng ([ATM1]): Nhận diện và đánh giá độ tin cậy của thông tin sức khỏe trên mạng xã hội; phân biệt được thông tin khoa học chính thống với các quảng cáo thực phẩm chức năng thổi phồng, tin giả mạo về giảm cân hoặc mẹo chữa bệnh phản khoa học.",
            "objective_qualities": "- Chăm chỉ, có ý thức an toàn thực phẩm và bảo vệ sức khỏe cho bản thân và gia đình.",
            "equipment_teacher": "- Máy tính, máy chiếu, bài trình chiếu có liên kết đến các trang kiểm chứng thông tin y tế.",
            "equipment_student": "- SGK Hóa học 12, điện thoại thông minh để tra cứu tài liệu.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Kích thích hứng thú về mùi thơm tự nhiên của hoa quả và nước hoa.",
                    "content": "Cho HS ngửi mẫu mùi chuối chín nhân tạo và đoán xem chất gì tạo nên mùi thơm hấp dẫn đó?",
                    "product": "Dự đoán của học sinh về este/hương liệu.",
                    "implementation_gv": "GV cho HS trải nghiệm giác quan và giới thiệu vào bài học.",
                    "implementation_hs": "HS ngửi mẫu, nêu cảm nhận."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Nắm vững định nghĩa, danh pháp, phản ứng thủy phân của ester và lipid.",
                    "content": "Nghiên cứu SGK, so sánh phản ứng thủy phân ester trong môi trường acid (thuận nghịch) và môi trường kiềm (một chiều - xà phòng hóa). Tra cứu nguồn tin y tế về trans-fat (chất béo chuyển hóa).",
                    "product": "Phiếu học tập hoàn thiện phương trình phản ứng hóa học và bảng tổng kết về chất béo tốt / chất béo xấu.",
                    "implementation_gv": "GV hướng dẫn HS quy trình SIFT để kiểm chứng bài viết về 'dầu ăn chiên đi chiên lại' trên mạng xã hội ([ATM1]).",
                    "implementation_hs": "HS áp dụng quy trình kiểm chứng nguồn tin số, thảo luận và hoàn thiện phiếu."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Củng cố kĩ năng viết đồng phân và phương trình thủy phân ester.",
                    "content": "Viết các đồng phân ester có công thức phân tử C4H8O2 và gọi tên tương ứng.",
                    "product": "4 đồng phân: HCOOCH2CH2CH3, HCOOCH(CH3)2, CH3COOC2H5, C2H5COOCH3.",
                    "implementation_gv": "GV chiếu đáp án chuẩn, giải đáp thắc mắc.",
                    "implementation_hs": "HS hoàn thành bài tập vào vở."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Hiểu tác hại của dầu mỡ chiên nhiều lần đối với sức khỏe.",
                    "content": "Tìm hiểu dự án tái chế dầu ăn thừa thành xà phòng hoặc dầu biodiesel, viết thông điệp truyền thông chia sẻ với gia đình.",
                    "product": "Bản ghi chú thông điệp an toàn dinh dưỡng.",
                    "implementation_gv": "Giao nhiệm vụ về nhà.",
                    "implementation_hs": "Nhận nhiệm vụ và lập kế hoạch thực hiện tại nhà."
                }
            ],
            "appendix": "PHIẾU BÀI TẬP:\n1. Thủy phân hoàn toàn 8,8g ethyl acetate bằng dung dịch NaOH vừa đủ. Tính khối lượng muối thu được.\n2. So sánh nhiệt độ sôi của ester so với alcohol và carboxylic acid có cùng số C."
        }
    }

    @staticmethod
    def get_lesson_list():
        """Lấy danh mục các bài học mẫu chuẩn 5512 tích hợp NLS và ATM"""
        res = []
        for key, val in KHBDProcessor.LESSON_DATABASE.items():
            res.append({
                "id": key,
                "grade": val.get("grade", "10"),
                "title": val.get("lesson_title", "")
            })
        return res

    @staticmethod
    def get_lesson(lesson_id):
        """Lấy chi tiết một bài học từ cơ sở dữ liệu mẫu"""
        if lesson_id in KHBDProcessor.LESSON_DATABASE:
            return KHBDProcessor.LESSON_DATABASE[lesson_id]
        return KHBDProcessor.LESSON_DATABASE["hoa10_bai3"]

    @staticmethod
    def generate_smart_outline(lesson_title, grade="10", subject="HÓA HỌC"):
        """Tự động sinh bộ khung giáo án chuẩn 5512 tích hợp NLS & ATM cho bài học mới bất kì"""
        clean_name = lesson_title.upper().strip()
        
        # Gợi ý thông minh cho NLS và ATM dựa trên nội dung bài học
        digital_suggestion = (
            f"- Năng lực số (Theo CV 3456/BGDĐT-GDPT: Mã 1.1.NC1a, 5.2.NC1a, 3.1.NC1a): "
            f"Khai thác và sử dụng các công cụ học tập số (mô phỏng thí nghiệm ảo PhET, phần mềm vẽ cấu trúc hóa học KingDraw/MolView, hoặc bảng tính Excel) "
            f"để mô hình hóa kiến thức, thu thập và xử lý số liệu của bài học {clean_name}.\n"
            f"- An toàn mạng (Theo QĐ 6206/BGDĐT-GDPT: Mã [ATM1, ATM4, ATM6]): "
            f"Kiểm chứng độ tin cậy của thông tin và học liệu số trước khi tiếp nhận; "
            f"tuân thủ an toàn thiết bị, bảo mật tài khoản cá nhân; tôn trọng bản quyền khi trích dẫn tài liệu, hình ảnh mô phỏng trong sản phẩm học tập."
        )

        return {
            "school_name": "TRƯỜNG THPT HUỲNH THÚC KHÁNG",
            "department": "TỔ HÓA HỌC",
            "teacher_name": "NGUYỄN HOÀI BẢO",
            "lesson_title": clean_name,
            "subject": subject,
            "grade": str(grade),
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": f"- Nêu được các khái niệm, định nghĩa và nội dung trọng tâm của {clean_name}.\n- Trình bày được tính chất, quy luật hoặc cấu tạo liên quan đến nội dung bài học.\n- Viết được các phương trình hóa học minh họa hoặc thực hiện được các bước giải bài tập liên quan.",
            "objective_competence_general": "- Tự chủ và tự học: Chủ động đọc SGK, hoàn thành phiếu học tập cá nhân.\n- Giao tiếp và hợp tác: Tích cực thảo luận nhóm, phân công nhiệm vụ và báo cáo kết quả.\n- Giải quyết vấn đề và sáng tạo: Vận dụng kiến thức bài học để giải quyết các tình huống thực tiễn.",
            "objective_competence_subject": f"- Nhận thức hóa học: Hiểu rõ bản chất các hiện tượng, phản ứng và quy luật trong {clean_name}.\n- Tìm hiểu thế giới tự nhiên dưới góc độ hóa học.\n- Vận dụng kiến thức, kĩ năng đã học vào đời sống và bảo vệ môi trường.",
            "objective_competence_digital_safety": digital_suggestion,
            "objective_qualities": "- Chăm chỉ, trung thực trong học tập và báo cáo kết quả.\n- Có tinh thần trách nhiệm với công việc chung của tập thể.",
            "equipment_teacher": f"- Máy tính, máy chiếu, bài giảng điện tử có tích hợp liên kết mô phỏng thí nghiệm ảo/học liệu số cho bài học {clean_name}.\n- Phiếu học tập số 1, số 2 và dụng cụ/hóa chất thí nghiệm (nếu có).",
            "equipment_student": f"- Sách giáo khoa {subject} {grade}, vở ghi chép, thiết bị thông minh có kết nối mạng phục vụ tra cứu có hướng dẫn.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (5 - 7 phút)",
                    "goal": f"Tạo hứng thú, kích thích trí tò mò của học sinh về nội dung {clean_name}.",
                    "content": "GV đưa ra hình ảnh/video thực tế hoặc câu hỏi tình huống mâu thuẫn nhận thức để học sinh suy ngẫm.",
                    "product": "Câu trả lời, dự đoán hoặc thắc mắc của học sinh dẫn dắt vào bài mới.",
                    "implementation_gv": "1. Chuyển giao: GV nêu tình huống gợi mở.\n2. Thực hiện: HS suy nghĩ cá nhân hoặc trao đổi cặp đôi.\n3. Báo cáo: Gọi 1-2 HS phát biểu.\n4. Kết luận: GV ghi nhận và giới thiệu vào bài mới.",
                    "implementation_hs": "HS quan sát, lắng nghe và tích cực phát biểu xây dựng bài."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (28 - 30 phút)",
                    "goal": f"Chiếm lĩnh kiến thức trọng tâm của {clean_name} kết hợp ứng dụng công cụ số an toàn.",
                    "content": "HS làm việc với SGK và Phiếu học tập số 1 theo nhóm; khai thác học liệu số/mô phỏng thí nghiệm dưới sự định hướng của giáo viên.",
                    "product": "Sản phẩm hoàn thành phiếu học tập hoặc bảng phụ của các nhóm có trích dẫn nguồn số an toàn.",
                    "implementation_gv": "1. Chuyển giao: Chia nhóm, cung cấp đường link tài nguyên số đã kiểm chứng an toàn ([ATM1, ATM4]).\n2. Thực hiện: Theo dõi, hỗ trợ các nhóm thực hành.\n3. Báo cáo: Đại diện nhóm trình bày sản phẩm, nhóm khác nhận xét.\n4. Kết luận: GV chốt kiến thức chuẩn xác vào vở.",
                    "implementation_hs": "HS tích cực thảo luận nhóm, sử dụng công cụ số văn minh, ghi nhận kiến thức trọng tâm."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": f"Củng cố, khắc sâu kiến thức vừa học về {clean_name}.",
                    "content": "Học sinh giải quyết các câu hỏi trắc nghiệm hoặc bài tập vận dụng nhanh trên bảng/nền tảng học tập số.",
                    "product": "Lời giải của học sinh trên bảng con hoặc vở bài tập.",
                    "implementation_gv": "GV chiếu câu hỏi, quan sát và nhận xét đánh giá.",
                    "implementation_hs": "HS tích cực làm bài và đối chiếu đáp án."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG VÀ MỞ RỘNG (5 phút)",
                    "goal": "Liên hệ kiến thức bài học với ứng dụng thực tiễn trong đời sống và sản xuất.",
                    "content": "Giải quyết câu hỏi thực tế hoặc nhiệm vụ nghiên cứu nhỏ tại nhà, tra cứu thông tin trực tuyến có kiểm chứng độ tin cậy.",
                    "product": "Bản báo cáo ngắn gọn hoặc phần chuẩn bị cho tiết học sau có dẫn nguồn hợp lệ ([ATM6]).",
                    "implementation_gv": "GV giao nhiệm vụ về nhà và nhắc nhở các quy tắc bảo vệ an toàn thông tin cá nhân khi tương tác trực tuyến.",
                    "implementation_hs": "Ghi chép nhiệm vụ và thực hiện ở nhà."
                }
            ],
            "appendix": f"PHIẾU HỌC TẬP:\nNội dung các câu hỏi hướng dẫn tự học và phiếu bài tập phục vụ bài dạy {clean_name}."
        }

    @staticmethod
    def export_word(khbd_data, output_path):
        """Xuất file KHBD chuẩn Công văn 5512 ra định dạng docx tích hợp NLS và ATM"""
        doc = Document()

        for section in doc.sections:
            section.top_margin = Inches(0.79)
            section.bottom_margin = Inches(0.79)
            section.left_margin = Inches(1.18)
            section.right_margin = Inches(0.79)

        normal_style = doc.styles['Normal']
        normal_style.font.name = 'Times New Roman'
        normal_style.font.size = Pt(13)
        normal_style.paragraph_format.line_spacing = 1.2
        normal_style.paragraph_format.space_after = Pt(4)

        # 1. BẢNG TIÊU ĐỀ TRƯỜNG & TỔ CHUYÊN MÔN
        table_hdr = doc.add_table(rows=1, cols=2)
        table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_hdr.columns[0].width = Inches(3.2)
        table_hdr.columns[1].width = Inches(3.5)

        c_left = table_hdr.cell(0, 0)
        p_l = c_left.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.add_run(f"{khbd_data.get('school_name', 'TRƯỜNG THPT HUỲNH THÚC KHÁNG').upper()}\n").font.bold = True
        p_l.add_run(f"{khbd_data.get('department', 'TỔ HÓA HỌC').upper()}\n").font.bold = True
        p_l.add_run("--------------------")

        c_right = table_hdr.cell(0, 1)
        p_r = c_right.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n").font.bold = True
        p_r.add_run("Độc lập - Tự do - Hạnh phúc\n").font.bold = True
        p_r.add_run("------------------------------------")

        doc.add_paragraph()

        # 2. TÊN BÀI HỌC VÀ THÔNG TIN CHUNG
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = p_title.add_run("KẾ HOẠCH BÀI DẠY\n")
        r_title.font.bold = True
        r_title.font.size = Pt(15)
        r_title.font.color.rgb = RGBColor(0, 32, 96)

        r_sub = p_title.add_run("(TÍCH HỢP NĂNG LỰC SỐ THEO CV 3456 & AN TOÀN MẠNG THEO QĐ 6206)\n")
        r_sub.font.bold = True
        r_sub.font.size = Pt(11)
        r_sub.font.color.rgb = RGBColor(0, 128, 64)

        r_lesson = p_title.add_run(f"{khbd_data.get('lesson_title', 'TÊN BÀI DẠY').upper()}\n")
        r_lesson.font.bold = True
        r_lesson.font.size = Pt(14)
        
        info_str = f"Môn học: {khbd_data.get('subject', 'Hóa học')} - Lớp: {khbd_data.get('grade', '10')} | Thời lượng: {khbd_data.get('duration', '2 tiết')}\nGiáo viên thực hiện: {khbd_data.get('teacher_name', 'NGUYỄN HOÀI BẢO')}"
        r_info = p_title.add_run(info_str)
        r_info.font.italic = True
        r_info.font.size = Pt(12)

        doc.add_paragraph()

        # 3. MỤC TIÊU (I)
        h1 = doc.add_paragraph()
        h1.add_run("I. MỤC TIÊU").font.bold = True
        
        # 1. Kiến thức
        p_k = doc.add_paragraph()
        p_k.paragraph_format.left_indent = Inches(0.2)
        p_k.add_run("1. Về kiến thức:\n").font.bold = True
        p_k.add_run(khbd_data.get("objective_knowledge", ""))

        # 2. Năng lực (Gồm 2.1, 2.2, 2.3)
        p_c = doc.add_paragraph()
        p_c.paragraph_format.left_indent = Inches(0.2)
        p_c.add_run("2. Về năng lực:\n").font.bold = True
        
        p_cg = doc.add_paragraph()
        p_cg.paragraph_format.left_indent = Inches(0.4)
        p_cg.add_run("2.1. Năng lực chung (Tự chủ, Giao tiếp - Hợp tác, Giải quyết vấn đề):\n").font.bold = True
        p_cg.add_run(khbd_data.get("objective_competence_general", ""))

        p_cs = doc.add_paragraph()
        p_cs.paragraph_format.left_indent = Inches(0.4)
        p_cs.add_run("2.2. Năng lực đặc thù (Hóa học):\n").font.bold = True
        p_cs.add_run(khbd_data.get("objective_competence_subject", ""))

        # 2.3. Năng lực số và An toàn mạng (Yêu cầu bổ sung đặc biệt)
        digital_safety_text = khbd_data.get("objective_competence_digital_safety", "").strip()
        if digital_safety_text:
            p_cd = doc.add_paragraph()
            p_cd.paragraph_format.left_indent = Inches(0.4)
            p_cd.add_run("2.3. Năng lực số và An toàn mạng (Theo CV 3456/BGDĐT-GDPT & QĐ 6206/BGDĐT-GDPT):\n").font.bold = True
            p_cd.add_run(digital_safety_text)

        # 3. Phẩm chất
        p_q = doc.add_paragraph()
        p_q.paragraph_format.left_indent = Inches(0.2)
        p_q.add_run("3. Về phẩm chất:\n").font.bold = True
        p_q.add_run(khbd_data.get("objective_qualities", ""))

        # 4. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (II)
        h2 = doc.add_paragraph()
        h2.paragraph_format.space_before = Pt(8)
        h2.add_run("II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU").font.bold = True

        p_eq = doc.add_paragraph()
        p_eq.paragraph_format.left_indent = Inches(0.2)
        p_eq.add_run("1. Giáo viên:\n").font.bold = True
        p_eq.add_run(khbd_data.get("equipment_teacher", ""))

        p_es = doc.add_paragraph()
        p_es.paragraph_format.left_indent = Inches(0.2)
        p_es.add_run("2. Học sinh:\n").font.bold = True
        p_es.add_run(khbd_data.get("equipment_student", ""))

        # 5. TIẾN TRÌNH DẠY HỌC (III)
        h3 = doc.add_paragraph()
        h3.paragraph_format.space_before = Pt(8)
        h3.add_run("III. TIẾN TRÌNH DẠY HỌC").font.bold = True

        for act in khbd_data.get("activities", []):
            p_act_title = doc.add_paragraph()
            p_act_title.paragraph_format.space_before = Pt(6)
            r_at = p_act_title.add_run(act.get("title", "HOẠT ĐỘNG"))
            r_at.font.bold = True
            r_at.font.color.rgb = RGBColor(0, 32, 96)

            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Inches(0.2)
            p_sub.add_run("a) Mục tiêu: ").font.bold = True
            p_sub.add_run(f"{act.get('goal', '')}\n")
            p_sub.add_run("b) Nội dung: ").font.bold = True
            p_sub.add_run(f"{act.get('content', '')}\n")
            p_sub.add_run("c) Sản phẩm: ").font.bold = True
            p_sub.add_run(f"{act.get('product', '')}\n")
            p_sub.add_run("d) Tổ chức thực hiện:").font.bold = True

            tbl_act = doc.add_table(rows=2, cols=2)
            tbl_act.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl_act.columns[0].width = Inches(3.8)
            tbl_act.columns[1].width = Inches(2.9)

            cell_h1 = tbl_act.cell(0, 0)
            cell_h1.text = "HOẠT ĐỘNG CỦA GV VÀ HS"
            cell_h1.paragraphs[0].runs[0].font.bold = True
            cell_h1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_background(cell_h1, "F2F2F2")

            cell_h2 = tbl_act.cell(0, 1)
            cell_h2.text = "DỰ KIẾN SẢN PHẨM"
            cell_h2.paragraphs[0].runs[0].font.bold = True
            cell_h2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_background(cell_h2, "F2F2F2")

            cell_b1 = tbl_act.cell(1, 0)
            set_cell_margins(cell_b1, 80, 80, 100, 100)
            p_b1 = cell_b1.paragraphs[0]
            p_b1.paragraph_format.line_spacing = 1.15
            p_b1.add_run(act.get("implementation_gv", ""))

            cell_b2 = tbl_act.cell(1, 1)
            set_cell_margins(cell_b2, 80, 80, 100, 100)
            p_b2 = cell_b2.paragraphs[0]
            p_b2.paragraph_format.line_spacing = 1.15
            p_b2.add_run(act.get("implementation_hs", act.get("product", "")))

            doc.add_paragraph()

        # 6. PHỤ LỤC (IV)
        appendix_text = khbd_data.get("appendix", "").strip()
        if appendix_text:
            h4 = doc.add_paragraph()
            h4.paragraph_format.space_before = Pt(8)
            h4.add_run("IV. PHỤ LỤC / HỒ SƠ DẠY HỌC").font.bold = True
            p_app = doc.add_paragraph()
            p_app.paragraph_format.left_indent = Inches(0.2)
            p_app.add_run(appendix_text)

        # Chữ ký
        doc.add_paragraph()
        tbl_sign = doc.add_table(rows=1, cols=2)
        tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_sign.columns[0].width = Inches(3.5)
        tbl_sign.columns[1].width = Inches(3.2)

        c_s1 = tbl_sign.cell(0, 0)
        p_s1 = c_s1.paragraphs[0]
        p_s1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_s1.add_run("DUYỆT CỦA TỔ CHUYÊN MÔN\n").font.bold = True
        p_s1.add_run("(Ký và ghi rõ họ tên)\n\n\n\n")

        c_s2 = tbl_sign.cell(0, 1)
        p_s2 = c_s2.paragraphs[0]
        p_s2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_s2.add_run("GIÁO VIÊN SOẠN BÀI\n").font.bold = True
        p_s2.add_run("(Ký và ghi rõ họ tên)\n\n\n\n")
        p_s2.add_run(khbd_data.get("teacher_name", "NGUYỄN HOÀI BẢO")).font.bold = True

        doc.save(output_path)
        return output_path
