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
    Bổ sung nội dung chi tiết ở cột Dự kiến sản phẩm để học sinh viết bài vào vở
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
            "equipment_student": "- Sách giáo khoa Hóa học 10, vở ghi chép bài.\n- Điện thoại thông minh/máy tính bảng có kết nối Internet để truy cập mô phỏng PhET.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (7 phút)",
                    "goal": "Tạo mâu thuẫn nhận thức về sự phân bố và chuyển động của electron quanh hạt nhân nguyên tử.",
                    "content": "GV chiếu hình ảnh so sánh mô hình hành tinh nguyên tử của Rutherford - Bohr với mô hình đám mây xác suất hiện đại. HS trả lời câu hỏi: Electron chuyển động theo quỹ đạo cố định hay trong một vùng không gian xác suất?",
                    "product": "Câu trả lời của HS: nhận thức được sự giới hạn của mô hình cổ điển và tính ưu việt của mô hình hiện đại.",
                    "implementation_gv": "GV trình chiếu video ngắn (1 phút) mô phỏng chuyển động hỗn loạn của electron quanh hạt nhân. Nêu câu hỏi gợi mở.",
                    "implementation_hs": "* Câu trả lời dự kiến của HS:\n- Theo mô hình Rutherford - Bohr, electron chuyển động theo quỹ đạo tròn/bầu dục cố định xung quanh hạt nhân.\n- Theo mô hình hiện đại, electron chuyển động rất nhanh và không theo quỹ đạo xác định.\n* Vấn đề bài học:\nCần tìm hiểu mô hình hiện đại, bản chất orbital nguyên tử và cách viết cấu hình electron của nguyên tử."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (30 phút)",
                    "goal": "Khám phá hình dạng AO s, p qua mô phỏng 3D; quy tắc viết cấu hình e.",
                    "content": "HS làm việc nhóm: truy cập mô phỏng PhET trên máy tính bảng/điện thoại, xoay các góc nhìn 3D của AO s và AO p. Điền Phiếu học tập số 1 về số lượng e tối đa và quy tắc điền electron.",
                    "product": "Phiếu học tập số 1 hoàn chỉnh; viết đúng cấu hình e của 20 nguyên tố đầu tiên.",
                    "implementation_gv": "1. Chuyển giao: GV hướng dẫn HS quét mã QR truy cập trang mô phỏng PhET an toàn, không bấm vào liên kết lạ ([ATM4]). Giao Phiếu học tập số 1.\n2. Thực hiện: Các nhóm thao tác trên mô phỏng 3D, quan sát hình dạng AO s, AO p và thảo luận hoàn thành phiếu.\n3. Báo cáo: Đại diện nhóm trình chiếu kết quả trên màn hình lớp.\n4. Kết luận: GV nhận xét, chuẩn hóa kiến thức trọng tâm để HS ghi bài vào vở.",
                    "implementation_hs": "* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. MÔ HÌNH NGUYÊN TỬ VÀ ORBITAL NGUYÊN TỬ (AO)\n1. Mô hình hiện đại về sự chuyển động của electron trong nguyên tử:\n- Các electron chuyển động rất nhanh xung quanh hạt nhân, không theo những quỹ đạo xác định, tạo nên các đám mây electron.\n2. Orbital nguyên tử (AO):\n- Khái niệm: Là khu vực không gian xung quanh hạt nhân mà tại đó xác suất tìm thấy electron là lớn nhất (khoảng 90%).\n- Hình dạng AO:\n  + AO s có dạng hình cầu đối xứng trong không gian.\n  + AO p có dạng hình số 8 nổi, gồm 3 orbital px, py, pz định hướng theo 3 trục tọa độ vuông góc trong không gian (Ox, Oy, Oz).\n- Mỗi AO chứa tối đa 2 electron có chiều tự quay ngược nhau (spin ngược dấu ↑↓ - electron ghép đôi).\n\nII. LỚP VÀ PHÂN LỚP ELECTRON\n1. Lớp electron:\n- Gồm các electron có mức năng lượng gần bằng nhau.\n- Các lớp được đánh số từ trong ra ngoài: n = 1, 2, 3, 4, 5, 6, 7 tương ứng với các lớp K, L, M, N, O, P, Q.\n- Số electron tối đa trong lớp thứ n là 2n² (với n ≤ 4).\n2. Phân lớp electron:\n- Các lớp được chia thành các phân lớp kí hiệu bằng chữ cái thường: s, p, d, f.\n- Các electron trong cùng một phân lớp có mức năng lượng bằng nhau.\n- Số electron tối đa trong từng phân lớp: s² (tối đa 2e), p⁶ (tối đa 6e), d¹⁰ (tối đa 10e), f¹⁴ (tối đa 14e).\n\nIII. CẤU HÌNH ELECTRON NGUYÊN TỬ\n1. Các nguyên lí và quy tắc phân bố electron:\n- Nguyên lí vững bền: Ở trạng thái cơ bản, các electron phân bố vào các AO theo thứ tự mức năng lượng tăng dần: 1s 2s 2p 3s 3p 4s 3d 4p 5s...\n- Nguyên lí Pauli: Trên mỗi AO chỉ có tối đa 2 electron có spin ngược nhau (↑↓).\n- Quy tắc Hund: Trong cùng một phân lớp chưa bão hòa, các electron phân bố sao cho số electron độc thân là lớn nhất.\n2. Các bước viết cấu hình electron nguyên tử:\n- Bước 1: Xác định tổng số electron của nguyên tử (bằng số hiệu nguyên tử Z).\n- Bước 2: Phân bố electron theo thứ tự mức năng lượng tăng dần.\n- Bước 3: Viết cấu hình electron theo thứ tự từng lớp (từ trong ra ngoài: 1 đến n).\n3. Dự đoán tính chất hóa học cơ bản của nguyên tố:\n- Nguyên tử có 1, 2, 3 e ở lớp ngoài cùng: Thường là kim loại (trừ H, He, B).\n- Nguyên tử có 5, 6, 7 e ở lớp ngoài cùng: Thường là phi kim.\n- Nguyên tử có 8 e ở lớp ngoài cùng (hoặc 2e ở He): Là khí hiếm (cấu hình bền vững, trơ về mặt hóa học)."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (5 phút)",
                    "goal": "Rèn luyện kĩ năng viết cấu hình electron và xác định tính chất nguyên tố.",
                    "content": "Làm bài tập: viết cấu hình electron cho Na (Z=11), Cl (Z=17), Fe (Z=26).",
                    "product": "Kết quả làm bài tập vào vở của học sinh.",
                    "implementation_gv": "GV chiếu 3 bài tập vận dụng viết cấu hình e của Na (Z=11), Cl (Z=17), Fe (Z=26). Cho HS làm cá nhân vào vở, gọi 3 HS lên bảng làm. Sau đó chữa bài chi tiết.",
                    "implementation_hs": "* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\n1. Na (Z = 11):\n- Thứ tự phân bố: 1s² 2s² 2p⁶ 3s¹.\n- Cấu hình electron: 1s² 2s² 2p⁶ 3s¹ (viết gọn: [Ne]3s¹).\n- Nhận xét: Có 3 lớp electron; lớp ngoài cùng (lớp M) có 1e → Là nguyên tố kim loại.\n\n2. Cl (Z = 17):\n- Thứ tự phân bố: 1s² 2s² 2p⁶ 3s² 3p⁵.\n- Cấu hình electron: 1s² 2s² 2p⁶ 3s² 3p⁵ (viết gọn: [Ne]3s² 3p⁵).\n- Nhận xét: Có 3 lớp electron; lớp ngoài cùng có 7e → Là nguyên tố phi kim.\n\n3. Fe (Z = 26):\n- Thứ tự mức năng lượng: 1s² 2s² 2p⁶ 3s² 3p⁶ 4s² 3d⁶.\n- Cấu hình electron: 1s² 2s² 2p⁶ 3s² 3p⁶ 3d⁶ 4s² (viết gọn: [Ar]3d⁶ 4s²).\n- Nhận xét: Lớp ngoài cùng (lớp N) có 2e; là nguyên tố d → Là kim loại chuyển tiếp."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG VÀ MỞ RỘNG (3 phút)",
                    "goal": "Tìm hiểu ứng dụng của cấu hình electron bền vững của khí hiếm.",
                    "content": "Tìm kiếm thông tin trực tuyến có dẫn nguồn về ứng dụng của khí hiếm Argon trong hàn kim loại và đèn Neon quảng cáo.",
                    "product": "Bản ghi chép ngắn gọn có ghi rõ nguồn tài liệu số tham khảo ([ATM6]).",
                    "implementation_gv": "Giao nhiệm vụ vận dụng thực tiễn về nhà: Tìm hiểu ứng dụng cấu hình e bền vững của khí Argon (Ar) trong công nghệ hàn kim loại và bóng đèn điện. Nhắc nhở quy tắc trích dẫn nguồn số an toàn ([ATM1, ATM6]).",
                    "implementation_hs": "* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n1. Giải thích ứng dụng khí Argon (Ar, Z=18):\n- Cấu hình electron của Ar: 1s² 2s² 2p⁶ 3s² 3p⁶. Lớp ngoài cùng bão hòa 8e (bát thế) nên phân tử Ar trơ, cực kì bền vững về mặt hóa học.\n- Ứng dụng thực tiễn:\n  + Làm khí bảo vệ trong công nghệ hàn kim loại nóng chảy (hàn TIG/MIG) để kim loại không bị oxy hóa bởi O₂ trong không khí.\n  + Bơm vào bóng đèn sợi đốt để bảo vệ dây tóc wolfram ở nhiệt độ cao không bị oxy hóa và bay hơi.\n2. Nguồn tài liệu số kiểm chứng ([ATM1, ATM6]): Viện Hóa học Việt Nam - Viện Hàn lâm KH&CN (vast.gov.vn)."
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
            "equipment_student": "- Thiết bị thông minh có cài ứng dụng Bảng tuần hoàn, SGK Hóa học 10, vở ghi.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (5 phút)",
                    "goal": "Tạo nhu cầu phân loại và sắp xếp các nguyên tố hóa học một cách khoa học.",
                    "content": "GV đưa ra danh sách 20 nguyên tố hóa học lộn xộn và đặt thử thách: Hãy sắp xếp các nguyên tố này sao cho những nguyên tố có tính chất tương tự nhau nằm cùng một cột.",
                    "product": "Các phương án sắp xếp sơ khởi của học sinh.",
                    "implementation_gv": "GV nêu câu hỏi thử thách, kích thích nhu cầu tìm hiểu bảng tuần hoàn của Mendeleev.",
                    "implementation_hs": "* Câu trả lời dự kiến của HS:\n- Có thể sắp xếp các nguyên tố theo chiều tăng dần của khối lượng nguyên tử hoặc số proton (điện tích hạt nhân).\n- Những nguyên tố có tính chất tương tự nhau (như kim loại kiềm Li, Na, K) được xếp vào cùng một cột dọc để tiện theo dõi."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (30 phút)",
                    "goal": "Tìm hiểu 3 nguyên tắc sắp xếp và cấu tạo ô, chu kì, nhóm qua ứng dụng số.",
                    "content": "HS sử dụng ứng dụng Bảng tuần hoàn số (ptable.com) để tra cứu thông tin trong ô nguyên tố, tìm hiểu sự thay đổi số lớp e qua các chu kì và số e hóa trị trong các nhóm A.",
                    "product": "Bản báo cáo phiếu học tập đã hoàn thiện của các nhóm.",
                    "implementation_gv": "1. Chuyển giao: GV chiếu giao diện ptable.com, hướng dẫn tính năng lọc nguyên tố theo khối s, p, d, f.\n2. Thực hiện: HS thực hành tra cứu, đối chiếu cấu hình electron với vị trí ô, chu kì, nhóm.\n3. Báo cáo: Đại diện nhóm chia sẻ màn hình kết quả tra cứu.\n4. Kết luận: GV chuẩn hóa 3 nguyên tắc sắp xếp và quy tắc xác định vị trí để HS ghi chép.",
                    "implementation_hs": "* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. NGUYÊN TẮC SẮP XẾP CÁC NGUYÊN TỐ TRONG BẢNG TUẦN HOÀN\n1. Các nguyên tố được sắp xếp theo chiều tăng dần của điện tích hạt nhân nguyên tử.\n2. Các nguyên tố có cùng số lớp electron trong nguyên tử được xếp vào cùng một hàng (chu kì).\n3. Các nguyên tố có số electron hóa trị tương tự nhau được xếp vào cùng một cột (nhóm).\n\nII. CẤU TẠO CỦA BẢNG TUẦN HOÀN\n1. Ô nguyên tố:\n- Mỗi nguyên tố chiếm một ô trong bảng tuần hoàn.\n- Số thứ tự ô = Số hiệu nguyên tử (Z) = Số proton = Số electron.\n- Cho biết: Kí hiệu hóa học, tên nguyên tố, số hiệu Z, khối lượng nguyên tử, độ âm điện, cấu hình electron.\n2. Chu kì:\n- Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được sắp xếp theo chiều điện tích hạt nhân tăng dần.\n- Số thứ tự chu kì = Số lớp electron trong nguyên tử.\n- Bảng tuần hoàn gồm 7 chu kì (đánh số từ 1 đến 7):\n  + Chu kì nhỏ: Chu kì 1 (2 nguyên tố), chu kì 2 (8 nguyên tố), chu kì 3 (8 nguyên tố).\n  + Chu kì lớn: Chu kì 4 (18 nguyên tố), chu kì 5 (18 nguyên tố), chu kì 6 (32 nguyên tố), chu kì 7 (chưa hoàn tất).\n3. Nhóm nguyên tố:\n- Nhóm là tập hợp các nguyên tố có cấu hình electron tương tự nhau, do đó có tính chất hóa học gần giống nhau và được xếp thành một cột.\n- Bảng tuần hoàn gồm 18 cột, chia thành 8 nhóm A (từ IA đến VIIIA) và 8 nhóm B (từ IB đến VIIIB).\n- Đối với nhóm A: Số thứ tự nhóm A = Số electron hóa trị = Số electron lớp ngoài cùng.\n- Phân loại: Khối s, p là nhóm A; Khối d, f là nhóm B (kim loại chuyển tiếp)."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Xác định vị trí nguyên tố trong bảng tuần hoàn từ số hiệu Z.",
                    "content": "Xác định vị trí (ô, chu kì, nhóm) của các nguyên tố có Z = 12, 16, 19 trong bảng tuần hoàn.",
                    "product": "Lời giải của học sinh trên bảng và trong vở.",
                    "implementation_gv": "GV gọi 3 học sinh lên bảng làm 3 câu, hướng dẫn nhận xét và chuẩn hóa đáp án.",
                    "implementation_hs": "* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\n1. Nguyên tố Mg (Z = 12):\n- Cấu hình e: 1s² 2s² 2p⁶ 3s².\n- Vị trí: Ô số 12 (do Z=12); Chu kì 3 (do có 3 lớp e); Nhóm IIA (do có 2e lớp ngoài cùng, nguyên tố s).\n\n2. Nguyên tố S (Z = 16):\n- Cấu hình e: 1s² 2s² 2p⁶ 3s² 3p⁴.\n- Vị trí: Ô số 16 (do Z=16); Chu kì 3 (do có 3 lớp e); Nhóm VIA (do có 6e hóa trị, nguyên tố p).\n\n3. Nguyên tố K (Z = 19):\n- Cấu hình e: [Ar]4s¹.\n- Vị trí: Ô số 19; Chu kì 4; Nhóm IA (kim loại kiềm)."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Tìm hiểu câu chuyện phát minh bảng tuần hoàn.",
                    "content": "Sưu tầm bài viết hoặc video ngắn từ nguồn tin cậy về giấc mơ thấy bảng tuần hoàn của Mendeleev.",
                    "product": "Bài chia sẻ ngắn kèm đường dẫn uy tín ([ATM1]).",
                    "implementation_gv": "Giao nhiệm vụ về nhà tìm hiểu lịch sử bảng tuần hoàn qua các nguồn kiểm chứng.",
                    "implementation_hs": "* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n- Lịch sử phát minh: Năm 1869, nhà hóa học D. I. Mendeleev công bố bảng tuần hoàn đầu tiên gồm 63 nguyên tố sắp xếp theo chiều tăng khối lượng nguyên tử.\n- Điểm vĩ đại: Mendeleev đã để trống các ô và dự đoán chính xác sự tồn tại cùng tính chất của các nguyên tố chưa được tìm thấy thời bấy giờ (như Gallium, Germanium, Scandium)."
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
            "equipment_student": "- SGK, thiết bị di động cài đặt ứng dụng KingDraw, vở ghi.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Kích thích hứng thú về cách các nguyên tử phi kim liên kết với nhau.",
                    "content": "Vì sao hai nguyên tử Chlorine (Cl) đều thiếu 1 electron nhưng lại có thể kết hợp với nhau thành phân tử Cl2 bền vững?",
                    "product": "Dự đoán của HS: dùng chung cặp electron.",
                    "implementation_gv": "Nêu tình huống xuất phát từ quy tắc octet.",
                    "implementation_hs": "* Câu trả lời dự kiến của HS:\n- Nguyên tử Cl (Z=17) có 7e lớp ngoài cùng, cần thêm 1e để đạt cấu hình octet bền vững của khí hiếm Ar.\n- Hai nguyên tử Cl góp chung 1 cặp electron: Cl:Cl → hình thành phân tử Cl₂ bền vững."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Viết công thức Lewis và dựng mô hình 3D trên KingDraw/MolView.",
                    "content": "HS sử dụng ứng dụng KingDraw hoặc MolView để vẽ cấu trúc Lewis và chuyển sang chế độ xem không gian 3D của các phân tử H2O, NH3, CH4.",
                    "product": "Ảnh chụp màn hình cấu trúc 3D của các phân tử do HS dựng.",
                    "implementation_gv": "GV hướng dẫn các thao tác cơ bản trên thanh công cụ MolView: chọn nguyên tố, nối liên kết đơn/đôi, bấm nút 3D.",
                    "implementation_hs": "* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. SỰ HÌNH THÀNH LIÊN KẾT CỘNG HÓA TRỊ\n1. Khái niệm:\n- Liên kết cộng hóa trị là liên kết được hình thành giữa hai nguyên tử bằng một hay nhiều cặp electron dùng chung.\n2. Công thức electron, công thức Lewis và công thức cấu tạo:\n- Phân tử H₂: H : H  →  H – H (1 liên kết đơn σ).\n- Phân tử O₂: O :: O  →  O = O (1 liên kết đôi gồm 1 liên kết σ và 1 liên kết π).\n- Phân tử N₂: N ::: N  →  N ≡ N (1 liên kết ba gồm 1 liên kết σ và 2 liên kết π).\n- Phân tử HCl: H : Cl̈:  →  H – Cl (liên kết đơn phân cực, cặp e lệch về phía Cl).\n\nII. PHÂN LOẠI LIÊN KẾT CỘNG HÓA TRỊ THEO HIỆU ĐỘ ÂM ĐIỆN (Δχ)\n- 0 ≤ Δχ < 0,4: Liên kết cộng hóa trị không phân cực (cặp e dùng chung nằm ở giữa hai nguyên tử, ví dụ: H₂, Cl₂, O₂, N₂).\n- 0,4 ≤ Δχ < 1,7: Liên kết cộng hóa trị có phân cực (cặp e dùng chung lệch về nguyên tử có độ âm điện lớn hơn, ví dụ: HCl, H₂O, NH₃).\n- Δχ ≥ 1,7: Liên kết ion (ví dụ: NaCl, CaO, KBr).\n\nIII. SỰ XEN PHỦ CÁC ORBITAL NGUYÊN TỬ TẠO LIÊN KẾT\n1. Liên kết sigma (σ): Hình thành do sự xen phủ trục (trục liên kết trùng trục đối xứng của các orbital), rất bền vững.\n2. Liên kết pi (π): Hình thành do sự xen phủ bên giữa 2 orbital p song song, kém bền hơn liên kết σ."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Phân loại liên kết dựa vào hiệu độ âm điện.",
                    "content": "Tính hiệu độ âm điện và xác định loại liên kết trong các phân tử: Na2O, AlCl3, HCl, N2.",
                    "product": "Bảng phân loại liên kết hoàn chỉnh.",
                    "implementation_gv": "GV nhận xét và chuẩn hóa bài làm.",
                    "implementation_hs": "* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\nXác định loại liên kết dựa vào hiệu độ âm điện:\n1. Trong phân tử N₂: Δχ = 3,04 - 3,04 = 0 → Liên kết cộng hóa trị không phân cực.\n2. Trong phân tử HCl: Δχ = 3,16 - 2,20 = 0,96 (thuộc khoảng 0,4 - 1,7) → Liên kết cộng hóa trị phân cực (lệch về Cl).\n3. Trong phân tử AlCl₃: Δχ = 3,16 - 1,61 = 1,55 (thuộc khoảng 0,4 - 1,7) → Liên kết cộng hóa trị phân cực mạnh.\n4. Trong phân tử Na₂O: Δχ = 3,44 - 0,93 = 2,51 (≥ 1,7) → Liên kết ion."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Xuất tệp hình ảnh công thức hóa học chuẩn.",
                    "content": "Xuất hình ảnh công thức cấu tạo của phân tử CO2 và chèn vào tệp tài liệu bài tập cá nhân kèm nguồn phần mềm ([ATM6]).",
                    "product": "Tệp bài làm có chèn hình ảnh công thức số hóa.",
                    "implementation_gv": "Hướng dẫn kĩ thuật lưu ảnh PNG trong suốt.",
                    "implementation_hs": "* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n- Cấu tạo phân tử CO₂: O = C = O (cấu trúc đường thẳng, góc liên kết 180°).\n- Mặc dù hai liên kết C=O là liên kết CHT phân cực, nhưng do cấu trúc thẳng đối xứng nên hai vector momen lưỡng cực triệt tiêu nhau → Toàn bộ phân tử CO₂ là phân tử không phân cực.\n- Ứng dụng: CO₂ không phân cực nên ít tan trong nước ở điều kiện thường, ứng dụng nén làm đá khô để bảo quản thực phẩm."
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
            "equipment_student": "- Điện thoại/máy tính có kết nối Internet, vở ghi chép.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Nhận biết các hiện tượng oxi hóa - khử quanh ta.",
                    "content": "Quan sát video đinh sắt bị gỉ, sự cháy của than củi, sự đổi màu quả táo sau khi gọt. Điểm chung của các quá trình này là gì?",
                    "product": "Phát hiện của HS: đều có sự tương tác với oxygen và thay đổi bản chất chất.",
                    "implementation_gv": "Chiếu video trực quan.",
                    "implementation_hs": "* Câu trả lời dự kiến của HS:\n- Quá trình đinh sắt gỉ (Fe tác dụng O₂ và H₂O), than cháy (C tác dụng O₂), quả táo đổi màu đều là phản ứng hóa học có sự chuyển dịch electron giữa các chất."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Nắm vững quy tắc xác định số oxi hóa và 4 bước cân bằng thăng bằng e.",
                    "content": "Nghiên cứu SGK, thảo luận nhóm thống nhất 4 bước cân bằng: Xác định số oxi hóa -> Viết quá trình nhường/nhận e -> Tìm hệ số -> Điền hệ số và kiểm tra.",
                    "product": "Bản nháp 4 bước cân bằng phản ứng Cu + HNO3 loãng.",
                    "implementation_gv": "GV làm mẫu 1 phương trình, sau đó chia nhóm thiết kế Infographic tóm tắt quy trình.",
                    "implementation_hs": "* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. CÁC KHÁI NIỆM CƠ BẢN\n1. Chất khử (chất bị oxi hóa): Là chất nhường electron (số oxi hóa tăng sau phản ứng).\n2. Chất oxi hóa (chất bị khử): Là chất nhận electron (số oxi hóa giảm sau phản ứng).\n   Quy tắc ghi nhớ: 'Khử cho - O nhận' (Chất khử cho electron, chất oxi hóa nhận electron).\n3. Quá trình oxi hóa (sự oxi hóa): Là quá trình nhường electron.\n4. Quá trình khử (sự khử): Là quá trình nhận electron.\n5. Phản ứng oxi hóa - khử: Là phản ứng hóa học trong đó có sự chuyển dịch electron giữa các chất phản ứng (làm thay đổi số oxi hóa của một số nguyên tố).\n\nII. 4 BƯỚC CÂN BẰNG PHẢN ỨNG OXI HÓA - KHỬ (PHƯƠNG PHÁP THĂNG BẰNG ELECTRON)\n* Ví dụ: Cân bằng phản ứng Cu + HNO₃ (loãng) → Cu(NO₃)₂ + NO↑ + H₂O\n- Bước 1: Xác định số oxi hóa của các nguyên tố thay đổi:\n  Cu⁰ → Cu⁺²;  N⁺⁵ (trong HNO₃) → N⁺² (trong NO).\n- Bước 2: Viết các quá trình nhường và nhận electron:\n  Quá trình oxi hóa: Cu⁰ → Cu⁺² + 2e\n  Quá trình khử:     N⁺⁵ + 3e → N⁺²\n- Bước 3: Tìm hệ số thích hợp sao cho tổng e nhường = tổng e nhận (BCNN = 6):\n  3 × | Cu⁰ → Cu⁺² + 2e\n  2 × | N⁺⁵ + 3e → N⁺²\n- Bước 4: Đặt hệ số vào phương trình và kiểm tra (thứ tự: Kim loại → Gốc acid/phi kim → H → O):\n  3Cu + 8HNO₃ → 3Cu(NO₃)₂ + 2NO↑ + 4H₂O\n  (Kiểm tra số nguyên tử O: Vế trái 8 × 3 = 24; Vế phải 3 × 6 + 2 + 4 = 24 → Đúng)."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Luyện tập cân bằng nhanh phản ứng oxi hóa khử.",
                    "content": "Cân bằng các phản ứng: Fe + H2SO4 đặc nóng; Al + HNO3 tạo N2O.",
                    "product": "Lời giải chi tiết của học sinh.",
                    "implementation_gv": "Theo dõi, sửa lỗi xác định sai số oxi hóa.",
                    "implementation_hs": "* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\nCân bằng các phản ứng sau:\n1. Fe + H₂SO₄ (đặc, nóng) → Fe₂(SO₄)₃ + SO₂↑ + H₂O\n   2 × | Fe⁰ → Fe⁺³ + 3e  (nhường 6e)\n   3 × | S⁺⁶ + 2e → S⁺⁴   (nhận 6e)\n   Phương trình: 2Fe + 6H₂SO₄ (đặc, nóng) → Fe₂(SO₄)₃ + 3SO₂↑ + 6H₂O.\n\n2. Al + HNO₃ (loãng) → Al(NO₃)₃ + N₂O↑ + H₂O\n   8 × | Al⁰ → Al⁺³ + 3e\n   3 × | 2N⁺⁵ + 8e → 2N⁺¹ (trong N₂O)\n   Phương trình: 8Al + 30HNO₃ → 8Al(NO₃)₃ + 3N₂O↑ + 15H₂O."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Nộp sản phẩm Infographic nhóm lên không gian lưu trữ chung an toàn.",
                    "content": "Xuất file Infographic định dạng PDF/PNG, nộp bài vào thư mục Google Drive của lớp với phân quyền an toàn ([ATM4]).",
                    "product": "Liên kết sản phẩm số Infographic đã chia sẻ quyền xem hợp lệ.",
                    "implementation_gv": "Nhắc nhở HS không chia sẻ công khai dữ liệu cá nhân (SĐT, địa chỉ) trên bài tập nộp.",
                    "implementation_hs": "* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n- Ứng dụng phản ứng oxi hóa - khử trong đời sống:\n  + Sự cháy của nhiên liệu (xăng dầu, than đá, khí gas) cung cấp năng lượng cho sinh hoạt và công nghiệp.\n  + Sự ăn mòn kim loại (gỉ sắt): Gây thiệt hại lớn. Biện pháp: sơn phủ bề mặt, bôi dầu mỡ, mạ kẽm (bảo vệ điện hóa cho vỏ tàu biển).\n  + Pin và ắc quy: Chuyển đổi hóa năng của phản ứng oxi hóa - khử thành điện năng."
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
            "equipment_student": "- SGK Hóa học 11, máy tính xách tay/điện thoại thông minh có ứng dụng bảng tính, vở ghi.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Phân biệt phản ứng một chiều và phản ứng thuận nghịch.",
                    "content": "Quan sát hiện tượng đốt cháy củi (phản ứng 1 chiều) và quá trình tạo thạch nhũ trong hang động (thuận nghịch).",
                    "product": "Phát biểu ban đầu của học sinh về phản ứng xảy ra theo 2 chiều trái ngược.",
                    "implementation_gv": "Chiếu video hang động Phong Nha - Kẻ Bàng.",
                    "implementation_hs": "* Câu trả lời dự kiến của HS:\n- Phản ứng cháy của củi chỉ xảy ra theo 1 chiều (không thể thu hồi gỗ từ tro và khói).\n- Quá trình hòa tan và kết tinh đá vôi trong hang động diễn ra theo hai chiều trái ngược: CaCO₃ + CO₂ + H₂O ⇌ Ca(HCO₃)₂."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Tìm hiểu đặc điểm cân bằng động, hằng số Kc và nguyên lí Le Chatelier qua mô phỏng PhET.",
                    "content": "HS thao tác trên mô phỏng PhET: thay đổi nhiệt độ bình chứa NO2/N2O4, quan sát màu sắc thay đổi. Mở tệp Excel nhập số liệu nồng độ ban đầu và nồng độ cân bằng để tự động tính KC.",
                    "product": "Bảng tính Excel chứa công thức tính KC và đồ thị biểu diễn tốc độ phản ứng vt, vn theo thời gian.",
                    "implementation_gv": "GV hướng dẫn HS thiết lập công thức trong ô Excel: =C2/(A2*B2^3). Nhắc nhở quy tắc phân quyền bảng tính chia sẻ ([ATM4]).",
                    "implementation_hs": "* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. PHẢN ỨNG MỘT CHIỀU VÀ PHẢN ỨNG THUẬN NGHỊCH\n1. Phản ứng một chiều: Là phản ứng chỉ xảy ra theo một chiều từ chất đầu tạo thành sản phẩm (dùng mũi tên đơn →).\n2. Phản ứng thuận nghịch: Là phản ứng trong cùng điều kiện xác định, xảy ra đồng thời theo cả hai chiều trái ngược nhau (dùng mũi tên hai chiều ⇌).\n\nII. TRẠNG THÁI CÂN BẰNG HÓA HỌC\n1. Khái niệm: Trạng thái cân bằng của phản ứng thuận nghịch là trạng thái tại đó tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch (vt = vn > 0).\n2. Đặc điểm:\n- Cân bằng hóa học là một cân bằng động (các phản ứng vẫn tiếp tục diễn ra với tốc độ bằng nhau).\n- Ở trạng thái cân bằng, nồng độ của các chất trong hệ phản ứng là không đổi.\n\nIII. HẰNG SỐ CÂN BẰNG NỒNG ĐỘ (KC)\n- Xét phản ứng tổng quát: aA + bB ⇌ cC + dD (ở nhiệt độ xác định)\n- Biểu thức hằng số cân bằng: KC = ([C]^c · [D]^d) / ([A]^a · [B]^b)\n  (Trong đó [A], [B], [C], [D] là nồng độ mol/L của các chất ở trạng thái cân bằng; không ghi nồng độ chất rắn vào biểu thức).\n- Ý nghĩa: KC chỉ phụ thuộc vào bản chất phản ứng và nhiệt độ. KC càng lớn thì hiệu suất phản ứng thuận càng cao.\n\nIV. NGUYÊN LÍ CHUYỂN DỊCH CÂN BẰNG LE CHATELIER\n'Một phản ứng thuận nghịch đang ở trạng thái cân bằng, khi chịu một tác động từ bên ngoài (nồng độ, nhiệt độ, áp suất) thì cân bằng sẽ chuyển dịch theo chiều làm giảm tác động đó.'\n- Ảnh hưởng của nhiệt độ: Tăng nhiệt độ → Cân bằng chuyển dịch theo chiều thu nhiệt (ΔH > 0); Giảm nhiệt độ → Cân bằng chuyển dịch theo chiều tỏa nhiệt (ΔH < 0).\n- Ảnh hưởng của nồng độ: Tăng nồng độ chất nào → Cân bằng chuyển dịch theo chiều làm giảm nồng độ chất đó (và ngược lại).\n- Ảnh hưởng của áp suất (áp dụng cho hệ có chất khí): Tăng áp suất → Cân bằng chuyển dịch theo chiều giảm số mol khí (và ngược lại)."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Dự đoán chiều chuyển dịch cân bằng trong tổng hợp ammonia (quá trình Haber - Bosch).",
                    "content": "N2(g) + 3H2(g) <=> 2NH3(g) (Delta H < 0). Cần điều chỉnh áp suất và nhiệt độ như thế nào để tăng hiệu suất tạo NH3?",
                    "product": "Kết luận: Cần tăng áp suất và hạ nhiệt độ phù hợp.",
                    "implementation_gv": "GV hướng dẫn HS lập luận theo nguyên lí Le Chatelier.",
                    "implementation_hs": "* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\nXét cân bằng tổng hợp ammonia: N₂(g) + 3H₂(g) ⇌ 2NH₃(g)   ΔrH°₂₉₈ = -92 kJ < 0 (phản ứng tỏa nhiệt).\nĐể tăng hiệu suất tạo thành NH₃ (chuyển dịch theo chiều thuận):\n1. Về áp suất: Phía chất phản ứng có 1 + 3 = 4 mol khí; phía sản phẩm có 2 mol khí.\n   → Chiều thuận làm giảm số mol khí → Cần TĂNG ÁP SUẤT của hệ.\n2. Về nhiệt độ: Chiều thuận là chiều tỏa nhiệt (ΔH < 0) → Cần HẠ NHIỆT ĐỘ.\n   (Trong thực tế công nghiệp, chọn nhiệt độ tối ưu khoảng 400 - 450 °C kết hợp xúc tác Fe để đảm bảo tốc độ phản ứng đủ nhanh)."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Giải thích vai trò của cân bằng hóa học trong sản xuất công nghiệp và đời sống.",
                    "content": "Tại sao khi mở nắp chai nước ngọt có ga lại thấy bọt khí sủi lên?",
                    "product": "Giải thích cân bằng khí CO2 tan trong nước ngọt phụ thuộc áp suất.",
                    "implementation_gv": "GV đặt câu hỏi mở rộng.",
                    "implementation_hs": "* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n- Giải thích hiện tượng nước ngọt có ga:\n  Trong chai đóng kín có cân bằng: CO₂(aq) ⇌ CO₂(g). Khí CO₂ được nén dưới áp suất cao để tăng độ tan trong nước.\n  Khi mở nắp chai: Áp suất giảm đột ngột → Cân bằng chuyển dịch theo chiều nghịch (tạo khí CO₂) → Khí CO₂ thoát ra ngoài tạo bọt khí sủi lên."
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
            "equipment_student": "- SGK Hóa học 12, điện thoại thông minh để tra cứu tài liệu, vở ghi.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Kích thích hứng thú về mùi thơm tự nhiên của hoa quả và nước hoa.",
                    "content": "Cho HS ngửi mẫu mùi chuối chín nhân tạo và đoán xem chất gì tạo nên mùi thơm hấp dẫn đó?",
                    "product": "Dự đoán của học sinh về este/hương liệu.",
                    "implementation_gv": "GV cho HS trải nghiệm giác quan và giới thiệu vào bài học.",
                    "implementation_hs": "* Câu trả lời dự kiến của HS:\n- Mùi thơm đặc trưng của quả chuối chín, quả dứa, hoa hồng... là do các hợp chất hữu cơ thuộc loại ester tạo nên (ví dụ: isoamyl acetate có mùi chuối chín, ethyl butyrate có mùi dứa)."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Nắm vững định nghĩa, danh pháp, phản ứng thủy phân của ester và lipid.",
                    "content": "Nghiên cứu SGK, so sánh phản ứng thủy phân ester trong môi trường acid (thuận nghịch) và môi trường kiềm (một chiều - xà phòng hóa). Tra cứu nguồn tin y tế về trans-fat (chất béo chuyển hóa).",
                    "product": "Phiếu học tập hoàn thiện phương trình phản ứng hóa học và bảng tổng kết về chất béo tốt / chất béo xấu.",
                    "implementation_gv": "GV hướng dẫn HS quy trình SIFT để kiểm chứng bài viết về 'dầu ăn chiên đi chiên lại' trên mạng xã hội ([ATM1]).",
                    "implementation_hs": "* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. KHÁI NIỆM VÀ CẤU TRÚC PHÂN TỬ ESTER\n1. Khái niệm: Khi thay thế nhóm –OH ở nhóm carboxyl (–COOH) của carboxylic acid bằng nhóm –OR' thì được ester.\n2. Công thức chung của ester đơn chức: R–COO–R' (R là H hoặc gốc hydrocarbon; R' là gốc hydrocarbon).\n- Ester no, đơn chức, mạch hở: CₙH₂ₙO₂ (n ≥ 2).\n3. Danh pháp (Tên gọi):\n- Tên ester = Tên gốc hydrocarbon R' + Tên gốc carboxylate (đổi đuôi ic thành ate).\n- Ví dụ:\n  + CH₃COOC₂H₅: Ethyl acetate\n  + HCOOCH₃: Methyl formate\n  + CH₃COOCH=CH₂: Vinyl acetate\n  + CH₃COOCH₂C₆H₅: Benzyl acetate (mùi hoa nhài).\n\nII. TÍNH CHẤT HÓA HỌC ĐẶC TRƯNG CỦA ESTER\n1. Phản ứng thủy phân trong môi trường acid (Phản ứng thuận nghịch):\n   RCOOR' + H₂O ⇌ (H₂SO₄ đặc, t°) RCOOH + R'OH\n   Ví dụ: CH₃COOC₂H₅ + H₂O ⇌ CH₃COOH + C₂H₅OH\n2. Phản ứng thủy phân trong môi trường kiềm (Phản ứng xà phòng hóa - Một chiều):\n   RCOOR' + NaOH → (t°) RCOONa + R'OH\n   Ví dụ: CH₃COOC₂H₅ + NaOH → (t°) CH₃COONa + C₂H₅OH\n\nIII. LIPID VÀ CHẤT BÉO (TRIGLYCERIDE)\n1. Khái niệm: Chất béo là triester của glycerol với các acid béo, gọi chung là triglyceride.\n- Công thức cấu tạo chung: (RCOO)₃C₃H₅\n2. Các acid béo thường gặp:\n- Acid béo no: Acid palmitic (C₁₅H₃₁COOH), Acid stearic (C₁₇H₃₅COOH).\n  → Tạo chất béo no: Trứng, mỡ động vật (trạng thái rắn ở nhiệt độ thường).\n- Acid béo không no: Acid oleic (C₁₇H₃₃COOH, có 1 liên kết đôi C=C), Acid linoleic (C₁₇H₃₁COOH).\n  → Tạo chất béo không no: Dầu thực vật, dầu cá (trạng thái lỏng ở nhiệt độ thường).\n3. Phản ứng xà phòng hóa chất béo:\n   (RCOO)₃C₃H₅ + 3NaOH → (t°) 3RCOONa (xà phòng) + C₃H₅(OH)₃ (glycerol)"
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (7 phút)",
                    "goal": "Củng cố kĩ năng viết đồng phân và phương trình thủy phân ester.",
                    "content": "Viết các đồng phân ester có công thức phân tử C4H8O2 và gọi tên tương ứng.",
                    "product": "4 đồng phân: HCOOCH2CH2CH3, HCOOCH(CH3)2, CH3COOC2H5, C2H5COOCH3.",
                    "implementation_gv": "GV chiếu đáp án chuẩn, giải đáp thắc mắc.",
                    "implementation_hs": "* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\nViết 4 đồng phân ester của C₄H₈O₂ và gọi tên:\n1. HCOOCH₂CH₂CH₃: Propyl formate\n2. HCOOCH(CH₃)₂: Isopropyl formate\n3. CH₃COOC₂H₅: Ethyl acetate\n4. C₂H₅COOCH₃: Methyl propionate"
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (3 phút)",
                    "goal": "Hiểu tác hại của dầu mỡ chiên nhiều lần đối với sức khỏe.",
                    "content": "Tìm hiểu dự án tái chế dầu ăn thừa thành xà phòng hoặc dầu biodiesel, viết thông điệp truyền thông chia sẻ với gia đình.",
                    "product": "Bản ghi chú thông điệp an toàn dinh dưỡng.",
                    "implementation_gv": "Giao nhiệm vụ về nhà.",
                    "implementation_hs": "* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n1. Nguy cơ từ chất béo chuyển hóa (Trans-fat) và dầu mỡ chiên rán nhiều lần:\n- Ở nhiệt độ cao kéo dài, các liên kết đôi dạng cis trong dầu thực vật bị biến đổi thành dạng trans hoặc bị oxi hóa tạo các hợp chất aldehyde, polymer độc hại gây nguy cơ xơ vữa động mạch, tim mạch và ung thư.\n2. Kiểm chứng thông tin y tế trên Internet ([ATM1]):\n- Áp dụng quy trình SIFT (Dừng lại - Điều tra nguồn tin - Tìm bài đối chiếu - Lần về ngữ cảnh gốc) để phân biệt quảng cáo thực phẩm chức năng thổi phồng với khuyến cáo khoa học của Viện Dinh dưỡng Quốc gia."
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
        """Tự động sinh bộ khung giáo án chuẩn 5512 tích hợp NLS & ATM cho bài học mới bất kì có nội dung ghi vở chi tiết"""
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
            "equipment_student": f"- Sách giáo khoa {subject} {grade}, vở ghi chép bài, thiết bị thông minh có kết nối mạng phục vụ tra cứu có hướng dẫn.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (5 - 7 phút)",
                    "goal": f"Tạo hứng thú, kích thích trí tò mò của học sinh về nội dung {clean_name}.",
                    "content": "GV đưa ra hình ảnh/video thực tế hoặc câu hỏi tình huống mâu thuẫn nhận thức để học sinh suy ngẫm.",
                    "product": "Câu trả lời, dự đoán hoặc thắc mắc của học sinh dẫn dắt vào bài mới.",
                    "implementation_gv": "1. Chuyển giao: GV nêu tình huống xuất phát gắn liền thực tiễn.\n2. Thực hiện: HS suy nghĩ cá nhân hoặc trao đổi cặp đôi.\n3. Báo cáo: Gọi 1-2 đại diện HS phát biểu dự đoán.\n4. Kết luận: GV nhận xét, dẫn dắt vào bài học mới.",
                    "implementation_hs": f"* Câu trả lời dự kiến của HS:\n- Nêu nhận xét ban đầu dựa trên hiện tượng thực tế quan sát được.\n- Dự đoán về tính chất, quy luật hoặc cấu tạo liên quan đến {clean_name}.\n* Vấn đề cần tìm hiểu của bài học:\n- Đặt ra mâu thuẫn nhận thức cần khám phá và giải quyết trong tiết học."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (28 - 30 phút)",
                    "goal": f"Chiếm lĩnh kiến thức trọng tâm của {clean_name} kết hợp ứng dụng công cụ số an toàn.",
                    "content": "HS làm việc với SGK và Phiếu học tập số 1 theo nhóm; khai thác học liệu số/mô phỏng thí nghiệm dưới sự định hướng của giáo viên.",
                    "product": "Sản phẩm hoàn thành phiếu học tập hoặc bảng phụ của các nhóm có trích dẫn nguồn số an toàn.",
                    "implementation_gv": "1. Chuyển giao: Chia nhóm, giao Phiếu học tập số 1 kèm hướng dẫn đường link tài nguyên số đã kiểm chứng an toàn ([ATM1, ATM4]).\n2. Thực hiện: Theo dõi, hỗ trợ các nhóm thực hành thảo luận.\n3. Báo cáo: Đại diện nhóm trình bày sản phẩm, nhóm khác nhận xét, bổ sung.\n4. Kết luận: GV chuẩn hóa kiến thức khoa học chính xác để học sinh ghi vào vở.",
                    "implementation_hs": f"* NỘI DUNG HỌC SINH GHI VỞ:\n\nI. KHÁI NIỆM VÀ ĐỊNH NGHĨA TRỌNG TÂM CỦA {clean_name}\n- Nêu định nghĩa khoa học chuẩn xác theo Chương trình GDPT 2018.\n- Đặc điểm cấu tạo, bản chất và ý nghĩa hóa học liên quan.\n\nII. TÍNH CHẤT VÀ QUY LUẬT CỐT LÕI\n- Các tính chất vật lí và tính chất hóa học đặc trưng của đối tượng nghiên cứu.\n- Viết các phương trình hóa học minh họa chuẩn xác (ghi rõ trạng thái, điều kiện xúc tác, nhiệt độ nếu có).\n\nIII. NGUYÊN TẮC VẬN DỤNG VÀ QUY TẮC BÀI HỌC\n- Các quy tắc thực hiện, công thức toán học hoặc phương pháp giải bài tập.\n- Những nhận xét, lưu ý quan trọng cần ghi nhớ để tránh sai sót."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": f"Củng cố, khắc sâu kiến thức vừa học về {clean_name}.",
                    "content": "Học sinh giải quyết các câu hỏi trắc nghiệm hoặc bài tập vận dụng nhanh trên bảng/nền tảng học tập số.",
                    "product": "Lời giải của học sinh trên bảng con hoặc vở bài tập.",
                    "implementation_gv": "GV trình chiếu hệ thống bài tập trắc nghiệm và tự luận định hướng; gọi học sinh lên bảng làm; theo dõi, hướng dẫn học sinh sửa bài.",
                    "implementation_hs": f"* BÀI TẬP LUYỆN TẬP GHI VỞ:\n\n1. Bài tập 1 (Nhận biết - Thông hiểu):\n- Câu hỏi củng cố định nghĩa và viết phương trình phản ứng của {clean_name}.\n- Lời giải chi tiết và kết luận chuẩn xác của học sinh.\n\n2. Bài tập 2 (Vận dụng):\n- Bài toán định lượng hoặc giải thích hiện tượng thực tế.\n- Các bước giải chi tiết và đáp số chuẩn xác để HS chữa vào vở."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG VÀ MỞ RỘNG (5 phút)",
                    "goal": "Liên hệ kiến thức bài học với ứng dụng thực tiễn trong đời sống và sản xuất.",
                    "content": "Giải quyết câu hỏi thực tế hoặc nhiệm vụ nghiên cứu nhỏ tại nhà, tra cứu thông tin trực tuyến có kiểm chứng độ tin cậy.",
                    "product": "Bản báo cáo ngắn gọn hoặc phần chuẩn bị cho tiết học sau có dẫn nguồn hợp lệ ([ATM6]).",
                    "implementation_gv": "GV giao nhiệm vụ về nhà và nhắc nhở các quy tắc bảo vệ an toàn thông tin cá nhân khi tương tác trực tuyến ([ATM1, ATM6]).",
                    "implementation_hs": f"* NỘI DUNG VẬN DỤNG GHI VỞ:\n\n1. Ứng dụng thực tiễn của {clean_name}:\n- Liên hệ vai trò, ứng dụng thực tế trong đời sống hàng ngày, y học, nông nghiệp và sản xuất công nghiệp hiện đại.\n- Biện pháp an toàn và bảo vệ môi trường liên quan đến nội dung bài học.\n\n2. Hướng dẫn tự học và tra cứu tài liệu số ([ATM1, ATM6]):\n- Sử dụng các cổng thông tin chính thống để tra cứu mở rộng; trích dẫn nguồn học liệu số rõ ràng."
                }
            ],
            "appendix": f"PHIẾU HỌC TẬP:\nNội dung các câu hỏi hướng dẫn tự học và phiếu bài tập phục vụ bài dạy {clean_name}."
        }

    @staticmethod
    def export_word(khbd_data, output_path):
        """Xuất file KHBD chuẩn Công văn 5512 ra định dạng docx tích hợp NLS và ATM, bổ sung nội dung học sinh ghi vở"""
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
            tbl_act.columns[0].width = Inches(3.4)
            tbl_act.columns[1].width = Inches(3.4)

            cell_h1 = tbl_act.cell(0, 0)
            cell_h1.text = "HOẠT ĐỘNG CỦA GV VÀ HS"
            cell_h1.paragraphs[0].runs[0].font.bold = True
            cell_h1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_background(cell_h1, "F2F2F2")

            cell_h2 = tbl_act.cell(0, 1)
            cell_h2.text = "DỰ KIẾN SẢN PHẨM\n(NỘI DUNG HỌC SINH GHI VỞ)"
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
