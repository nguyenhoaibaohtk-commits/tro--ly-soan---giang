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
    """Xử lý tạo và xuất Kế hoạch bài dạy theo Công văn 5512/BGDĐT-GDTrH"""

    LESSON_DATABASE = {
        "hoa10_bai4": {
            "id": "hoa10_bai4",
            "grade": "10",
            "lesson_title": "BÀI 4: CẤU TRÚC LỚP VỎ ELECTRON NGUYÊN TỬ",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Trình bày được mô hình của Rutherford - Bohr và mô hình hiện đại về orbital nguyên tử (AO).\n- Nêu được khái niệm về orbital nguyên tử (AO s, AO p) và số lượng electron trong một AO.\n- Trình bày được cấu trúc lớp vỏ electron: lớp, phân lớp electron và số lượng electron tối đa trong một phân lớp, một lớp.\n- Viết được cấu hình electron nguyên tử của 20 nguyên tố đầu tiên theo ô orbital và dạng số.",
            "objective_competence_general": "- Tự chủ và tự học: Chủ động đọc SGK, quan sát video mô phỏng và hoàn thành phiếu học tập cá nhân.\n- Giao tiếp và hợp tác: Tích cực thảo luận nhóm, phân công nhiệm vụ rõ ràng và chia sẻ kết quả.\n- Giải quyết vấn đề và sáng tạo: Vận dụng quy tắc viết cấu hình electron để dự đoán tính chất cơ bản của nguyên tố.",
            "objective_competence_subject": "- Nhận thức hóa học: Phân biệt được mô hình hành tinh nguyên tử cổ điển và mô hình hiện đại về đám mây electron; hiểu được ý nghĩa các lớp và phân lớp electron.\n- Tìm hiểu thế giới tự nhiên: Sử dụng tranh ảnh, mô hình 3D để hình dung không gian của AO s (hình cầu) và AO p (hình số 8 nổi).\n- Vận dụng kiến thức: Dự đoán được nguyên tố kim loại, phi kim hay khí hiếm dựa vào số electron lớp ngoài cùng.",
            "objective_qualities": "- Chăm chỉ: Nghiêm túc hoàn thành các nhiệm vụ học tập trên lớp và bài tập về nhà.\n- Trung thực: Khách quan, trung thực trong quá trình báo cáo kết quả thảo luận nhóm.\n- Trách nhiệm: Có ý thức hợp tác trách nhiệm với các thành viên trong nhóm.",
            "equipment_teacher": "- Máy tính, máy chiếu, bài giảng PowerPoint có video mô phỏng chuyển động electron và hình ảnh không gian của AO s, AO p.\n- Phiếu học tập số 1, 2 và bảng phụ cho các nhóm hoạt động.",
            "equipment_student": "- Sách giáo khoa Hóa học 10, vở ghi chép.\n- Ôn tập kiến thức về hạt proton, neutron, electron đã học ở bài trước.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: MỞ ĐẦU / KHỞI ĐỘNG (5 - 7 phút)",
                    "goal": "Tạo hứng thú, kích thích tư duy học sinh về sự phân bố và chuyển động của electron quanh hạt nhân nguyên tử.",
                    "content": "GV chiếu hình ảnh hệ Mặt Trời và đặt câu hỏi: Liệu các electron có chuyển động xung quanh hạt nhân theo quỹ đạo tròn cố định giống như các hành tinh quay quanh Mặt Trời hay không?",
                    "product": "Học sinh nêu suy nghĩ, dự đoán cá nhân; nhận thấy sự mâu thuẫn giữa mô hình cổ điển và hiện đại.",
                    "implementation_gv": "1. Chuyển giao nhiệm vụ: GV trình chiếu hình ảnh, nêu câu hỏi gợi mở.\n2. Thực hiện nhiệm vụ: HS quan sát, trao đổi nhanh theo cặp đôi trong 2 phút.\n3. Báo cáo, thảo luận: Đại diện 1-2 HS phát biểu ý kiến trước lớp.\n4. Kết luận, nhận định: GV ghi nhận các ý kiến, dẫn dắt vào bài học mới.",
                    "implementation_hs": "HS lắng nghe, quan sát tranh ảnh, ghi nhận câu hỏi khởi động và tích cực nêu dự đoán cá nhân."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (30 phút)",
                    "goal": "Giúp HS nắm vững khái niệm AO s, AO p; cấu trúc lớp, phân lớp và quy tắc viết cấu hình electron nguyên tử.",
                    "content": "HS làm việc với SGK và Phiếu học tập số 1 để tìm hiểu:\n- Mô hình hiện đại về đám mây electron và Orbital nguyên tử (AO).\n- Khái niệm lớp và phân lớp electron.\n- Nguyên lí Pauli, nguyên lí vững bền và quy tắc Hund khi điền electron.",
                    "product": "Bản điền đầy đủ thông tin vào Phiếu học tập số 1; viết đúng cấu hình electron của một số nguyên tử quen thuộc (C, N, O, Na, Al).",
                    "implementation_gv": "1. Chuyển giao: GV chia lớp thành 4 nhóm, phát Phiếu học tập số 1 yêu cầu hoàn thiện trong 10 phút.\n2. Thực hiện: Các nhóm thảo luận, GV quan sát hỗ trợ các nhóm gặp khó khăn.\n3. Báo cáo: Đại diện nhóm 1 và nhóm 3 treo bảng phụ báo cáo; nhóm 2 và 4 nhận xét, phản biện.\n4. Kết luận: GV nhận xét, chuẩn hóa kiến thức, lưu ý các lỗi sai thường gặp khi viết cấu hình e.",
                    "implementation_hs": "HS tích cực phối hợp cùng nhóm, nghiên cứu mục tương ứng trong SGK để hoàn thành phiếu học tập."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": "Khắc sâu kiến thức về viết cấu hình electron và xác định tính chất cơ bản của nguyên tố.",
                    "content": "HS làm bài tập viết cấu hình electron cho các nguyên tử có Z = 11, 15, 17 và cho biết chúng là kim loại hay phi kim.",
                    "product": "Đáp án của HS trên bảng con hoặc bảng nhóm; lời giải chi tiết giải thích vì sao là kim loại/phi kim.",
                    "implementation_gv": "1. Chuyển giao: GV chiếu câu hỏi luyện tập.\n2. Thực hiện: HS làm bài cá nhân vào vở trong 4 phút.\n3. Báo cáo: 2 HS lên bảng trình bày, các HS dưới lớp đổi vở chấm chéo.\n4. Kết luận: GV chốt điểm số và tuyên dương các em làm nhanh, chính xác.",
                    "implementation_hs": "HS độc lập làm bài, đối chiếu kết quả với bạn và sửa sai nếu có."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG VÀ MỞ RỘNG (5 phút)",
                    "goal": "Vận dụng hiểu biết về cấu hình electron để giải thích tính trơ của khí hiếm và ứng dụng trong thực tiễn.",
                    "content": "Câu hỏi: Tại sao khí Helium hay Neon lại được dùng để bơm vào bóng bay quảng cáo hoặc làm đèn led phát sáng mà không sợ cháy nổ?",
                    "product": "Bản ghi chú vào vở về cấu hình electron bão hòa bền vững của khí hiếm.",
                    "implementation_gv": "1. Chuyển giao: GV giao câu hỏi vận dụng làm tại nhà.\n2. Hướng dẫn HS tìm kiếm thông tin trên internet và sách báo khoa học.\n3. Nhắc nhở HS chuẩn bị bài học tiếp theo.",
                    "implementation_hs": "HS ghi chép nhiệm vụ về nhà vào vở và chủ động tra cứu tài liệu."
                }
            ],
            "appendix": "PHIẾU HỌC TẬP SỐ 1:\n1. Orbital nguyên tử (AO) là gì? Hình dạng AO s và AO p khác nhau như thế nào?\n2. Lớp thứ n có tối đa bao nhiêu phân lớp? Lớp M (n=3) có các phân lớp nào?\n3. Viết cấu hình electron của nguyên tử Na (Z=11) và Cl (Z=17)."
        },

        "hoa10_bai5": {
            "id": "hoa10_bai5",
            "grade": "10",
            "lesson_title": "BÀI 5: CẤU TẠO BẢNG TUẦN HOÀN CÁC NGUYÊN TỐ HÓA HỌC",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Nêu được lịch sử phát minh bảng tuần hoàn của Mendeleev và nguyên tắc sắp xếp các nguyên tố.\n- Mô tả được cấu tạo bảng tuần hoàn gồm: ô nguyên tố, chu kì (chu kì nhỏ, chu kì lớn) và nhóm (nhóm A, nhóm B).\n- Xác định được mối liên hệ giữa cấu hình electron nguyên tử với vị trí (ô, chu kì, nhóm) trong bảng tuần hoàn.",
            "objective_competence_general": "- Tự chủ và tự học: Tự tra cứu bảng tuần hoàn để xác định vị trí các nguyên tố.\n- Giao tiếp và hợp tác: Làm việc nhóm để phân tích đặc điểm cấu tạo của các chu kì và nhóm nguyên tố.\n- Giải quyết vấn đề: Suy luận vị trí của nguyên tố khi biết số hiệu nguyên tử Z hoặc cấu hình electron.",
            "objective_competence_subject": "- Nhận thức hóa học: Hiểu được nguyên tắc số thứ tự chu kì bằng số lớp electron; số thứ tự nhóm A bằng số electron hóa trị.\n- Tìm hiểu thế giới tự nhiên: Khám phá trật tự hài hòa của các nguyên tố trong vũ trụ thông qua bảng tuần hoàn.",
            "objective_qualities": "- Chăm chỉ: Chủ động thực hiện các bài tập tra cứu bảng tuần hoàn.\n- Trách nhiệm: Tham gia tích cực vào các hoạt động nhóm.",
            "equipment_teacher": "- Bảng tuần hoàn các nguyên tố hóa học khổ lớn treo tường hoặc bản điện tử.\n- Phiếu học tập hướng dẫn tìm hiểu cấu tạo ô, chu kì, nhóm.",
            "equipment_student": "- Bảng tuần hoàn cá nhân, SGK Hóa học 10.",
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
                    "goal": "Tìm hiểu 3 nguyên tắc sắp xếp và 3 thành phần cấu tạo của bảng tuần hoàn (ô, chu kì, nhóm).",
                    "content": "HS nghiên cứu SGK, quan sát Bảng tuần hoàn và hoàn thành Phiếu học tập số 1 về:\n- Thông tin trong 1 ô nguyên tố.\n- Khái niệm chu kì và số lượng nguyên tố trong mỗi chu kì.\n- Khái niệm nhóm, phân biệt nhóm A và nhóm B.",
                    "product": "Bản báo cáo phiếu học tập đã hoàn thiện của các nhóm.",
                    "implementation_gv": "GV giao nhiệm vụ theo từng trạm/nhóm: Trạm 1 (Ô nguyên tố), Trạm 2 (Chu kì), Trạm 3 (Nhóm).",
                    "implementation_hs": "HS luân chuyển trạm, ghi nhận thông tin vào phiếu học tập."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": "Rèn luyện kĩ năng xác định vị trí của nguyên tố trong bảng tuần hoàn từ cấu hình e.",
                    "content": "Xác định vị trí (ô, chu kì, nhóm) của các nguyên tố có Z = 12, 16, 19 trong bảng tuần hoàn.",
                    "product": "Lời giải của học sinh trên bảng và trong vở.",
                    "implementation_gv": "GV gọi 3 học sinh lên bảng làm 3 câu, hướng dẫn nhận xét.",
                    "implementation_hs": "HS làm bài tập vào vở và nhận xét bài của bạn trên bảng."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (5 phút)",
                    "goal": "Tìm hiểu ý nghĩa thực tiễn và cách học thuộc bảng tuần hoàn nhanh bằng thơ/thần chú.",
                    "content": "Sưu tầm các câu khẩu quyết ghi nhớ các nhóm IA, IIA, VIIA và giới thiệu với cả lớp.",
                    "product": "Bản ghi các câu thơ vui nhớ bảng tuần hoàn.",
                    "implementation_gv": "GV giới thiệu ví dụ ghi nhớ nhóm IA (Lâu nay không rảnh coi phim...).",
                    "implementation_hs": "HS ghi nhớ nhiệm vụ về nhà sáng tạo thêm khẩu quyết."
                }
            ],
            "appendix": "PHIẾU HỌC TẬP:\n1. Số thứ tự của ô nguyên tố bằng đại lượng nào?\n2. Vì sao chu kì 1 chỉ có 2 nguyên tố, chu kì 2 và 3 có 8 nguyên tố?\n3. Cho nguyên tử X có cấu hình: 1s2 2s2 2p6 3s2 3p4. Xác định ô, chu kì, nhóm của X."
        },

        "hoa10_bai6": {
            "id": "hoa10_bai6",
            "grade": "10",
            "lesson_title": "BÀI 6: XU HƯỚNG BIẾN ĐỔI TÍNH CHẤT NGUYÊN TỬ VÀ ĐƠN CHẤT",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Giải thích được xu hướng biến đổi bán kính nguyên tử, độ âm điện trong một chu kì và trong một nhóm A.\n- Trình bày được xu hướng biến đổi tính kim loại, tính phi kim của các nguyên tố trong một chu kì và một nhóm A.\n- Nêu được quy luật biến đổi tính acid - base của các oxide và hydroxide theo chu kì.",
            "objective_competence_general": "- Giải quyết vấn đề: Vận dụng điện tích hạt nhân và số lớp e để giải thích quy luật biến đổi tính chất.",
            "objective_competence_subject": "- Nhận thức hóa học: Hiểu sâu sắc bản chất lực hút giữa hạt nhân và electron lớp ngoài cùng.",
            "objective_qualities": "- Trung thực, chăm chỉ trong học tập.",
            "equipment_teacher": "- Biểu đồ biến đổi bán kính nguyên tử và độ âm điện.",
            "equipment_student": "- SGK, bảng tuần hoàn.",
            "activities": [
                {
                    "title": "HOẠT ĐỘNG 1: KHỞI ĐỘNG (5 phút)",
                    "goal": "Khơi gợi sự tò mò về kích thước của các nguyên tử.",
                    "content": "Tại sao đi từ Trái sang Phải trong một chu kì, số hạt trong nguyên tử tăng lên mà bán kính nguyên tử lại co nhỏ lại?",
                    "product": "Dự đoán của học sinh.",
                    "implementation_gv": "GV nêu mâu thuẫn nhận thức để tạo hứng thú.",
                    "implementation_hs": "HS thảo luận cặp đôi."
                },
                {
                    "title": "HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC (30 phút)",
                    "goal": "Tìm hiểu quy luật biến đổi bán kính, độ âm điện, tính kim loại - phi kim.",
                    "content": "Hoàn thành bảng tổng kết quy luật biến đổi trong chu kì và trong nhóm A.",
                    "product": "Bảng ma trận quy luật biến đổi tính chất.",
                    "implementation_gv": "GV chia nhóm hoàn thành sơ đồ tư duy.",
                    "implementation_hs": "HS nghiên cứu SGK, vẽ sơ đồ tư duy."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": "Sắp xếp dãy nguyên tố theo thứ tự tăng/giảm tính chất.",
                    "content": "Sắp xếp tính kim loại tăng dần của: Na, Mg, Al, K.",
                    "product": "Đáp án: Al < Mg < Na < K.",
                    "implementation_gv": "GV gọi HS trình bày lời giải thích.",
                    "implementation_hs": "HS trình bày cách so sánh."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (5 phút)",
                    "goal": "Giải thích tính acid của nước mưa hoặc ứng dụng của oxide cao nhất.",
                    "content": "So sánh tính acid của H2SO4 và H3PO4 dựa vào vị trí nguyên tố.",
                    "product": "Báo cáo ngắn của HS.",
                    "implementation_gv": "Giao nhiệm vụ về nhà.",
                    "implementation_hs": "Ghi nhận bài tập về nhà."
                }
            ],
            "appendix": "PHIẾU TỔNG KẾT:\n- Trong 1 chu kì (Trái -> Phải): R giảm, ĐÂĐ tăng, Tính KL giảm, Tính PK tăng.\n- Trong 1 nhóm A (Trên -> Dưới): R tăng, ĐÂĐ giảm, Tính KL tăng, Tính PK giảm."
        },

        "hoa11_bai1": {
            "id": "hoa11_bai1",
            "grade": "11",
            "lesson_title": "BÀI 1: KHÁI NIỆM VỀ CÂN BẰNG HÓA HỌC",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Trình bày được khái niệm phản ứng thuận nghịch và trạng thái cân bằng hóa học.\n- Viết được biểu thức hằng số cân bằng Kc của phản ứng thuận nghịch.\n- Vận dụng nguyên lí chuyển dịch cân bằng Le Chatelier để dự đoán chiều chuyển dịch cân bằng khi thay đổi nhiệt độ, nồng độ, áp suất.",
            "objective_competence_general": "- Năng lực tự học, hợp tác và giải quyết vấn đề.",
            "objective_competence_subject": "- Nhận thức hóa học: Hiểu cân bằng hóa học là cân bằng động (vt = vn > 0).",
            "objective_qualities": "- Chăm chỉ, cẩn thận, yêu thích khám phá khoa học.",
            "equipment_teacher": "- Dụng cụ thí nghiệm điều chế khí NO2 / N2O4 ngâm vào nước nóng và nước đá.",
            "equipment_student": "- SGK Hóa học 11, vở ghi chép.",
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
                    "goal": "Tìm hiểu đặc điểm cân bằng động, hằng số Kc và nguyên lí Le Chatelier.",
                    "content": "HS làm việc nhóm phân tích đồ thị tốc độ phản ứng theo thời gian và tác động của nhiệt độ, nồng độ, áp suất.",
                    "product": "Bản tổng kết nguyên lí Le Chatelier.",
                    "implementation_gv": "GV biểu diễn thí nghiệm NO2 (màu nâu đỏ) chuyển thành N2O4 (không màu).",
                    "implementation_hs": "Quan sát sự đổi màu của bình khí khi ngâm nước đá/nước nóng."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": "Dự đoán chiều chuyển dịch cân bằng trong tổng hợp ammonia (quá trình Haber - Bosch).",
                    "content": "N2(g) + 3H2(g) <=> 2NH3(g) (Delta H < 0). Cần điều chỉnh áp suất và nhiệt độ như thế nào để tăng hiệu suất tạo NH3?",
                    "product": "Kết luận: Cần tăng áp suất và hạ nhiệt độ phù hợp.",
                    "implementation_gv": "GV hướng dẫn HS lập luận theo nguyên lí Le Chatelier.",
                    "implementation_hs": "HS suy luận và ghi chép."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (5 phút)",
                    "goal": "Giải thích vai trò của cân bằng hóa học trong sản xuất công nghiệp và đời sống.",
                    "content": "Tại sao khi mở nắp chai nước ngọt có ga lại thấy bọt khí sủi lên?",
                    "product": "Giải thích cân bằng khí CO2 tan trong nước ngọt phụ thuộc áp suất.",
                    "implementation_gv": "GV đặt câu hỏi mở rộng.",
                    "implementation_hs": "Ghi chép câu hỏi và hoàn thành ở nhà."
                }
            ],
            "appendix": "PHIẾU BÀI TẬP: Cho cân bằng: CaCO3(s) <=> CaO(s) + CO2(g) (Delta H > 0). Cân bằng chuyển dịch theo chiều nào khi: a) Thêm CaCO3? b) Tăng nhiệt độ? c) Tăng áp suất?"
        },

        "hoa12_bai1": {
            "id": "hoa12_bai1",
            "grade": "12",
            "lesson_title": "BÀI 1: ESTER VÀ LIPID",
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": "- Nêu được khái niệm, đặc điểm cấu tạo phân tử của ester và chất béo (triglyceride).\n- Trình bày được tính chất vật lí (mùi thơm hoa quả, tính tan) và tính chất hóa học đặc trưng (thủy phân trong môi trường acid và base - xà phòng hóa).\n- Nêu được ứng dụng của ester và chất béo trong công nghiệp thực phẩm, mỹ phẩm và đời sống.",
            "objective_competence_general": "- Tự chủ, giao tiếp, hợp tác và ứng dụng kiến thức vào thực tế đời sống.",
            "objective_competence_subject": "- Nhận thức hóa học: Viết đúng phương trình phản ứng este hóa và xà phòng hóa chất béo.\n- Vận dụng kiến thức: Giải thích quy trình sản xuất xà phòng và sử dụng dầu mỡ hợp lí.",
            "objective_qualities": "- Chăm chỉ, có ý thức an toàn thực phẩm và bảo vệ sức khỏe.",
            "equipment_teacher": "- Mẫu một số ester có mùi thơm (isoamyl acetate mùi chuối chín, ethyl butyrate mùi dứa).\n- Tranh ảnh quy trình sản xuất xà phòng.",
            "equipment_student": "- SGK Hóa học 12, bảng ôn tập về carboxylic acid và alcohol.",
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
                    "content": "Nghiên cứu SGK, so sánh phản ứng thủy phân ester trong môi trường acid (thuận nghịch) và môi trường kiềm (một chiều - xà phòng hóa).",
                    "product": "Phiếu học tập hoàn thiện phương trình phản ứng hóa học.",
                    "implementation_gv": "GV chia nhóm làm việc, hướng dẫn cách gọi tên ester RCOOR'.",
                    "implementation_hs": "HS luyện tập gọi tên và viết phương trình phản ứng."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": "Củng cố kĩ năng viết đồng phân và phương trình thủy phân ester.",
                    "content": "Viết các đồng phân ester có công thức phân tử C4H8O2 và gọi tên tương ứng.",
                    "product": "4 đồng phân: HCOOCH2CH2CH3, HCOOCH(CH3)2, CH3COOC2H5, C2H5COOCH3.",
                    "implementation_gv": "GV chiếu đáp án chuẩn, giải đáp thắc mắc.",
                    "implementation_hs": "HS hoàn thành bài tập vào vở."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG (5 phút)",
                    "goal": "Hiểu tác hại của dầu mỡ chiên đi chiên lại nhiều lần đối với sức khỏe.",
                    "content": "Tại sao không nên sử dụng dầu ăn đã bị đen và có mùi khét? Đề xuất cách tái chế dầu ăn thừa thành xà phòng hoặc nhiên liệu sinh học.",
                    "product": "Báo cáo ngắn về hiện tượng ôi thiu dầu mỡ và an toàn thực phẩm.",
                    "implementation_gv": "Giao nhiệm vụ nghiên cứu STEM về làm xà phòng handmade.",
                    "implementation_hs": "Nhận nhiệm vụ và lập kế hoạch thực hiện tại nhà."
                }
            ],
            "appendix": "PHIẾU BÀI TẬP:\n1. Thủy phân hoàn toàn 8,8g ethyl acetate bằng dung dịch NaOH vừa đủ. Tính khối lượng muối thu được.\n2. So sánh nhiệt độ sôi của ester so với alcohol và carboxylic acid có cùng số C."
        }
    }

    @staticmethod
    def get_lesson_list():
        """Lấy danh sách các bài học có sẵn trong hệ thống"""
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
        return KHBDProcessor.LESSON_DATABASE["hoa10_bai4"]

    @staticmethod
    def generate_smart_outline(lesson_title, grade="10", subject="HÓA HỌC"):
        """Tự động sinh bộ khung giáo án chuẩn 5512 cho bài học mới bất kì"""
        clean_name = lesson_title.upper().strip()
        return {
            "school_name": "TRƯỜNG THPT HUỲNH THÚC KHÁNG",
            "department": "TỔ HÓA HỌC",
            "teacher_name": "NGUYỄN VĂN BẢO",
            "lesson_title": clean_name,
            "subject": subject,
            "grade": str(grade),
            "duration": "2 tiết",
            "curriculum": "Kết nối tri thức / Cánh Diều / Chân trời sáng tạo",
            "objective_knowledge": f"- Nêu được các khái niệm, định nghĩa và nội dung trọng tâm của {clean_name}.\n- Trình bày được tính chất, quy luật hoặc cấu tạo liên quan đến nội dung bài học.\n- Viết được các phương trình hóa học minh họa hoặc thực hiện được các bước giải bài tập liên quan.",
            "objective_competence_general": "- Tự chủ và tự học: Chủ động đọc SGK, hoàn thành phiếu học tập cá nhân.\n- Giao tiếp và hợp tác: Tích cực thảo luận nhóm, phân công nhiệm vụ và báo cáo kết quả.\n- Giải quyết vấn đề và sáng tạo: Vận dụng kiến thức bài học để giải quyết các tình huống thực tiễn.",
            "objective_competence_subject": f"- Nhận thức hóa học: Hiểu rõ bản chất các hiện tượng, phản ứng và quy luật trong {clean_name}.\n- Tìm hiểu thế giới tự nhiên dưới góc độ hóa học.\n- Vận dụng kiến thức, kĩ năng đã học vào đời sống và bảo vệ môi trường.",
            "objective_qualities": "- Chăm chỉ, trung thực trong học tập và báo cáo kết quả.\n- Có tinh thần trách nhiệm với công việc chung của tập thể.",
            "equipment_teacher": f"- Máy chiếu, giáo án điện tử, hình ảnh/video minh họa cho bài học {clean_name}.\n- Phiếu học tập số 1, số 2 và dụng cụ/hóa chất thí nghiệm (nếu có).",
            "equipment_student": f"- Sách giáo khoa {subject} {grade}, vở ghi chép, chuẩn bị trước bài học theo hướng dẫn.",
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
                    "goal": f"Chiếm lĩnh kiến thức trọng tâm của {clean_name}.",
                    "content": "HS làm việc với SGK và Phiếu học tập số 1 theo nhóm để tìm hiểu các nội dung cốt lõi.",
                    "product": "Sản phẩm hoàn thành phiếu học tập hoặc bảng phụ của các nhóm.",
                    "implementation_gv": "1. Chuyển giao: Chia nhóm, phát phiếu học tập và quy định thời gian.\n2. Thực hiện: Theo dõi, hỗ trợ các nhóm gặp khó khăn.\n3. Báo cáo: Đại diện nhóm trình bày, nhóm khác nhận xét.\n4. Kết luận: GV chốt kiến thức chuẩn xác vào vở.",
                    "implementation_hs": "HS tích cực thảo luận nhóm, ghi nhận kiến thức trọng tâm."
                },
                {
                    "title": "HOẠT ĐỘNG 3: LUYỆN TẬP (8 phút)",
                    "goal": f"Củng cố, khắc sâu kiến thức vừa học về {clean_name}.",
                    "content": "Học sinh giải quyết các câu hỏi trắc nghiệm hoặc bài tập vận dụng nhanh.",
                    "product": "Lời giải của học sinh trên bảng con hoặc vở bài tập.",
                    "implementation_gv": "GV chiếu câu hỏi, quan sát và nhận xét đánh giá.",
                    "implementation_hs": "HS tích cực làm bài và đối chiếu đáp án."
                },
                {
                    "title": "HOẠT ĐỘNG 4: VẬN DỤNG VÀ MỞ RỘNG (5 phút)",
                    "goal": "Liên hệ kiến thức bài học với ứng dụng thực tiễn trong đời sống và sản xuất.",
                    "content": "Giải quyết câu hỏi thực tế hoặc nhiệm vụ nghiên cứu nhỏ tại nhà.",
                    "product": "Bản báo cáo ngắn gọn hoặc phần chuẩn bị cho tiết học sau.",
                    "implementation_gv": "GV giao nhiệm vụ về nhà và hướng dẫn tài liệu tham khảo.",
                    "implementation_hs": "Ghi chép nhiệm vụ và thực hiện ở nhà."
                }
            ],
            "appendix": f"PHIẾU HỌC TẬP:\nNội dung các câu hỏi hướng dẫn tự học và phiếu bài tập phục vụ bài dạy {clean_name}."
        }

    @staticmethod
    def export_word(khbd_data, output_path):
        """Xuất file KHBD chuẩn Công văn 5512 ra định dạng docx"""
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
        p_l.add_run(f"{khbd_data.get('school_name', 'TRƯỜNG THPT ...').upper()}\n").font.bold = True
        p_l.add_run(f"{khbd_data.get('department', 'TỔ CHUYÊN MÔN').upper()}\n").font.bold = True
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
        r_title = p_title.add_run("KẾ HOẠCH BÀI DẠY (CV 5512)\n")
        r_title.font.bold = True
        r_title.font.size = Pt(15)
        r_title.font.color.rgb = RGBColor(0, 32, 96)

        r_lesson = p_title.add_run(f"{khbd_data.get('lesson_title', 'TÊN BÀI DẠY').upper()}\n")
        r_lesson.font.bold = True
        r_lesson.font.size = Pt(14)
        
        info_str = f"Môn học: {khbd_data.get('subject', 'Hóa học')} - Lớp: {khbd_data.get('grade', '10')} | Thời lượng: {khbd_data.get('duration', '2 tiết')}\nGiáo viên thực hiện: {khbd_data.get('teacher_name', '...')}"
        r_info = p_title.add_run(info_str)
        r_info.font.italic = True
        r_info.font.size = Pt(12)

        doc.add_paragraph()

        # 3. MỤC TIÊU (I)
        h1 = doc.add_paragraph()
        h1.add_run("I. MỤC TIÊU").font.bold = True
        
        p_k = doc.add_paragraph()
        p_k.paragraph_format.left_indent = Inches(0.2)
        p_k.add_run("1. Về kiến thức:\n").font.bold = True
        p_k.add_run(khbd_data.get("objective_knowledge", ""))

        p_c = doc.add_paragraph()
        p_c.paragraph_format.left_indent = Inches(0.2)
        p_c.add_run("2. Về năng lực:\n").font.bold = True
        
        p_cg = doc.add_paragraph()
        p_cg.paragraph_format.left_indent = Inches(0.4)
        p_cg.add_run("a) Năng lực chung:\n").font.bold = True
        p_cg.add_run(khbd_data.get("objective_competence_general", ""))

        p_cs = doc.add_paragraph()
        p_cs.paragraph_format.left_indent = Inches(0.4)
        p_cs.add_run("b) Năng lực đặc thù (Hóa học):\n").font.bold = True
        p_cs.add_run(khbd_data.get("objective_competence_subject", ""))

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
        p_s2.add_run(khbd_data.get("teacher_name", "")).font.bold = True

        doc.save(output_path)
        return output_path
