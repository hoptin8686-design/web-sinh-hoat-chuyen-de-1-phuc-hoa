# -*- coding: utf-8 -*-
import os
import sys
import shutil
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="CCCCCC", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def format_run(run, font_name="Times New Roman", size_pt=12, bold=False, italic=False, color_rgb=(0,0,0), underline=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.color.rgb = RGBColor(*color_rgb)

def set_paragraph_spacing(p, before_pt=2, after_pt=4, line_spacing=1.15):
    p.paragraph_format.space_before = Pt(before_pt)
    p.paragraph_format.space_after = Pt(after_pt)
    p.paragraph_format.line_spacing = line_spacing

def add_styled_heading(doc, text, level=1):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before_pt=10, after_pt=4)
    run = p.add_run(text)
    if level == 1:
        format_run(run, size_pt=14, bold=True, color_rgb=(180, 20, 20))
    elif level == 2:
        format_run(run, size_pt=13, bold=True, color_rgb=(30, 64, 175))
    else:
        format_run(run, size_pt=12, bold=True, color_rgb=(15, 23, 42))
    return p

# -------------------------------------------------------------
# 1. TẠO TÀI LIỆU BÁO CÁO SINH HOẠT CHUYÊN ĐỀ 1
# -------------------------------------------------------------
def build_report_doc(file_path):
    doc = docx.Document()
    
    # Page setup: A4, Margins 2cm
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.79)

    # Header Table: Quốc hiệu & Cơ quan ban hành
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="FFFFFF", sz="0", val="none")
    
    row = table.rows[0]
    cell_l = row.cells[0]
    cell_r = row.cells[1]
    cell_l.width = Inches(3.2)
    cell_r.width = Inches(3.3)
    
    p_l = cell_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_l, 0, 2)
    r1 = p_l.add_run("ĐẢNG BỘ XÃ PHỤC HÒA\n")
    format_run(r1, size_pt=11, bold=True)
    r2 = p_l.add_run("CHI BỘ TRƯỜNG THPT PHỤC HÒA\n")
    format_run(r2, size_pt=11, bold=True)
    r3 = p_l.add_run("*\n")
    format_run(r3, size_pt=10, bold=True)
    r_date = p_l.add_run("Số:    -BC/CB")
    format_run(r_date, size_pt=11, italic=True)

    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_r, 0, 2)
    r4 = p_r.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n\n")
    format_run(r4, size_pt=12, bold=True, color_rgb=(180, 20, 20))
    r5 = p_r.add_run("Phục Hòa, ngày 28 tháng 10 năm 2026")
    format_run(r5, size_pt=11, italic=True)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_title, before_pt=14, after_pt=4)
    r_title1 = p_title.add_run("BÁO CÁO\n")
    format_run(r_title1, size_pt=16, bold=True, color_rgb=(180, 20, 20))
    r_title2 = p_title.add_run("Chuyên đề 1: Phát huy vai trò tiên phong, gương mẫu của đảng viên trong chuyển đổi số và ứng dụng trí tuệ nhân tạo trong quản lý, dạy học năm học 2026 – 2027\n")
    format_run(r_title2, size_pt=13, bold=True, color_rgb=(15, 23, 42))
    r_subtitle = p_title.add_run("(Tài liệu phục vụ sinh hoạt chuyên đề Chi bộ tháng 11/2026)")
    format_run(r_subtitle, size_pt=11, italic=True)

    # I. NỘI DUNG CHI TIẾT
    add_styled_heading(doc, "I. NỘI DUNG CHI TIẾT BUỔI TRAO ĐỔI CHUYÊN ĐỀ", 1)

    # Phần 1
    add_styled_heading(doc, "Phần 1: Đặt vấn đề & Khảo sát thực trạng tại Trường THPT Phục Hòa", 2)
    
    p = doc.add_paragraph()
    set_paragraph_spacing(p)
    r = p.add_run("* Thực trạng giảng dạy tại trường THPT Phục Hòa:")
    format_run(r, bold=True)
    
    items_p1 = [
        "Giáo viên mất nhiều thời gian tìm kiếm tài liệu, soạn giáo án thủ công, thiết kế đồ dùng dạy học trực quan còn hạn chế do thiếu trang thiết bị chuyên dụng.",
        "Học sinh (đặc biệt là học sinh đồng bào dân tộc thiểu số tại địa bàn huyện vùng cao Quảng Hòa) đôi khi còn rụt rè, khả năng tự học ở nhà gặp khó khăn do thiếu sự kèm cặp sát sao và môi trường học liệu phong phú."
    ]
    for it in items_p1:
        p_it = doc.add_paragraph(style='List Bullet')
        set_paragraph_spacing(p_it, 1, 2)
        r_it = p_it.add_run(it)
        format_run(r_it, size_pt=12)

    p = doc.add_paragraph()
    set_paragraph_spacing(p)
    r = p.add_run("* Cơ hội từ Trí tuệ nhân tạo (AI):")
    format_run(r, bold=True)

    items_ai = [
        "AI giúp giáo viên tiết kiệm từ 40% - 50% thời gian hành chính và chuẩn bị tài liệu, bài giảng.",
        "Tạo ra các bài giảng sinh động, hình ảnh minh họa trực quan, video mô phỏng giúp học sinh vùng cao dễ dàng tiếp thu kiến thức trừu tượng, thúc đẩy hứng thú học tập."
    ]
    for it in items_ai:
        p_it = doc.add_paragraph(style='List Bullet')
        set_paragraph_spacing(p_it, 1, 2)
        r_it = p_it.add_run(it)
        format_run(r_it, size_pt=12)

    # Phần 2
    add_styled_heading(doc, "Phần 2: Các ứng dụng cụ thể của AI trong các khâu dạy học", 2)
    p_note = doc.add_paragraph()
    set_paragraph_spacing(p_note)
    r_note = p_note.add_run("(Tập trung vào các công cụ miễn phí, dễ sử dụng, phù hợp với giáo viên phổ thông vùng cao)")
    format_run(r_note, italic=True)

    app_sections = [
        ("1. AI hỗ trợ soạn giáo án và thiết kế kế hoạch bài dạy (Kế hoạch bài dạy theo CV 5512)",
         "Gemini, ChatGPT, Claude, Copilot.",
         [
             "Giáo viên yêu cầu AI viết khung giáo án theo đúng công văn hiện hành của Bộ GD&ĐT (CV 5512/BGDĐT): xác định rõ 04 hoạt động dạy học (Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng) với các tiêu chí Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện.",
             "Điều chỉnh nội dung cho sát với năng lực thực tế của học sinh trường THPT Phục Hòa: lược bớt phần nâng cao phức tạp, tăng cường hoạt động trực quan, thảo luận nhóm, trò chơi học tập."
         ]),
        ("2. AI hỗ trợ tạo học liệu, hình ảnh minh họa và video bài giảng trực quan",
         "Canva AI, Microsoft Copilot Designer, Gamma App, CapCut AI.",
         [
             "Với các môn học khó hình dung (Địa lý, Lịch sử, Sinh học, Vật lý, Tin học), giáo viên dùng AI tạo ra sơ đồ tư duy, tranh minh họa lịch sử hoặc slide bài giảng tự động chỉ trong vòng vài phút.",
             "Giúp các tiết học trở nên sinh động, trực quan hóa các khái niệm trừu tượng, thu hút học sinh tập trung vào bài học."
         ]),
        ("3. AI hỗ trợ ra đề thi, kiểm tra và ma trận đề tự động chuẩn quy cách",
         "Quizizz AI, Gemini, Claude, kết hợp định dạng Word chuẩn hóa.",
         [
             "Giáo viên đưa nội dung bài học vào AI và yêu cầu tạo câu hỏi theo ma trận 4 mức độ: Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao; phân loại theo định dạng đề thi mới (Trắc nghiệm nhiều lựa chọn, Đúng/Sai, Điền khuyết).",
             "Tiết kiệm thời gian biên soạn ngân hàng câu hỏi, chuẩn hóa nhanh trên Word để import vào các nền tảng thi trực tuyến (Azota, Submite, Google Forms) hoặc in giấy kiểm tra trực tiếp."
         ]),
        ("4. AI làm 'Trợ lý học tập cá nhân' cho học sinh vùng cao",
         "Gemini, ChatGPT trên điện thoại thông minh hoặc máy tính phòng học bộ môn.",
         [
             "Khuyến khích học sinh sử dụng AI để giải đáp bài toán khó, giải thích thuật ngữ ngữ văn, hoặc luyện phát âm và giao tiếp tiếng Anh cơ bản.",
             "Giáo viên hướng dẫn học sinh cách đặt câu hỏi chuẩn mực, rèn luyện tư duy phản biện, biến AI thành 'người gia sư tận tụy' ngoài giờ lên lớp."
         ])
    ]

    for title_s, tools_s, steps_s in app_sections:
        p_s = doc.add_paragraph()
        set_paragraph_spacing(p_s, before_pt=6, after_pt=2)
        r_st = p_s.add_run(title_s + "\n")
        format_run(r_st, bold=True, color_rgb=(30, 64, 175))
        r_tl = p_s.add_run(f"- Công cụ gợi ý: {tools_s}\n")
        format_run(r_tl, italic=True)
        r_th = p_s.add_run("- Cách thực hiện hiệu quả tại trường:")
        format_run(r_th, bold=True)
        for st in steps_s:
            p_step = doc.add_paragraph(style='List Bullet')
            set_paragraph_spacing(p_step, 1, 2)
            r_step = p_step.add_run(st)
            format_run(r_step, size_pt=12)

    # Phần 3
    add_styled_heading(doc, "Phần 3: Thảo luận, thực hành tại chỗ và chia sẻ kinh nghiệm", 2)
    p3_items = [
        ("Thực hành mẫu (Demo trực tiếp): ", "Cán bộ phụ trách CNTT hoặc giáo viên cốt cán thực hiện thao tác trực tiếp trên máy chiếu: từ việc nhập câu lệnh (prompt) chuẩn trên Gemini/ChatGPT đến khi thu về sản phẩm hoàn chỉnh (kế hoạch bài dạy, bộ câu hỏi trắc nghiệm Word)."),
        ("Thảo luận nhóm / Trao đổi ý kiến: ", "Tập trung thảo luận 2 vấn đề then chốt: (1) Khó khăn lớn nhất khi giáo viên tiếp cận AI là gì? (Ngại công nghệ, thiếu thiết bị, đường truyền mạng yếu); (2) Môn học của tôi có thể ứng dụng AI vào khâu nào ngay trong tuần tới?"),
        ("Đúc rút nguyên tắc cốt lõi: ", "Thống nhất nguyên tắc: AI là công cụ hỗ trợ, quyết định cuối cùng vẫn thuộc về tư duy và trách nhiệm sư phạm của giáo viên. Tuyệt đối không sao chép máy móc, giáo viên phải kiểm chứng lại toàn bộ thông tin do AI cung cấp trước khi đưa vào giảng dạy.")
    ]
    for lbl, desc in p3_items:
        p = doc.add_paragraph(style='List Bullet')
        set_paragraph_spacing(p, 2, 2)
        r1 = p.add_run(lbl)
        format_run(r1, bold=True)
        r2 = p.add_run(desc)
        format_run(r2)

    # Phần 4
    add_styled_heading(doc, "Phần 4: Nâng cao trách nhiệm an toàn thông tin và đạo đức số", 2)
    p_an = doc.add_paragraph()
    set_paragraph_spacing(p_an)
    r_an = p_an.add_run("Trong bối cảnh chuyển đổi số và ứng dụng trí tuệ nhân tạo (AI) mạnh mẽ, bên cạnh những cơ hội mang lại, nhà trường cũng đối diện với không ít rủi ro về lộ thông tin, mã độc, tin giả (fake news), hay việc lạm dụng AI sai mục đích. Vì vậy, việc nâng cao nhận thức về an toàn thông tin và đạo đức số là nhiệm vụ cốt lõi, bảo đảm môi trường giáo dục an lành, văn minh.")
    format_run(r_an)

    p_qp = doc.add_paragraph()
    set_paragraph_spacing(p_qp)
    r_qp1 = p_qp.add_run("Quan điểm chỉ đạo: ")
    format_run(r_qp1, bold=True)
    r_qp2 = p_qp.add_run("“An toàn thông tin đi trước một bước; đạo đức số là gốc rễ của người học trong kỷ nguyên số”.")
    format_run(r_qp2, italic=True)

    sec4_items = [
        ("Bảo mật dữ liệu nhà trường: ", "Quản lý chặt chẽ hệ thống CSDL học sinh, điểm số, hồ sơ điện tử (vnEdu, CSDL ngành, CSDL Đảng). Phân quyền truy cập rõ ràng đối với Ban Giám hiệu, giáo viên bộ môn và nhân viên giáo vụ (tránh rò rỉ hoặc chỉnh sửa trái phép)."),
        ("Quản lý tài khoản và mật khẩu an toàn: ", "Hướng dẫn 100% giáo viên bật tính năng xác thực hai lớp (2FA) trên tài khoản công việc (Google Workspace, Microsoft 365, SmartCA, định danh VNeID, vnEdu). Cảnh báo chiêu trò lừa đảo mạo danh nhà trường, tin nhắn giả."),
        ("Phòng ngừa mã độc và virus: ", "Nhắc nhở giáo viên cẩn trọng khi cắm USB, thiết bị lưu trữ lạ vào máy tính văn phòng/phòng học thông minh để phòng chống virus, mã độc tống tiền (ransomware) làm mất dữ liệu bài giảng, đề thi."),
        ("Quy tắc ứng xử trên mạng xã hội cho học sinh: ", "Tuyên truyền thông điệp: 'Phát ngôn chuẩn mực - Không chia sẻ tin giả - Nói không với bạo lực mạng (cyberbullying)'. Đặc biệt quan tâm bảo vệ học sinh đồng bào dân tộc thiểu số trước các cạm bẫy trực tuyến."),
        ("Đạo đức trong sử dụng Trí tuệ nhân tạo (AI): ", "Đối với giáo viên: Tôn trọng bản quyền tác giả, kiểm chứng thông tin, không dùng AI thay thế trách nhiệm người thầy. Đối với học sinh: Trung thực trong học tập, dùng AI để gợi ý tư duy, cấm sao chép văn mẫu và gian lận trong kiểm tra.")
    ]
    for lbl, desc in sec4_items:
        p = doc.add_paragraph(style='List Bullet')
        set_paragraph_spacing(p, 2, 2)
        r1 = p.add_run(lbl)
        format_run(r1, bold=True)
        r2 = p.add_run(desc)
        format_run(r2)

    # Phần 5
    add_styled_heading(doc, "Phần 5: Phát huy vai trò tiên phong, gương mẫu của đảng viên trong chuyển đổi số và AI", 2)
    p_dv = doc.add_paragraph()
    set_paragraph_spacing(p_dv)
    r_dv = p_dv.add_run("Trong công cuộc chuyển đổi số và ứng dụng AI vào quản lý, dạy học tại đơn vị, yếu tố quyết định thành công không chỉ nằm ở hạ tầng công nghệ mà cốt lõi là đội ngũ con người. Đội ngũ cán bộ, đảng viên trong chi bộ nhà trường phải luôn giữ vai trò là 'ngọn đuốc đi đầu', tiên phong trong nhận thức, hành động và đổi mới sáng tạo.")
    format_run(r_dv)

    p_dv_role = doc.add_paragraph()
    set_paragraph_spacing(p_dv_role)
    r_role1 = p_dv_role.add_run("Quan điểm chính trị: ")
    format_run(r_role1, bold=True)
    r_role2 = p_dv_role.add_run("Chuyển đổi số và ứng dụng AI không chỉ là nhiệm vụ chuyên môn thuần túy mà là nhiệm vụ chính trị quan trọng của mỗi đảng viên trong thời kỳ mới. Việc ứng dụng công nghệ để nâng cao chất lượng giáo dục vùng cao chính là cụ thể hóa tinh thần đổi mới, dám nghĩ, dám làm.")
    format_run(r_role2)

    dv_contents = [
        ("Tiên phong trong tư duy và tự học nâng cao năng lực số: ", "Chủ động tự nghiên cứu các công cụ AI (Gemini, ChatGPT, Canva AI); mỗi đảng viên trực tiếp giảng dạy phải là người đầu tiên tại tổ chuyên môn đăng ký tiết dạy mẫu ứng dụng hiệu quả CNTT và AI."),
        ("Gương mẫu trong kỷ luật thông tin và đạo đức số: ", "Gương mẫu tuyệt đối trong phát ngôn mạng xã hội; không chia sẻ thông tin chưa kiểm chứng; tuân thủ nghiêm ngặt bảo mật dữ liệu nhà trường; xây dựng văn hóa mạng văn minh."),
        ("Tiên phong hỗ trợ, kèm cặp quần chúng và đồng nghiệp: ", "Thực hiện phương châm 'Đảng viên đi trước, làng nước theo sau' - các đảng viên am hiểu công nghệ sẵn sàng cầm tay chỉ việc, hỗ trợ đồng nghiệp lớn tuổi hoặc còn e ngại công nghệ cùng tiến bộ.")
    ]
    for lbl, desc in dv_contents:
        p = doc.add_paragraph(style='List Bullet')
        set_paragraph_spacing(p, 2, 2)
        r1 = p.add_run(lbl)
        format_run(r1, bold=True)
        r2 = p.add_run(desc)
        format_run(r2)

    # KẾT QUẢ, SẢN PHẨM CẦN ĐẠT SAU CHUYÊN ĐỀ
    add_styled_heading(doc, "II. KẾT QUẢ VÀ SẢN PHẨM CỤ THỂ CẦN ĐẠT SAU CHUYÊN ĐỀ", 1)
    
    # Table of products
    table_p = doc.add_table(rows=1, cols=3)
    table_p.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_p, color="B0B0B0", sz="4", val="single")
    
    hdr = table_p.rows[0].cells
    hdr[0].width = Inches(1.8)
    hdr[1].width = Inches(3.2)
    hdr[2].width = Inches(1.7)
    headers = ["Cấp thực hiện", "Danh mục sản phẩm cụ thể", "Thời hạn hoàn thành"]
    for i, h in enumerate(headers):
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(hdr[i], "1E40AF")
        set_cell_margins(hdr[i], 120, 120, 120, 120)
        r = p.add_run(h)
        format_run(r, bold=True, color_rgb=(255, 255, 255))

    prod_rows = [
        ("Cấp Nhà trường\n(Chi bộ & BGH)", 
         "- 01 Báo cáo đánh giá thực trạng và kế hoạch chiến lược ứng dụng AI, chuyển đổi số 2026-2027.\n- Ban hành Quy chế an toàn thông tin & bảo mật dữ liệu.\n- 100% cán bộ, đảng viên ký cam kết đạo đức số.",
         "Tháng 11/2026"),
        ("Cấp Tổ chuyên môn & Giáo viên",
         "- Mỗi tổ chuyên môn đóng góp 03 - 05 Kế hoạch bài dạy mẫu (CV 5512) ứng dụng AI.\n- Ngân hàng câu hỏi kiểm tra đánh giá số hóa chuẩn quy cách Word.\n- 04 - 06 tiết dạy chuyên đề minh họa ứng dụng CNTT/AI toàn trường.",
         "Học kỳ I\n(Năm học 2026-2027)"),
        ("Đối với Học sinh",
         "- 100% học sinh được hướng dẫn sử dụng trợ lý AI an toàn, trung thực.\n- Thành lập các 'Đội xung kích công nghệ số và an toàn mạng' do Đoàn trường phụ trách.",
         "Xuyên suốt năm học")
    ]
    for c1, c2, c3 in prod_rows:
        row = table_p.add_row()
        for i, text in enumerate([c1, c2, c3]):
            cell = row.cells[i]
            set_cell_margins(cell, 100, 100, 120, 120)
            p = cell.paragraphs[0]
            if i != 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, 2, 2)
            r = p.add_run(text)
            format_run(r, size_pt=11)

    # Footer Signature
    p_concl = doc.add_paragraph()
    set_paragraph_spacing(p_concl, before_pt=12, after_pt=6)
    r_c = p_concl.add_run("Trên đây là Báo cáo Chuyên đề 1: 'Phát huy vai trò tiên phong, gương mẫu của đảng viên trong chuyển đổi số và ứng dụng trí tuệ nhân tạo trong quản lý, dạy học năm học 2026 – 2027' của Chi bộ trường THPT Phục Hòa./.")
    format_run(r_c, italic=True)

    table_sign = doc.add_table(rows=1, cols=2)
    table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sign.autofit = False
    set_table_borders(table_sign, color="FFFFFF", sz="0", val="none")
    row_s = table_sign.rows[0]
    cell_sl = row_s.cells[0]
    cell_sr = row_s.cells[1]
    cell_sl.width = Inches(3.2)
    cell_sr.width = Inches(3.3)

    p_sl = cell_sl.paragraphs[0]
    set_paragraph_spacing(p_sl, 2, 2)
    r_sl = p_sl.add_run("Nơi nhận:\n- Đảng ủy viên phụ trách chi bộ;\n- Các đảng viên trong chi bộ;\n- Lưu: CB.")
    format_run(r_sl, size_pt=10, italic=True)

    p_sr = cell_sr.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_sr, 2, 2)
    r_sr1 = p_sr.add_run("T/M CHI BỘ\nBÍ THƯ\n\n\n\n")
    format_run(r_sr1, size_pt=12, bold=True)
    r_sr2 = p_sr.add_run("Hoàng Trung Kiên")
    format_run(r_sr2, size_pt=13, bold=True)

    doc.save(file_path)
    print(f"[OK] Đã tạo file: {file_path}")

# -------------------------------------------------------------
# 2. TẠO TÀI LIỆU BẢN CHUẨN QUY TRÌNH TẠO ĐỀ KIỂM TRA TRÊN WORD & AI
# -------------------------------------------------------------
def build_exam_guide_doc(file_path):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.79)

    # Header
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="FFFFFF", sz="0", val="none")
    row = table.rows[0]
    cell_l = row.cells[0]
    cell_r = row.cells[1]
    cell_l.width = Inches(3.2)
    cell_r.width = Inches(3.3)
    
    p_l = cell_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_l, 0, 2)
    r1 = p_l.add_run("SỞ GD&ĐT CAO BẰNG\n")
    format_run(r1, size_pt=11)
    r2 = p_l.add_run("TRƯỜNG THPT PHỤC HÒA\n")
    format_run(r2, size_pt=11, bold=True)
    r3 = p_l.add_run("TỔ CHUYÊN MÔN")
    format_run(r3, size_pt=10, bold=True)

    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_r, 0, 2)
    r4 = p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    format_run(r4, size_pt=11, bold=True)
    r5 = p_r.add_run("Độc lập - Tự do - Hạnh phúc\n")
    format_run(r5, size_pt=11, bold=True)
    r6 = p_r.add_run("Phục Hòa, ngày 29 tháng 10 năm 2026")
    format_run(r6, size_pt=10, italic=True)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_title, before_pt=14, after_pt=4)
    r_title1 = p_title.add_run("HƯỚNG DẪN BẢN CHUẨN\n")
    format_run(r_title1, size_pt=15, bold=True, color_rgb=(180, 20, 20))
    r_title2 = p_title.add_run("QUY TRÌNH 5 BƯỚC TẠO NHANH BÀI KIỂM TRA TRÊN WORD ỨNG DỤNG TRÍ TUỆ NHÂN TẠO (AI)\n")
    format_run(r_title2, size_pt=13, bold=True, color_rgb=(30, 64, 175))
    r_sub = p_title.add_run("Chuẩn hóa Ma trận, Bản đặc tả, Cú pháp Word và Tích hợp Khảo thí Số (Azota, Submite, Quizizz)")
    format_run(r_sub, size_pt=11, italic=True)

    # PHẦN I
    add_styled_heading(doc, "PHẦN I. MỤC TIÊU VÀ NGUYÊN TẮC CHUẨN HÓA", 1)
    p_m1 = doc.add_paragraph()
    set_paragraph_spacing(p_m1)
    r_m1 = p_m1.add_run("Quy trình này được biên soạn nhằm chuẩn hóa các bước tạo bài kiểm tra thường xuyên và định kỳ cho giáo viên trường THPT Phục Hòa, giải quyết triệt để 3 vấn đề lớn:\n"
                        "1. Tiết kiệm 70% thời gian ra đề và định dạng Word (từ 2-3 giờ xuống còn 10-15 phút).\n"
                        "2. Đảm bảo 100% đề thi bám sát ma trận và đặc tả theo định hướng phát triển phẩm chất, năng lực học sinh.\n"
                        "3. Định dạng Word chuẩn tuyệt đối để import trực tiếp 1-click vào các phần mềm khảo thí (Azota, Submite, Google Form, Quizizz, McMIX) mà không bị lỗi font hay nhảy dòng.")
    format_run(r_m1)

    # PHẦN II
    add_styled_heading(doc, "PHẦN II. QUY TRÌNH 5 BƯỚC THẦN TỐC RA ĐỀ KIỂM TRA TRÊN WORD KẾT HỢP AI", 1)

    steps = [
        ("BƯỚC 1: XÂY DỰNG MA TRẬN VÀ BẢN ĐẶC TẢ ĐỀ KIỂM TRA CHUẨN",
         "1.1. Khung tỉ lệ 4 mức độ tư duy chuẩn của Bộ GD&ĐT:\n"
         "- Mức 1: Nhận biết (40% tổng số điểm) - Nhận diện, nhắc lại kiến thức cơ bản.\n"
         "- Mức 2: Thông hiểu (30% tổng số điểm) - Giải thích, phân biệt, làm rõ khái niệm.\n"
         "- Mức 3: Vận dụng (20% tổng số điểm) - Áp dụng kiến thức giải quyết tình huống cụ thể.\n"
         "- Mức 4: Vận dụng cao (10% tổng số điểm) - Phân tích, tổng hợp, giải quyết vấn đề mới.\n\n"
         "1.2. Cấu trúc 3 Phần theo Định dạng Đề thi Mới (Từ năm 2025/2026):\n"
         "- Phần I: Câu hỏi trắc nghiệm nhiều phương án lựa chọn (Chọn 1 phương án đúng trong 4 đáp án A, B, C, D).\n"
         "- Phần II: Câu hỏi trắc nghiệm Đúng/Sai (Mỗi câu gồm 4 lệnh hỏi a, b, c, d; học sinh chọn Đúng hoặc Sai cho từng ý).\n"
         "- Phần III: Câu hỏi trắc nghiệm trả lời ngắn hoặc Tự luận (Học sinh điền kết quả hoặc giải chi tiết)."),

        ("BƯỚC 2: THIẾT KẾ PROMPT AI SƯ PHẠM (KỸ THUẬT PROMPT ENGINEERING)",
         "Để AI (Gemini, ChatGPT, Claude) không bị ảo giác và xuất đúng định dạng, giáo viên cần dùng công thức Prompt 5 thành tố: [VAI TRÒ] + [NHIỆM VỤ] + [NỘI DUNG/MA TRẬN] + [RÀO CHẮN SƯ PHẠM] + [ĐỊNH DẠNG ĐẦU RA].\n\n"
         "MẪU PROMPT CHUẨN TẠO PHẦN I (Nhiều lựa chọn):\n"
         "\"Bạn là một giáo viên chuyên môn môn [Tên Môn], hãy tạo cho tôi 10 câu trắc nghiệm 4 phương án cho chủ đề [Tên Chủ Đề] lớp [Khối Lớp]. Ma trận gồm: 4 câu Nhận biết, 3 câu Thông hiểu, 2 câu Vận dụng, 1 câu Vận dụng cao. "
         "Yêu cầu định dạng đầu ra chuẩn Microsoft Word: Mỗi câu bắt đầu bằng 'Câu X: ', 4 đáp án viết lần lượt là A. , B. , C. , D. Mỗi đáp án trên một dòng riêng. Cuối cùng lập bảng đáp án chi tiết và hướng dẫn giải.\"\n\n"
         "MẪU PROMPT CHUẨN TẠO PHẦN II (Đúng/Sai):\n"
         "\"Hãy tạo 2 câu hỏi trắc nghiệm Đúng/Sai môn [Tên Môn] chủ đề [Tên Chủ Đề]. Cấu trúc mỗi câu: Cho một ngữ cảnh/tình huống thực tiễn, sau đó đưa ra 4 mệnh đề a), b), c), d). Xác định rõ mệnh đề nào Đúng, mệnh đề nào Sai và giải thích ngắn gọn.\""),

        ("BƯỚC 3: QUY CHUẨN SOẠN THẢO VÀ ĐỊNH DẠNG TRÊN MICROSOFT WORD",
         "3.1. Thiết lập trang Word chuẩn (Page Setup):\n"
         "- Cỡ giấy: A4 (210 x 297 mm).\n"
         "- Căn lề chuẩn: Top: 2cm, Bottom: 2cm, Left: 2.5cm, Right: 2cm.\n"
         "- Phông chữ: Times New Roman, Cỡ chữ 12pt (hoặc 13pt), Màu chữ: Tự động (Black).\n"
         "- Dãn dòng: Line spacing: 1.15 lines, Spacing Before: 0pt, After: 3pt.\n\n"
         "3.2. Cú pháp viết câu hỏi và đáp án bắt buộc để phần mềm tự động nhận diện:\n"
         "- Câu hỏi: Bắt đầu bằng chữ 'Câu 1:' hoặc 'Câu 1.' (in đậm).\n"
         "- Các phương án lựa chọn: 'A. ', 'B. ', 'C. ', 'D. ' (chữ cái in hoa, có dấu chấm và một dấu cách liền kề).\n"
         "- Đánh dấu đáp án đúng: Sử dụng gạch chân dưới chữ cái đáp án đúng (Ví dụ: <u>A.</u>) HOẶC tạo bảng đáp án chuẩn ở cuối trang Word.\n"
         "- Lưu ý: Không dùng số thứ tự tự động (AutoNumbering) của Word vì khi import vào Azota/Submite dễ bị nhảy số."),

        ("BƯỚC 4: MẸO VÀ PHÍM TẮT WORD XỬ LÝ NHANH THẦN TỐC TRONG 30 GIÂY",
         "Sử dụng hộp thoại Find & Replace (Ctrl + H) để chuẩn hóa đề thi siêu tốc:\n"
         "1. Xóa toàn bộ khoảng trắng thừa trong bài thi:\n"
         "   - Find what: ^w\n"
         "   - Replace with: (gõ 1 phím cách)\n"
         "   - Bấm Replace All (Loại bỏ triệt để các lỗi thụt đầu dòng, cách dòng lệch).\n"
         "2. Xóa các dòng trống thừa:\n"
         "   - Find what: ^p^p\n"
         "   - Replace with: ^p\n"
         "   - Bấm Replace All 2-3 lần.\n"
         "3. Chuẩn hóa đáp án thành 4 cột bằng bảng (Table):\n"
         "   - Nếu muốn 4 đáp án A, B, C, D nằm ngang thẳng hàng: Tạo bảng 1 hàng 4 cột, dán 4 đáp án vào rồi chọn chế độ No Border (Ẩn viền bảng). Đảm bảo in ra vô cùng đẹp mắt và tiết kiệm giấy in!"),

        ("BƯỚC 5: KIỂM SOÁT TÍNH CHUẨN XÁC VÀ ĐẨY LÊN NỀN TẢNG KHẢO THÍ SỐ",
         "5.1. Rào chắn sư phạm (Teacher Verification):\n"
         "- Kiểm tra công thức, thuật ngữ, số liệu đảm bảo tính chính xác khoa học.\n"
         "- Điều chỉnh độ khó phù hợp với đặc thù học sinh trường THPT Phục Hòa.\n\n"
         "5.2. Chuyển đổi và Khảo thí đa kênh:\n"
         "- Kênh 1: In trực tiếp từ Word ra giấy kiểm tra cho học sinh làm bài tại lớp.\n"
         "- Kênh 2: Import file Word (.docx) vào hệ thống Azota / Submite chỉ bằng 1 thao tác kéo thả.\n"
         "- Kênh 3: Dùng ứng dụng chấm trắc nghiệm bằng camera điện thoại (Azota, QM, ZipGrade) - tự động chấm điểm và xuất bảng điểm Excel vào sổ điểm điện tử.")
    ]

    for title_st, desc_st in steps:
        add_styled_heading(doc, title_st, 2)
        p = doc.add_paragraph()
        set_paragraph_spacing(p)
        r = p.add_run(desc_st)
        format_run(r, size_pt=11.5)

    # PHẦN III: ĐỀ THI MẪU MINH HỌA
    add_styled_heading(doc, "PHẦN III. BẢN MẪU ĐỀ KIỂM TRA CHUẨN QUY CÁCH TRÊN WORD", 1)
    
    p_demo_box = doc.add_paragraph()
    set_paragraph_spacing(p_demo_box, before_pt=4, after_pt=4)
    r_demo_box = p_demo_box.add_run("Dưới đây là cấu trúc minh họa một đề kiểm tra chuẩn quy cách (Môn Tin học 10 - KNTT) sẵn sàng in ấn hoặc đẩy lên Azota:")
    format_run(r_demo_box, italic=True)

    # Box sample
    table_sample = doc.add_table(rows=1, cols=1)
    table_sample.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_box = table_sample.rows[0].cells[0]
    cell_box.width = Inches(6.5)
    set_cell_background(cell_box, "F8FAFC")
    set_cell_margins(cell_box, 150, 150, 150, 150)
    set_table_borders(table_sample, color="94A3B8", sz="6", val="single")

    p_in_box = cell_box.paragraphs[0]
    set_paragraph_spacing(p_in_box, 2, 2)
    sample_text = (
        "SỞ GD&ĐT CAO BẰNG                           KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027\n"
        "TRƯỜNG THPT PHỤC HÒA                        Môn: TIN HỌC - KHỐI 10\n"
        "ĐỀ KIỂM TRA CHÍNH THỨC                     Thời gian làm bài: 45 phút\n"
        "Họ và tên học sinh: .................................................... Lớp: 10A...   SBD: ............\n"
        "--------------------------------------------------------------------------------------------------\n\n"
        "PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (6,0 điểm)\n\n"
        "Câu 1: Thiết bị nào sau đây là thiết bị vào (Input) của máy tính?\n"
        "A. Màn hình máy tính.\n"
        "B. Máy in laser.\n"
        "C. Bàn phím và chuột.\n"
        "D. Tai nghe.\n\n"
        "Câu 2: Phát biểu nào sau đây thể hiện đúng nhất về vai trò của Trí tuệ nhân tạo (AI)?\n"
        "A. AI có thể thay thế hoàn toàn vai trò của con người trong mọi lĩnh vực.\n"
        "B. AI là công cụ hỗ trợ thông minh, giúp con người nâng cao năng suất và chất lượng công việc.\n"
        "C. AI chỉ được ứng dụng trong các trò chơi điện tử.\n"
        "D. AI không có bất kỳ rủi ro nào về an toàn thông tin.\n\n"
        "PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG / SAI (4,0 điểm)\n\n"
        "Câu 3: Khi tham gia không gian mạng và sử dụng các công cụ AI, học sinh cần tuân thủ quy tắc:\n"
        "a) Tuyệt đối không chia sẻ thông tin cá nhân, mật khẩu tài khoản cho người lạ.\n"
        "b) Được phép sao chép nguyên văn bài văn của AI vào bài làm để nộp lấy điểm cao.\n"
        "c) Luôn kiểm chứng lại thông tin do AI cung cấp trước khi sử dụng.\n"
        "d) Sử dụng mạng xã hội văn minh, nói không với tin giả và bạo lực học đường trực tuyến.\n\n"
        "-------------------------------------- HẾT --------------------------------------\n\n"
        "BẢNG ĐÁP ÁN CHUẨN:\n"
        "Phần I: 1-C; 2-B.\n"
        "Phần II: Câu 3: a - Đúng; b - Sai; c - Đúng; d - Đúng."
    )
    r_smp = p_in_box.add_run(sample_text)
    format_run(r_smp, font_name="Courier New", size_pt=10)

    # Footer
    p_end = doc.add_paragraph()
    set_paragraph_spacing(p_end, before_pt=12, after_pt=4)
    r_end = p_end.add_run("Tổ Chuyên môn Trường THPT Phục Hòa cam kết đồng hành, hỗ trợ cầm tay chỉ việc cho 100% giáo viên trong nhà trường ứng dụng thành thạo quy trình này trong năm học 2026 - 2027.")
    format_run(r_end, italic=True)

    doc.save(file_path)
    print(f"[OK] Đã tạo file: {file_path}")

if __name__ == "__main__":
    out_dir = r"d:\Du-an-web\web-sinh-hoat-chuyen-de-1-phuc-hoa"
    desktop_dir = r"D:\Desktop"
    
    f1_name = "Bao_Cao_Sinh_Hoat_Chuyen_De_1_THPT_Phuc_Hoa.docx"
    f2_name = "Ban_Chuan_Quy_Trinh_Tao_De_Kiem_Tra_Word_AI.docx"
    
    p1 = os.path.join(out_dir, f1_name)
    p2 = os.path.join(out_dir, f2_name)
    
    build_report_doc(p1)
    build_exam_guide_doc(p2)
    
    # Copy to Desktop
    shutil.copy(p1, os.path.join(desktop_dir, f1_name))
    shutil.copy(p2, os.path.join(desktop_dir, f2_name))
    print(f"[OK] Đã sao chép 02 file Word ra màn hình Desktop ({desktop_dir})!")
