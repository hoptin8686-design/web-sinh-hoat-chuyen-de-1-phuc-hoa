@echo off
chcp 65001 >nul
title TỰ ĐỘNG ĐẨY WEB SINH HOẠT CHUYÊN ĐỀ 1 LÊN GITHUB & VERCEL - THPT PHỤC HÒA
color 0A

echo ======================================================================
echo   HỆ THỐNG TỰ ĐỘNG ĐỒNG BỘ GITHUB & VERCEL
echo   Nội dung Sinh hoạt Chuyên đề 1: Phát huy vai trò tiên phong của đảng viên
echo   trong chuyển đổi số và ứng dụng AI - Bản chuẩn tạo đề kiểm tra trên Word
echo   Trường THPT Phục Hòa - Tỉnh Cao Bằng
echo ======================================================================
echo.

set "PATH=C:\Program Files\Python39;C:\Users\DMX HOA THUAN\AppData\Local\Programs\Python\Python39;C:\Users\DMX HOA THUAN\AppData\Local\Microsoft\WindowsApps;C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd;C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin;%PATH%"
cd /d "d:\Du-an-web\web-sinh-hoat-chuyen-de-1-phuc-hoa"

echo [1/3] Đang cập nhật 02 file Word ra Desktop và kiểm tra mã nguồn...
python deploy.py

echo.
echo ======================================================================
echo   ✓ ĐÃ ĐỒNG BỘ THÀNH CÔNG LÊN GITHUB VÀ VERCEL!
echo ======================================================================
echo.
echo [LINK TRUY CẬP TRỰC TUYẾN]:
echo   1. Kho mã nguồn GitHub: https://github.com/hoptin8686-design/web-sinh-hoat-chuyen-de-1-phuc-hoa
echo   2. Link Web Vercel:     https://web-sinh-hoat-chuyen-de-1-phuc-hoa.vercel.app/
echo   3. Vercel Project:      https://vercel.com/new/clone?repository-url=https%%3A%%2F%%2Fgithub.com%%2Fhoptin8686-design%%2Fweb-sinh-hoat-chuyen-de-1-phuc-hoa
echo.
echo [02 FILE WORD ĐÃ CẬP NHẬT TRÊN MÀN HÌNH MÁY TÍNH (DESKTOP)]:
echo   - D:\Desktop\Ban_Chuan_Quy_Trinh_Tao_De_Kiem_Tra_Word_AI.docx
echo   - D:\Desktop\Bao_Cao_Sinh_Hoat_Chuyen_De_1_THPT_Phuc_Hoa.docx
echo.
echo Bấm phím bất kỳ để kết thúc...
pause >nul
