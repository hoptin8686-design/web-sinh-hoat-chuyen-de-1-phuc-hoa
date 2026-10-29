# -*- coding: utf-8 -*-
"""
Script tự động đẩy Web Sinh Hoạt Chuyên Đề 1 lên GitHub & Vercel
Chi Bộ Trường THPT Phục Hòa
"""
import os
import subprocess
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

GIT = r'C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd\git.exe'
GH = r'C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin\gh.exe'
CWD = r'd:\Du-an-web\web-sinh-hoat-chuyen-de-1-phuc-hoa'

env = os.environ.copy()
env['PATH'] = os.path.dirname(GIT) + ';' + os.path.dirname(GH) + ';' + env.get('PATH', '')

def run_cmd(args):
    print("RUNNING:", " ".join(args))
    r = subprocess.run(args, cwd=CWD, env=env, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.stdout:
        print("STDOUT:", r.stdout.strip())
    if r.stderr:
        print("STDERR:", r.stderr.strip())
    return r.returncode

print("=== [1/4] TẠO / CẬP NHẬT 02 TÀI LIỆU WORD CHUẨN RA MÀN HÌNH MÁY TÍNH ===")
subprocess.run([sys.executable, os.path.join(CWD, 'create_documents.py')], cwd=CWD, env=env)

print("\n=== [2/4] KHỞI TẠO VÀ CẤU HÌNH GIT REPOSITORY ===")
run_cmd([GIT, 'init'])
run_cmd([GIT, 'config', 'user.name', 'hoptin8686-design'])
run_cmd([GIT, 'config', 'user.email', 'hoptin8686@gmail.com'])
run_cmd([GIT, 'branch', '-M', 'main'])

print("\n=== [3/4] THÊM TẬP TIN VÀ TẠO BẢN LƯU (COMMIT) ===")
run_cmd([GIT, 'add', '.'])
commit_code = run_cmd([GIT, 'commit', '-m', 'feat: Khoi tao Web Noi dung Sinh hoat Chuyen de 1 va Ban chuan tao de kiem tra Word bang AI - THPT Phuc Hoa'])

print("\n=== [4/4] TẠO VÀ ĐẨY LÊN GITHUB (TỰ ĐỘNG ĐẨY SANG VERCEL) ===")
repo_name = 'web-sinh-hoat-chuyen-de-1-phuc-hoa'
create_code = run_cmd([GH, 'repo', 'create', repo_name, '--public', '--source=.', '--remote=origin', '--push'])

if create_code != 0:
    print("Repo có thể đã tồn tại, tiến hành push trực tiếp...")
    run_cmd([GIT, 'remote', 'set-url', 'origin', f'https://github.com/hoptin8686-design/{repo_name}.git'])
    run_cmd([GIT, 'push', '-u', 'origin', 'main'])

print("\n======================================================================")
print("  ✓ ĐÃ ĐẨY LÊN GITHUB THÀNH CÔNG!")
print("  ✓ KHO MÃ NGUỒN: https://github.com/hoptin8686-design/" + repo_name)
print("  ✓ LINK WEB TRỰC TUYẾN: https://" + repo_name + ".vercel.app/")
print("  ✓ FILE WORD TRÊN DESKTOP: D:\\Desktop\\Ban_Chuan_Quy_Trinh_Tao_De_Kiem_Tra_Word_AI.docx")
print("======================================================================")
