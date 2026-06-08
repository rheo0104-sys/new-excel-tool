import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import re

# --- [로고 컬러 테마 정의] ---
COLOR_BG = "#F4F7F9"        # 가볍고 깔끔한 배경 라이트 그레이/블루
COLOR_WHITE = "#FFFFFF"     # 카드 배경
COLOR_PRIMARY = "#00AEEF"   # Dream Blue (로고 메인 컬러)
COLOR_SECONDARY = "#00A99D" # Teal Green (로고 포인트 컬러)
COLOR_TEXT = "#333333"      # 텍스트 다크 그레이
COLOR_BORDER = "#D1D9E0"    # 부드러운 테두리 색상

def process_logic(source_path, ref_path):
    try:
        source_wb = openpyxl.load_workbook(source_path)
        ref_df = pd.read_excel(ref_path)

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_font = Font(color="FF0000", bold=True)

        # ---------------------------------------------------------
        # [데이터 매칭 최적화] "비교 기준" 캐싱 (DOMAIN/VARIABLE_ID vs Entry ID/Sub Item)
        # ---------------------------------------------------------
        ref_data_map = {}
        for _, r_row in ref_df.iterrows():
            eid = str(r_row.get('Entry ID', '')).strip()
            sub = str(r_row.get('Sub Item', '')).strip()
            dmq = str(r_row.get('Default Missing Query', '')).strip().upper()
            
            비교_기준 = f"{eid}/{sub}"
            ref_data_map[비교_기준] = {"DMQ": dmq, "Sub Item": sub}
            
            if eid and not sub:
                ref_data_map[f"{eid}/"] = {"DMQ": dmq, "Sub Item": ""}

        # ---------------------------------------------------------
        # [헤더 인식 강화] 줄바꿈, 대소문자, 한글/영문 혼용 처리
        # ---------------------------------------------------------
        def get_col_idx(headers, target_name):
            target = re.sub(r'\s+', '', target_name.strip().upper())
            for i, h in enumerate(headers):
                if h:
                    cleaned_h = re.sub(r'\s+', '', str(h)).upper()
                    if target in cleaned_h:
                        return i
            raise ValueError(f"'{target_name}' 열을 찾을 수 없습니다.")

        # --- [모듈 1: DB Specification] ---
        if "DB Specification" in source_wb.sheetnames:
            ws1 = source_wb["DB Specification"]
            h1 = [cell.value for cell in ws1[1]] 
            
            try:
                v1_idx = get_col_idx(h1, 'VARIABLE_ID')
                l1_idx = get_col_idx(h1, 'LAYOUT')
                m1_idx = get_col_idx(h1, 'Mandatory')
                d1_idx = get_col_idx(h1, 'DOMAIN')
            except ValueError as e:
                return f"DB Specification 시트 헤더 오류: {str(e)}"

            for row in ws1.iter_rows(min_row=2):
                domain = str(row[d1_idx].value).strip() if row[d1_idx].value else ""
                vid = str(row[v1_idx].value).strip() if row[v1_idx].value else ""
                layout = str(row[l1_idx].value).strip() if row[l1_idx].value else ""
                target = row[m1_idx]
                
                old_val = str(target.value).strip() if target.value else ""
                new_val = old_val

                비교_기준_원본 = f"{domain}/{vid}"

                if layout == "SYSDEFINED":
                    new_val = "N"
                elif 비교_기준_원본 in ref_data_map:
                    ref = ref_data_map[비교_기준_원본]
                    if "." in ref["Sub Item"]:
                        new_val = "Y/N"
                    else:
                        if ref["DMQ"] == "YES": new_val = "Y"
                        elif ref["DMQ"] == "NO": new_val = "N"

                if old_val != new_val:
                    target.value = new_val
                    target.fill, target.font = yellow_fill, red_font
        else:
            return "오류: 필수 시트 'DB Specification'이 원본에 없습니다."

        # --- [모듈 2: DB Specification_CAT] ---
        if "DB Specification_CAT" in source_wb.sheetnames:
            ws2 = source_wb["DB Specification_CAT"]
            h2 = [cell.value for cell in ws2[1]]
            
            try:
                v2_idx = get_col_idx(h2, 'VARIABLE_ID')
                s2_idx = get_col_idx(h2, 'SUB_ITEM')
                m2_idx = get_col_idx(h2, 'Mandatory')
                d2_idx = get_col_idx(h2, 'DOMAIN')
            except ValueError as e:
                pass 
            else:
                for row in ws2.iter_rows(min_row=2):
                    domain_val = str(row[d2_idx].value).strip() if row[d2_idx].value else ""
                    vid_val = str(row[v2_idx].value).strip() if row[v2_idx].value else ""
                    sub_val = str(row[s2_idx].value).strip() if row[s2_idx].value else ""
                    
                    비교_기준_원본 = f"{domain_val}/{vid_val}"
                    if sub_val:
                         비교_기준_원본 = f"{domain_val}/{sub_val}"
                    
                    target = row[m2_idx]
                    old_val = str(target.value).strip() if target.value else ""
                    new_val = old_val

                    if 비교_기준_원본 in ref_data_map:
                        ref = ref_data_map[비교_기준_원본]
                        if ref["DMQ"] == "YES": new_val = "Y"
                        elif ref["DMQ"] == "NO": new_val = "N"

                    if old_val != new_val:
                        target.value = new_val
                        target.fill, target.font = yellow_fill, red_font

        output_name = f"{os.path.splitext(source_path)[0]}_updated.xlsx"
        source_wb.save(output_name)
        return output_name

    except Exception as e:
        return f"실행 중 오류 발생: {str(e)}"

# --- [UI 디자인 고도화 (DreamCIS Theme)] ---
root = TkinterDnD.Tk()
root.title("DreamCIS DB Specification Optimizer")
root.geometry("640x520")
root.configure(bg=COLOR_BG)
root.resizable(False, False)

# 상단 로고 영역
header_frame = tk.Frame(root, bg=COLOR_WHITE, height=60)
header_frame.pack(fill="x", side="top")
header_frame.pack_propagate(False)

logo_label_1 = tk.Label(header_frame, text="dream", font=('Arial', 24, 'bold'), fg=COLOR_PRIMARY, bg=COLOR_WHITE)
logo_label_1.pack(side="left", padx=(20, 0), pady=10)
logo_label_2 = tk.Label(header_frame, text="cis", font=('Arial', 24, 'bold'), fg=COLOR_SECONDARY, bg=COLOR_WHITE)
logo_label_2.pack(side="left", pady=10)

title_desc = tk.Label(header_frame, text="Data Management System", font=('Arial', 10), fg="#A0A0A0", bg=COLOR_WHITE)
title_desc.pack(side="right", padx=20, pady=20)

# 메인 프레임
main_frame = tk.Frame(root, bg=COLOR_BG)
main_frame.pack(fill="both", expand=True, padx=30, pady=20)

# 작업 모드 선택
func_frame = tk.Frame(main_frame, bg=COLOR_BG)
func_frame.pack(fill="x", pady=(0, 20))

tk.Label(func_frame, text="● 작업 모드", font=('맑은 고딕', 10, 'bold'), fg=COLOR_TEXT, bg=COLOR_BG).pack(anchor="w", pady=(0, 5))

style = ttk.Style()
style.theme_use('clam')
style.configure("TCombobox", padding=5, relief="flat", borderwidth=1, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER)
mode_combo = ttk.Combobox(func_frame, values=["DB Specification Mandatory Auto-Update"], state="readonly", font=('맑은 고딕', 10))
mode_combo.pack(fill="x")
mode_combo.current(0)

# 드래그 앤 드롭 영역 (컨테이너)
drop_container = tk.Frame(main_frame, bg=COLOR_BG)
drop_container.pack(fill="x", pady=10)

path_source, path_ref = "", ""

def create_drop_box(parent, title, drop_func, side_pack):
    # 가로(좌우) 배치를 위해 고정 너비 설정
    border_frame = tk.Frame(parent, bg=COLOR_BORDER, width=280, height=160)
    border_frame.pack_propagate(False)
    # side_pack 파라미터에 따라 left / right로 밀착 정렬 수행
    border_frame.pack(side=side_pack, padx=(0, 10) if side_pack=="left" else (10, 0))
    
    inner_frame = tk.Frame(border_frame, bg=COLOR_WHITE)
    inner_frame.pack(fill="both", expand=True, padx=1, pady=1)
    
    tk.Label(inner_frame, text=title, font=('맑은 고딕', 11, 'bold'), fg=COLOR_TEXT, bg=COLOR_WHITE).pack(pady=(18, 5))
    
    icon_label = tk.Label(inner_frame, text="⊕", font=('Arial', 32), fg=COLOR_PRIMARY, bg=COLOR_WHITE)
    icon_label.pack(expand=True)
    
    name_label = tk.Label(inner_frame, text="이곳으로 엑셀 파일을 드래그하세요", font=('맑은 고딕', 9), fg="#95A5A6", bg=COLOR_WHITE, wraplength=250)
    name_label.pack(side="bottom", pady=15)
    
    inner_frame.drop_target_register(DND_FILES)
    inner_frame.dnd_bind('<<Drop>>', lambda e: drop_func(e, border_frame, icon_label, name_label))
    
    return inner_frame

def drop_src(e, border, icon, name):
    global path_source
    path_source = e.data.strip('{?}')
    border.config(bg=COLOR_SECONDARY)
    icon.config(text="✓", fg=COLOR_SECONDARY)
    name.config(text=os.path.basename(path_source), fg=COLOR_TEXT, font=('맑은 고딕', 9, 'bold'))

def drop_ref(e, border, icon, name):
    global path_ref
    path_ref = e.data.strip('{?}')
    border.config(bg=COLOR_SECONDARY)
    icon.config(text="✓", fg=COLOR_SECONDARY)
    name.config(text=os.path.basename(path_ref), fg=COLOR_TEXT, font=('맑은 고딕', 9, 'bold'))

# 변경 요청 사항 적용: 라벨명 변경 및 좌(left)/우(right) 구조 배치 명시
create_drop_box(drop_container, "DB specification", drop_src, "left")
create_drop_box(drop_container, "Entry search", drop_ref, "right")

# 실행 버튼
def run_process():
    if not path_source or not path_ref:
        messagebox.showwarning("파일 확인", "모든 파일을 드롭해주세요.")
        return
        
    btn_run.config(text="업데이트 중...", state="disabled", bg="#95A5A6")
    root.update()
    
    res = process_logic(path_source, path_ref)
    
    if res.endswith(".xlsx"):
        messagebox.showinfo("작업 완료", f"작업이 성공적으로 완료되었습니다.\n\n결과 파일:\n{os.path.basename(res)}")
    else:
        messagebox.showerror("오류 발생", res)
        
    btn_run.config(text="업데이트 실행하기", state="normal", bg=COLOR_PRIMARY)

btn_run = tk.Button(main_frame, text="업데이트 실행하기", command=run_process, 
                    bg=COLOR_PRIMARY, fg=COLOR_WHITE, font=('맑은 고딕', 12, 'bold'), 
                    relief="flat", activebackground=COLOR_SECONDARY, activeforeground=COLOR_WHITE, cursor="hand2", pady=12)
btn_run.pack(fill="x", pady=(25, 0))

root.mainloop()
