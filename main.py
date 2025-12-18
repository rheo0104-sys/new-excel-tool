import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def process_logic(source_path, ref_path):
    try:
        # 1. 파일 로드
        source_wb = openpyxl.load_workbook(source_path)
        ref_df = pd.read_excel(ref_path)

        # 서식 설정 (노란 배경, 빨간 굵은 글씨)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_font = Font(color="FF0000", bold=True)

        # 참고파일 데이터 캐싱 (속도 및 비교 최적화)
        ref_data_map = {}
        for _, r_row in ref_df.iterrows():
            eid = str(r_row['Entry ID']).strip()
            sub = str(r_row['Sub Item']).strip()
            dmq = str(r_row['Default Missing Query']).strip().upper()
            
            # 모듈 1용 (ID 기준), 모듈 2용 (ID/SubItem 기준)
            ref_data_map[eid] = {"Sub Item": sub, "DMQ": dmq}
            ref_data_map[f"{eid}/{sub}"] = {"DMQ": dmq}

        # 열 인덱스 찾기 함수 (대소문자/공백 무시)
        def get_col_idx(headers, target_name):
            target = target_name.strip().upper()
            for i, h in enumerate(headers):
                if h and str(h).strip().upper() == target:
                    return i
            raise ValueError(f"'{target_name}' 열을 찾을 수 없습니다.")

        # --- [모듈 1: DB Specification - 필수] ---
        if "DB Specification" in source_wb.sheetnames:
            ws1 = source_wb["DB Specification"]
            h1 = [cell.value for cell in ws1[1]] # 1행 헤더
            
            try:
                v1_idx = get_col_idx(h1, 'VARIABLE_ID')
                l1_idx = get_col_idx(h1, 'LAYOUT')
                m1_idx = get_col_idx(h1, 'Mandatory')
            except ValueError as e:
                return f"DB Specification 시트 오류: {str(e)}"

            for row in ws1.iter_rows(min_row=2):
                vid = str(row[v1_idx].value).strip() if row[v1_idx].value else ""
                layout = str(row[l1_idx].value).strip() if row[l1_idx].value else ""
                target = row[m1_idx]
                old_val = str(target.value).strip() if target.value else ""
                new_val = old_val

                # [모듈 1 로직 순서]
                # 1. LAYOUT 확인
                if layout == "SYSDEFINED":
                    new_val = "N"
                # 2. VARIABLE_ID 매칭 확인
                elif vid in ref_data_map:
                    ref = ref_data_map[vid]
                    # Sub Item에 "." 포함 여부
                    if "." in ref["Sub Item"]:
                        new_val = "Y/N"
                    # Default Missing Query 변환
                    else:
                        if ref["DMQ"] == "YES": new_val = "Y"
                        elif ref["DMQ"] == "NO": new_val = "N"

                # 값 변경 시 서식 적용
                if old_val != new_val:
                    target.value = new_val
                    target.fill, target.font = yellow_fill, red_font
        else:
            return "오류: 필수 시트 'DB Specification'이 원본에 없습니다."

        # --- [모듈 2: DB Specification_CAT - 선택] ---
        if "DB Specification_CAT" in source_wb.sheetnames:
            ws2 = source_wb["DB Specification_CAT"]
            h2 = [cell.value for cell in ws2[1]]
            
            try:
                v2_idx = get_col_idx(h2, 'VARIABLE_ID')
                s2_idx = get_col_idx(h2, 'SUB_ITEM')
                m2_idx = get_col_idx(h2, 'Mandatory')
            except ValueError as e:
                return f"DB Specification_CAT 시트 오류: {str(e)}"

            for row in ws2.iter_rows(min_row=2):
                vid_val = str(row[v2_idx].value).strip() if row[v2_idx].value else ""
                sub_val = str(row[s2_idx].value).strip() if row[s2_idx].value else ""
                vid_sub_key = f"{vid_val}/{sub_val}"
                
                target = row[m2_idx]
                old_val = str(target.value).strip() if target.value else ""
                new_val = old_val

                # [모듈 2 로직] ID/SubItem 키 매칭 시 DMQ 변환
                if vid_sub_key in ref_data_map:
                    ref = ref_data_map[vid_sub_key]
                    if ref["DMQ"] == "YES": new_val = "Y"
                    elif ref["DMQ"] == "NO": new_val = "N"

                if old_val != new_val:
                    target.value = new_val
                    target.fill, target.font = yellow_fill, red_font

        # 3. 결과 저장
        output_name = f"{os.path.splitext(source_path)[0]}_updated.xlsx"
        source_wb.save(output_name)
        return output_name

    except Exception as e:
        return f"실행 중 알 수 없는 오류: {str(e)}"

# --- [GUI 구성: 요청하신 레이아웃 반영] ---
root = TkinterDnD.Tk()
root.title("DB Specification Update Tool")
root.geometry("600x450")

# ① 기능 선택
tk.Label(root, text="① 기능 선택", font=('Arial', 10, 'bold')).pack(pady=10)
mode_combo = ttk.Combobox(root, values=["DB specification_mandatory update"], state="readonly", width=50)
mode_combo.pack(); mode_combo.current(0)

file_frame = tk.Frame(root)
file_frame.pack(pady=20)

path_source, path_ref = "", ""

def create_drop_box(parent, title, drop_func):
    frame = tk.LabelFrame(parent, text=title, width=250, height=180)
    frame.pack_propagate(False); frame.pack(side="left", padx=15)
    label = tk.Label(frame, text="⊕", font=('Arial', 40), fg="#bdc3c7")
    label.pack(expand=True)
    name_label = tk.Label(frame, text="", font=('Arial', 8), fg="blue", wraplength=200)
    name_label.pack(side="bottom", pady=5)
    frame.drop_target_register(DND_FILES)
    frame.dnd_bind('<<Drop>>', lambda e: drop_func(e, label, name_label))
    return frame

def drop_src(e, l, n):
    global path_source
    path_source = e.data.strip('{?}')
    l.config(text="✔", fg="#2ecc71")
    n.config(text=os.path.basename(path_source))

def drop_ref(e, l, n):
    global path_ref
    path_ref = e.data.strip('{?}')
    l.config(text="✔", fg="#2ecc71")
    n.config(text=os.path.basename(path_ref))

create_drop_box(file_frame, "② 원본파일", drop_src)
create_drop_box(file_frame, "③ 참고파일", drop_ref)

def run():
    if not path_source or not path_ref:
        messagebox.showwarning("알림", "두 파일을 모두 입력해주세요.")
        return
    res = process_logic(path_source, path_ref)
    if res.endswith(".xlsx"):
        messagebox.showinfo("성공", f"업데이트 완료!\n결과 파일: {os.path.basename(res)}")
    else:
        messagebox.showerror("확인 필요", res)

tk.Button(root, text="④ 실행하기", command=run, bg="#2ecc71", fg="white", font=('Arial', 12, 'bold'), width=15).pack(pady=30)

root.mainloop()
