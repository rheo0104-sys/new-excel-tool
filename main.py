import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def process_logic(source_path, ref_path, mode):
    try:
        source_wb = openpyxl.load_workbook(source_path)
        ref_df = pd.read_excel(ref_path)

        # 비교 기준: DOMAIN/VARIABLE_ID vs Entry ID/Sub Item
        ref_df['match_key'] = ref_df['Entry ID'].astype(str) + "/" + ref_df['Sub Item'].astype(str)
        ref_dict = dict(zip(ref_df['match_key'], ref_df['Status']))

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_bold_font = Font(color="FF0000", bold=True)

        # 원본파일의 시트 2, 3 작업
        for i in [1, 2]:
            sheet = source_wb.worksheets[i]
            headers = [cell.value for cell in sheet[1]]
            
            d_idx = headers.index('DOMAIN')
            v_idx = headers.index('VARIABLE_ID')
            t_idx = headers.index('VAL')

            for row in sheet.iter_rows(min_row=2):
                m_key = f"{row[d_idx].value}/{row[v_idx].value}"
                if m_key in ref_dict:
                    cell = row[t_idx]
                    cell.value = ref_dict[m_key]
                    cell.fill = yellow_fill
                    cell.font = red_bold_font

        # --- 파일명 설정 부분 ---
        base_name = os.path.splitext(source_path)[0] # 확장자 제거
        output_name = f"{base_name}_updated.xlsx"    # _updated 붙이기
        # -----------------------
        
        source_wb.save(output_name)
        return output_name
    except Exception as e:
        return str(e)

# --- GUI 레이아웃 (그림 반영) ---
root = TkinterDnD.Tk()
root.title("Excel Professional Tool")
root.geometry("600x450")

# ① 기능 선택
tk.Label(root, text="① 수행할 기능을 선택하세요", font=('Arial', 10, 'bold')).pack(pady=10)
mode_combo = ttk.Combobox(root, values=["데이터 업데이트", "값 비교 검증"], state="readonly", width=50)
mode_combo.pack()
mode_combo.current(0)

file_frame = tk.Frame(root)
file_frame.pack(pady=20)

# ② 원본파일 (2번 파일)
frame2 = tk.LabelFrame(file_frame, text="② 원본파일 드래그", width=250, height=150)
frame2.pack_propagate(False)
frame2.pack(side="left", padx=10)
label2 = tk.Label(frame2, text="원본 파일을 던지세요", fg="gray")
label2.pack(expand=True)

# ③ 참고파일 (3번 파일)
frame3 = tk.LabelFrame(file_frame, text="③ 참고파일 드래그", width=250, height=150)
frame3.pack_propagate(False)
frame3.pack(side="left", padx=10)
label3 = tk.Label(frame3, text="참고 파일을 던지세요", fg="gray")
label3.pack(expand=True)

path_source = ""
path_ref = ""

def handle_drop_source(event):
    global path_source
    path_source = event.data.strip('{?}')
    label2.config(text=os.path.basename(path_source), fg="blue")

def handle_drop_ref(event):
    global path_ref
    path_ref = event.data.strip('{?}')
    label3.config(text=os.path.basename(path_ref), fg="blue")

frame2.drop_target_register(DND_FILES)
frame2.dnd_bind('<<Drop>>', handle_drop_source)
frame3.drop_target_register(DND_FILES)
frame3.dnd_bind('<<Drop>>', handle_drop_ref)

# ④ 실행 버튼
def start_action():
    if not path_source or not path_ref:
        messagebox.showwarning("알림", "파일을 모두 입력해주세요.")
        return
    res = process_logic(path_source, path_ref, mode_combo.get())
    if res.endswith(".xlsx"):
        messagebox.showinfo("성공", f"작업 완료!\n파일명: {os.path.basename(res)}")
    else:
        messagebox.showerror("오류", f"에러: {res}")

btn_run = tk.Button(root, text="④ 실행하기", command=start_action, bg="#2ecc71", fg="white", font=('Arial', 12, 'bold'), width=15)
btn_run.pack(pady=20)

root.mainloop()
