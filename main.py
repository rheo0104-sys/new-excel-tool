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

        # 비교 기준: DOMAIN/VARIABLE_ID (원본) vs Entry ID/Sub Item (참고)
        ref_df['match_key'] = ref_df['Entry ID'].astype(str) + "/" + ref_df['Sub Item'].astype(str)
        # 기능 1: DB specification_mandatory update (Status 값을 가져와 업데이트)
        ref_dict = dict(zip(ref_df['match_key'], ref_df['Status']))

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_bold_font = Font(color="FF0000", bold=True)

        # 원본파일의 시트 2, 3 작업
        for i in [1, 2]:
            sheet = source_wb.worksheets[i]
            headers = [cell.value for cell in sheet[1]]
            
            # 필요한 열 위치 찾기
            d_idx = headers.index('DOMAIN')
            v_id_idx = headers.index('VARIABLE_ID')
            val_idx = headers.index('VAL')

            for row in sheet.iter_rows(min_row=2):
                m_key = f"{row[d_idx].value}/{row[v_id_idx].value}"
                if m_key in ref_dict:
                    cell = row[val_idx]
                    cell.value = ref_dict[m_key]
                    cell.fill = yellow_fill
                    cell.font = red_bold_font

        base_name = os.path.splitext(source_path)[0]
        output_name = f"{base_name}_updated.xlsx"
        source_wb.save(output_name)
        return output_name
    except Exception as e:
        return str(e)

# --- GUI 레이아웃 ---
root = TkinterDnD.Tk()
root.title("DB Specification Update Tool")
root.geometry("600x450")

# ① 기능 선택 (단일 기능으로 수정)
tk.Label(root, text="① 기능 선택", font=('Arial', 10, 'bold')).pack(pady=10)
mode_combo = ttk.Combobox(root, values=["DB specification_mandatory update"], state="readonly", width=50)
mode_combo.pack()
mode_combo.current(0)

file_frame = tk.Frame(root)
file_frame.pack(pady=20)

# 공통 스타일 설정 (원 안의 + 기호)
plus_style = {"font": ('Arial', 40), "fg": "#bdc3c7"}

# ② 원본파일 드래그 영역
frame2 = tk.LabelFrame(file_frame, text="② 원본파일", width=250, height=180)
frame2.pack_propagate(False)
frame2.pack(side="left", padx=15)
label2 = tk.Label(frame2, text="⊕", **plus_style)
label2.pack(expand=True)
label2_name = tk.Label(frame2, text="", font=('Arial', 8), fg="blue")
label2_name.pack(side="bottom", pady=5)

# ③ 참고파일 드래그 영역
frame3 = tk.LabelFrame(file_frame, text="③ 참고파일", width=250, height=180)
frame3.pack_propagate(False)
frame3.pack(side="left", padx=15)
label3 = tk.Label(frame3, text="⊕", **plus_style)
label3.pack(expand=True)
label3_name = tk.Label(frame3, text="", font=('Arial', 8), fg="blue")
label3_name.pack(side="bottom", pady=5)

path_source = ""
path_ref = ""

def handle_drop_source(event):
    global path_source
    path_source = event.data.strip('{?}')
    label2.config(text="✔", fg="#2ecc71") # 완료 시 체크 표시로 변경
    label2_name.config(text=os.path.basename(path_source))

def handle_drop_ref(event):
    global path_ref
    path_ref = event.data.strip('{?}')
    label3.config(text="✔", fg="#2ecc71")
    label3_name.config(text=os.path.basename(path_ref))

frame2.drop_target_register(DND_FILES)
frame2.dnd_bind('<<Drop>>', handle_drop_source)
frame3.drop_target_register(DND_FILES)
frame3.dnd_bind('<<Drop>>', handle_drop_ref)

# ④ 실행 버튼
def start_action():
    if not path_source or not path_ref:
        messagebox.showwarning("알림", "파일을 모두 드래그하여 입력해주세요.")
        return
    res = process_logic(path_source, path_ref, mode_combo.get())
    if res.endswith(".xlsx"):
        messagebox.showinfo("완료", f"업데이트 성공!\n파일: {os.path.basename(res)}")
    else:
        messagebox.showerror("오류", f"에러 내용: {res}")

btn_run = tk.Button(root, text="④ 실행하기", command=start_action, bg="#2ecc71", fg="white", font=('Arial', 12, 'bold'), width=15)
btn_run.pack(pady=30)

root.mainloop()
