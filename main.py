import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def process_logic(source_path, ref_path):
    try:
        source_wb = openpyxl.load_workbook(source_path)
        ref_df = pd.read_excel(ref_path)

        # 서식 설정
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_font = Font(color="FF0000", bold=True)

        # 참고파일 데이터 캐싱 (속도 향상 및 비교 용이)
        ref_data_map = {}
        for _, r_row in ref_df.iterrows():
            eid = str(r_row['Entry ID']).strip()
            sub = str(r_row['Sub Item']).strip()
            dmq = str(r_row['Default Missing Query']).strip().upper()
            
            # 모듈 1용 (ID 기준), 모듈 2용 (ID/SubItem 기준) 데이터 저장
            ref_data_map[eid] = {"Sub Item": sub, "DMQ": dmq}
            ref_data_map[f"{eid}/{sub}"] = {"DMQ": dmq}

        # --- [모듈 1: DB Specification] ---
        if "DB Specification" in source_wb.sheetnames:
            ws1 = source_wb["DB Specification"]
            h1 = [cell.value for cell in ws1[1]]
            v_idx, l_idx, m_idx = h1.index('VARIABLE_ID'), h1.index('LAYOUT'), h1.index('Mandatory')

            for row in ws1.iter_rows(min_row=2):
                vid = str(row[v_idx].value).strip()
                layout = str(row[l_idx].value).strip()
                target = row[m_idx]
                old_val = str(target.value).strip() if target.value else ""
                new_val = old_val

                if layout == "SYSDEFINED":
                    new_val = "N"
                elif vid in ref_data_map:
                    ref = ref_data_map[vid]
                    if "." in ref["Sub Item"]: new_val = "Y/N"
                    else: new_val = "Y" if ref["DMQ"] == "YES" else "N" if ref["DMQ"] == "NO" else old_val

                if old_val != new_val:
                    target.value = new_val
                    target.fill, target.font = yellow_fill, red_font
        else:
            return "오류: 필수 시트 'DB Specification'이 없습니다."

        # --- [모듈 2: DB Specification_CAT] ---
        if "DB Specification_CAT" in source_wb.sheetnames:
            ws2 = source_wb["DB Specification_CAT"]
            h2 = [cell.value for cell in ws2[1]]
            v_idx, s_idx, m_idx = h2.index('VARIABLE_ID'), h2.index('SUB_ITEM'), h2.index('Mandatory')

            for row in ws2.iter_rows(min_row=2):
                vid_sub = f"{str(row[v_idx].value).strip()}/{str(row[s_idx].value).strip()}"
                target = row[m_idx]
                old_val = str(target.value).strip() if target.value else ""
                new_val = old_val

                if vid_sub in ref_data_map:
                    ref = ref_data_map[vid_sub]
                    new_val = "Y" if ref["DMQ"] == "YES" else "N" if ref["DMQ"] == "NO" else old_val

                if old_val != new_val:
                    target.value = new_val
                    target.fill, target.font = yellow_fill, red_font

        # 파일 저장
        output_name = f"{os.path.splitext(source_path)[0]}_updated.xlsx"
        source_wb.save(output_name)
        return output_name

    except Exception as e:
        return f"실행 중 오류: {str(e)}"

# --- [GUI 구성] ---
root = TkinterDnD.Tk()
root.title("DB Specification Update Tool")
root.geometry("600x450")

tk.Label(root, text="① 기능 선택", font=('Arial', 10, 'bold')).pack(pady=10)
mode_combo = ttk.Combobox(root, values=["DB specification_mandatory update"], state="readonly", width=50)
mode_combo.pack(); mode_combo.current(0)

file_frame = tk.Frame(root); file_frame.pack(pady=20)
path_source, path_ref = "", ""

def create_drop_box(parent, title, drop_func):
    frame = tk.LabelFrame(parent, text=title, width=250, height=180)
    frame.pack_propagate(False); frame.pack(side="left", padx=15)
    label = tk.Label(frame, text="⊕", font=('Arial', 40), fg="#bdc3c7")
    label.pack(expand=True)
    name_label = tk.Label(frame, text="", font=('Arial', 8), fg="blue")
    name_label.pack(side="bottom", pady=5)
    frame.drop_target_register(DND_FILES)
    frame.dnd_bind('<<Drop>>', lambda e: drop_func(e, label, name_label))
    return frame

def drop_src(e, l, n): global path_source; path_source = e.data.strip('{?}'); l.config(text="✔", fg="#2ecc71"); n.config(text=os.path.basename(path_source))
def drop_ref(e, l, n): global path_ref; path_ref = e.data.strip('{?}'); l.config(text="✔", fg="#2ecc71"); n.config(text=os.path.basename(path_ref))

create_drop_box(file_frame, "② 원본파일", drop_src)
create_drop_box(file_frame, "③ 참고파일", drop_ref)

def run():
    if not path_source or not path_ref: messagebox.showwarning("알림", "파일을 모두 넣어주세요."); return
    res = process_logic(path_source, path_ref)
    if res.endswith(".xlsx"): messagebox.showinfo("완료", f"저장 성공!\n{os.path.basename(res)}")
    else: messagebox.showerror("오류", res)

tk.Button(root, text="④ 실행하기", command=run, bg="#2ecc71", fg="white", font=('Arial', 12, 'bold'), width=15).pack(pady=30)
root.mainloop()
