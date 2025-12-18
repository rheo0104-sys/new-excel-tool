import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import tkinter as tk
from tkinter import filedialog, messagebox

def process_excel(source_path, ref_path):
    try:
        # 1. 파일 읽기 (원본은 시트2, 시트3 사용 / 참고파일은 시트1 사용)
        source_wb = openpyxl.load_workbook(source_path)
        ref_df = pd.read_excel(ref_path)

        # 2. 비교 기준(match_key) 생성 함수
        def make_key(df, col1, col2):
            return df[col1].astype(str) + "/" + df[col2].astype(str)

        # 참고파일 키 생성 (Entry ID / Sub Item)
        ref_df['match_key'] = make_key(ref_df, 'Entry ID', 'Sub Item')
        ref_dict = dict(zip(ref_df['match_key'], ref_df['Status'])) # Status 값을 가져온다고 가정

        # 3. 원본파일 시트 수정 (Sheet 2, Sheet 3 순회)
        target_sheets = [source_wb.worksheets[1], source_wb.worksheets[2]] # 2번째, 3번째 시트
        
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_bold_font = Font(color="FF0000", bold=True)

        for sheet in target_sheets:
            # 헤더에서 열 위치 찾기
            headers = [cell.value for cell in sheet[1]]
            try:
                domain_idx = headers.index('DOMAIN')
                var_id_idx = headers.index('VARIABLE_ID')
                target_col_idx = headers.index('VAL') # 업데이트할 컬럼 이름이 'VAL'이라 가정
            except ValueError:
                continue

            for row in sheet.iter_rows(min_row=2):
                domain_val = str(row[domain_idx].value)
                var_id_val = str(row[var_id_idx].value)
                match_key = f"{domain_val}/{var_id_val}"

                # 참고파일에 해당 키가 있으면 업데이트
                if match_key in ref_dict:
                    cell = row[target_col_idx]
                    cell.value = ref_dict[match_key] # 값 변경
                    cell.fill = yellow_fill # 노란 배경
                    cell.font = red_bold_font # 빨간 굵은 글씨

        # 4. 저장
        output_name = "Final_Result.xlsx"
        source_wb.save(output_name)
        return output_name

    except Exception as e:
        return str(e)

# --- GUI 부분 ---
def select_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry.delete(0, tk.END)
    entry.insert(0, path)

def run():
    s = entry_source.get()
    r = entry_ref.get()
    if not s or not r:
        messagebox.showwarning("경고", "파일을 둘 다 선택해주세요.")
        return
    
    result = process_excel(s, r)
    if result.endswith(".xlsx"):
        messagebox.showinfo("완료", f"저장 완료: {result}")
    else:
        messagebox.showerror("에러", f"오류 발생: {result}")

root = tk.Tk()
root.title("Excel Updater")
root.geometry("400x200")

tk.Label(root, text="원본 파일 (시트 2,3):").pack()
entry_source = tk.Entry(root, width=50)
entry_source.pack()
tk.Button(root, text="찾기", command=lambda: select_file(entry_source)).pack()

tk.Label(root, text="참고 파일 (Entry ID/Sub Item):").pack()
entry_ref = tk.Entry(root, width=50)
entry_ref.pack()
tk.Button(root, text="찾기", command=lambda: select_file(entry_ref)).pack()

tk.Button(root, text="실행하기", command=run, bg="green", fg="white").pack(pady=10)

root.mainloop()
