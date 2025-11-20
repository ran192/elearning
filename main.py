import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import utils
import os
import pandas as pd

# ------------------------------
# 多選欄位對話框（支援 Shift/Ctrl 選取）
# ------------------------------
def select_columns_dialog(columns, title="選擇欄位"):
    selected_cols = []
    def on_ok():
        for i in listbox.curselection():
            selected_cols.append(columns[i])
        top.destroy()
    top = tk.Toplevel()
    top.title(title)
    top.geometry("300x300")
    tk.Label(top, text="請選擇欄位:").pack(pady=5)
    listbox = tk.Listbox(top, selectmode=tk.EXTENDED, exportselection=0)
    for col in columns:
        listbox.insert(tk.END, col)
    listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tk.Button(top, text="確定", command=on_ok).pack(pady=5)
    top.grab_set()
    top.wait_window()
    return selected_cols

# ------------------------------
# 原本統計分析功能
# ------------------------------
def run_reliability_analysis():
    file_path = utils.select_file(title="選擇 Excel 檔案")
    if not file_path: return
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in [".xlsx", ".xls"]:
        messagebox.showerror("錯誤", "僅能選擇 Excel 檔案進行分析！")
        return
    try:
        df = pd.read_excel(file_path)
        df_numeric = df.apply(pd.to_numeric, errors='coerce').dropna(how='all')
    except Exception as e:
        messagebox.showerror("讀取錯誤", f"無法讀取檔案：{e}")
        return
    result_df = utils.reliability_analysis(df_numeric)
    if result_df.empty:
        messagebox.showinfo("訊息", "未產生任何結果。")
        return
    utils.save_data(df=result_df, original_file_path=file_path)

def run_generate_reports():
    file_paths = filedialog.askopenfilenames(title="選擇一個或多個 Excel 檔案", filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_paths: return
    utils.generate_reports(file_paths)

def run_one_sample_ttest():
    file_paths = filedialog.askopenfilenames(title="選擇一個或多個 Excel 檔案 (單一樣本T檢定)", filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_paths: return
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        value_cols = select_columns_dialog(df.columns.tolist(), title="選擇要分析的欄位")
        if not value_cols: continue
        utils.one_sample_analysis(file_path, df, value_cols)

def run_independent_ttest():
    file_paths = filedialog.askopenfilenames(title="選擇一個或多個 Excel 檔案 (獨立樣本T檢定)", filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_paths: return
    group_col = simpledialog.askstring("輸入群組欄位", "請輸入分組變項欄位名稱：")
    if not group_col: return
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        value_cols = select_columns_dialog(df.columns.tolist(), title="選擇要分析的欄位")
        if not value_cols: continue
        utils.independent_ttest_analysis(file_path, group_col, value_cols)

def run_paired_ttest():
    file_paths = filedialog.askopenfilenames(title="選擇一個或多個 Excel 檔案 (成對樣本T檢定)", filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_paths: return
    group_col = simpledialog.askstring("輸入分組欄位", "請輸入分組欄位名稱（用於分組資料對照）:")
    if not group_col: return
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        target_cols = select_columns_dialog(df.columns.tolist(), title="選擇要做成對T檢定的欄位（多選）")
        if not target_cols: continue
        utils.paired_ttest_analysis(file_path, target_cols, group_col)
        
def run_data_transformation():
    utils.run_transform_process()
    
def run_merge_process():
    utils.run_merge_process()

# ------------------------------
# GUI 主程式
# ------------------------------
root = tk.Tk()
root.title("統計分析工具")
root.geometry("450x500")

tk.Label(root, text="選擇要執行的功能：", font=("Arial", 12), wraplength=400).pack(pady=10)

# 按鈕與功能對應
functions_dict = {
    "生成圖表與PDF報告": run_generate_reports,
    "信度分析 (Reliability Analysis)": run_reliability_analysis,
    "單一樣本T檢定 (One-Sample T-Test)": run_one_sample_ttest,
    "獨立樣本T檢定 (Independent T-Test)": run_independent_ttest,
    "成對樣本T檢定 (Paired T-Test)": run_paired_ttest,
    "資料轉換 (多欄多數值)": run_data_transformation,
    "欄位合併 (前綴加總)": run_merge_process
}

for name, func in functions_dict.items():
    tk.Button(root, text=name, width=40, command=func).pack(pady=5)

tk.Button(root, text="退出", width=40, command=root.quit).pack(pady=20)

root.mainloop()