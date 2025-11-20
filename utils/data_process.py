import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import os
import re

def merge_columns_by_prefix(df, new_col_prefix=None):
    """
    自動抓取所有欄位，依前綴自動分組加總生成新欄位。
    嚴格區分大小寫，欄位名稱格式例: e1, e2, f1, f2
    new_col_prefix: 若 None，使用原前綴；若有輸入，使用該前綴加後綴區分
    """
    pattern = re.compile(r"([a-zA-Z]+)\d+$")  # 前綴為字母，後面跟數字
    matched_cols = [col for col in df.columns if pattern.match(col)]
    
    if not matched_cols:
        messagebox.showwarning("警告", "未找到符合格式的欄位 (例: e1, e2)。")
        return df

    prefixes = {}
    for col in matched_cols:
        prefix = pattern.match(col).group(1)
        prefixes.setdefault(prefix, []).append(col)

    for prefix, cols_to_merge in prefixes.items():
        col_name = new_col_prefix if new_col_prefix else prefix
        if new_col_prefix:
            col_name = f"{new_col_prefix}_{prefix}"
        df[col_name] = df[cols_to_merge].sum(axis=1)

    return df

def run_merge_process():
    """
    GUI：選擇檔案並自動合併同前綴欄位
    """
    file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("讀取錯誤", f"無法讀取檔案：{e}")
        return

    # 詢問使用者是否要自訂前綴
    use_custom_prefix = messagebox.askyesno("自訂前綴", "是否要自訂合併後的新欄位前綴？\n選「否」將直接使用原前綴名稱。")
    new_col_prefix = None
    if use_custom_prefix:
        new_col_prefix = simpledialog.askstring("新欄位名稱前綴", "請輸入合併後新欄位名稱前綴:")

    df = merge_columns_by_prefix(df, new_col_prefix)

    folder = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_dir = os.path.join(folder, f"{base_name}_merged")
    os.makedirs(output_dir, exist_ok=True)
    save_path = os.path.join(output_dir, f"{base_name}_merged.xlsx")
    df.to_excel(save_path, index=False)

    messagebox.showinfo("完成", f"欄位自動合併完成！結果儲存至：\n{save_path}")
