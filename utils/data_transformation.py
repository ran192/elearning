import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def transform_columns(df, columns, mapping):
    df_copy = df.copy()
    for col in columns:
        if col in df_copy.columns:
            df_copy[col] = df_copy[col].map(lambda x: mapping.get(x, x))
    return df_copy

def run_transform_process():
    # 選檔案
    file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel Files","*.xlsx *.xls")])
    if not file_path:
        return
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("讀取錯誤", f"無法讀取檔案：{e}")
        return

    # --- 第一步：選擇欄位 ---
    def step1_select_columns():
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showerror("未選欄位", "請先選擇要轉換的欄位")
            return
        selected_columns = [df.columns[i] for i in selected_indices]
        top.destroy()
        step2_rules_input(selected_columns)

    top = tk.Toplevel()
    top.title("步驟1：選擇要轉換的欄位")
    top.geometry("400x300")
    listbox = tk.Listbox(top, selectmode=tk.EXTENDED, exportselection=0)
    for col in df.columns:
        listbox.insert(tk.END, col)
    listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tk.Button(top, text="下一步", width=20, command=step1_select_columns).pack(pady=10)

    # --- 第二步：輸入規則 ---
    def step2_rules_input(selected_columns):
        rules_win = tk.Toplevel()
        rules_win.title("步驟2：輸入轉換規則")
        rules_win.geometry("500x400")

        tk.Label(rules_win, text=f"已選欄位: {', '.join(selected_columns)}").pack(pady=5)

        frame_rules = tk.Frame(rules_win)
        frame_rules.pack(pady=5)

        tk.Label(frame_rules, text="初始值").grid(row=0, column=0, padx=5)
        tk.Label(frame_rules, text="變換值").grid(row=0, column=1, padx=5)

        rules_listbox = tk.Listbox(rules_win, height=10, width=50)
        rules_listbox.pack(pady=5)

        entry_old = tk.Entry(rules_win, width=20)
        entry_old.pack(pady=2)
        entry_old.insert(0, "初始值")
        entry_new = tk.Entry(rules_win, width=20)
        entry_new.pack(pady=2)
        entry_new.insert(0, "變換值")

        def add_rule():
            old_val = entry_old.get().strip()
            new_val = entry_new.get().strip()
            if old_val == "" or new_val == "":
                messagebox.showerror("錯誤", "請輸入初始值與變換值")
                return
            try:
                old_val_f = float(old_val)
                new_val_f = float(new_val)
            except:
                messagebox.showerror("錯誤", "請輸入數字")
                return
            rules_listbox.insert(tk.END, f"{old_val_f} -> {new_val_f}")
            entry_old.delete(0, tk.END)
            entry_new.delete(0, tk.END)

        def execute_transform():
            mapping = {}
            for item in rules_listbox.get(0, tk.END):
                old, new = item.split("->")
                mapping[float(old.strip())] = float(new.strip())
            df_transformed = transform_columns(df, selected_columns, mapping)
            folder = os.path.dirname(file_path)
            base = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(folder, f"{base}_transformed.xlsx")
            df_transformed.to_excel(output_path, index=False)
            messagebox.showinfo("完成", f"轉換完成，檔案已儲存至：{output_path}")
            rules_win.destroy()

        tk.Button(rules_win, text="新增規則", width=20, command=add_rule).pack(pady=5)
        tk.Button(rules_win, text="執行轉換", width=20, command=execute_transform).pack(pady=5)
        tk.Button(rules_win, text="取消", width=20, command=rules_win.destroy).pack(pady=5)
