import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
import tkinter as tk
from tkinter import filedialog, messagebox

# ====== 字型設定 ======
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False


def generate_report(data: pd.DataFrame, name: str, base_dir: str, is_global=False):
    """根據表格欄位動態生成圖表與PDF報告"""
    prefix = "World" if is_global else name
    charts_dir = os.path.join(base_dir, "charts")
    reports_dir = os.path.join(base_dir, "reports")
    os.makedirs(charts_dir, exist_ok=True)
    os.makedirs(reports_dir, exist_ok=True)

    # ====== 自動偵測數值欄位 ======
    numeric_cols = data.select_dtypes(include='number').columns.tolist()
    if '年份' in numeric_cols:
        numeric_cols.remove('年份')

    image_paths = []

    # ====== 為每個欄位繪圖 ======
    for col in numeric_cols:
        plt.figure(figsize=(8, 5))
        try:
            # 折線圖（年份為橫軸）
            sns.lineplot(data=data, x='年份', y=col, marker="o")
            plt.title(f"{prefix}：{col} 變化趨勢")
            plt.tight_layout()
            img_path = os.path.join(charts_dir, f"{prefix}_{col}.png")
            plt.savefig(img_path)
            plt.close()
            image_paths.append(img_path)
        except Exception as e:
            print(f"⚠️ 無法繪製欄位 {col}: {e}")
            plt.close()

    # ====== 若至少有兩個數值欄位，嘗試繪製散點關係圖 ======
    if len(numeric_cols) >= 2:
        x_col = numeric_cols[0]
        for y_col in numeric_cols[1:]:
            plt.figure(figsize=(6, 5))
            sns.scatterplot(data=data, x=x_col, y=y_col)
            plt.title(f"{prefix}：{x_col} 與 {y_col} 關係圖")
            plt.tight_layout()
            img_path = os.path.join(charts_dir, f"{prefix}_{x_col}_vs_{y_col}.png")
            plt.savefig(img_path)
            plt.close()
            image_paths.append(img_path)

    # ====== PDF 報告 ======
    pdf_name = "全球總覽報告.pdf" if is_global else f"{name}_報告.pdf"
    pdf_path = os.path.join(reports_dir, pdf_name)
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    title_text = f"<b>{'🌍 全球總覽報告' if is_global else name + ' 經濟與人口發展報告'}</b>"
    story.append(Paragraph(title_text, styles["Title"]))
    story.append(Spacer(1, 20))
    story.append(Paragraph(f"資料筆數：{len(data)}", styles["Normal"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("以下圖表依據資料欄位自動生成：", styles["Normal"]))
    story.append(Spacer(1, 20))

    for img_path in image_paths:
        story.append(Image(img_path, width=400, height=300))
        story.append(Spacer(1, 20))

    if not image_paths:
        story.append(Paragraph("⚠️ 未找到可視覺化的數值欄位。", styles["Normal"]))

    story.append(Paragraph("報告生成完畢。", styles["Italic"]))
    doc.build(story)

    print(f"✅ 已生成報告 → {pdf_path}")


def generate_reports(file_paths):
    """主流程：支援多檔案分析"""
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    root = tk.Tk()
    root.withdraw()

    for file_path in file_paths:
        if not os.path.isfile(file_path):
            messagebox.showerror("錯誤", f"找不到檔案：{file_path}")
            continue

        base_dir = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)

        if not file_path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror("格式錯誤", f"{file_name} 不是 Excel 檔案！")
            continue

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("讀取錯誤", f"無法讀取檔案：{file_path}\n{e}")
            continue

        if '國家' not in df.columns:
            messagebox.showerror("格式錯誤", f"檔案 {file_name} 缺少『國家』欄位！")
            continue

        countries = df['國家'].dropna().unique()
        for country in countries:
            generate_report(df[df['國家'] == country], name=country, base_dir=base_dir)

        # === 全球平均 ===
        world_data = df.groupby('年份').mean(numeric_only=True).reset_index()
        generate_report(world_data, name="World", base_dir=base_dir, is_global=True)

    messagebox.showinfo("完成", "🎉 所有報告已生成，可於各檔案所在資料夾內查看。")


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="請選擇要分析的 Excel 檔案",
        filetypes=[("Excel 檔案", "*.xlsx *.xls")]
    )
    if file_paths:
        generate_reports(file_paths)
    else:
        messagebox.showinfo("取消", "未選取任何檔案。")
