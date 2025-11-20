import pandas as pd
import numpy as np
from scipy import stats
from statsmodels.stats.power import TTestIndPower
import os, math

def independent_ttest_analysis(file_paths, group_col, value_cols):
    """
    獨立樣本 t 檢定（Independent T-Test）
    file_paths: 單檔或多檔路徑（可為 str 或 list）
    group_col: 分組欄位名稱（0,1）
    value_cols: 欄位名稱 list
    """
    # 統一轉成 list
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    for file_path in file_paths:
        df = pd.read_excel(file_path)

        # 使用檔案所在資料夾作為基底
        folder_path = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = os.path.join(folder_path, f"{base_name}_independent")
        os.makedirs(output_dir, exist_ok=True)

        results = []
        power_calc = TTestIndPower()

        # 確保組別按數字小到大
        groups = sorted(df[group_col].dropna().unique())
        if len(groups) != 2:
            raise ValueError("分組欄位必須恰好兩組")

        group1, group2 = groups
        g1_data = df[df[group_col] == group1]
        g2_data = df[df[group_col] == group2]

        for col in value_cols:
            x1 = g1_data[col].dropna()
            x2 = g2_data[col].dropna()

            t_stat, p_val = stats.ttest_ind(x1, x2, equal_var=False)
            mean1, mean2 = np.mean(x1), np.mean(x2)
            sd1, sd2 = np.std(x1, ddof=1), np.std(x2, ddof=1)
            diff = mean1 - mean2
            se = math.sqrt(sd1**2/len(x1) + sd2**2/len(x2))
            df_t = len(x1) + len(x2) - 2
            ci = stats.t.interval(0.95, df_t, loc=diff, scale=se)
            ll, ul = ci

            # Cohen's d（使用合併標準差）
            pooled_sd = math.sqrt(((len(x1)-1)*sd1**2 + (len(x2)-1)*sd2**2) / df_t)
            d = diff / pooled_sd
            # power
            power = power_calc.solve_power(effect_size=abs(d), nobs1=len(x1), alpha=0.05)

            results.append({
                "Variable": col,
                "M1": mean1,
                "SD1": sd1,
                "M2": mean2,
                "SD2": sd2,
                "t": t_stat,
                "p": p_val,
                "95%CI_LL": ll,
                "95%CI_UL": ul,
                "Cohen_d": d,
                "1-β": power
            })

        df_result = pd.DataFrame(results)
        excel_path = os.path.join(output_dir, "independent_results.xlsx")
        text_path = os.path.join(output_dir, "summary.txt")

        df_result.to_excel(excel_path, index=False)

        with open(text_path, "w", encoding="utf-8") as f:
            f.write("=== 獨立樣本 t 檢定結果 ===\n")
            f.write(df_result.to_string(index=False))

        print(f"✅ 獨立樣本檢定結果已輸出至：{output_dir}")
