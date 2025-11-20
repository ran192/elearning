import pandas as pd
import numpy as np
from scipy import stats
from statsmodels.stats.power import TTestPower
import os

def paired_ttest_analysis(file_paths, target_cols, group_col=None):
    """
    成對樣本 T 檢定（依分組欄位）
    file_paths: 單檔或多檔
    target_cols: 欄位列表，對每個欄位做成對 T 檢定
    group_col: 若指定，依這個欄位分組
    """
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    all_results = []

    for file_path in file_paths:
        df = pd.read_excel(file_path)

        folder_path = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_dir = os.path.join(folder_path, f"{base_name}_paired")
        os.makedirs(output_dir, exist_ok=True)

        results = []
        power_calc = TTestPower()

        if group_col and group_col in df.columns:
            groups = sorted(df[group_col].dropna().unique())
            if len(groups) != 2:
                raise ValueError("分組欄位必須恰好兩組")
            g1, g2 = groups
            df1 = df[df[group_col] == g1]
            df2 = df[df[group_col] == g2]
        else:
            # 若沒有指定分組欄位，直接對整個欄位做成對
            df1 = df2 = df

        for col in target_cols:
            if col not in df.columns:
                continue
            x1 = df1[col].fillna(0).to_numpy()
            x2 = df2[col].fillna(0).to_numpy()

            # 對齊長度，不足補 0
            max_len = max(len(x1), len(x2))
            if len(x1) < max_len:
                x1 = np.append(x1, [0]*(max_len - len(x1)))
            if len(x2) < max_len:
                x2 = np.append(x2, [0]*(max_len - len(x2)))

            t_stat, p_val = stats.ttest_rel(x1, x2)
            m1, m2 = np.mean(x1), np.mean(x2)
            sd1, sd2 = np.std(x1, ddof=1), np.std(x2, ddof=1)
            diff = m1 - m2
            se = np.sqrt(sd1**2/len(x1) + sd2**2/len(x2))
            dfree = len(x1) - 1
            ci = stats.t.interval(0.95, dfree, loc=diff, scale=se)
            pooled_sd = np.sqrt((sd1**2 + sd2**2)/2) if (sd1>0 or sd2>0) else 0
            d = diff / pooled_sd if pooled_sd != 0 else 0
            power = power_calc.solve_power(effect_size=abs(d), nobs=len(x1), alpha=0.05)

            results.append({
                "Variable": col,
                "M1": m1,
                "SD1": sd1,
                "M2": m2,
                "SD2": sd2,
                "t": t_stat,
                "p": p_val,
                "95%CI_LL": ci[0],
                "95%CI_UL": ci[1],
                "Cohen_d": d,
                "1-β": power
            })

        df_result = pd.DataFrame(results)
        excel_path = os.path.join(output_dir, "paired_results.xlsx")
        text_path = os.path.join(output_dir, "summary.txt")
        df_result.to_excel(excel_path, index=False)
        with open(text_path, "w", encoding="utf-8") as f:
            f.write("=== 成對樣本 t 檢定結果 ===\n")
            f.write(df_result.to_string(index=False))

        print(f"✅ 成對樣本檢定結果已輸出至：{output_dir}")
        all_results.append(df_result)

    return all_results
