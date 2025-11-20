import pandas as pd
import numpy as np
from scipy import stats
from statsmodels.stats.power import TTestPower
import os

def one_sample_analysis(file_path, data, value_cols, popmeans=None):
    """
    單一樣本 t 檢定 (支援多欄位)
    Parameters:
        file_path (str): 原始檔案路徑
        data (pd.DataFrame): 待分析資料
        value_cols (list[str]): 欄位名稱
        popmeans (dict, optional): {欄位名稱: 母體平均數}，若 None，則自動使用各欄位平均數
    Returns:
        pd.DataFrame: 分析結果
    """
    # 取得檔案所在資料夾
    base_folder = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # 輸出資料夾放在檔案所在資料夾
    output_dir = os.path.join(base_folder, f"{base_name}_one_sample")
    os.makedirs(output_dir, exist_ok=True)

    if popmeans is None:
        popmeans = {col: data[col].mean() for col in value_cols}

    results = []
    power_calc = TTestPower()

    for col in value_cols:
        x = data[col].dropna()
        popmean = popmeans.get(col, np.mean(x))

        # 計算 t 檢定
        t_stat, p_val = stats.ttest_1samp(x, popmean)
        mean, sd = np.mean(x), np.std(x, ddof=1)
        se = sd / np.sqrt(len(x))
        df = len(x) - 1
        ci = stats.t.interval(0.95, df, loc=mean, scale=se)
        ll, ul = ci
        d = (mean - popmean) / sd if sd != 0 else 0
        power = power_calc.solve_power(effect_size=abs(d), nobs=len(x), alpha=0.05)

        results.append({
            "Variable": col,
            "Mean": mean,
            "SD": sd,
            "t": t_stat,
            "p": p_val,
            "95%CI_LL": ll,
            "95%CI_UL": ul,
            "Cohen_d": d,
            "1-β": power,
        })

    df_result = pd.DataFrame(results)

    # 儲存 Excel
    excel_path = os.path.join(output_dir, "one_sample_results.xlsx")
    df_result.to_excel(excel_path, index=False)

    # 儲存文字檔
    text_path = os.path.join(output_dir, "summary.txt")
    with open(text_path, "w", encoding="utf-8") as f:
        f.write("=== 單一樣本 t 檢定結果 ===\n")
        f.write(df_result.to_string(index=False))

    print(f"✅ 單一樣本檢定結果已輸出至：{output_dir}")
    return df_result
