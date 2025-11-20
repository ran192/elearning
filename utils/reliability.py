import pandas as pd
import numpy as np
import re
from tkinter import messagebox

def cronbach_alpha(df_subset):
    """Calculate Cronbach's Alpha; return 0 if cannot calculate"""
    try:
        df_subset = df_subset.dropna()
        k = len(df_subset.columns)
        if k < 2:
            return 0
        item_var = df_subset.var(axis=0, ddof=1)
        total_scores = df_subset.sum(axis=1)
        total_var = total_scores.var(ddof=1)
        if total_var == 0:
            return 0
        alpha = (k / (k - 1)) * (1 - item_var.sum() / total_var)
        return alpha
    except:
        return 0

def filter_columns_by_prefix(df, target_prefixes):
    """Return dict {prefix: [columns]} for matching prefixes"""
    groups = {}
    for col in df.columns:
        clean_col = str(col).replace(" ", "").replace("\u3000", "")
        m = re.match(r"([A-Za-z]+)", clean_col)
        if m:
            prefix = m.group(1).upper()
            if prefix in target_prefixes:
                if prefix not in groups:
                    groups[prefix] = []
                groups[prefix].append(col)
    return groups

def reliability_table(df_subset, scale_name):
    """Generate reliability table for a single scale"""
    alpha_total = cronbach_alpha(df_subset)
    results = []
    first = True
    for col in df_subset.columns:
        try:
            if len(df_subset.columns) > 1:
                total_minus_item = df_subset.drop(columns=[col]).sum(axis=1)
                corr = df_subset[col].corr(total_minus_item)
                alpha_if_deleted = cronbach_alpha(df_subset.drop(columns=[col]))
            else:
                corr = 0
                alpha_if_deleted = 0
        except:
            corr = 0
            alpha_if_deleted = 0
        results.append({
            "Scale": scale_name if first else "",
            "Item": col,
            "Corrected Item-Total Correlation": round(corr,4) if not pd.isna(corr) else 0,
            "Cronbach's Alpha if Item Deleted": round(alpha_if_deleted,4) if not pd.isna(alpha_if_deleted) else 0,
            "Overall Cronbach's Alpha": round(alpha_total,4) if first else ""
        })
        first = False
    return pd.DataFrame(results)

def reliability_analysis(df, target_prefixes=None):
    """Analyze multiple scales and return combined table"""
    if target_prefixes is None:
        target_prefixes = {"E", "CE", "PP", "FWS", "R"}
    groups = filter_columns_by_prefix(df, target_prefixes)
    if not groups:
        messagebox.showerror("Error", "No columns matched the target prefixes!")
        return pd.DataFrame()
    
    all_tables = []
    for name, cols in groups.items():
        sub_df = df[cols]
        table = reliability_table(sub_df, name)
        all_tables.append(table)
    
    combined = pd.concat(all_tables, ignore_index=True)
    return combined
