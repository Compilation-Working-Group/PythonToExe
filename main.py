import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import sys

# --------------------------
# 核心算法逻辑
# --------------------------

def get_grade_config():
    """定义赋分等级和区间标准 (甘肃/通用 3+1+2)"""
    return [
        {'grade': 'A', 'percent': 0.15, 't_max': 100, 't_min': 86},
        {'grade': 'B', 'percent': 0.35, 't_max': 85,  't_min': 71},
        {'grade': 'C', 'percent': 0.35, 't_max': 70,  't_min': 56},
        {'grade': 'D', 'percent': 0.13, 't_max': 55,  't_min': 41},
        {'grade': 'E', 'percent': 0.02, 't_max': 40,  't_min': 30},
    ]

def calculate_assigned_score(series):
    """
    对单科原始成绩进行赋分计算
    :param series: Pandas Series (原始分数)
    :return: Pandas Series (赋分后分数)
    """
    # --- 修复核心开始 ---
    # 1. 强制转换为数字类型
    # errors='coerce' 会将无法转换的内容（如"缺考", " ", 文本）直接变成 NaN (空值)
    series_numeric = pd.to_numeric(series, errors='coerce')
    
    # 2. 过滤掉 NaN 空值
    valid_scores = series_numeric.dropna()
    # --- 修复核心结束 ---

    total_count = len(valid_scores)
    
    if total_count == 0:
        # 如果整列都没有有效数字，返回全空的Series
        return pd.Series(index=series.index, dtype=float)

    # 3. 排序：降序
    sorted_scores = valid_scores.sort_values(ascending=False)
    
    # 4. 划分等级并计算
    assigned_result = pd.Series(index=valid_scores.index, dtype=float)
    
    current_idx = 0
    configs = get_grade_config()
    
    for cfg in configs:
        # 计算该等级的人数
        count = int(np.round(total_count * cfg['percent']))
        
        # 修正E等级人数，确保覆盖剩余所有人
        if cfg['grade'] == 'E':
            count = total_count - current_idx
        
        if count <= 0:
            continue

        # 获取该等级内的所有学生索引
        # 注意处理索引越界，虽然理论上不会，但加个切片保护更安全
        end_idx = min(current_idx + count, total_count)
        if current_idx >= end_idx:
            break

        grade_indices = sorted_scores.iloc[current_idx : end_idx].index
        grade_raw_scores = sorted_scores.iloc[current_idx : end_idx]
        
        # 获取该等级原始分的 Max (Y2) 和 Min (Y1)
        Y2 = grade_raw_scores.max()
        Y1 = grade_raw_scores.min()
        T2 = cfg['t_max']
        T1 = cfg['t_min']
        
        # 赋分公式计算
        def calculate_single(Y):
            if Y2 == Y1: 
                return (T2 + T1) / 2
            else:
                return T1 + ((Y - Y1) * (T2 - T1)) / (Y2 - Y1)

        assigned_vals = grade_raw_scores.apply(calculate_single)
        assigned_result.loc[grade_indices] = assigned_vals
        
        current_idx = end_idx

    # 四舍五入保留整数，并填回原长度的Series中（未参与计算的保持NaN）
    return assigned_result.round()

# --------------------------
# GUI 与 业务流程
# --------------------------

def run_app():
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口

    try:
        messagebox.showinfo("甘肃新高考赋分工具", "欢迎使用！\n请准备好Excel文件，确保成绩列尽量为纯数字。\n（系统会自动跳过“缺考”或文字标注的学生）")

        # 1. 选择文件
        file_path = filedialog.askopenfilename(
            title="选择学生成绩表 (Excel)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return

        try:
            # 读取所有列为对象，防止pandas自作聪明把考号前面的0去掉
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {e}")
            return

        # 2. 识别科目列
        all_columns = df.columns.tolist()
        msg = f"检测到以下列名:\n{all_columns}\n\n请输入需要【赋分】的科目名称（用中文逗号或英文逗号分隔）:\n例如: 化学,生物"
        
        assigned_subs_str = simpledialog.askstring("选择赋分科目", msg)
        if not assigned_subs_str:
            return
            
        assigned_subs = [s.strip() for s in assigned_subs_str.replace("，", ",").split(",") if s.strip()]
        
        # 3. 处理数据
        output_df = df.copy()
        
        # 验证列是否存在
        for sub in assigned_subs:
            if sub not in df.columns:
                messagebox.showerror("错误", f"找不到列名: {sub}")
                return

        # 3.1 计算赋分
        for sub in assigned_subs:
            new_col = f"{sub}_赋分"
            try:
                # 调用修复后的函数
                result_series = calculate_assigned_score(df[sub])
                output_df[new_col] = result_series
            except Exception as e:
                import traceback
                err_msg = traceback.format_exc()
                messagebox.showerror("计算错误", f"计算科目【{sub}】时出错:\n{str(e)}\n\n详情:\n{err_msg}")
                return

        # 3.2 计算总分
        raw_subs_str = simpledialog.askstring("选择原始计入科目", 
            f"请输入【直接计入总分】的原始科目 (语数外+首选科目):\n例如: 语文,数学,英语,物理")
        
        if raw_subs_str:
            raw_subs = [s.strip() for s in raw_subs_str.replace("，", ",").split(",") if s.strip()]
            
            # 这里的计算也要小心，先转数字
            temp_sum_df = pd.DataFrame()
            for col in raw_subs:
                if col in output_df.columns:
                    # 同样强制转数字，防止原始分里也有空格
                    temp_sum_df[col] = pd.to_numeric(output_df[col], errors='coerce').fillna(0)
            
            # 计算原始分总和
            total_score_col = temp_sum_df.sum(axis=1)
            
            # 加上赋分科目的分
            for sub in assigned_subs:
                # fillna(0) 把没赋分（因为缺考等）的当0分处理
                score_to_add = output_df[f"{sub}_赋分"].fillna(0)
                total_score_col += score_to_add
                
            output_df["总分"] = total_score_col
            
            # 3.3 排名
            output_df["排名"] = output_df["总分"].rank(ascending=False, method='min')
            output_df = output_df.sort_values("排名")

        # 4. 导出
        save_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="赋分结果.xlsx"
        )
        
        if save_path:
            output_df.to_excel(save_path, index=False)
            messagebox.showinfo("成功", f"处理完成！\n文件已保存至: {save_path}")
            try:
                os.startfile(os.path.dirname(save_path))
            except:
                pass
                
    except Exception as e:
        messagebox.showerror("未知错误", f"程序发生意外错误:\n{e}")

if __name__ == "__main__":
    run_app()
