import customtkinter as ctk
import pandas as pd
import numpy as np
import threading
import os
import sys
from tkinter import filedialog, messagebox

# --- 全局外观设置 ---
ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")  

class GaokaoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 1. 窗口基础设置
        self.title("甘肃新高考赋分系统 Pro Max | 俞晋全制作")
        self.geometry("1100x850")
        self.minsize(900, 700)
        
        # 数据变量
        self.file_path = None
        self.df_raw = None
        self.sheet_names = []
        
        # 布局配置
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧边栏 ===
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(9, weight=1) 

        # Logo
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="高考赋分工具", font=ctk.CTkFont(size=22, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 20))

        # 1. 导入
        self.btn_load = ctk.CTkButton(self.sidebar_frame, text="1. 导入Excel成绩表", height=40, command=self.load_file_action)
        self.btn_load.grid(row=1, column=0, padx=20, pady=10)

        # 2. Sheet选择
        self.lbl_sheet = ctk.CTkLabel(self.sidebar_frame, text="选择工作表 (Sheet):", anchor="w")
        self.lbl_sheet.grid(row=2, column=0, padx=20, pady=(15, 0), sticky="w")
        self.sheet_dropdown = ctk.CTkOptionMenu(self.sidebar_frame, values=[], command=self.change_sheet_event)
        self.sheet_dropdown.grid(row=3, column=0, padx=20, pady=(5, 10))
        self.sheet_dropdown.set("等待导入...")
        self.sheet_dropdown.configure(state="disabled")

        # 3. 班级列
        self.lbl_class = ctk.CTkLabel(self.sidebar_frame, text="指定班级列 (计算班排):", anchor="w")
        self.lbl_class.grid(row=4, column=0, padx=20, pady=(15, 0), sticky="w")
        self.class_col_dropdown = ctk.CTkOptionMenu(self.sidebar_frame, values=[])
        self.class_col_dropdown.grid(row=5, column=0, padx=20, pady=(5, 10))
        self.class_col_dropdown.set("等待加载...")

        # 底部按钮区
        self.btn_calc = ctk.CTkButton(self.sidebar_frame, text="开始赋分计算", height=50, fg_color="green", font=ctk.CTkFont(size=16, weight="bold"), command=self.start_calculation)
        self.btn_calc.grid(row=10, column=0, padx=20, pady=15)
        self.btn_calc.configure(state="disabled")

        self.btn_export = ctk.CTkButton(self.sidebar_frame, text="导出结果 Excel", height=40, command=self.export_file)
        self.btn_export.grid(row=11, column=0, padx=20, pady=(0, 30))
        self.btn_export.configure(state="disabled")

        # === 右侧内容区 ===
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        # 状态栏
        self.status_label = ctk.CTkLabel(self.main_frame, text="欢迎使用！请点击左侧按钮导入数据。", anchor="w", font=("Microsoft YaHei UI", 16))
        self.status_label.pack(fill="x", pady=(0, 15))

        # 滚动设置区
        self.scroll_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="科目勾选设置")
        self.scroll_frame.pack(fill="both", expand=True)

        # 原始计入科目区
        self.lbl_raw = ctk.CTkLabel(self.scroll_frame, text="【直接计入总分】 (语数外 + 物理/历史):", anchor="w", font=("Microsoft YaHei UI", 13, "bold"), text_color=("gray30", "gray80"))
        self.lbl_raw.pack(fill="x", pady=(10, 5), padx=10)
        self.raw_checkboxes_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.raw_checkboxes_frame.pack(fill="x", pady=5, padx=10)
        self.raw_checkboxes = []

        # 赋分科目区
        self.lbl_assign = ctk.CTkLabel(self.scroll_frame, text="【等级赋分科目】 (化生政地):", anchor="w", font=("Microsoft YaHei UI", 13, "bold"), text_color=("gray30", "gray80"))
        self.lbl_assign.pack(fill="x", pady=(25, 5), padx=10)
        self.assign_checkboxes_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.assign_checkboxes_frame.pack(fill="x", pady=5, padx=10)
        self.assign_checkboxes = []

        # 进度条
        self.progressbar = ctk.CTkProgressBar(self.main_frame, height=15)
        self.progressbar.pack(fill="x", pady=(20, 0))
        self.progressbar.set(0)

    # --- 文件与界面逻辑 ---

    def load_file_action(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path: return
        
        self.file_path = file_path
        self.status_label.configure(text=f"正在分析文件: {os.path.basename(file_path)}...")
        self.progressbar.start()
        threading.Thread(target=self.read_excel_sheets).start()

    def read_excel_sheets(self):
        try:
            excel_file = pd.ExcelFile(self.file_path)
            self.sheet_names = excel_file.sheet_names
            self.after(0, self.update_sheet_ui)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("错误", f"读取失败: {e}"))
            self.after(0, self.progressbar.stop)

    def update_sheet_ui(self):
        self.progressbar.stop()
        self.progressbar.set(1)
        self.status_label.configure(text=f"已就绪: {os.path.basename(self.file_path)}")
        self.sheet_dropdown.configure(values=self.sheet_names, state="normal")
        self.sheet_dropdown.set(self.sheet_names[0])
        self.change_sheet_event(self.sheet_names[0])

    def change_sheet_event(self, sheet_name):
        try:
            self.df_raw = pd.read_excel(self.file_path, sheet_name=sheet_name)
            columns = self.df_raw.columns.tolist()
            
            self.class_col_dropdown.configure(values=columns)
            default_class = next((c for c in columns if "班" in str(c)), columns[0] if columns else "")
            self.class_col_dropdown.set(default_class)

            self.create_subject_checkboxes(columns)
            
            self.btn_calc.configure(state="normal")
            self.status_label.configure(text=f"当前工作表: {sheet_name} | 请勾选对应科目")
        except Exception as e:
            messagebox.showerror("错误", f"加载工作表失败: {e}")

    def create_subject_checkboxes(self, columns):
        for cb in self.raw_checkboxes + self.assign_checkboxes: cb.destroy()
        self.raw_checkboxes.clear()
        self.assign_checkboxes.clear()
        
        common_raw = ["语文", "数学", "英语", "物理", "历史", "外语"]
        common_assign = ["化学", "生物", "地理", "政治", "思想政治"]

        def add_cb(parent, text, storage, keywords):
            cb = ctk.CTkCheckBox(parent, text=text, font=("Microsoft YaHei UI", 12))
            cb.grid(row=len(storage)//5, column=len(storage)%5, sticky="w", padx=10, pady=8)
            if any(k in str(text) for k in keywords): cb.select()
            storage.append(cb)

        for col in columns:
            add_cb(self.raw_checkboxes_frame, col, self.raw_checkboxes, common_raw)
        for col in columns:
            add_cb(self.assign_checkboxes_frame, col, self.assign_checkboxes, common_assign)

    # --- 核心计算逻辑 (新增原始总分计算) ---
    
    def start_calculation(self):
        self.selected_raw = [cb.cget("text") for cb in self.raw_checkboxes if cb.get() == 1]
        self.selected_assign = [cb.cget("text") for cb in self.assign_checkboxes if cb.get() == 1]
        self.selected_class_col = self.class_col_dropdown.get()

        if not self.selected_raw and not self.selected_assign:
            messagebox.showwarning("提示", "请至少勾选一个科目！")
            return

        self.btn_calc.configure(state="disabled")
        self.status_label.configure(text="正在全速计算中...")
        self.progressbar.configure(mode="indeterminate")
        self.progressbar.start()
        
        threading.Thread(target=self.run_math_logic).start()

    def run_math_logic(self):
        try:
            df = self.df_raw.copy()
            
            # 定义赋分标准
            grade_configs = [
                {'grade': 'A', 'percent': 0.15, 't_max': 100, 't_min': 86},
                {'grade': 'B', 'percent': 0.35, 't_max': 85,  't_min': 71},
                {'grade': 'C', 'percent': 0.35, 't_max': 70,  't_min': 56},
                {'grade': 'D', 'percent': 0.13, 't_max': 55,  't_min': 41},
                {'grade': 'E', 'percent': 0.02, 't_max': 40,  't_min': 30},
            ]

            def calculate_assigned_score(series):
                series_num = pd.to_numeric(series, errors='coerce')
                valid = series_num.dropna()
                if len(valid) == 0: return pd.Series(index=series.index, dtype=float)
                
                sorted_scores = valid.sort_values(ascending=False)
                result = pd.Series(index=valid.index, dtype=float)
                curr = 0
                for cfg in grade_configs:
                    cnt = int(np.round(len(valid) * cfg['percent']))
                    if cfg['grade'] == 'E': cnt = len(valid) - curr
                    if cnt <= 0: continue
                    end = min(curr + cnt, len(valid))
                    if curr >= end: break
                    chunk = sorted_scores.iloc[curr:end]
                    Y2, Y1 = chunk.max(), chunk.min()
                    T2, T1 = cfg['t_max'], cfg['t_min']
                    def linear(Y): return (T2+T1)/2 if Y2==Y1 else T1 + ((Y-Y1)*(T2-T1))/(Y2-Y1)
                    result.loc[chunk.index] = chunk.apply(linear)
                    curr = end
                return result.round()

            def calc_ranks(dframe, target_col, rank_base_name):
                yr_rk = f"{rank_base_name}年排"
                cl_rk = f"{rank_base_name}班排"
                dframe[yr_rk] = dframe[target_col].rank(ascending=False, method='min')
                if self.selected_class_col in dframe.columns:
                    dframe[cl_rk] = dframe.groupby(self.selected_class_col)[target_col].rank(ascending=False, method='min')
                else:
                    dframe[cl_rk] = None
                return yr_rk, cl_rk

            # === 准备收集列 ===
            # 用于计算总分的列表
            cols_for_raw_total = []    # 存放所有科目的【原始】分数列名
            cols_for_final_total = []  # 存放原始科目+赋分科目【赋分后】的分数列名
            
            output_cols_order = []     # 记录中间成绩列的展示顺序

            # 1. 处理原始科目 (语数外+物理历史)
            for sub in self.selected_raw:
                df[sub] = pd.to_numeric(df[sub], errors='coerce')
                yr_rk, cl_rk = calc_ranks(df, sub, sub)
                
                # 原始科目：既算在原始总分里，也算在最终总分里
                cols_for_raw_total.append(sub)
                cols_for_final_total.append(sub)
                
                output_cols_order.extend([sub, yr_rk, cl_rk])

            # 2. 处理赋分科目 (化生政地)
            for sub in self.selected_assign:
                # 原始成绩
                df[sub] = pd.to_numeric(df[sub], errors='coerce')
                
                # 赋分成绩
                assigned_col_name = f"{sub}赋分"
                df[assigned_col_name] = calculate_assigned_score(df[sub])
                
                # 计算赋分排 (题目要求：赋分科目只排赋分后的)
                yr_rk, cl_rk = calc_ranks(df, assigned_col_name, assigned_col_name)
                
                # 归档
                cols_for_raw_total.append(sub)            # 原始总分用这一列
                cols_for_final_total.append(assigned_col_name) # 最终总分用这一列
                
                # 展示顺序: 原始 -> 赋分 -> 年排 -> 班排
                output_cols_order.extend([sub, assigned_col_name, yr_rk, cl_rk])

            # 3. 计算【原始总分】及排名
            df["原始总分"] = df[cols_for_raw_total].sum(axis=1, min_count=1)
            raw_yr_rk, raw_cl_rk = calc_ranks(df, "原始总分", "原始总分")
            raw_total_group = ["原始总分", raw_yr_rk, raw_cl_rk]

            # 4. 计算【最终总分】(赋分后) 及排名
            df["总分"] = df[cols_for_final_total].sum(axis=1, min_count=1)
            final_yr_rk, final_cl_rk = calc_ranks(df, "总分", "总分")
            final_total_group = ["总分", final_yr_rk, final_cl_rk]

            # 默认按最终总分年排排序
            df = df.sort_values(final_yr_rk)

            # 5. 构建最终表格顺序
            # 基础信息 + 单科详情 + 原始总分(新) + 最终总分
            all_generated_cols = set(output_cols_order + raw_total_group + final_total_group)
            base_info_cols = [c for c in df.columns if c not in all_generated_cols]
            
            # 关键：将 raw_total_group 放在 final_total_group 前面
            final_order = base_info_cols + output_cols_order + raw_total_group + final_total_group
            
            final_order = [c for c in final_order if c in df.columns]
            self.df_result = df[final_order]

            self.after(0, self.finish_calculation)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("计算错误", str(e)))
            self.after(0, self.stop_loading_ui)

    def finish_calculation(self):
        self.stop_loading_ui()
        self.status_label.configure(text="✅ 计算完成！包含原始总分与赋分总分。")
        self.btn_export.configure(state="normal", fg_color="#2CC985", text="导出 Excel 结果")
        messagebox.showinfo("成功", "计算完成！\n已新增【原始总分】及其排名，并排在【赋分总分】之前。")

    def stop_loading_ui(self):
        self.progressbar.stop()
        self.progressbar.configure(mode="determinate")
        self.progressbar.set(1)
        self.btn_calc.configure(state="normal")

    def export_file(self):
        save_path = filedialog.asksaveasfilename(title="保存结果", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="赋分结果_ProMax.xlsx")
        if save_path:
            try:
                self.df_result.to_excel(save_path, index=False)
                messagebox.showinfo("导出成功", f"文件已保存至:\n{save_path}")
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                messagebox.showerror("保存失败", str(e))

if __name__ == "__main__":
    app = GaokaoApp()
    app.mainloop()
