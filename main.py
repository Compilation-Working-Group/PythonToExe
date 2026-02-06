import customtkinter as ctk
import pandas as pd
import numpy as np
import threading
import os
import sys
import platform     
import subprocess   
from tkinter import filedialog, messagebox

# --- 全局外观设置 ---
ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")  

class GaokaoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 1. 窗口基础设置
        self.title("甘肃新高考赋分系统 Pro Max (自定义参数版) | 俞晋全名师工作室")
        self.geometry("1200x850")
        self.minsize(1000, 750)
        
        # 数据变量
        self.file_path = None
        self.df_raw = None
        self.sheet_names = []
        self.param_entries = {} 
        
        # 布局配置
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ==========================
        # === 左侧边栏 (操作区) ===
        # ==========================
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1) # 让中间空出空间

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

        # 底部按钮区 (布局行号调整以适应弹性空间)
        self.btn_calc = ctk.CTkButton(self.sidebar_frame, text="开始赋分计算", height=50, fg_color="green", font=ctk.CTkFont(size=16, weight="bold"), command=self.start_calculation)
        self.btn_calc.grid(row=11, column=0, padx=20, pady=15)
        self.btn_calc.configure(state="disabled")

        self.btn_export = ctk.CTkButton(self.sidebar_frame, text="导出结果 Excel", height=40, command=self.export_file)
        self.btn_export.grid(row=12, column=0, padx=20, pady=(0, 30))
        self.btn_export.configure(state="disabled")

        # ==========================
        # === 右侧主内容区 (Tab) ===
        # ==========================
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        # 状态栏
        self.status_label = ctk.CTkLabel(self.main_frame, text="欢迎使用！请先导入数据，然后确认【赋分标准】。", anchor="w", font=("Microsoft YaHei UI", 16))
        self.status_label.pack(fill="x", pady=(0, 10))

        # 创建选项卡
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True)
        
        # 添加三个 Tab
        self.tabview.add("科目设置")
        self.tabview.add("赋分标准设置")
        self.tabview.add("说明与规则") # 新增 Tab
        
        # --- Tab 1: 科目设置 ---
        self.setup_subject_tab()

        # --- Tab 2: 赋分参数设置 ---
        self.setup_params_tab()
        
        # --- Tab 3: 说明与规则 (新增) ---
        self.setup_help_tab()

        # 进度条
        self.progressbar = ctk.CTkProgressBar(self.main_frame, height=15)
        self.progressbar.pack(fill="x", pady=(15, 0))
        self.progressbar.set(0)

    # --------------------------
    # 界面构建辅助函数
    # --------------------------
    def setup_subject_tab(self):
        tab = self.tabview.tab("科目设置")
        
        # 滚动设置区
        self.scroll_frame = ctk.CTkScrollableFrame(tab, label_text="勾选对应列名")
        self.scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

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

    def setup_params_tab(self):
        tab = self.tabview.tab("赋分标准设置")
        
        info_lbl = ctk.CTkLabel(tab, text="请根据实际需求修改参数（默认值为甘肃省标准）。\n人数比例请输入整数（如15代表15%）。", font=("Microsoft YaHei UI", 13))
        info_lbl.pack(pady=10)

        # 参数网格容器
        grid_frame = ctk.CTkFrame(tab)
        grid_frame.pack(padx=20, pady=10)

        # 表头
        headers = ["等级", "人数比例 (%)", "赋分上限 (T2)", "赋分下限 (T1)"]
        for col, text in enumerate(headers):
            ctk.CTkLabel(grid_frame, text=text, font=("Arial", 12, "bold")).grid(row=0, column=col, padx=15, pady=10)

        # 默认数据 (甘肃标准)
        default_data = [
            ('A', '15', '100', '86'),
            ('B', '35', '85',  '71'),
            ('C', '35', '70',  '56'),
            ('D', '13', '55',  '41'),
            ('E', '2',  '40',  '30')
        ]

        self.param_entries = {} 

        for row, (grade, pct, tmax, tmin) in enumerate(default_data, start=1):
            ctk.CTkLabel(grid_frame, text=grade, font=("Arial", 14, "bold")).grid(row=row, column=0, pady=5)
            
            e_pct = ctk.CTkEntry(grid_frame, width=80, justify="center")
            e_pct.insert(0, pct)
            e_pct.grid(row=row, column=1, pady=5)
            
            e_max = ctk.CTkEntry(grid_frame, width=80, justify="center")
            e_max.insert(0, tmax)
            e_max.grid(row=row, column=2, pady=5)
            
            e_min = ctk.CTkEntry(grid_frame, width=80, justify="center")
            e_min.insert(0, tmin)
            e_min.grid(row=row, column=3, pady=5)

            self.param_entries[f"{grade}_percent"] = e_pct
            self.param_entries[f"{grade}_max"] = e_max
            self.param_entries[f"{grade}_min"] = e_min

    def setup_help_tab(self):
        """新增：说明与规则 Tab 页面的内容"""
        tab = self.tabview.tab("说明与规则")
        
        # 创建只读文本框
        textbox = ctk.CTkTextbox(tab, font=("Microsoft YaHei UI", 14), wrap="word")
        textbox.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 说明文档内容
        help_content = """【一、甘肃省新高考赋分规则说明】

1. 基本原理
   思想政治、地理、化学、生物学 4 门科目作为等级赋分科目。
   将每门科目的原始成绩从高到低划分为 A、B、C、D、E 五个等级。
   
2. 等级划分比例与赋分区间 (默认标准)
   ------------------------------------------------------------
   等级 |  人数比例  |  原始分区间   |  赋分区间 (分)
   ------------------------------------------------------------
    A   |    约15%   |  [Y2, Y1]    |  100 ~ 86
    B   |    约35%   |  [Y2, Y1]    |   85 ~ 71
    C   |    约35%   |  [Y2, Y1]    |   70 ~ 56
    D   |    约13%   |  [Y2, Y1]    |   55 ~ 41
    E   |    约 2%   |  [Y2, Y1]    |   40 ~ 30
   ------------------------------------------------------------

3. 计算公式 (等比例转换)
   假设某同学原始分为 Y，该等级原始分最高为 Y2，最低为 Y1；
   对应赋分等级最高为 T2，最低为 T1，则赋分成绩 T 计算如下：

       (Y - Y1) / (Y2 - Y1) = (T - T1) / (T2 - T1)

   最后结果 T 四舍五入取整。

============================================================

【二、软件使用注意事项】

1. Excel 文件格式要求：
   - 第一行必须是【表头】（如：姓名、班级、语文、数学...）。
   - 不能有【合并单元格】，否则可能导致读取数据错位。
   - 分数列应为纯数字，缺考请留空或填0。

2. 操作流程：
   第一步：点击左侧“导入 Excel”。
   第二步：选择正确的 Sheet 工作表。
   第三步：在“科目设置”中勾选列（区分直接计入和赋分科目）。
   第四步：点击“开始赋分计算”。
   第五步：导出结果。

3. 常见报错解决：
   - 如果提示“权限错误”，请检查是否已经在 Excel 中打开了该文件，请先关闭 Excel 再导入。
   - 如果计算结果全为空，请检查是否选对了“班级列”。

【俞晋全名师工作室 出品 | 2026】
"""
        textbox.insert("0.0", help_content)
        textbox.configure(state="disabled") # 设置为只读模式

    # --------------------------
    # 文件加载与 UI 更新逻辑
    # --------------------------
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
            self.status_label.configure(text=f"当前工作表: {sheet_name} | 请在【科目设置】页勾选")
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

    # --------------------------
    # 核心计算逻辑
    # --------------------------
    def get_user_configs(self):
        configs = []
        grades = ['A', 'B', 'C', 'D', 'E']
        try:
            for g in grades:
                pct = float(self.param_entries[f"{g}_percent"].get()) / 100.0
                t_max = int(self.param_entries[f"{g}_max"].get())
                t_min = int(self.param_entries[f"{g}_min"].get())
                
                configs.append({
                    'grade': g,
                    'percent': pct,
                    't_max': t_max,
                    't_min': t_min
                })
            return configs
        except ValueError:
            messagebox.showerror("参数错误", "赋分标准中请输入有效的数字！")
            return None

    def start_calculation(self):
        self.selected_raw = [cb.cget("text") for cb in self.raw_checkboxes if cb.get() == 1]
        self.selected_assign = [cb.cget("text") for cb in self.assign_checkboxes if cb.get() == 1]
        self.selected_class_col = self.class_col_dropdown.get()

        if not self.selected_raw and not self.selected_assign:
            messagebox.showwarning("提示", "请至少勾选一个科目！")
            return
        
        self.user_configs = self.get_user_configs()
        if not self.user_configs:
            return

        self.btn_calc.configure(state="disabled")
        self.status_label.configure(text="正在根据自定义参数计算...")
        self.progressbar.configure(mode="indeterminate")
        self.progressbar.start()
        
        threading.Thread(target=self.run_math_logic).start()

    def run_math_logic(self):
        try:
            df = self.df_raw.copy()
            grade_configs = self.user_configs 

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

            cols_for_raw_total = []    
            cols_for_final_total = []  
            output_cols_order = []     

            # 1. 原始科目
            for sub in self.selected_raw:
                df[sub] = pd.to_numeric(df[sub], errors='coerce')
                yr_rk, cl_rk = calc_ranks(df, sub, sub)
                cols_for_raw_total.append(sub)
                cols_for_final_total.append(sub)
                output_cols_order.extend([sub, yr_rk, cl_rk])

            # 2. 赋分科目
            for sub in self.selected_assign:
                df[sub] = pd.to_numeric(df[sub], errors='coerce')
                assigned_col_name = f"{sub}赋分"
                df[assigned_col_name] = calculate_assigned_score(df[sub])
                
                yr_rk, cl_rk = calc_ranks(df, assigned_col_name, assigned_col_name)
                
                cols_for_raw_total.append(sub)            
                cols_for_final_total.append(assigned_col_name) 
                output_cols_order.extend([sub, assigned_col_name, yr_rk, cl_rk])

            # 3. 原始总分
            df["原始总分"] = df[cols_for_raw_total].sum(axis=1, min_count=1)
            raw_yr_rk, raw_cl_rk = calc_ranks(df, "原始总分", "原始总分")
            raw_total_group = ["原始总分", raw_yr_rk, raw_cl_rk]

            # 4. 最终总分
            df["总分"] = df[cols_for_final_total].sum(axis=1, min_count=1)
            final_yr_rk, final_cl_rk = calc_ranks(df, "总分", "总分")
            final_total_group = ["总分", final_yr_rk, final_cl_rk]

            df = df.sort_values(final_yr_rk)

            all_generated_cols = set(output_cols_order + raw_total_group + final_total_group)
            base_info_cols = [c for c in df.columns if c not in all_generated_cols]
            
            final_order = base_info_cols + output_cols_order + raw_total_group + final_total_group
            final_order = [c for c in final_order if c in df.columns]
            self.df_result = df[final_order]

            self.after(0, self.finish_calculation)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("计算错误", str(e)))
            self.after(0, self.stop_loading_ui)

    def finish_calculation(self):
        self.stop_loading_ui()
        self.status_label.configure(text="✅ 计算完成！数据已应用当前赋分标准。")
        self.btn_export.configure(state="normal", fg_color="#2CC985", text="导出 Excel 结果")
        messagebox.showinfo("成功", "计算完成！\n请注意：本次计算使用了您在【赋分标准设置】中填写的参数。")

    def stop_loading_ui(self):
        self.progressbar.stop()
        self.progressbar.configure(mode="determinate")
        self.progressbar.set(1)
        self.btn_calc.configure(state="normal")

    def export_file(self):
        save_path = filedialog.asksaveasfilename(
            title="保存结果", 
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")], 
            initialfile="赋分结果_自定义参数.xlsx"
        )
        if save_path:
            try:
                self.df_result.to_excel(save_path, index=False)
                messagebox.showinfo("导出成功", f"文件已保存至:\n{save_path}")
                
                # --- 跨平台打开文件夹逻辑 ---
                folder_path = os.path.dirname(save_path)
                system_name = platform.system()
                
                try:
                    if system_name == "Windows":
                        os.startfile(folder_path)
                    elif system_name == "Darwin": # macOS
                        subprocess.call(["open", folder_path])
                    else: # Linux
                        subprocess.call(["xdg-open", folder_path])
                except Exception as open_err:
                    print(f"尝试打开文件夹失败: {open_err}")
                # -------------------------

            except Exception as e:
                messagebox.showerror("保存失败", str(e))

if __name__ == "__main__":
    app = GaokaoApp()
    app.mainloop()
