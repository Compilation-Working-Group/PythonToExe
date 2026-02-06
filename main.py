import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import sys
import threading

# 设置外观模式 (System, Dark, Light)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SmartDividerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 窗口设置
        self.title("SmartDivider - 新高考智选分班助手")
        self.geometry("800x600")
        
        # 数据容器
        self.df = None
        self.file_path = None

        # 布局容器
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧边栏 (设置区) ===
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="分班规则设定", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # 规则1：班额设置
        self.size_label = ctk.CTkLabel(self.sidebar_frame, text="目标班级人数:", anchor="w")
        self.size_label.grid(row=1, column=0, padx=20, pady=(10, 0))
        self.class_size_entry = ctk.CTkEntry(self.sidebar_frame, placeholder_text="例如: 50")
        self.class_size_entry.insert(0, "50")
        self.class_size_entry.grid(row=2, column=0, padx=20, pady=(0, 10))

        # 规则2：选科列名识别 (用逗号分隔)
        self.col_label = ctk.CTkLabel(self.sidebar_frame, text="选科列名 (逗号隔开):", anchor="w")
        self.col_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.cols_entry = ctk.CTkEntry(self.sidebar_frame, placeholder_text="科目1,科目2,科目3")
        self.cols_entry.insert(0, "首选,再选1,再选2") # 默认假设列名
        self.cols_entry.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="n")

        # === 右侧主区域 ===
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)

        # 步骤1：导入文件
        self.step1_label = ctk.CTkLabel(self.main_frame, text="步骤 1: 导入学生基础信息表 (Excel)", font=ctk.CTkFont(size=16))
        self.step1_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.import_btn = ctk.CTkButton(self.main_frame, text="选择 Excel 文件 (.xlsx)", command=self.load_excel)
        self.import_btn.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        # 日志输出区
        self.log_textbox = ctk.CTkTextbox(self.main_frame, width=400)
        self.log_textbox.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.log_textbox.insert("0.0", "等待导入数据...\n请确保Excel包含'姓名'及选科列。\n")

        # 底部操作栏
        self.action_frame = ctk.CTkFrame(self, height=50, corner_radius=0)
        self.action_frame.grid(row=1, column=1, sticky="ew", padx=0, pady=0)
        
        self.run_btn = ctk.CTkButton(self.action_frame, text="开始自动分班", fg_color="green", hover_color="darkgreen", command=self.start_processing)
        self.run_btn.pack(side="right", padx=20, pady=10)
        self.run_btn.configure(state="disabled")

    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.file_path = file_path
                self.df = pd.read_excel(file_path)
                self.log(f"成功加载文件: {os.path.basename(file_path)}")
                self.log(f"数据行数: {len(self.df)}")
                self.log(f"包含列: {', '.join(self.df.columns)}")
                self.run_btn.configure(state="normal")
            except Exception as e:
                messagebox.showerror("错误", f"无法读取文件: {e}")

    def start_processing(self):
        # 使用线程避免界面卡顿
        threading.Thread(target=self.process_classes).start()

    def process_classes(self):
        try:
            self.run_btn.configure(state="disabled")
            self.log("-" * 30)
            self.log("开始计算分班...")

            # 1. 获取参数
            try:
                max_size = int(self.class_size_entry.get())
                subject_cols = [c.strip() for c in self.cols_entry.get().split(",")]
            except ValueError:
                self.log("错误: 班级人数必须是整数。")
                return

            # 2. 验证列是否存在
            missing_cols = [c for c in subject_cols if c not in self.df.columns]
            if missing_cols:
                self.log(f"错误: Excel中找不到以下列: {missing_cols}")
                self.log("请检查左侧'选科列名'设置是否与Excel表头一致。")
                self.run_btn.configure(state="normal")
                return

            # 3. 生成组合标签 (如 "物理+化学+生物")
            self.df['选科组合'] = self.df[subject_cols].apply(lambda x: '+'.join(x.astype(str)), axis=1)
            
            # 4. 核心分班算法
            results = []
            
            # 按组合分组
            grouped = self.df.groupby('选科组合')
            
            for combo, group in grouped:
                count = len(group)
                num_classes = (count // max_size) + (1 if count % max_size > 0 else 0)
                self.log(f"组合 [{combo}]: {count}人 -> 拆分为 {num_classes} 个班")
                
                # 简单切分 (实际教学中可能需要按成绩S型分班，此处暂按名单顺序切分)
                shuffled_group = group.sample(frac=1, random_state=42).reset_index(drop=True) # 随机打乱防止同名扎堆
                
                for i in range(num_classes):
                    start = i * max_size
                    end = (i + 1) * max_size
                    sub_df = shuffled_group.iloc[start:end].copy()
                    class_name = f"{combo}-{i+1}班"
                    sub_df['拟定班级'] = class_name
                    results.append(sub_df)
            
            # 5. 合并结果
            final_df = pd.concat(results)
            
            # 6. 导出
            save_path = os.path.splitext(self.file_path)[0] + "_分班结果.xlsx"
            final_df.to_excel(save_path, index=False)
            
            self.log("-" * 30)
            self.log(f"处理完成！")
            self.log(f"结果已保存至: {save_path}")
            messagebox.showinfo("成功", "分班完成！结果文件已生成。")

        except Exception as e:
            self.log(f"发生错误: {e}")
            messagebox.showerror("运行错误", str(e))
        finally:
            self.run_btn.configure(state="normal")

if __name__ == "__main__":
    app = SmartDividerApp()
    app.mainloop()
