import sys
import os

# --- 【关键修复】针对 Linux/PyInstaller 丢失模块的强制导入 ---
try:
    import PIL._tkinter_finder
except ImportError:
    pass
# -------------------------------------------------------

import threading
import json
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
import requests
from docx import Document
from docx.shared import Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 字体自动适配 (防止Linux乱码) ---
DEFAULT_FONT = "Helvetica"
SYSTEM_PLATFORM = sys.platform
if SYSTEM_PLATFORM.startswith('win'):
    MAIN_FONT_NAME = "微软雅黑"
elif SYSTEM_PLATFORM.startswith('darwin'): # macOS
    MAIN_FONT_NAME = "PingFang SC"
else: # Linux
    MAIN_FONT_NAME = "WenQuanYi Micro Hei" # Linux 常用中文字体

class LessonPlanWriter(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero") 
        self.title("金塔县中学教案智能生成助手 - 多课时专业版")
        self.geometry("1280x900")
        
        # 核心数据存储：{ 1: {data...}, 2: {data...} }
        self.lesson_data = {} 
        self.active_period = 1 # 当前正在编辑的课时指针
        
        # 状态变量
        self.is_generating = False
        self.stop_flag = False
        self.api_key_var = tk.StringVar()
        self.total_periods_var = tk.IntVar(value=1)
        self.current_period_disp_var = tk.StringVar(value="1") # 用于Combobox显示
        
        self.setup_ui()
        # 初始化第一课时的空白数据结构
        self.save_current_data_to_memory(1)

    def setup_ui(self):
        # --- 顶部：全局设置 ---
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill=X)
        
        # API Key
        ttk.Label(top_frame, text="API Key:", width=8).pack(side=LEFT)
        ttk.Entry(top_frame, textvariable=self.api_key_var, show="*", width=20).pack(side=LEFT, padx=5)
        
        # 课题
        ttk.Label(top_frame, text="课题:", width=6).pack(side=LEFT, padx=(10, 0))
        self.topic_entry = ttk.Entry(top_frame, width=20)
        self.topic_entry.pack(side=LEFT, padx=5)
        self.topic_entry.insert(0, "离子反应")

        # --- 课时管理区域 ---
        period_frame = ttk.Labelframe(top_frame, text="课时进度管理", padding=(5, 2), bootstyle="primary")
        period_frame.pack(side=LEFT, padx=20)
        
        ttk.Label(period_frame, text="本课题共").pack(side=LEFT)
        # 总课时调整
        self.total_spin = ttk.Spinbox(period_frame, from_=1, to=10, width=2, textvariable=self.total_periods_var, command=self.update_period_list)
        self.total_spin.pack(side=LEFT, padx=2)
        ttk.Label(period_frame, text="课时  |  正在编辑第").pack(side=LEFT)
        
        # 当前课时切换 (Combobox)
        self.period_combo = ttk.Combobox(period_frame, values=[1], width=2, state="readonly", textvariable=self.current_period_disp_var)
        self.period_combo.current(0)
        self.period_combo.pack(side=LEFT, padx=2)
        self.period_combo.bind("<<ComboboxSelected>>", self.handle_period_switch) # 绑定切换事件
        
        ttk.Label(period_frame, text="课时").pack(side=LEFT)

        # 教案类型
        ttk.Label(top_frame, text="类型:", width=5).pack(side=LEFT, padx=(20, 0))
        self.type_combo = ttk.Combobox(top_frame, values=["详案", "简案"], state="readonly", width=6)
        self.type_combo.current(0)
        self.type_combo.pack(side=LEFT, padx=5)

        # --- 中间主体：分两栏 ---
        # 修复：Linux下必须用 Panedwindow (小写w)
        main_pane = ttk.Panedwindow(self, orient=HORIZONTAL)
        main_pane.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        # --- 左侧：框架 ---
        left_frame = ttk.Labelframe(main_pane, text="1. 本课时教学框架", padding=10)
        main_pane.add(left_frame, weight=1)
        
        # 滚动条支持
        left_canvas = tk.Canvas(left_frame)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        self.scrollable_frame = ttk.Frame(left_canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=scrollbar.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.fields = {}
        # 字段定义
        labels = [
            ("章节名称", "chapter", 2),
            ("本课时教学目标 (纯文本)", "objectives", 8),
            ("本课时重点", "key_points", 4),
            ("本课时难点", "difficulties", 4),
            ("教学方法", "methods", 3),
            ("作业设计", "homework", 4),
        ]
        
        # 字体样式
        font_bold = (MAIN_FONT_NAME, 9, "bold")
        font_norm = (MAIN_FONT_NAME, 9)

        for text, key, height in labels:
            lbl = ttk.Label(self.scrollable_frame, text=text, font=font_bold)
            lbl.pack(anchor=W, pady=(5, 0))
            txt = tk.Text(self.scrollable_frame, height=height, width=40, font=font_norm)
            txt.pack(fill=X, pady=(0, 5))
            self.fields[key] = txt
        
        ttk.Button(left_frame, text="生成当前课时框架", command=self.generate_framework, bootstyle="info").pack(fill=X, pady=5)

        # --- 右侧：过程 ---
        right_frame = ttk.Labelframe(main_pane, text="2. 本课时教学过程 (40分钟/纯文本)", padding=10)
        main_pane.add(right_frame, weight=2)
        
        # 额外指令
        cmd_frame = ttk.Frame(right_frame)
        cmd_frame.pack(fill=X, pady=5)
        ttk.Label(cmd_frame, text="额外指令:").pack(side=LEFT)
        self.instruction_entry = ttk.Entry(cmd_frame)
        self.instruction_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
        self.instruction_entry.insert(0, "环节清晰，学生活动具体")

        # 文本框
        self.process_text = ScrolledText(right_frame, font=(MAIN_FONT_NAME, 10))
        self.process_text.pack(fill=BOTH, expand=True, pady=5)
        
        # 底部按钮
        ctrl_frame = ttk.Frame(right_frame)
        ctrl_frame.pack(fill=X, pady=5)
        ttk.Button(ctrl_frame, text="撰写当前课时过程", command=self.start_writing_process, bootstyle="success").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="停止", command=self.stop_generation, bootstyle="danger").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="清空当前页", command=self.clear_current, bootstyle="secondary").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="导出所有课时 (Word)", command=self.export_word, bootstyle="warning").pack(side=RIGHT, padx=5)

        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(self, textvariable=self.status_var, relief=SUNKEN, anchor=W).pack(fill=X, side=BOTTOM)

    # --- 逻辑处理 ---

    def update_period_list(self):
        """更新总课时列表"""
        try:
            total = int(self.total_spin.get())
            current_vals = [i for i in range(1, total + 1)]
            self.period_combo['values'] = current_vals
            # 如果当前所在课时 > 新的总课时，重置回1
            if self.active_period > total:
                self.period_combo.current(0)
                self.handle_period_switch(None)
        except:
            pass

    def handle_period_switch(self, event):
        """切换课时逻辑：保存旧的 -> 加载新的"""
        try:
            new_period = int(self.period_combo.get())
        except ValueError:
            return

        if new_period == self.active_period:
            return

        # 1. 保存当前界面内容到内存
        self.save_current_data_to_memory(self.active_period)
        
        # 2. 从内存加载新课时内容
        self.load_data_from_memory(new_period)
        
        # 3. 更新指针
        self.active_period = new_period

    def save_current_data_to_memory(self, period):
        data = {
            'chapter': self.fields['chapter'].get("1.0", END).strip(),
            'objectives': self.fields['objectives'].get("1.0", END).strip(),
            'key_points': self.fields['key_points'].get("1.0", END).strip(),
            'difficulties': self.fields['difficulties'].get("1.0", END).strip(),
            'methods': self.fields['methods'].get("1.0", END).strip(),
            'homework': self.fields['homework'].get("1.0", END).strip(),
            'process': self.process_text.get("1.0", END).strip()
        }
        self.lesson_data[period] = data

    def load_data_from_memory(self, period):
        data = self.lesson_data.get(period, {})
        
        # 先清空
        for key in self.fields:
            self.fields[key].delete("1.0", END)
        self.process_text.delete("1.0", END)
        
        # 后填入
        if data:
            for key in self.fields:
                if key in data:
                    self.fields[key].insert("1.0", data[key])
            if 'process' in data:
                self.process_text.insert("1.0", data['process'])

    def clean_text(self, text):
        """深度清洗 Markdown 符号"""
        # 替换加粗
        text = text.replace("**", "").replace("__", "")
        # 替换代码块
        text = text.replace("```json", "").replace("```", "")
        # 逐行处理标题符号
        lines = []
        for line in text.split('\n'):
            clean_line = line.strip()
            # 去除开头的 # (标题)
            while clean_line.startswith("#"):
                clean_line = clean_line[1:].strip()
            lines.append(clean_line)
        return "\n".join(lines)

    def get_api_key(self):
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showerror("错误", "请输入 DeepSeek API Key")
            return None
        return key

    def stop_generation(self):
        if self.is_generating:
            self.stop_flag = True
            self.status_var.set("正在停止...")

    def clear_current(self):
        if messagebox.askyesno("确认", f"确定要清空第 {self.active_period} 课时的内容吗？"):
            for key in self.fields:
                self.fields[key].delete("1.0", END)
            self.process_text.delete("1.0", END)

    # --- AI 生成逻辑 ---

    def generate_framework(self):
        api_key = self.get_api_key()
        if not api_key: return
        
        topic = self.topic_entry.get()
        current_p = self.active_period
        total_p = self.total_periods_var.get()
        
        self.is_generating = True
        self.stop_flag = False
        threading.Thread(target=self._thread_generate_framework, args=(api_key, topic, current_p, total_p)).start()

    def _thread_generate_framework(self, api_key, topic, current_p, total_p):
        self.status_var.set(f"正在生成第 {current_p}/{total_p} 课时框架...")
        
        prompt = f"""
        任务：为高中化学课题《{topic}》设计第 {current_p} 课时的教案框架（全课共 {total_p} 课时）。
        【严格要求】
        1. 必须使用纯文本，严禁使用Markdown（不要用**加粗**，不要用#标题）。
        2. 请严格按照以下JSON格式返回，Key必须保持一致：
        {{
            "chapter": "必修第一册 第二章...",
            "objectives": "1. 知识与技能...\\n2. 过程与方法...",
            "key_points": "本节课的重点...",
            "difficulties": "本节课的难点...",
            "methods": "讲授法、实验法...",
            "homework": "课后习题..."
        }}
        """
        
        try:
            url = "https://api.deepseek.com/chat/completions"
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
            data = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "stream": False
            }
            
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 200:
                raw_content = response.json()['choices'][0]['message']['content']
                # 预处理 JSON 字符串
                json_str = raw_content.replace("```json", "").replace("```", "").strip()
                data = json.loads(json_str)
                
                # 二次清洗内容文本
                for k, v in data.items():
                    data[k] = self.clean_text(v)
                
                self.after(0, lambda: self._update_framework_ui(data))
                self.status_var.set("框架生成完毕")
            else:
                self.status_var.set(f"API错误: {response.status_code}")
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
        finally:
            self.is_generating = False

    def _update_framework_ui(self, data):
        for key, value in data.items():
            if key in self.fields:
                self.fields[key].delete("1.0", END)
                self.fields[key].insert("1.0", value)

    def start_writing_process(self):
        api_key = self.get_api_key()
        if not api_key: return
        
        context = {k: v.get("1.0", END).strip() for k, v in self.fields.items()}
        topic = self.topic_entry.get()
        instruction = self.instruction_entry.get()
        plan_type = self.type_combo.get()
        current_p = self.active_period
        
        self.is_generating = True
        self.stop_flag = False
        threading.Thread(target=self._thread_write_process, args=(api_key, topic, context, instruction, plan_type, current_p)).start()

    def _thread_write_process(self, api_key, topic, context, instruction, plan_type, current_p):
        self.status_var.set(f"正在撰写第 {current_p} 课时过程...")
        
        prompt = f"""
        任务：撰写高中化学《{topic}》第 {current_p} 课时的“教学过程与师生活动”。
        
        【输入信息】
        目标：{context['objectives']}
        重难点：{context['key_points']}
        
        【严格限制 - 必须遵守】
        1. 输出格式：纯文本！绝对不要使用Markdown（禁止**加粗**，禁止###标题）。
        2. 时间控制：本节课时长严格控制在40分钟。请合理分配各环节时间。
        3. 写作风格：{plan_type}。{instruction}
        
        【内容结构】
        请按顺序撰写：
        一、导入新课（约x分钟）
        二、...
        三、...
        四、课堂小结
        """

        url = "https://api.deepseek.com/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "stream": True
        }

        try:
            response = requests.post(url, headers=headers, json=data, stream=True)
            for line in response.iter_lines():
                if self.stop_flag: break
                if line:
                    decoded_line = line.decode('utf-8').replace("data: ", "")
                    if decoded_line != "[DONE]":
                        try:
                            json_line = json.loads(decoded_line)
                            content = json_line['choices'][0]['delta'].get('content', '')
                            if content:
                                # 实时清洗
                                content = self.clean_text(content)
                                self.after(0, lambda c=content: self.process_text.insert(END, c))
                                self.after(0, lambda: self.process_text.see(END))
                        except:
                            pass
            self.status_var.set("撰写完成")
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
        finally:
            self.is_generating = False

    def export_word(self):
        """导出所有课时到同一个Word文档"""
        # 1. 强制保存当前界面数据
        self.save_current_data_to_memory(self.active_period)
        
        filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if not filename: return

        try:
            doc = Document()
            doc.styles['Normal'].font.name = u'宋体'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            topic = self.topic_entry.get()
            total_p = self.total_periods_var.get()
            
            # 循环导出每一课时
            for i in range(1, total_p + 1):
                data = self.lesson_data.get(i, {})
                if not data: continue # 空课时跳过
                
                if i > 1: doc.add_page_break() # 分页
                
                # 标题
                p_title = doc.add_heading(f"第 {i} 课时教案", level=1)
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # 创建符合金塔县中学模版的表格 (8行4列)
                table = doc.add_table(rows=8, cols=4)
                table.style = 'Table Grid'
                table.autofit = False
                
                # 设置默认行高
                for row in table.rows:
                    row.height = Cm(1.2)

                # R1: 课题 | 时间
                table.cell(0, 0).text = "课题"
                table.cell(0, 1).text = topic
                table.cell(0, 2).text = "时间"
                table.cell(0, 3).text = datetime.now().strftime("%Y-%m-%d")

                # R2: 章节 | 课时
                table.cell(1, 0).text = "课程章节"
                table.cell(1, 1).text = data.get('chapter', '')
                table.cell(1, 2).text = "本节课时"
                table.cell(1, 3).text = f"第 {i} 课时 (共 {total_p} 课时)"

                # R3: 课标 (合并)
                table.cell(2, 0).merge(table.cell(2, 3))
                table.cell(2, 0).text = f"课程标准:\n{data.get('standard', '')}" # 兼容旧数据

                # R4: 目标 (合并)
                table.cell(3, 0).merge(table.cell(3, 3))
                table.cell(3, 0).text = f"教学目标:\n{data.get('objectives', '')}"

                # R5: 重点/难点/方法 (合并显示)
                table.cell(4, 0).merge(table.cell(4, 3))
                p = table.cell(4, 0).paragraphs[0]
                p.add_run("教学重点：").bold = True
                p.add_run(f"{data.get('key_points', '')}\n")
                p.add_run("教学难点：").bold = True
                p.add_run(f"{data.get('difficulties', '')}\n")
                p.add_run("教学方法：").bold = True
                p.add_run(f"{data.get('methods', '')}")

                # R6: 过程 (合并)
                table.cell(5, 0).merge(table.cell(5, 3))
                cell = table.cell(5, 0)
                cell.text = "教学过程与师生活动 (40分钟)"
                # 写入过程内容
                cell.add_paragraph(data.get('process', ''))

                # R7: 作业 (合并)
                table.cell(6, 0).merge(table.cell(6, 3))
                table.cell(6, 0).text = f"作业设计:\n{data.get('homework', '')}"

                # R8: 反思 (合并)
                table.cell(7, 0).merge(table.cell(7, 3))
                table.cell(7, 0).text = "课后反思:\n"

            doc.save(filename)
            messagebox.showinfo("成功", f"已成功导出 {total_p} 个课时的教案到 Word！")
            
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    app = LessonPlanWriter()
    app.mainloop()
