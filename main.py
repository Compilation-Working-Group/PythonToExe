import sys
import os

# --- 针对 Linux/PyInstaller 丢失模块的强制导入 ---
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
from datetime import datetime

# --- 字体自动适配 ---
DEFAULT_FONT = "Helvetica"
SYSTEM_PLATFORM = sys.platform
if SYSTEM_PLATFORM.startswith('win'):
    MAIN_FONT_NAME = "微软雅黑"
elif SYSTEM_PLATFORM.startswith('darwin'):
    MAIN_FONT_NAME = "PingFang SC"
else:
    MAIN_FONT_NAME = "WenQuanYi Micro Hei"

class LessonPlanWriter(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero") 
        self.title("金塔县中学教案助手 - 新课标素养版")
        self.geometry("1300x950")
        
        # 核心数据存储
        self.lesson_data = {} 
        self.active_period = 1 
        
        # 状态变量
        self.is_generating = False
        self.stop_flag = False
        self.api_key_var = tk.StringVar()
        self.total_periods_var = tk.IntVar(value=1)
        self.current_period_disp_var = tk.StringVar(value="1")
        
        self.setup_ui()
        self.save_current_data_to_memory(1)

    def setup_ui(self):
        # --- 顶部：全局设置 ---
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill=X)
        
        ttk.Label(top_frame, text="API Key:", width=8).pack(side=LEFT)
        ttk.Entry(top_frame, textvariable=self.api_key_var, show="*", width=20).pack(side=LEFT, padx=5)
        
        ttk.Label(top_frame, text="课题:", width=5).pack(side=LEFT, padx=(10, 0))
        self.topic_entry = ttk.Entry(top_frame, width=18)
        self.topic_entry.pack(side=LEFT, padx=5)
        self.topic_entry.insert(0, "离子反应")

        # --- 课时管理区域 ---
        period_frame = ttk.Labelframe(top_frame, text="进度控制", padding=(5, 2), bootstyle="primary")
        period_frame.pack(side=LEFT, padx=15)
        
        ttk.Label(period_frame, text="共").pack(side=LEFT)
        self.total_spin = ttk.Spinbox(period_frame, from_=1, to=10, width=2, textvariable=self.total_periods_var, command=self.update_period_list)
        self.total_spin.pack(side=LEFT, padx=2)
        ttk.Label(period_frame, text="课时 | 编辑第").pack(side=LEFT)
        
        self.period_combo = ttk.Combobox(period_frame, values=[1], width=2, state="readonly", textvariable=self.current_period_disp_var)
        self.period_combo.current(0)
        self.period_combo.pack(side=LEFT, padx=2)
        self.period_combo.bind("<<ComboboxSelected>>", self.handle_period_switch)
        ttk.Label(period_frame, text="课时").pack(side=LEFT)

        # 教案类型
        ttk.Label(top_frame, text="类型:", width=5).pack(side=LEFT, padx=(15, 0))
        self.type_combo = ttk.Combobox(top_frame, values=["详案", "简案"], state="readonly", width=5)
        self.type_combo.current(0)
        self.type_combo.pack(side=LEFT, padx=5)

        # --- 中间主体 ---
        main_pane = ttk.Panedwindow(self, orient=HORIZONTAL)
        main_pane.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        # --- 左侧：框架设计 ---
        left_frame = ttk.Labelframe(main_pane, text="1. 本课时设计框架", padding=10)
        main_pane.add(left_frame, weight=1)
        
        # 滚动区域
        left_canvas = tk.Canvas(left_frame)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        self.scrollable_frame = ttk.Frame(left_canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=scrollbar.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.fields = {}
        
        # 字体
        font_bold = (MAIN_FONT_NAME, 9, "bold")
        font_norm = (MAIN_FONT_NAME, 9)

        # --- 新增：自定义本课时内容 ---
        lbl_custom = ttk.Label(self.scrollable_frame, text="★ 本课时自定义教学内容 (可选，留空则AI自动规划):", font=font_bold, bootstyle="danger")
        lbl_custom.pack(anchor=W, pady=(0, 0))
        txt_custom = tk.Text(self.scrollable_frame, height=3, width=40, font=font_norm, bg="#f0f0f0")
        txt_custom.pack(fill=X, pady=(0, 10))
        self.fields['custom_content'] = txt_custom
        
        # 其他字段
        labels = [
            ("章节名称", "chapter", 1),
            ("素养导向教学目标 (通过...培养...)", "objectives", 8),
            ("教学重点", "key_points", 3),
            ("教学难点", "difficulties", 3),
            ("教学方法", "methods", 2),
            ("作业设计", "homework", 3),
        ]
        
        for text, key, height in labels:
            lbl = ttk.Label(self.scrollable_frame, text=text, font=font_bold)
            lbl.pack(anchor=W, pady=(5, 0))
            txt = tk.Text(self.scrollable_frame, height=height, width=40, font=font_norm)
            txt.pack(fill=X, pady=(0, 5))
            self.fields[key] = txt
        
        ttk.Button(left_frame, text="生成当前课时框架 (按自定义或自动)", command=self.generate_framework, bootstyle="info").pack(fill=X, pady=5)

        # --- 右侧：过程撰写 ---
        right_frame = ttk.Labelframe(main_pane, text="2. 教学过程 (40分钟/纯文本)", padding=10)
        main_pane.add(right_frame, weight=2)
        
        cmd_frame = ttk.Frame(right_frame)
        cmd_frame.pack(fill=X, pady=5)
        ttk.Label(cmd_frame, text="额外指令:").pack(side=LEFT)
        self.instruction_entry = ttk.Entry(cmd_frame)
        self.instruction_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
        self.instruction_entry.insert(0, "体现学生主体地位，探究活动详实")

        self.process_text = ScrolledText(right_frame, font=(MAIN_FONT_NAME, 10))
        self.process_text.pack(fill=BOTH, expand=True, pady=5)
        
        # 底部按钮
        ctrl_frame = ttk.Frame(right_frame)
        ctrl_frame.pack(fill=X, pady=5)
        ttk.Button(ctrl_frame, text="撰写过程", command=self.start_writing_process, bootstyle="success").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="停止", command=self.stop_generation, bootstyle="danger").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="清空当前", command=self.clear_current, bootstyle="secondary").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="导出Word (全课时)", command=self.export_word, bootstyle="warning").pack(side=RIGHT, padx=5)

        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(self, textvariable=self.status_var, relief=SUNKEN, anchor=W).pack(fill=X, side=BOTTOM)

    # --- 逻辑处理 ---

    def update_period_list(self):
        try:
            total = int(self.total_spin.get())
            current_vals = [i for i in range(1, total + 1)]
            self.period_combo['values'] = current_vals
            if self.active_period > total:
                self.period_combo.current(0)
                self.handle_period_switch(None)
        except:
            pass

    def handle_period_switch(self, event):
        try:
            new_period = int(self.period_combo.get())
        except ValueError:
            return
        if new_period == self.active_period:
            return
        self.save_current_data_to_memory(self.active_period)
        self.load_data_from_memory(new_period)
        self.active_period = new_period

    def save_current_data_to_memory(self, period):
        # 保存所有字段，包括新增的 custom_content
        data = {key: self.fields[key].get("1.0", END).strip() for key in self.fields}
        data['process'] = self.process_text.get("1.0", END).strip()
        self.lesson_data[period] = data

    def load_data_from_memory(self, period):
        data = self.lesson_data.get(period, {})
        for key in self.fields:
            self.fields[key].delete("1.0", END)
        self.process_text.delete("1.0", END)
        
        if data:
            for key in self.fields:
                if key in data:
                    self.fields[key].insert("1.0", data[key])
            if 'process' in data:
                self.process_text.insert("1.0", data['process'])

    def clean_text(self, text):
        """清洗 Markdown"""
        text = text.replace("**", "").replace("__", "")
        text = text.replace("```json", "").replace("```", "")
        lines = []
        for line in text.split('\n'):
            clean_line = line.strip()
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
        if messagebox.askyesno("确认", f"清空第 {self.active_period} 课时？"):
            for key in self.fields:
                self.fields[key].delete("1.0", END)
            self.process_text.delete("1.0", END)

    # --- AI 生成逻辑 (核心修改) ---

    def generate_framework(self):
        api_key = self.get_api_key()
        if not api_key: return
        
        topic = self.topic_entry.get()
        current_p = self.active_period
        total_p = self.total_periods_var.get()
        
        # 获取用户自定义内容
        custom_content = self.fields['custom_content'].get("1.0", END).strip()
        
        self.is_generating = True
        self.stop_flag = False
        threading.Thread(target=self._thread_generate_framework, args=(api_key, topic, current_p, total_p, custom_content)).start()

    def _thread_generate_framework(self, api_key, topic, current_p, total_p, custom_content):
        self.status_var.set(f"正在生成第 {current_p} 课时框架...")
        
        # 构建动态 Prompt
        content_instruction = ""
        if custom_content:
            content_instruction = f"【特别注意】用户已指定本课时(第{current_p}课时)的教学内容为：『{custom_content}』。请务必只围绕此内容设计，不要涉及其他课时的内容。"
        else:
            content_instruction = f"请根据高中化学常规教学逻辑，自行规划第{current_p}课时（共{total_p}课时）的核心内容。"

        prompt = f"""
        任务：为高中化学课题《{topic}》设计第 {current_p} 课时的教案框架。
        {content_instruction}

        【核心要求 - 必须严格执行】
        1. **教学目标改革**：严禁使用旧的“三维目标”（知识与技能等）。必须采用**新课标素养导向目标**。
           - 格式范例：“通过实验探究……，能从微观角度辨析……，培养宏观辨识与微观探析的素养。”
           - 语气：“通过……（活动），学生能够……（成果）”。
        2. 纯文本格式：禁止Markdown，禁止加粗。
        3. 请返回标准JSON格式：
        {{
            "chapter": "所属章节",
            "objectives": "素养导向的教学目标（不要分三维，写成一段或几点）",
            "key_points": "本课时重点",
            "difficulties": "本课时难点",
            "methods": "教学方法",
            "homework": "作业"
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
                json_str = raw_content.replace("```json", "").replace("```", "").strip()
                data = json.loads(json_str)
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
        # 更新除 custom_content 外的其他字段
        for key, value in data.items():
            if key in self.fields and key != 'custom_content':
                self.fields[key].delete("1.0", END)
                self.fields[key].insert("1.0", value)

    def start_writing_process(self):
        api_key = self.get_api_key()
        if not api_key: return
        
        # 收集上下文
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
        
        # 将自定义内容也加入 Prompt
        custom_content = context.get('custom_content', '')
        custom_hint = ""
        if custom_content:
            custom_hint = f"本课时核心内容锁定为：{custom_content}。"

        prompt = f"""
        任务：撰写高中化学《{topic}》第 {current_p} 课时的“教学过程与师生活动”。
        
        【输入信息】
        {custom_hint}
        素养目标：{context['objectives']}
        重难点：{context['key_points']}
        
        【严格限制】
        1. 格式：纯文本，无Markdown。
        2. 时长：严格控制在40分钟。
        3. 风格：{plan_type}。{instruction}
        4. 理念：体现新课标，注重“教-学-评”一体化，突出学生探究。
        
        【输出结构】
        按：环节名称（时间）- 教师活动 - 学生活动 - 设计意图（体现素养培养） 撰写。
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
        self.save_current_data_to_memory(self.active_period)
        filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if not filename: return

        try:
            doc = Document()
            doc.styles['Normal'].font.name = u'宋体'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            topic = self.topic_entry.get()
            total_p = self.total_periods_var.get()
            
            for i in range(1, total_p + 1):
                data = self.lesson_data.get(i, {})
                if not data: continue 
                
                if i > 1: doc.add_page_break() 
                
                p_title = doc.add_heading(f"第 {i} 课时教案", level=1)
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                table = doc.add_table(rows=8, cols=4)
                table.style = 'Table Grid'
                table.autofit = False
                
                for row in table.rows:
                    row.height = Cm(1.2)

                table.cell(0, 0).text = "课题"
                table.cell(0, 1).text = topic
                table.cell(0, 2).text = "时间"
                table.cell(0, 3).text = datetime.now().strftime("%Y-%m-%d")

                table.cell(1, 0).text = "课程章节"
                table.cell(1, 1).text = data.get('chapter', '')
                table.cell(1, 2).text = "本节课时"
                
                # 如果有自定义内容，最好在导出时也体现一下
                custom_info = data.get('custom_content', '')
                info_text = f"第 {i} 课时 (共 {total_p} 课时)"
                if custom_info:
                    info_text += f"\n内容：{custom_info}"
                table.cell(1, 3).text = info_text

                # 课标
                table.cell(2, 0).merge(table.cell(2, 3))
                table.cell(2, 0).text = f"课程标准:\n{data.get('standard', '（AI自动匹配相关课标）')}" 

                # 教学目标 (素养导向)
                table.cell(3, 0).merge(table.cell(3, 3))
                table.cell(3, 0).text = f"素养导向教学目标:\n{data.get('objectives', '')}"

                # 重点难点
                table.cell(4, 0).merge(table.cell(4, 3))
                p = table.cell(4, 0).paragraphs[0]
                p.add_run("教学重点：").bold = True
                p.add_run(f"{data.get('key_points', '')}\n")
                p.add_run("教学难点：").bold = True
                p.add_run(f"{data.get('difficulties', '')}\n")
                p.add_run("教学方法：").bold = True
                p.add_run(f"{data.get('methods', '')}")

                # 过程
                table.cell(5, 0).merge(table.cell(5, 3))
                cell = table.cell(5, 0)
                cell.text = "教学过程与师生活动 (40分钟)"
                cell.add_paragraph(data.get('process', ''))

                # 作业
                table.cell(6, 0).merge(table.cell(6, 3))
                table.cell(6, 0).text = f"作业设计:\n{data.get('homework', '')}"

                # 反思
                table.cell(7, 0).merge(table.cell(7, 3))
                table.cell(7, 0).text = "课后反思:\n"

            doc.save(filename)
            messagebox.showinfo("成功", f"已导出 {total_p} 个课时的教案！")
            
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    app = LessonPlanWriter()
    app.mainloop()
