import threading
import json
import time
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
import requests
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

class LessonPlanWriter(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero") # 可选 themes: cosmo, flatly, journal, superhero
        self.title("金塔县中学教案智能生成助手 - DeepSeek驱动")
        self.geometry("1200x850")
        
        # 状态变量
        self.is_generating = False
        self.stop_flag = False
        self.api_key_var = tk.StringVar()
        
        self.setup_ui()

    def setup_ui(self):
        # --- 顶部：设置与课题输入 ---
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill=X)
        
        ttk.Label(top_frame, text="DeepSeek API Key:", width=15).pack(side=LEFT)
        ttk.Entry(top_frame, textvariable=self.api_key_var, show="*", width=30).pack(side=LEFT, padx=5)
        
        ttk.Label(top_frame, text="课题名称:", width=10).pack(side=LEFT, padx=(20, 0))
        self.topic_entry = ttk.Entry(top_frame, width=30)
        self.topic_entry.pack(side=LEFT, padx=5)
        self.topic_entry.insert(0, "离子反应") # 默认示例

        ttk.Label(top_frame, text="教案类型:", width=10).pack(side=LEFT, padx=(20, 0))
        self.type_combo = ttk.Combobox(top_frame, values=["详案 (详细师生互动)", "简案 (提纲挈领)"], state="readonly", width=15)
        self.type_combo.current(0)
        self.type_combo.pack(side=LEFT, padx=5)

        # --- 中间：分两栏 ---
        # 左侧：基本信息框架（可修改）
        # 右侧：教学过程撰写（AI生成 + 人工修改）
        main_pane = ttk.PanedWindow(self, orient=HORIZONTAL)
        main_pane.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        # --- 左侧面板：教案框架 ---
        left_frame = ttk.Labelframe(main_pane, text="1. 教案框架 (自动生成/手动修改)", padding=10)
        main_pane.add(left_frame, weight=1)
        
        # 滚动容器
        left_canvas = tk.Canvas(left_frame)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        self.scrollable_frame = ttk.Frame(left_canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        left_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=scrollbar.set)
        
        left_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 框架字段
        self.fields = {}
        labels = [
            ("章节", "chapter", 3),
            ("课时信息 (如: 共2课时 第1课时)", "hours", 3),
            ("课程标准", "standard", 5),
            ("教学目标", "objectives", 6),
            ("教学重点", "key_points", 4),
            ("教学难点", "difficulties", 4),
            ("教学方法", "methods", 3),
            ("作业设计", "homework", 4),
        ]
        
        for text, key, height in labels:
            lbl = ttk.Label(self.scrollable_frame, text=text, font=("微软雅黑", 9, "bold"))
            lbl.pack(anchor=W, pady=(5, 0))
            txt = tk.Text(self.scrollable_frame, height=height, width=40, font=("微软雅黑", 9))
            txt.pack(fill=X, pady=(0, 5))
            self.fields[key] = txt
        
        # 框架操作按钮
        frame_btn_area = ttk.Frame(left_frame)
        frame_btn_area.pack(fill=X, pady=5)
        ttk.Button(frame_btn_area, text="Step 1: 生成框架", command=self.generate_framework, bootstyle="info").pack(fill=X)

        # --- 右侧面板：核心撰写 ---
        right_frame = ttk.Labelframe(main_pane, text="2. 教学过程撰写 & 导出", padding=10)
        main_pane.add(right_frame, weight=2)
        
        # 额外指令
        cmd_frame = ttk.Frame(right_frame)
        cmd_frame.pack(fill=X, pady=5)
        ttk.Label(cmd_frame, text="额外撰写指令 (可选):").pack(side=LEFT)
        self.instruction_entry = ttk.Entry(cmd_frame)
        self.instruction_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
        self.instruction_entry.insert(0, "体现新课标理念，注重实验探究")

        # 教学过程文本框
        ttk.Label(right_frame, text="教学过程与师生活动:", font=("微软雅黑", 10, "bold")).pack(anchor=W)
        self.process_text = ScrolledText(right_frame, font=("微软雅黑", 10))
        self.process_text.pack(fill=BOTH, expand=True, pady=5)
        
        # 底部控制栏
        ctrl_frame = ttk.Frame(right_frame)
        ctrl_frame.pack(fill=X, pady=5)
        
        ttk.Button(ctrl_frame, text="Step 2: 开始撰写/继续撰写", command=self.start_writing_process, bootstyle="success").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="停止", command=self.stop_generation, bootstyle="danger").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="清空内容", command=self.clear_all, bootstyle="secondary").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="导出Word教案", command=self.export_word, bootstyle="warning").pack(side=RIGHT, padx=5)

        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(self, textvariable=self.status_var, relief=SUNKEN, anchor=W).pack(fill=X, side=BOTTOM)

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

    def clear_all(self):
        if messagebox.askyesno("确认", "确定要清空所有内容吗？"):
            for key in self.fields:
                self.fields[key].delete("1.0", END)
            self.process_text.delete("1.0", END)
            self.topic_entry.delete(0, END)

    def generate_framework(self):
        """生成左侧的框架信息"""
        api_key = self.get_api_key()
        if not api_key: return
        
        topic = self.topic_entry.get()
        if not topic:
            messagebox.showwarning("提示", "请输入课题名称")
            return

        self.is_generating = True
        self.stop_flag = False
        threading.Thread(target=self._thread_generate_framework, args=(api_key, topic)).start()

    def _thread_generate_framework(self, api_key, topic):
        self.status_var.set("正在分析课题并生成框架...")
        
        prompt = f"""
        请为高中化学课题《{topic}》设计一个教案框架。
        请严格按照以下JSON格式返回内容，不要包含markdown代码块标记：
        {{
            "chapter": "所属章节名称",
            "hours": "本节共X课时，本节课为第Y课时",
            "standard": "课程标准要求",
            "objectives": "1. 知识与技能...\\n2. 过程与方法...\\n3. 情感态度与价值观...",
            "key_points": "教学重点...",
            "difficulties": "教学难点...",
            "methods": "讲授法、实验探究法等",
            "homework": "作业布置内容"
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
                content = response.json()['choices'][0]['message']['content']
                # 清理可能存在的代码块标记
                content = content.replace("```json", "").replace("```", "").strip()
                data = json.loads(content)
                
                # 更新UI
                self.after(0, lambda: self._update_framework_ui(data))
                self.status_var.set("框架生成完毕，请检查修改")
            else:
                self.status_var.set(f"API错误: {response.status_code}")
        except Exception as e:
            self.status_var.set(f"发生错误: {str(e)}")
        finally:
            self.is_generating = False

    def _update_framework_ui(self, data):
        for key, value in data.items():
            if key in self.fields:
                self.fields[key].delete("1.0", END)
                self.fields[key].insert("1.0", value)

    def start_writing_process(self):
        """生成右侧的详细教学过程"""
        api_key = self.get_api_key()
        if not api_key: return
        
        # 收集上下文
        context = {k: v.get("1.0", END).strip() for k, v in self.fields.items()}
        topic = self.topic_entry.get()
        instruction = self.instruction_entry.get()
        plan_type = self.type_combo.get()
        
        self.is_generating = True
        self.stop_flag = False
        threading.Thread(target=self._thread_write_process, args=(api_key, topic, context, instruction, plan_type)).start()

    def _thread_write_process(self, api_key, topic, context, instruction, plan_type):
        self.status_var.set("正在撰写教学过程...")
        
        prompt = f"""
        你是一位经验丰富的高中化学教师。请根据以下框架信息，撰写《{topic}》的详细“教学过程与师生活动”。
        
        【基本框架】
        教学目标：{context['objectives']}
        重点难点：{context['key_points']} & {context['difficulties']}
        教学方法：{context['methods']}
        
        【要求】
        1. 类型：{plan_type}
        2. 额外指令：{instruction}
        3. 格式：请按“教学环节”、“教师活动”、“学生活动”、“设计意图”进行组织，内容要详实具体。
        4. 直接输出教学过程内容，不要重复前面的目标等信息。
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
                if self.stop_flag:
                    break
                if line:
                    decoded_line = line.decode('utf-8').replace("data: ", "")
                    if decoded_line != "[DONE]":
                        try:
                            json_line = json.loads(decoded_line)
                            content = json_line['choices'][0]['delta'].get('content', '')
                            if content:
                                self.after(0, lambda c=content: self.process_text.insert(END, c))
                                self.after(0, lambda: self.process_text.see(END))
                        except:
                            pass
            self.status_var.set("撰写完成" if not self.stop_flag else "已停止")
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
        finally:
            self.is_generating = False

    def export_word(self):
        """导出符合金塔县中学模版的Word文档"""
        filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if not filename: return

        try:
            doc = Document()
            # 设置中文字体基础
            doc.styles['Normal'].font.name = u'宋体'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 读取数据
            data = {k: v.get("1.0", END).strip() for k, v in self.fields.items()}
            topic = self.topic_entry.get()
            process = self.process_text.get("1.0", END).strip()
            
            # --- 创建表格 ---
            # 这是一个复杂的表格，我们需要手动构建每一行来匹配您的PDF模版
            # 模版结构大约是：
            # R1: 课题 | 时间
            # R2: 章节 | 课时
            # R3: 课标
            # R4: 目标
            # R5: 重点 | 难点 | 方法 (这里PDF看似是分开的，我们用嵌套或拆分单元格模拟)
            
            table = doc.add_table(rows=8, cols=4)
            table.style = 'Table Grid'
            table.autofit = False
            
            # 设置列宽 (大致比例)
            for row in table.rows:
                row.height = Cm(1.0) # 基础高度

            # --- Row 1: 课题 & 时间 ---
            # 单元格 0: "课题" Label
            table.cell(0, 0).text = "课题"
            # 单元格 1: 课题内容 (合并一点)
            table.cell(0, 1).text = topic
            # 单元格 2: "时间" Label
            table.cell(0, 2).text = "时间"
            # 单元格 3: 时间内容
            table.cell(0, 3).text = datetime.now().strftime("%Y-%m-%d")

            # --- Row 2: 章节 & 课时 ---
            table.cell(1, 0).text = "课程章节"
            table.cell(1, 1).text = data.get('chapter', '')
            table.cell(1, 2).text = "课时安排"
            table.cell(1, 3).text = data.get('hours', '')

            # --- Row 3: 课程标准 (跨全列) ---
            table.cell(2, 0).merge(table.cell(2, 3))
            table.cell(2, 0).text = f"课程标准:\n{data.get('standard', '')}"

            # --- Row 4: 教学目标 (跨全列) ---
            table.cell(3, 0).merge(table.cell(3, 3))
            table.cell(3, 0).text = f"教学目标:\n{data.get('objectives', '')}"

            # --- Row 5: 教学分析 (重点、难点、方法) ---
            # 这行比较特殊，根据PDF，可能是几个并列的框
            # 我们将这行拆分为三个主要区域，利用换行符区分
            table.cell(4, 0).merge(table.cell(4, 3))
            # 或者我们手动重构这行：
            p = table.cell(4, 0).paragraphs[0]
            p.add_run(f"教学重点：\n{data.get('key_points', '')}\n\n").bold = True
            p.add_run(f"教学难点：\n{data.get('difficulties', '')}\n\n").bold = True
            p.add_run(f"教学方法：\n{data.get('methods', '')}").bold = True

            # --- Row 6: 教学过程 (最长的一块) ---
            table.cell(5, 0).merge(table.cell(5, 3))
            cell = table.cell(5, 0)
            cell.text = "教学过程与师生活动"
            # 添加具体内容
            p = cell.add_paragraph(process)
            
            # --- Row 7: 作业设计 ---
            table.cell(6, 0).merge(table.cell(6, 3))
            table.cell(6, 0).text = f"作业设计:\n{data.get('homework', '')}"

            # --- Row 8: 课后反思 ---
            table.cell(7, 0).merge(table.cell(7, 3))
            table.cell(7, 0).text = "课后反思:\n (课后手动填写)"

            doc.save(filename)
            messagebox.showinfo("成功", "教案已导出！")
            
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    app = LessonPlanWriter()
    app.mainloop()
