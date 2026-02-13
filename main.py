import sys
import os
import json # 新增：用于处理配置文件

# --- 兼容性修复 ---
try:
    import PIL._tkinter_finder
except ImportError:
    pass
import PIL.ImageTk 
# -----------------

import threading
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
import requests
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

# --- 字体自动适配 ---
DEFAULT_FONT = "Helvetica"
SYSTEM_PLATFORM = sys.platform
if SYSTEM_PLATFORM.startswith('win'):
    MAIN_FONT_NAME = "微软雅黑"
    UI_FONT_SIZE = 9
elif SYSTEM_PLATFORM.startswith('darwin'): 
    MAIN_FONT_NAME = "PingFang SC"
    UI_FONT_SIZE = 11
else: 
    MAIN_FONT_NAME = "WenQuanYi Micro Hei" 
    UI_FONT_SIZE = 10

# --- 配置文件路径 (用户主目录) ---
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".jinta_lesson_config.json")

class LessonPlanWriter(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly") 
        self.title("金塔县中学教案智能生成系统 v3.4 (自动保存授权版)")
        self.geometry("1350x950")
        
        self.lesson_data = {} 
        self.active_period = 1 
        
        self.is_generating = False
        self.stop_flag = False
        
        # 变量
        self.api_key = "" # 内部存储 Key，不再展示在界面上
        self.api_status_var = tk.StringVar(value="❌ 未配置")
        self.total_periods_var = tk.IntVar(value=1)
        self.current_period_disp_var = tk.StringVar(value="1")
        
        self.author_info = "设计与开发：金塔县中学化学教研组 · 俞晋全 (Yu JinQuan) | 核心驱动：DeepSeek-V3"
        
        self.load_config() # 启动时自动加载 Key
        self.setup_ui()
        self.save_current_data_to_memory(1)

    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.api_key = config.get("api_key", "")
                    if self.api_key:
                        self.api_status_var.set("✅ 已就绪 (自动加载)")
        except Exception:
            pass

    def save_config(self):
        """保存配置文件"""
        try:
            config = {"api_key": self.api_key}
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f)
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置: {str(e)}")

    def open_api_settings(self):
        """打开二级弹窗输入 Key"""
        # 使用 simpledialog 获取输入，默认值为当前 Key
        new_key = simpledialog.askstring(
            title="配置 API Key",
            prompt="请输入 DeepSeek API Key:\n(输入后将自动保存，下次无需再次输入)",
            initialvalue=self.api_key,
            parent=self
        )
        
        if new_key is not None: # 点击了确定（可能是空字符串，代表清空）
            self.api_key = new_key.strip()
            self.save_config() # 立即保存
            
            if self.api_key:
                self.api_status_var.set("✅ 已就绪")
                messagebox.showinfo("成功", "API Key 已保存！下次打开软件可直接使用。")
            else:
                self.api_status_var.set("❌ 未配置")

    def setup_ui(self):
        # ================= 顶部控制区 =================
        header_frame = ttk.Frame(self, padding=(15, 15))
        header_frame.pack(fill=X)
        
        # 1. API 设置 (改为按钮 + 状态标签)
        api_frame = ttk.Labelframe(header_frame, text="🔑 授权管理", padding=10, bootstyle="info")
        api_frame.pack(side=LEFT, fill=Y, padx=(0, 10))
        
        ttk.Button(api_frame, text="⚙️ 配置 API Key", command=self.open_api_settings, bootstyle="info").pack(side=LEFT, padx=5)
        ttk.Label(api_frame, textvariable=self.api_status_var, font=(MAIN_FONT_NAME, 9)).pack(side=LEFT, padx=5)

        # 2. 课题与进度
        topic_frame = ttk.Labelframe(header_frame, text="📚 课题与进度规划", padding=10, bootstyle="primary")
        topic_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5)
        
        f1 = ttk.Frame(topic_frame)
        f1.pack(fill=X, pady=(0, 5))
        ttk.Label(f1, text="课题名称:", font=(MAIN_FONT_NAME, UI_FONT_SIZE, "bold")).pack(side=LEFT)
        self.topic_entry = ttk.Entry(f1, width=30, bootstyle="primary")
        self.topic_entry.pack(side=LEFT, padx=5, fill=X, expand=True)
        self.topic_entry.insert(0, "离子反应")
        
        ttk.Label(f1, text="教案类型:", font=(MAIN_FONT_NAME, UI_FONT_SIZE)).pack(side=LEFT, padx=(15, 5))
        self.type_combo = ttk.Combobox(f1, values=["详案 (标准)", "简案 (提纲)"], state="readonly", width=10, bootstyle="primary")
        self.type_combo.current(0)
        self.type_combo.pack(side=LEFT)

        f2 = ttk.Frame(topic_frame)
        f2.pack(fill=X)
        ttk.Label(f2, text="总课时:", font=(MAIN_FONT_NAME, UI_FONT_SIZE)).pack(side=LEFT)
        self.total_spin = ttk.Spinbox(f2, from_=1, to=10, width=3, textvariable=self.total_periods_var, command=self.update_period_list, bootstyle="primary")
        self.total_spin.pack(side=LEFT, padx=5)
        
        ttk.Separator(f2, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=10)
        
        ttk.Label(f2, text="当前编辑:", font=(MAIN_FONT_NAME, UI_FONT_SIZE, "bold"), bootstyle="warning").pack(side=LEFT)
        ttk.Label(f2, text="第").pack(side=LEFT, padx=2)
        self.period_combo = ttk.Combobox(f2, values=[1], width=3, state="readonly", textvariable=self.current_period_disp_var, bootstyle="warning")
        self.period_combo.current(0)
        self.period_combo.pack(side=LEFT)
        self.period_combo.bind("<<ComboboxSelected>>", self.handle_period_switch)
        ttk.Label(f2, text="课时").pack(side=LEFT, padx=2)

        # 3. 全局操作区
        action_frame = ttk.Labelframe(header_frame, text="⚙️ 全局操作", padding=10, bootstyle="secondary")
        action_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))
        
        ttk.Button(action_frame, text="📥 导出全套Word教案", command=self.export_word, bootstyle="warning").pack(fill=X, pady=2)
        ttk.Button(action_frame, text="🗑️ 清空所有数据", command=self.clear_all_data, bootstyle="danger outline").pack(fill=X, pady=2)
        ttk.Button(action_frame, text="ℹ️ 关于作者", command=self.show_author, bootstyle="info outline").pack(fill=X, pady=2)

        # ================= 中间主体 =================
        main_pane = ttk.Panedwindow(self, orient=HORIZONTAL)
        main_pane.pack(fill=BOTH, expand=True, padx=15, pady=5)
        
        # 左侧框架
        left_frame = ttk.Labelframe(main_pane, text="1. 教学设计框架 (AI辅助)", padding=10, bootstyle="info")
        main_pane.add(left_frame, weight=1)
        
        # 滚动区域配置
        self.left_canvas = tk.Canvas(left_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.left_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.left_canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")))
        
        left_canvas_window = self.left_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        def configure_canvas(event):
            self.left_canvas.itemconfig(left_canvas_window, width=event.width)
        self.left_canvas.bind('<Configure>', configure_canvas)
        
        self.left_canvas.configure(yscrollcommand=scrollbar.set)
        self.left_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            self.left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        def _on_linux_scroll_up(event):
            self.left_canvas.yview_scroll(-1, "units")
        def _on_linux_scroll_down(event):
            self.left_canvas.yview_scroll(1, "units")
        def _bind_mouse(event):
            self.left_canvas.bind_all("<MouseWheel>", _on_mousewheel)
            self.left_canvas.bind_all("<Button-4>", _on_linux_scroll_up)
            self.left_canvas.bind_all("<Button-5>", _on_linux_scroll_down)
        def _unbind_mouse(event):
            self.left_canvas.unbind_all("<MouseWheel>")
            self.left_canvas.unbind_all("<Button-4>")
            self.left_canvas.unbind_all("<Button-5>")

        left_frame.bind('<Enter>', _bind_mouse)
        left_frame.bind('<Leave>', _unbind_mouse)

        self.fields = {}
        font_bold = (MAIN_FONT_NAME, UI_FONT_SIZE, "bold")
        font_norm = (MAIN_FONT_NAME, UI_FONT_SIZE)

        # 自定义内容区
        custom_frame = ttk.LabelFrame(self.scrollable_frame, text="★ 本课时自定义教学内容 (可选)", padding=5, bootstyle="danger")
        custom_frame.pack(fill=X, pady=(0, 10))
        ttk.Label(custom_frame, text="若填写，AI将严格围绕此内容设计；若留空，则自动规划。", font=(MAIN_FONT_NAME, UI_FONT_SIZE-1), bootstyle="secondary").pack(anchor=W)
        self.fields['custom_content'] = tk.Text(custom_frame, height=3, font=font_norm, bg="#fff0f0", fg="#000")
        self.fields['custom_content'].pack(fill=X, pady=2)
        
        labels = [
            ("📖 章节名称", "chapter", 1),
            ("📋 课程标准 (2017版2025修订)", "standard", 4), 
            ("🎯 素养导向目标", "objectives", 6),
            ("🔥 教学重点", "key_points", 3),
            ("💡 教学难点", "difficulties", 3),
            ("🛠️ 教学方法", "methods", 2),
            ("✍️ 作业设计", "homework", 3),
        ]
        
        for text, key, height in labels:
            lbl = ttk.Label(self.scrollable_frame, text=text, font=font_bold, bootstyle="primary")
            lbl.pack(anchor=W, pady=(5, 0))
            txt = tk.Text(self.scrollable_frame, height=height, font=font_norm)
            txt.pack(fill=X, pady=(0, 5))
            self.fields[key] = txt
        
        ttk.Button(left_frame, text="⚡ 生成当前课时框架", command=self.generate_framework, bootstyle="info").pack(fill=X, pady=5)

        # --- 右侧：过程撰写 ---
        right_frame = ttk.Labelframe(main_pane, text="2. 教学过程与活动 (40分钟)", padding=10, bootstyle="success")
        main_pane.add(right_frame, weight=2)
        
        cmd_frame = ttk.Frame(right_frame)
        cmd_frame.pack(fill=X, pady=5)
        ttk.Label(cmd_frame, text="💬 额外指令:", font=font_bold).pack(side=LEFT)
        self.instruction_entry = ttk.Entry(cmd_frame, bootstyle="success")
        self.instruction_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
        self.instruction_entry.insert(0, "环节清晰，体现学生探究，师生互动具体")

        self.process_text = ScrolledText(right_frame, font=(MAIN_FONT_NAME, 11), padding=10)
        self.process_text.pack(fill=BOTH, expand=True, pady=5)
        
        ctrl_frame = ttk.Frame(right_frame)
        ctrl_frame.pack(fill=X, pady=5)
        
        ttk.Button(ctrl_frame, text="🚀 开始撰写 (Stream)", command=self.start_writing_process, bootstyle="success").pack(side=LEFT, padx=5, fill=X, expand=True)
        ttk.Button(ctrl_frame, text="🛑 停止", command=self.stop_generation, bootstyle="danger").pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="🧹 清空当前页", command=self.clear_current, bootstyle="secondary outline").pack(side=LEFT, padx=5)

        # 底部状态栏
        footer_frame = ttk.Frame(self, bootstyle="light")
        footer_frame.pack(fill=X, side=BOTTOM)
        
        self.status_var = tk.StringVar(value="准备就绪")
        status_lbl = ttk.Label(footer_frame, textvariable=self.status_var, padding=(10, 5), font=(MAIN_FONT_NAME, 9))
        status_lbl.pack(side=LEFT)
        
        author_lbl = ttk.Label(footer_frame, text=self.author_info, padding=(10, 5), font=(MAIN_FONT_NAME, 9), foreground="gray")
        author_lbl.pack(side=RIGHT)

    # --- 逻辑处理 ---

    def show_author(self):
        messagebox.showinfo("关于作者", f"{self.author_info}\n\n版本：3.4.0 (Linux/Win/Mac)\n适用：金塔县中学教案模版标准")

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
        # 优先读取内部存储的 Key
        if not self.api_key:
            messagebox.showwarning("未配置 API Key", "请先点击左上角的【⚙️ 配置 API Key】按钮进行授权。\n配置后将自动保存，下次无需输入。")
            return None
        return self.api_key

    def stop_generation(self):
        if self.is_generating:
            self.stop_flag = True
            self.status_var.set("⛔ 已停止生成")

    def clear_current(self):
        if messagebox.askyesno("确认", f"确定清空【第 {self.active_period} 课时】的所有内容吗？"):
            for key in self.fields:
                self.fields[key].delete("1.0", END)
            self.process_text.delete("1.0", END)
            self.status_var.set(f"第 {self.active_period} 课时已清空")

    def clear_all_data(self):
        if messagebox.askyesno("危险操作", "确定要清空【所有课时】的所有数据吗？\n此操作不可恢复！"):
            self.lesson_data = {} 
            self.active_period = 1
            self.total_periods_var.set(1)
            self.period_combo['values'] = [1]
            self.period_combo.current(0)
            
            for key in self.fields:
                self.fields[key].delete("1.0", END)
            self.process_text.delete("1.0", END)
            self.topic_entry.delete(0, END)
            self.topic_entry.insert(0, "离子反应")
            
            self.status_var.set("⚠️ 所有数据已重置")

    # --- AI 生成逻辑 ---

    def generate_framework(self):
        api_key = self.get_api_key()
        if not api_key: return
        
        topic = self.topic_entry.get()
        current_p = self.active_period
        total_p = self.total_periods_var.get()
        custom_content = self.fields['custom_content'].get("1.0", END).strip()
        
        self.is_generating = True
        self.stop_flag = False
        threading.Thread(target=self._thread_generate_framework, args=(api_key, topic, current_p, total_p, custom_content)).start()

    def _thread_generate_framework(self, api_key, topic, current_p, total_p, custom_content):
        self.status_var.set(f"🤖 正在分析第 {current_p} 课时框架...")
        
        content_instruction = ""
        if custom_content:
            content_instruction = f"【特别指令】用户强制指定本课时(第{current_p}课时)内容为：『{custom_content}』。请只围绕此内容设计。"
        else:
            content_instruction = f"请根据教学逻辑，自动规划第{current_p}课时（共{total_p}课时）的核心内容。"

        prompt = f"""
        任务：为高中化学课题《{topic}》设计第 {current_p} 课时的教案框架。
        {content_instruction}

        【核心要求】
        1. **课程标准**：【必须】引用**《普通高中化学课程标准（2017年版2025年日常修订版）》**中与本课时内容直接相关的具体条目，严禁使用“匹配课标”等模糊词汇。
        2. **素养导向**：严禁使用“三维目标”分类。请用一段通顺的话描述“通过...培养...素养”。
        3. 格式：纯文本，无Markdown。
        4. 返回JSON格式，Key必须保持一致：
        {{
            "chapter": "所属章节",
            "standard": "在此处填写具体的2025日常修订版课标条目内容",
            "objectives": "素养导向目标",
            "key_points": "重点",
            "difficulties": "难点",
            "methods": "方法",
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
                self.status_var.set("✅ 框架生成完毕")
            else:
                self.status_var.set(f"❌ API错误: {response.status_code}")
        except Exception as e:
            self.status_var.set(f"❌ 错误: {str(e)}")
        finally:
            self.is_generating = False

    def _update_framework_ui(self, data):
        for key, value in data.items():
            if key in self.fields and key != 'custom_content':
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
        self.status_var.set(f"✍️ 正在撰写第 {current_p} 课时过程...")
        
        custom_content = context.get('custom_content', '')
        custom_hint = f"本课时核心锁定：{custom_content}。" if custom_content else ""

        prompt = f"""
        任务：撰写高中化学《{topic}》第 {current_p} 课时的“教学过程”。
        
        【输入信息】
        {custom_hint}
        素养目标：{context['objectives']}
        重难点：{context['key_points']}
        
        【严格限制】
        1. 格式：纯文本，严禁Markdown。
        2. 时长：40分钟。
        3. 风格：{plan_type}。{instruction}
        4. 理念：新课标“教-学-评”一体化。
        
        【输出结构】
        环节名称（时间）- 教师活动 - 学生活动 - 设计意图
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
            self.status_var.set("✅ 撰写完成")
        except Exception as e:
            self.status_var.set(f"❌ 错误: {str(e)}")
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
                
                # 标题
                p_title = doc.add_heading(f"第 {i} 课时教案", level=1)
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                table = doc.add_table(rows=8, cols=4)
                table.style = 'Table Grid'
                table.autofit = False
                
                for row in table.rows:
                    row.height = Cm(1.2)

                # R1
                table.cell(0, 0).text = "课题"
                table.cell(0, 1).text = topic
                table.cell(0, 2).text = "时间"
                table.cell(0, 3).text = datetime.now().strftime("%Y-%m-%d")

                # R2
                custom_info = data.get('custom_content', '')
                info_text = f"第 {i} 课时 (共 {total_p} 课时)"
                if custom_info: info_text += f"\n[自定义内容]: {custom_info}"
                
                table.cell(1, 0).text = "课程章节"
                table.cell(1, 1).text = data.get('chapter', '')
                table.cell(1, 2).text = "课时说明"
                table.cell(1, 3).text = info_text

                # R3 课标
                table.cell(2, 0).merge(table.cell(2, 3))
                table.cell(2, 0).text = f"课程标准:\n{data.get('standard', '（未生成，请点击生成框架）')}" 

                # R4 目标
                table.cell(3, 0).merge(table.cell(3, 3))
                table.cell(3, 0).text = f"素养导向目标:\n{data.get('objectives', '')}"

                # R5 重点难点方法
                table.cell(4, 0).merge(table.cell(4, 3))
                p = table.cell(4, 0).paragraphs[0]
                p.add_run("教学重点：").bold = True
                p.add_run(f"{data.get('key_points', '')}\n")
                p.add_run("教学难点：").bold = True
                p.add_run(f"{data.get('difficulties', '')}\n")
                p.add_run("教学方法：").bold = True
                p.add_run(f"{data.get('methods', '')}")

                # R6 过程
                table.cell(5, 0).merge(table.cell(5, 3))
                cell = table.cell(5, 0)
                cell.text = "教学过程与师生活动 (40分钟)"
                cell.add_paragraph(data.get('process', ''))

                # R7 作业
                table.cell(6, 0).merge(table.cell(6, 3))
                table.cell(6, 0).text = f"作业设计:\n{data.get('homework', '')}"

                # R8 反思
                table.cell(7, 0).merge(table.cell(7, 3))
                table.cell(7, 0).text = "课后反思:\n"

            doc.save(filename)
            messagebox.showinfo("导出成功", f"🎉 已成功导出 {total_p} 个课时的教案！")
            
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

if __name__ == "__main__":
    app = LessonPlanWriter()
    app.mainloop()
