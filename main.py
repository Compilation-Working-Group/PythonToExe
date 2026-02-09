import customtkinter as ctk
import threading
from openai import OpenAI
import os
import sys
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox, simpledialog
import json
import time
import re
import difflib # 新增：用于模糊匹配去重

# --- 依赖库检测 ---
try:
    import pypdf
except ImportError:
    pypdf = None
try:
    import openpyxl
except ImportError:
    openpyxl = None

# --- Linux 显示修正 ---
if sys.platform.startswith('linux'):
    try:
        import tkinter
        root = tkinter.Tk()
        root.destroy()
    except:
        if os.environ.get('DISPLAY','') == '':
            os.environ.__setitem__('DISPLAY', ':0')

# --- 配置区域 ---
APP_VERSION = "v32.0.0 (Smart Clean + Excel Filter)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 文体风格定义 ===
STYLE_GUIDE = {
    "期刊论文": {
        "desc": "学术严谨，理实结合，适合发表。",
        "default_topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "default_words": "3000",
        "default_instruction": "要求：\n1. 结合具体的化学教学案例。\n2. 数据详实，逻辑严密。\n3. 适合《化学教育》或《中化参》风格。",
        "writing_prompt": "语气学术、客观、务实。严禁堆砌空洞理论，必须用具体的化学知识点和教学片段来支撑观点。",
    },
    "教学反思": {
        "desc": "第一人称，深度剖析，真诚走心。",
        "default_topic": "高三化学二轮复习课后的深刻反思",
        "default_words": "2000",
        "default_instruction": "要求：\n1. 必须使用第一人称‘我’。\n2. 重点复盘课堂上的‘遗憾点’和‘生成性问题’。\n3. 剖析原因要深刻。",
        "writing_prompt": "使用第一人称。文风要诚恳、犀利。多描写课堂上的真实细节（如学生的错题、冷场的瞬间）。",
    },
    "教学案例": {
        "desc": "叙事风格，还原现场，生动具体。",
        "default_topic": "《钠与水反应》教学案例分析",
        "default_words": "2500",
        "default_instruction": "要求：\n1. 采用‘教育叙事’手法。\n2. 还原师生对话，描写实验现象。\n3. 突出‘意外’与‘机智化解’。",
        "writing_prompt": "采用叙事风格。大量使用对话描写、动作描写。还原真实的课堂冲突和教学灵感。",
    },
    "工作计划": {
        "desc": "行政公文，条理清晰，数据导向。",
        "default_topic": "2026年春季学期高二化学备课组工作计划",
        "default_words": "2000",
        "default_instruction": "要求：\n1. 语言简练，干脆利落。\n2. 包含具体的行事历。\n3. 目标要量化。",
        "writing_prompt": "行政公文风格。多用‘一要...二要...’句式。内容必须具体可执行，包含时间节点。",
    },
    "工作总结": {
        "desc": "汇报风格，亮点突出，分析透彻。",
        "default_topic": "2025年度个人教学工作总结",
        "default_words": "3000",
        "default_instruction": "要求：\n1. 用数据说话。\n2. 既要展示成绩，也要诚恳分析不足。\n3. 结构严谨。",
        "writing_prompt": "汇报风格。多用数据对比。对成绩要总结经验，对不足要分析原因并提出对策。",
    },
    "自由定制": {
        "desc": "完全根据指令生成。",
        "default_topic": "（在此输入题目）",
        "default_words": "1000",
        "default_instruction": "请详细描述您的要求...",
        "writing_prompt": "严格遵循用户的特殊指令，风格不限。",
    }
}

class MasterWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能写作系统 - {APP_VERSION}")
        self.geometry("1300x900")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()
        self.stop_event = threading.Event()
        self.reference_content = "" 

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_write = self.tabview.add("写作工作台")
        self.tab_settings = self.tabview.add("系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(6, weight=1) 

        # 顶部控制区
        ctrl_frame = ctk.CTkFrame(t, fg_color="transparent")
        ctrl_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(ctrl_frame, text="文体类型:", font=("bold", 14)).pack(side="left", padx=5)
        self.combo_mode = ctk.CTkComboBox(ctrl_frame, values=list(STYLE_GUIDE.keys()), width=180, command=self.on_mode_change)
        self.combo_mode.set("期刊论文")
        self.combo_mode.pack(side="left", padx=5)
        
        ctk.CTkLabel(ctrl_frame, text="目标字数:", font=("bold", 14)).pack(side="left", padx=(20, 5))
        self.entry_words = ctk.CTkEntry(ctrl_frame, width=100)
        self.entry_words.insert(0, "3000")
        self.entry_words.pack(side="left", padx=5)

        ctk.CTkLabel(t, text="文章标题:", font=("bold", 12)).grid(row=1, column=0, padx=10, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=600)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # --- 参考文档区 ---
        ctk.CTkLabel(t, text="参考资料:", font=("bold", 12)).grid(row=2, column=0, padx=10, sticky="e")
        ref_frame = ctk.CTkFrame(t, fg_color="transparent")
        ref_frame.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        self.btn_upload = ctk.CTkButton(ref_frame, text="📂 上传/筛选资料", command=self.load_reference_file, width=140, fg_color="#E67E22")
        self.btn_upload.pack(side="left", padx=5)
        
        self.btn_clear_ref = ctk.CTkButton(ref_frame, text="❌ 清除", command=self.clear_reference_file, width=60, fg_color="#C0392B")
        self.btn_clear_ref.pack(side="left", padx=5)
        
        self.lbl_ref_status = ctk.CTkLabel(ref_frame, text="未上传 (AI将基于通用知识写作)", text_color="gray")
        self.lbl_ref_status.pack(side="left", padx=10)

        ctk.CTkLabel(t, text="具体指令:", font=("bold", 12)).grid(row=3, column=0, padx=10, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=50, font=("Arial", 12))
        self.txt_instructions.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 双面板
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=2) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧大纲
        outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        outline_frame.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(outline_frame, text="Step 1: 智能大纲 (AI根据题目生成)", text_color="#1F6AA5", font=("bold", 13)).pack(side="left")
        
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 12)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_o_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_o_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_o_frame, text="生成/重置大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5)
        ctk.CTkButton(btn_o_frame, text="清空", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=60).pack(side="right", padx=5)

        # 右侧正文
        content_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        content_frame.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(content_frame, text="Step 2: 正文撰写 (自动清洗重复标题)", text_color="#2CC985", font=("bold", 13)).pack(side="left")
        self.status_label = ctk.CTkLabel(content_frame, text="就绪", text_color="gray")
        self.status_label.pack(side="right")

        self.txt_content = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 14))
        self.txt_content.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        btn_w_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_w_frame.grid(row=2, column=1, sticky="ew")
        self.btn_run_write = ctk.CTkButton(btn_w_frame, text="开始撰写全文", command=self.run_full_write, fg_color="#2CC985", font=("bold", 14))
        self.btn_run_write.pack(side="left", padx=5)
        self.btn_stop = ctk.CTkButton(btn_w_frame, text="🔴 停止", command=self.stop_writing, fg_color="#C0392B", width=80)
        self.btn_stop.pack(side="left", padx=5)
        self.btn_clear_all = ctk.CTkButton(btn_w_frame, text="🧹 清空", command=self.clear_all, fg_color="gray", width=80)
        self.btn_clear_all.pack(side="right", padx=5)
        self.btn_export = ctk.CTkButton(btn_w_frame, text="导出 Word", command=self.save_to_word, width=120)
        self.btn_export.pack(side="right", padx=5)

        self.progressbar = ctk.CTkProgressBar(t, mode="determinate", height=2)
        self.progressbar.grid(row=7, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self.progressbar.set(0)

        self.on_mode_change("期刊论文")

    def setup_settings_tab(self):
        t = self.tab_settings
        ctk.CTkLabel(t, text="API Key:").pack(pady=(20, 5))
        self.entry_key = ctk.CTkEntry(t, width=400, show="*")
        self.entry_key.insert(0, self.api_config.get("api_key", ""))
        self.entry_key.pack(pady=5)
        ctk.CTkLabel(t, text="Base URL:").pack(pady=5)
        self.entry_url = ctk.CTkEntry(t, width=400)
        self.entry_url.insert(0, self.api_config.get("base_url", ""))
        self.entry_url.pack(pady=5)
        ctk.CTkLabel(t, text="Model:").pack(pady=5)
        self.entry_model = ctk.CTkEntry(t, width=400)
        self.entry_model.insert(0, self.api_config.get("model", ""))
        self.entry_model.pack(pady=5)
        ctk.CTkButton(t, text="保存配置", command=self.save_config).pack(pady=20)

    # --- 核心升级：Excel 筛选与多格式读取 ---
    def load_reference_file(self):
        filetypes = [
            ("All Supported", "*.docx *.pdf *.xlsx *.txt *.md *.csv *.py *.json"),
            ("Excel Data", "*.xlsx *.xls"),
            ("Word", "*.docx"),
            ("PDF", "*.pdf"),
            ("Text", "*.txt *.md")
        ]
        filepath = filedialog.askopenfilename(filetypes=filetypes)
        if not filepath: return
        
        filename = os.path.basename(filepath)
        ext = os.path.splitext(filepath)[1].lower()
        content = ""
        filter_key = ""
        
        try:
            # 1. Excel (.xlsx) - 支持班级/关键词筛选
            if ext in [".xlsx", ".xls"]:
                if openpyxl is None: raise ImportError("缺少 openpyxl")
                
                # 弹出对话框，询问筛选关键词
                dialog = ctk.CTkInputDialog(text="【Excel数据预筛选】\n请输入要提取的关键词（如 '高二1班'）。\n若需全部分析，请留空直接点击OK。", title="数据筛选")
                filter_key = dialog.get_input()
                if filter_key is None: filter_key = "" # Cancel = empty
                filter_key = filter_key.strip()

                wb = openpyxl.load_workbook(filepath, data_only=True)
                for sheet in wb:
                    sheet_data = []
                    rows = list(sheet.iter_rows(values_only=True))
                    if not rows: continue
                    
                    # 永远保留表头
                    header = rows[0]
                    # 转CSV格式，更利于AI理解
                    sheet_data.append(",".join([str(c) if c else "" for c in header]))
                    
                    match_count = 0
                    for row in rows[1:]:
                        row_str = ",".join([str(c) if c else "" for c in row])
                        # 核心筛选逻辑：如果没有关键词，或者关键词在行内，则保留
                        if not filter_key or (filter_key in row_str):
                            sheet_data.append(row_str)
                            match_count += 1
                    
                    if len(sheet_data) > 1: # 有数据（不仅是表头）
                        content += f"\n--- Sheet: {sheet.title} (匹配到 {match_count} 行) ---\n"
                        content += "\n".join(sheet_data) + "\n"

            # 2. Word (.docx)
            elif ext == ".docx":
                doc = Document(filepath)
                content = "\n".join([p.text for p in doc.paragraphs])
            
            # 3. PDF (.pdf)
            elif ext == ".pdf":
                if pypdf is None: raise ImportError("缺少 pypdf")
                reader = pypdf.PdfReader(filepath)
                for page in reader.pages: content += page.extract_text() + "\n"
            
            # 4. 纯文本
            else:
                try:
                    with open(filepath, "r", encoding="utf-8") as f: content = f.read()
                except UnicodeDecodeError:
                    with open(filepath, "r", encoding="gbk") as f: content = f.read()

            content = content.strip()
            if not content: raise ValueError("文件内容为空或筛选后无数据")
                
            self.reference_content = content[:20000] # 放宽到2万字
            if len(content) > 20000: self.reference_content += "\n...(内容过长，已截取)"
            
            status_msg = f"已挂载: {filename}"
            if filter_key: status_msg += f" (筛选: {filter_key})"
            
            self.lbl_ref_status.configure(text=status_msg, text_color="green")
            self.btn_clear_ref.configure(state="normal")
            messagebox.showinfo("成功", f"文件解析成功！\nAI将基于此数据进行分析。")
            
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取: {str(e)}")

    def clear_reference_file(self):
        self.reference_content = ""
        self.lbl_ref_status.configure(text="未上传 (AI将基于通用知识写作)", text_color="gray")
        self.btn_clear_ref.configure(state="disabled")
        messagebox.showinfo("已清除", "参考资料已清空。")

    def on_mode_change(self, choice):
        config = STYLE_GUIDE.get(choice, STYLE_GUIDE["自由定制"])
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, config.get("default_topic", ""))
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", config.get("default_instruction", ""))
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, config.get("default_words", "3000"))
        
        self.txt_outline.delete("0.0", "end")
        self.txt_outline.insert("0.0", f"（已切换至【{choice}】模式，请点击“生成/重置大纲”...）")

    def stop_writing(self):
        self.stop_event.set()
        self.status_label.configure(text="已停止", text_color="red")

    def clear_all(self):
        self.txt_outline.delete("0.0", "end")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)
        self.status_label.configure(text="已清空")

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_gen_outline(self):
        self.stop_event.clear()
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        if not topic:
            self.status_label.configure(text="请输入标题！", text_color="red")
            return
        threading.Thread(target=self.thread_outline, args=(mode, topic, instr), daemon=True).start()

    def thread_outline(self, mode, topic, instr):
        client = self.get_client()
        if not client: return
        self.btn_gen_outline.configure(state="disabled")
        self.status_label.configure(text="正在分析题目并构建大纲...", text_color="#1F6AA5")
        
        style_cfg = STYLE_GUIDE.get(mode, STYLE_GUIDE["自由定制"])
        ref_hint = ""
        if self.reference_content:
            ref_hint = f"【资料背景】：用户提供了数据/资料（{len(self.reference_content)}字），请务必在构建大纲时，安排章节来分析这些数据。"

        prompt = f"""
        任务：为《{topic}》写一份【{mode}】的详细大纲。
        【参考风格】：{style_cfg['desc']}
        【用户指令】：{instr}
        {ref_hint}
        
        【要求】：
        1. 拒绝千篇一律。请根据题目内涵定制结构。
        2. 必须包含一级标题（如一、二、三）和二级标题（如（一）（二））。
        3. 不要包含Markdown符号。
        4. 直接输出大纲。
        """
        try:
            resp = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True
            )
            self.txt_outline.delete("0.0", "end")
            for chunk in resp:
                if self.stop_event.is_set(): break
                if chunk.choices[0].delta.content:
                    c = chunk.choices[0].delta.content
                    self.txt_outline.insert("end", c)
                    self.txt_outline.see("end")
            self.status_label.configure(text="大纲已生成，请检查并修改。", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal")

    def run_full_write(self):
        self.stop_event.clear()
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 5:
            self.status_label.configure(text="请先生成大纲", text_color="red")
            return
            
        lines = [l.strip() for l in outline_raw.split('\n') if l.strip()]
        if len(lines) > 0:
            first_line = lines[0]
            topic = self.entry_topic.get().strip()
            # 智能滤除第一行如果它像标题
            if len(topic) > 2 and topic[:4] in first_line:
                lines = lines[1:]

        tasks = []
        current_task = []
        for line in lines:
            is_header = False
            # 强化标题识别逻辑
            if re.match(r'^[一二三四五六七八九十]+、', line): is_header = True
            if re.match(r'^第[一二三四五六七八九十]+部分', line): is_header = True
            if "摘要" in line or "参考文献" in line: is_header = True
            
            if is_header:
                if current_task: tasks.append(current_task)
                current_task = [line]
            else:
                current_task.append(line)
        if current_task: tasks.append(current_task)

        if not tasks:
            self.status_label.configure(text="大纲格式无法识别", text_color="red")
            return

        topic = self.entry_topic.get()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        try: total_words = int(self.entry_words.get())
        except: total_words = 3000
        
        threading.Thread(target=self.thread_write, args=(tasks, mode, topic, instr, total_words), daemon=True).start()

    def thread_write(self, tasks, mode, topic, instr, total_words):
        client = self.get_client()
        if not client: return

        self.btn_run_write.configure(state="disabled")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)
        
        style_cfg = STYLE_GUIDE.get(mode, STYLE_GUIDE["自由定制"])
        core_tasks = [t for t in tasks if "摘要" not in t[0] and "参考文献" not in t[0]]
        core_count = len(core_tasks) if len(core_tasks) > 0 else 1
        
        reserved_words = 0
        if any("摘要" in t[0] for t in tasks): reserved_words += 300
        available_words = total_words - reserved_words
        if available_words < 500: available_words = 500
        avg_core_words = available_words // core_count

        last_paragraph = "（文章刚开始，暂无上文）"
        
        ref_prompt_block = ""
        if self.reference_content:
            ref_prompt_block = f"""
            【重要参考资料】：
            以下是用户提供的真实数据或资料。请务必：
            1. 分析这些数据。
            2. 在正文中引用数据作为论据。
            3. 保持数据真实性，不要编造。
            
            {self.reference_content}
            ------------------------------------------------
            """

        def get_core_text(t):
            # 提取汉字，用于比对
            return re.sub(r'[^\u4e00-\u9fa50-9]', '', t)

        try:
            for i, task_lines in enumerate(tasks):
                if self.stop_event.is_set(): break
                
                header = task_lines[0]
                sub_points = "\n".join(task_lines[1:])
                current_limit = avg_core_words
                prompt_suffix = ""
                
                if "摘要" in header: 
                    current_limit = 300
                    prompt_suffix = "【特殊要求】：必须在摘要下方另起一行，列出3-5个【关键词】。"
                elif "参考文献" in header: 
                    current_limit = 0
                elif any(x in header for x in ["一、", "引言", "结语"]): 
                    current_limit = int(avg_core_words * 0.6)
                else:
                    current_limit = int(avg_core_words * 1.2)
                
                self.status_label.configure(text=f"撰写: {header}...", text_color="#1F6AA5")
                self.progressbar.set(i / len(tasks))

                self.txt_content.insert("end", f"\n\n【{header}】\n")
                self.txt_content.see("end")

                sys_prompt = f"""
                你是一位资深教育专家。
                文体：{mode}
                风格：{style_cfg['writing_prompt']}
                {ref_prompt_block}
                【写作铁律】：
                1. 严禁复述章节标题！(系统已自动插入，请直接开始写正文)。
                2. 严禁Markdown（不要**加粗**，不要##标题）。
                3. 基于提供的资料进行真实分析。
                4. {prompt_suffix}
                """
                
                user_prompt = f"题目：{topic}\n当前章节：{header}\n包含要点：{sub_points}\n上下文：...{last_paragraph[-150:]}\n字数：约 {current_limit} 字。\n请直接输出正文内容。"

                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[{"role":"system","content":sys_prompt}, {"role":"user","content":user_prompt}],
                    temperature=0.7,
                    stream=True
                )
                
                current_section_text = ""
                header_core = get_core_text(header) # 提取系统标题的核心字

                for chunk in resp:
                    if self.stop_event.is_set(): break
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        temp_text = current_section_text + content
                        
                        # --- 核心修复：外科手术式去重 ---
                        if "摘要" in header:
                            if len(temp_text) < 10 and ("摘" in temp_text or "要" in temp_text):
                                current_section_text += content
                                continue 
                            clean_chunk = re.sub(r'^【?摘要】?[:：]?\s*', '', content)
                            self.txt_content.insert("end", clean_chunk)
                        else:
                            # 暂存前 100 字符，用于检测是否包含标题
                            if len(temp_text) < 100:
                                current_section_text += content
                            else:
                                # 缓冲区满了，或者已经过了开头。
                                # 检查缓冲区是否包含标题
                                if current_section_text != "SAFE":
                                    # 检查相似度
                                    similarity = difflib.SequenceMatcher(None, header_core, get_core_text(current_section_text)).ratio()
                                    if similarity > 0.6 or header_core in get_core_text(current_section_text):
                                        # 发现标题重复！尝试切除第一句
                                        parts = current_section_text.split('\n', 1)
                                        if len(parts) > 1:
                                            self.txt_content.insert("end", parts[1])
                                        else:
                                            # 只有一行且是标题，丢弃，只保留新来的content
                                            pass 
                                    else:
                                        # 没有重复，安全上屏
                                        self.txt_content.insert("end", current_section_text)
                                    
                                    current_section_text = "SAFE" # 标记为安全
                                
                                # 直接输出新内容
                                self.txt_content.insert("end", content)
                        
                        self.txt_content.see("end")
                        if len(temp_text) > 50: last_paragraph = temp_text
                
                # 循环结束，检查缓冲区是否还有未上屏的内容
                if current_section_text != "SAFE" and len(current_section_text) > 0:
                     # 同样的检查逻辑
                    similarity = difflib.SequenceMatcher(None, header_core, get_core_text(current_section_text)).ratio()
                    if similarity > 0.6 or header_core in get_core_text(current_section_text):
                        parts = current_section_text.split('\n', 1)
                        if len(parts) > 1: self.txt_content.insert("end", parts[1])
                    else:
                        self.txt_content.insert("end", current_section_text)

            if not self.stop_event.is_set():
                self.status_label.configure(text="撰写完成！", text_color="green")
                self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"API错误: {str(e)}", text_color="red")
        finally:
            self.btn_run_write.configure(state="normal")

    # --- 核心升级：导出 Word (智能清洗 Markdown) ---
    def save_to_word(self):
        content = self.txt_content.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_t = p_title.add_run(self.entry_topic.get())
            run_t.font.name = u'黑体'
            run_t._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run_t.font.size = Pt(18)
            run_t.bold = True
            
            p_auth = doc.add_paragraph()
            p_auth.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_a = p_auth.add_run(f"{DEV_NAME}\n({DEV_ORG})")
            run_a.font.name = u'楷体'
            run_a._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
            run_a.font.size = Pt(12)
            
            doc.add_paragraph() 

            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue

                # 识别系统插入的章节标记
                if line.startswith("【") and line.endswith("】"):
                    header = line.replace("【", "").replace("】", "")
                    
                    if "摘要" in header or "关键词" in header:
                        p = doc.add_paragraph()
                        run = p.add_run(header)
                        run.bold = True
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    elif re.match(r'^[一二三四五六七八九十]+、', header):
                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(12)
                        run = p.add_run(header)
                        run.bold = True
                        run.font.size = Pt(14)
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    else:
                        p = doc.add_paragraph(header)
                        p.runs[0].bold = True
                else:
                    # 智能清洗 Markdown 痕迹
                    clean_line = line
                    
                    # 1. 加粗处理：**加粗** -> 去掉星号，应用加粗样式 (简化版：直接去星号，避免复杂解析)
                    clean_line = re.sub(r'\*\*(.*?)\*\*', r'\1', clean_line) 
                    # 2. 标题处理：### 标题 -> 去掉井号
                    clean_line = re.sub(r'#{1,6}\s?', '', clean_line)
                    # 3. 列表处理：- 列表 -> 列表
                    if clean_line.startswith("- "): clean_line = clean_line[2:]
                    
                    p = doc.add_paragraph(clean_line)
                    p.paragraph_format.first_line_indent = Pt(24) 
                    p.paragraph_format.line_spacing = 1.25

            doc.save(file_path)
            self.status_label.configure(text=f"已导出纯净版: {os.path.basename(file_path)}", text_color="green")

    def load_config(self):
        try:
            with open("config.json", "r") as f: self.api_config = json.load(f)
        except: pass
    def save_config(self):
        self.api_config["api_key"] = self.entry_key.get().strip()
        self.api_config["base_url"] = self.entry_url.get().strip()
        self.api_config["model"] = self.entry_model.get().strip()
        with open("config.json", "w") as f: json.dump(self.api_config, f)

if __name__ == "__main__":
    app = MasterWriterApp()
    app.mainloop()
