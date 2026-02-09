import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox
import json
import time
import re

# --- 配置区域 ---
APP_VERSION = "v25.0.0 (Strict Word Count Control)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 文体风格定义 ===
STYLE_GUIDE = {
    "期刊论文": {
        "desc": "参照《虚拟仿真》、《热重分析》等范文。学术严谨，理实结合。",
        "default_topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "default_words": "3000",
        "default_instruction": "要求：\n1. 语气严谨学术，多用数据支撑。\n2. 策略部分必须结合具体的《氯气》或《氧化还原》实验案例。\n3. 摘要要写成连贯的短文，不要列条目。",
        "outline_prompt": "请设计一份标准的教育期刊论文大纲。必须包含：摘要、关键词、一、问题的提出；二、核心概念/理论；三、教学策略/模型建构（核心）；四、成效与反思；参考文献。",
        "writing_prompt": "【核心风格】：一线名师的经验总结。严禁写成硕博论文！\n1. 语言简练，拒绝宏大理论堆砌。\n2. 多用短句，多用“实词”。\n3. 策略部分必须“干货满满”，直接讲怎么上课、怎么做实验。\n4. 案例要具体到化学方程式、实验现象、学生原话。",
        "is_paper": True
    },
    "教学反思": {
        "desc": "参照《二轮复习反思》。第一人称，深度剖析。",
        "default_topic": "高三化学二轮复习课后的深刻反思",
        "default_words": "2000",
        "default_instruction": "要求：\n1. 使用第一人称‘我’。\n2. 拒绝套话，重点描写课堂上真实的遗憾、突发状况和学生的真实反应。\n3. 剖析要深刻，多找自身原因。",
        "outline_prompt": "请设计一份深度教学反思大纲。建议结构：一、教学初衷；二、课堂实录与问题；三、原因深度剖析；四、改进措施。",
        "writing_prompt": "使用第一人称‘我’。拒绝套话，重点描写课堂上真实的遗憾、突发状况。剖析要深刻。",
        "is_paper": False
    },
    "教学案例": {
        "desc": "叙事风格，还原课堂现场。",
        "default_topic": "《钠与水反应》教学案例分析",
        "default_words": "2500",
        "default_instruction": "要求：\n1. 采用‘叙事研究’风格。\n2. 像写故事一样描述课堂冲突、师生对话和实验现象。\n3. 重点突出“意外生成”的处理。",
        "outline_prompt": "请设计一份教学案例大纲。建议结构：一、案例背景；二、情境描述（片段）；三、案例分析；四、教学启示。",
        "writing_prompt": "采用‘叙事研究’风格。像写故事一样描述课堂冲突、师生对话和实验现象。",
        "is_paper": False
    },
    "工作计划": {
        "desc": "行政公文风格，条理清晰。",
        "default_topic": "2026年春季学期高二化学备课组工作计划",
        "default_words": "2000",
        "default_instruction": "要求：\n1. 语言简练，行政公文风。\n2. 措施要具体，多用数据（如周课时、目标分）。\n3. 包含具体的行事历。",
        "outline_prompt": "请设计一份工作计划大纲。包含：指导思想、工作目标、主要措施、行事历。",
        "writing_prompt": "语言简练，多用‘一要...二要...’的句式。措施要具体，多用数据。",
        "is_paper": False
    },
    "工作总结": {
        "desc": "汇报风格，数据详实。",
        "default_topic": "2025年度个人教学工作总结",
        "default_words": "3000",
        "default_instruction": "要求：\n1. 用数据说话（平均分、获奖数）。\n2. 既要展示亮点，也要诚恳分析不足。\n3. 结构严谨。",
        "outline_prompt": "请设计一份工作总结大纲。包含：工作概况、主要成绩、存在不足、未来展望。",
        "writing_prompt": "用数据说话（平均分、获奖数）。既要展示亮点，也要诚恳分析不足。",
        "is_paper": False
    },
    "自由定制": {
        "desc": "根据指令自动生成。",
        "default_topic": "（在此输入自定义文稿主题）",
        "default_words": "1000",
        "default_instruction": "请详细描述您的写作要求...",
        "outline_prompt": "请根据用户的具体指令设计最合理的大纲结构。",
        "writing_prompt": "严格遵循用户的特殊要求。",
        "is_paper": False
    }
}

class MasterWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"俞晋全名师工作室全能写作系统 - {APP_VERSION}")
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

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_write = self.tabview.add("写作工作台")
        self.tab_settings = self.tabview.add("系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(5, weight=1) 

        # --- 顶部控制区 ---
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

        ctk.CTkLabel(t, text="具体指令:", font=("bold", 12)).grid(row=2, column=0, padx=10, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=50, font=("Arial", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # --- 核心双面板区 ---
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=2) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧：大纲
        outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        outline_frame.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(outline_frame, text="Step 1: 生成并修改大纲", text_color="#1F6AA5", font=("bold", 13)).pack(side="left")
        
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 12)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_o_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_o_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_o_frame, text="生成/重置大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5)
        ctk.CTkButton(btn_o_frame, text="清空", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=60).pack(side="right", padx=5)

        # 右侧：正文
        content_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        content_frame.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(content_frame, text="Step 2: 撰写预览 (实时流式)", text_color="#2CC985", font=("bold", 13)).pack(side="left")
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
        self.progressbar.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
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

    # --- 交互逻辑 ---

    def on_mode_change(self, choice):
        config = STYLE_GUIDE.get(choice, STYLE_GUIDE["自由定制"])
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, config.get("default_topic", ""))
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", config.get("default_instruction", ""))
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, config.get("default_words", "3000"))
        self.txt_outline.delete("0.0", "end")
        self.txt_outline.insert("0.0", f"（请点击“生成大纲”按钮，AI将为您规划【{choice}】的结构...）")

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

    # --- 生成大纲 ---
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
        self.status_label.configure(text="正在规划结构...", text_color="#1F6AA5")
        
        style_cfg = STYLE_GUIDE.get(mode, STYLE_GUIDE["自由定制"])
        
        prompt = f"""
        任务：为《{topic}》写一份【{mode}】的详细大纲。
        【参考风格】：{style_cfg['desc']}
        【结构建议】：{style_cfg['outline_prompt']}
        【用户指令】：{instr}
        【要求】：
        1. 必须包含一级标题（如一、二、三）和二级标题（如（一）（二））。
        2. 不要包含Markdown符号。
        3. 直接输出大纲，不要废话。
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
            self.status_label.configure(text="大纲已生成，请手动修改。", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal")

    # --- 撰写全文 (字数精准控制 + 实时流式) ---
    def run_full_write(self):
        self.stop_event.clear()
        
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 5:
            self.status_label.configure(text="请先生成或输入大纲", text_color="red")
            return
            
        lines = [l.strip() for l in outline_raw.split('\n') if l.strip()]
        
        # 智能滤除标题行
        if len(lines) > 0:
            first_line = lines[0]
            topic = self.entry_topic.get().strip()
            if len(topic) > 2 and topic[:4] in first_line:
                lines = lines[1:]

        tasks = []
        current_task = []
        for line in lines:
            is_header = False
            if re.match(r'^[一二三四五六七八九十]+、', line): is_header = True
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
        
        # 动态计算每个核心章节应分配的字数
        core_tasks = [t for t in tasks if "摘要" not in t[0] and "参考文献" not in t[0]]
        core_count = len(core_tasks) if len(core_tasks) > 0 else 1
        
        reserved_words = 0
        if any("摘要" in t[0] for t in tasks): reserved_words += 300
