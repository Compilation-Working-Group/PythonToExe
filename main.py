import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox
import json
import time
import re

# --- 配置区域 ---
APP_VERSION = "v14.0.1 (Linux Fix)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 动态预设库 ===
PRESET_CONFIGS = {
    "自由定制 / 其它文稿": {
        "topic": "（在此输入任何文稿的主题，如：国旗下讲话、新闻通稿、申报材料）",
        "instruction": "请详细描述您的要求：\n1. 场景：比如‘全校师生大会’或‘教研组会议’。\n2. 语气：比如‘激昂感人’或‘严肃客观’。\n3. 特殊要求：比如‘要引用一句古诗’。",
        "words": "1000",
        "is_custom": True 
    },
    "期刊论文 (标准学术)": {
        "topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "instruction": "要求：\n1. 必须包含：摘要、关键词、引言、正文章节、结语、参考文献。\n2. 重点写‘教学策略’，结合具体的《氯气》实验案例。",
        "words": "4000",
        "is_custom": False
    },
    "教学案例 (叙事风格)": {
        "topic": "《钠与水反应》教学案例分析",
        "instruction": "要求：\n1. 结构：案例背景、情境描述（重点）、案例分析、反思。\n2. 像讲故事一样描述课堂冲突。",
        "words": "2500",
        "is_custom": False
    },
    "教学反思 (个人独白)": {
        "topic": "高三化学二轮复习课后的深刻反思",
        "instruction": "要求：\n1. 第一人称‘我’。\n2. 结构：初衷、效果、不足、改进。",
        "words": "1500",
        "is_custom": False
    },
    "工作计划 (行政公文)": {
        "topic": "2026年春季学期高二化学备课组工作计划",
        "instruction": "要求：\n1. 结构：指导思想、目标、措施、行事历。\n2. 条理清晰，多用数据。",
        "words": "2000",
        "is_custom": False
    },
    "工作总结 (汇报材料)": {
        "topic": "2025年度个人教学工作总结",
        "instruction": "要求：\n1. 结构：概况、成绩、不足、规划。\n2. 数据详实。",
        "words": "3000",
        "is_custom": False
    }
}

class UniversalWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能写作助手 (修复版) - {DEV_NAME}")
        self.geometry("1200x900")
        
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
        
        self.tab_write = self.tabview.add("智能写作工作台")
        self.tab_settings = self.tabview.add("系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 写作工作台 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(5, weight=1) 

        # 1. 文体选择
        ctk.CTkLabel(t, text="文体类型:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        modes = list(PRESET_CONFIGS.keys())
        self.combo_mode = ctk.CTkComboBox(t, values=modes, width=250, command=self.on_mode_change)
        self.combo_mode.set("自由定制 / 其它文稿")
        self.combo_mode.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # 2. 标题
        ctk.CTkLabel(t, text="标题/主题:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 3. 具体指令
        ctk.CTkLabel(t, text="指令要求:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=60, font=("Microsoft YaHei UI", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 4. 字数
        ctk.CTkLabel(t, text="目标字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_words = ctk.CTkEntry(t, width=150)
        self.entry_words.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # 分割线
        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 5. 双面板布局
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=2) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧：大纲区
        ctk.CTkLabel(self.paned_frame, text="第一步：生成大纲 (AI 自动规划结构)", text_color="#1F6AA5", font=("bold", 12)).grid(row=0, column=0, sticky="w", padx=5)
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 13)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_outline_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_outline_frame, text="1. 生成/重置大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5, pady=5)
        ctk.CTkButton(btn_outline_frame, text="清空", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=60).pack(side="right", padx=5)

        # 右侧：正文区
        ctk.CTkLabel(self.paned_frame, text="第二步：按大纲撰写全文", text_color="#2CC985", font=("bold", 12)).grid(row=0, column=1, sticky="w", padx=5)
        self.txt_content = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 14))
        self.txt_content.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        btn_write_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_write_frame.grid(row=2, column=1, sticky="ew")
        
        self.btn_run_write = ctk.CTkButton(btn_write_frame, text="2. 按大纲撰写全文", command=self.run_full_write, fg_color="#2CC985", font=("bold", 14))
        self.btn_run_write.pack(side="left", padx=5, pady=5)
        
        self.btn_stop = ctk.CTkButton(btn_write_frame, text="🔴 紧急停止", command=self.stop_writing, fg_color="#C0392B", width=100)
        self.btn_stop.pack(side="left", padx=5)

        self.btn_clear_all = ctk.CTkButton(btn_write_frame, text="🧹 清空全部", command=self.clear_all, fg_color="gray", width=80)
        self.btn_clear_all.pack(side="right", padx=5)
        
        self.btn_export = ctk.CTkButton(btn_write_frame, text="导出 Word", command=self.save_to_word, width=100)
        self.btn_export.pack(side="right", padx=5)

        # 初始化
        self.on_mode_change("自由定制 / 其它文稿")

    # === Tab 2: 设置 ===
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
        preset = PRESET_CONFIGS.get(choice, PRESET_CONFIGS["自由定制 / 其它文稿"])
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, preset["topic"])
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", preset["instruction"])
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, preset["words"])

    def clear_all(self):
        self.txt_outline.delete("0.0", "end")
        self.txt_content.delete("0.0", "end")
        self.status_label.configure(text="已清空所有内容", text_color="gray")
        self.progressbar.set(0)

    def stop_writing(self):
        self.stop_event.set()
        self.status_label.configure(text="已发送停止指令...", text_color="red")

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    # --- 任务：生成大纲 ---
    def run_gen_outline(self):
        self.stop_event.clear()
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        
        if not topic:
            self.status_label.configure(text="请先输入标题！", text_color="red")
            return

        threading.Thread(target=self.thread_outline, args=(mode, topic, instr), daemon=True).start()

    def thread_outline(self, mode, topic, instr):
        client = self.get_client()
        if not client: return

        self.btn_gen_outline.configure(state="disabled", text="规划中...")
        self.status_label.configure(text=f"正在根据您的指令规划【{mode}】结构...", text_color="#1F6AA5")
        
        if "自由定制" in mode:
            prompt = f"""
            你是一位专业文秘。请根据用户要求，设计一份文稿大纲。
            文章主题：{topic}
            用户具体指令：{instr}
            
            【要求】：
            1. 请分析这应该是什么文体（如演讲稿、新闻、申报材料等）。
            2. 根据文体设计 4-6 个合理的章节标题。
            3. 直接输出标题，不要 Markdown 符号。
            """
        elif "期刊论文" in mode:
            prompt = f"""
            请为高中化学论文《{topic}》设计大纲。
            【强制要求】：必须包含：摘要与关键词、一、引言；二、理论价值；三、教学策略；四、结语；参考文献。
            用户指令：{instr}
            直接输出标题，无Markdown。
            """
        else:
            prompt = f"""
            请为《{topic}》设计一份【{mode}】大纲。
            【强制要求】：
            1. 严禁包含‘摘要’、‘关键词’、‘参考文献’。
            2. 按照该文体的标准结构设计章节（如一、二、三）。
            用户指令：{instr}
            直接输出标题，无Markdown。
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
            
            self.status_label.configure(text="大纲已生成！您可以手动修改左侧大纲，然后点击'撰写全文'。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal", text="1. 生成/重置大纲")

    # --- 任务：撰写全文 ---
    def run_full_write(self):
        self.stop_event.clear()
        
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 2:
            self.status_label.configure(text="大纲为空！请先生成。", text_color="red")
            return
            
        sections = [line.strip() for line in outline_raw.split('\n') if line.strip()]
        if not sections: return

        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        try: total_words = int(self.entry_words.get())
        except: total_words = 3000
        
        threading.Thread(target=self.thread_write, args=(sections, mode, topic, instr, total_words), daemon=True).start()

    def thread_write(self, sections, mode, topic, instr, total_words):
        client = self.get_client()
        if not client: return

        self.btn_run_write.configure(state="disabled", text="写作中...")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)
        
        avg_words = int(total_words / len(sections))
        total_steps = len(sections)

        try:
            for i, section_title in enumerate(sections):
                if self.stop_event.is_set():
                    self.status_label.configure(text="已停止。", text_color="red")
                    break

                self.status_label.configure(text=f"正在撰写 ({i+1}/{total_steps}): {section_title}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)

                # 插入标题标记
                self.txt_content.insert("end", f"\n\n【{section_title}】\n")
                self.txt_content.see("end")

                # 构建 Prompt
                system_prompt = f"""
                你是一位专业文秘。
                当前任务：根据大纲，撰写文章的【{section_title}】部分。
                文体类型：{mode}
                
                【指令】：
                1. 严禁Markdown格式。
                2. 严禁重复输出章节标题（标题已自动插入）。
                3. 必须严格遵守用户指令：{instr}
                4. 如果是演讲稿，语气要口语化；如果是论文，语气要学术。
                """
                
                user_prompt = f"""
                标题：{topic}
                当前章节：{section_title}
                字数要求：约 {avg_words} 字
                请直接写正文。
                """

                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.8
                )

                for chunk in resp:
                    if self.stop_event.is_set(): break
                    if chunk.choices[0].delta.content:
                        c = chunk.choices[0].delta.content
                        self.txt_content.insert("end", c)
                        self.txt_content.see("end")
                
                time.sleep(0.5) 

            if not self.stop_event.is_set():
                self.status_label.configure(text="撰写完成！", text_color="green")
                self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_run_write.configure(state="normal", text="2. 按大纲撰写全文")
            self.btn_gen_outline.configure(state="normal")

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
            run_title = p_title.add_run(self.entry_topic.get())
            run_title.font.size = Pt(16)
            run_title.bold = True
            run_title.font.name = u'黑体'
            run_title._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            
            doc.add_paragraph()

            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue

                if line.startswith("【") and line.endswith("】"):
                    header = line.replace("【", "").replace("】", "")
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(12)
                    run = p.add_run(header)
                    run.bold = True
                    run.font.size = Pt(14)
                    run.font.name = u'黑体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                else:
                    clean_line = re.sub(r'\*\*|##|__|```', '', line)
                    if clean_line.startswith("- ") or clean_line.startswith("* "): clean_line = clean_line[2:]
                    p = doc.add_paragraph(clean_line)
                    p.paragraph_format.first_line_indent = Pt(24)

            doc.save(file_path)
            self.status_label.configure(text=f"已导出: {os.path.basename(file_path)}", text_color="green")

    def load_config(self):
        try:
            with open("config.json", "r") as f: self.api_config = json.load(f)
        except: pass
    def save_config(self):
        self.api_config["api_key"] = self.entry_key.get().strip()
        self.api_config["base_url"] = self.entry_url.get().strip()
        self.api_config["model"] = self.entry_model.get().strip()
        with open("config.json", "w") as f: json.dump(self.api_config, f)

# === 修正的核心 ===
if __name__ == "__main__":
    # 使用正确的类名进行实例化
    app = UniversalWriterApp()
    app.mainloop()
