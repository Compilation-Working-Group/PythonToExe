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
APP_VERSION = "v17.0.0 (Detailed Outline + Smart Expand)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 动态预设库 ===
PRESET_CONFIGS = {
    "期刊论文 (标准学术)": {
        "topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "instruction": "要求：\n1. 语气严谨学术，多用数据。\n2. 策略部分必须结合具体的《氯气》实验案例。\n3. 摘要要连贯。",
        "words": "4500",
        "structure_hint": "包含：摘要、关键词、一、引言；二、理论价值；三、教学策略；四、成效反思；参考文献。"
    },
    "教学反思 (深度实战)": {
        "topic": "高三化学二轮复习课后的深刻反思",
        "instruction": "要求：\n1. 第一人称‘我’。\n2. 拒绝套话，分析真实问题。\n3. 结构：现象->原因->措施。",
        "words": "2000",
        "structure_hint": "包含：一、背景；二、现象；三、原因；四、改进。"
    },
    "教学案例 (叙事风格)": {
        "topic": "《钠与水反应》教学案例分析",
        "instruction": "要求：\n1. 像写故事一样描述课堂冲突。\n2. 还原现场细节。",
        "words": "2500",
        "structure_hint": "包含：一、背景；二、片段描述；三、分析；四、反思。"
    },
    "工作计划 (务实版)": {
        "topic": "2026年春季学期高二化学备课组工作计划",
        "instruction": "要求：\n1. 条理清晰，多用数据。\n2. 具体到月份。",
        "words": "2000",
        "structure_hint": "包含：一、指导思想；二、目标；三、措施；四、行事历。"
    },
    "工作总结 (数据版)": {
        "topic": "2025年度个人教学工作总结",
        "instruction": "要求：\n1. 用数据说话。\n2. 举具体例子。",
        "words": "3000",
        "structure_hint": "包含：一、概况；二、成绩；三、不足；四、规划。"
    },
    "自由定制 / 其它文稿": {
        "topic": "（在此输入文稿主题）",
        "instruction": "请详细描述要求。",
        "words": "1500",
        "structure_hint": "请自动规划合理的结构。"
    }
}

class InteractiveWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能写作助手 (详细大纲版) - {DEV_NAME}")
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

        ctk.CTkLabel(t, text="选择文体:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        modes = list(PRESET_CONFIGS.keys())
        self.combo_mode = ctk.CTkComboBox(t, values=modes, width=250, command=self.on_mode_change)
        self.combo_mode.set("期刊论文 (标准学术)")
        self.combo_mode.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        ctk.CTkLabel(t, text="标题/主题:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(t, text="指令要求:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=60, font=("Microsoft YaHei UI", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(t, text="目标字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_words = ctk.CTkEntry(t, width=150)
        self.entry_words.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 双面板布局
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=2) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧：大纲
        ctk.CTkLabel(self.paned_frame, text="第一步：生成详细大纲", text_color="#1F6AA5", font=("bold", 12)).grid(row=0, column=0, sticky="w", padx=5)
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 13)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_outline_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_outline_frame, text="1. 生成详细大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5, pady=5)
        ctk.CTkButton(btn_outline_frame, text="清空", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=60).pack(side="right", padx=5)

        # 右侧：正文
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

        self.on_mode_change("期刊论文 (标准学术)")

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
        preset = PRESET_CONFIGS.get(choice, PRESET_CONFIGS["期刊论文 (标准学术)"])
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, preset["topic"])
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", preset["instruction"])
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, preset["words"])

    def clear_all(self):
        self.txt_outline.delete("0.0", "end")
        self.txt_content.delete("0.0", "end")
        self.status_label.configure(text="已清空", text_color="gray")
        self.progressbar.set(0)

    def stop_writing(self):
        self.stop_event.set()
        self.status_label.configure(text="已停止", text_color="red")

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    # --- 任务：生成详细大纲 (核心修复) ---
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
        self.status_label.configure(text=f"正在规划【{mode}】的详细结构...", text_color="#1F6AA5")
        
        # 获取结构建议
        preset = PRESET_CONFIGS.get(mode, {})
        hint = preset.get("structure_hint", "")

        # 核心提示词：强制要求二级标题
        prompt = f"""
        任务：为《{topic}》写一份【{mode}】的**详细大纲**。
        用户的指令：{instr}
        结构参考：{hint}
        
        【强制要求】：
        1. 必须包含一级标题（如“一、引言”）和 **二级标题**（如“（一）研究背景”）。
        2. 每一章下面至少要有 2-3 个小标题，让大纲看起来非常丰满。
        3. 如果是期刊论文，必须包含：摘要、关键词、参考文献。
        4. 直接输出大纲内容，不要 Markdown，不要多余解释。
        """
        
        try:
            resp = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True,
                temperature=0.8
            )
            
            self.txt_outline.delete("0.0", "end")
            for chunk in resp:
                if self.stop_event.is_set(): break
                if chunk.choices[0].delta.content:
                    c = chunk.choices[0].delta.content
                    self.txt_outline.insert("end", c)
                    self.txt_outline.see("end")
            
            self.status_label.configure(text="详细大纲已生成！请确认满意后点击'撰写全文'。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal", text="1. 生成详细大纲")

    # --- 任务：撰写全文 (逐条目撰写) ---
    def run_full_write(self):
        self.stop_event.clear()
        
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 5:
            self.status_label.configure(text="大纲为空！", text_color="red")
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
        
        # 智能分配字数：条目越多，单条字数越少，但总数达标
        avg_words = int(total_words / len(sections))
        if avg_words < 200: avg_words = 200 # 保证每个小节至少写点东西
        
        total_steps = len(sections)

        try:
            for i, section_title in enumerate(sections):
                if self.stop_event.is_set(): break

                self.status_label.configure(text=f"正在撰写 ({i+1}/{total_steps}): {section_title}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)

                # 插入标题 (区分一级和二级标题的格式)
                # 简单判断：如果是一、二、三，则空两行；如果是（一）、（二），则空一行
                if any(x in section_title for x in ['一、', '二、', '三、', '四、', '五、', '六、', '参考文献']):
                     self.txt_content.insert("end", f"\n\n【{section_title}】\n")
                else:
                     self.txt_content.insert("end", f"\n【{section_title}】\n")
                     
                self.txt_content.see("end")

                # 特殊处理：摘要
                is_abstract = "摘要" in section_title
                prompt_extra = "请撰写连贯的短文，严禁列条目。" if is_abstract else "内容要务实，结合具体案例。"

                system_prompt = f"""
                你是一位专业的高中化学教师文秘。
                当前任务：撰写【{section_title}】的内容。
                文体类型：{mode}
                
                【指令】：
                1. 严禁复述标题。
                2. 严禁 Markdown。
                3. {prompt_extra}
                4. {instr}
                """
                
                user_prompt = f"""
                标题：{topic}
                当前小节：{section_title}
                字数：约 {avg_words} 字
                请直接写正文。
                """

                # 使用非流式请求以便清洗
                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.75
                )
                
                raw = resp.choices[0].message.content
                
                # 清洗算法：去除开头的标题重复
                clean = raw.strip()
                pattern = r'^\s*(\#+|【|\*\*|)?\s*' + re.escape(section_title) + r'\s*(】|\*\*|)?\s*\n?'
                clean = re.sub(pattern, '', clean, flags=re.IGNORECASE).strip()
                
                self.txt_content.insert("end", clean)
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
                    
                    # 判断一级还是二级标题
                    if any(x in header for x in ['一、', '二、', '三、', '四、', '五、', '六、', '参考文献', '摘要']):
                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(12)
                        run = p.add_run(header)
                        run.bold = True
                        run.font.size = Pt(14) # 一级标题大一点
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    else:
                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(6)
                        run = p.add_run(header)
                        run.bold = True
                        run.font.size = Pt(12) # 二级标题小一点
                        run.font.name = u'楷体' # 二级标题用楷体区分
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
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

if __name__ == "__main__":
    app = InteractiveWriterApp()
    app.mainloop()
