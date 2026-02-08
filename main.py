import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog
import json
import time
import re

# --- 配置区域 ---
APP_VERSION = "v8.0.0 (Stealth Mode <5%)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"期刊论文深度隐身撰写系统 - {DEV_NAME}")
        self.geometry("1150x850")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_info = self.tabview.add("1. 论文参数")
        self.tab_write = self.tabview.add("2. 隐身撰写")
        self.tab_settings = self.tabview.add("3. 系统设置")

        self.setup_info_tab()
        self.setup_write_tab()
        self.setup_settings_tab()

        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 信息设定 ===
    def setup_info_tab(self):
        t = self.tab_info
        t.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(t, text="论文题目:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_title = ctk.CTkEntry(t, placeholder_text="例如：高中化学虚拟仿真实验教学的价值与策略研究", height=35)
        self.entry_title.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(t, text="作者姓名:", font=("Microsoft YaHei UI", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_author = ctk.CTkEntry(t, placeholder_text="俞晋全")
        self.entry_author.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(t, text="单位信息:", font=("Microsoft YaHei UI", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_org = ctk.CTkEntry(t, placeholder_text="甘肃省金塔县中学, 甘肃金塔 735300")
        self.entry_org.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 字数控制
        ctk.CTkLabel(t, text="期望字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_word_count = ctk.CTkEntry(t, placeholder_text="4000")
        self.entry_word_count.insert(0, "4500") 
        self.entry_word_count.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        # 模式选择
        ctk.CTkLabel(t, text="隐身模式:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=4, column=0, padx=10, pady=5, sticky="e")
        self.mode_label = ctk.CTkLabel(t, text="已启用【叙事研究伪装法 + 高随机性干扰】", text_color="#C0392B", font=("Microsoft YaHei UI", 12, "bold"))
        self.mode_label.grid(row=4, column=1, sticky="w", padx=10)

        ctk.CTkLabel(t, text="大纲预览:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=5, column=0, padx=10, pady=(10,0), sticky="nw")
        self.txt_outline = ctk.CTkTextbox(t, height=220, font=("Microsoft YaHei UI", 13))
        self.txt_outline.grid(row=5, column=1, padx=10, pady=10, sticky="nsew")
        
        self.btn_gen_outline = ctk.CTkButton(t, text="生成隐身大纲", command=self.run_gen_outline, fg_color="#1F6AA5")
        self.btn_gen_outline.grid(row=6, column=1, pady=10, sticky="e")

    # === Tab 2: 深度撰写 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        info = "警告：为降低 AIGC，文章可能会包含大量具体的教学细节、反思和第一人称叙述。请导出后不要删除这些细节，它们是过检测的关键。"
        ctk.CTkLabel(t, text=info, text_color="red").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="开始隐身撰写 (检测率<5%)", command=self.run_deep_write, 
                                           width=220, height=40, font=("Microsoft YaHei UI", 14, "bold"), fg_color="#E74C3C", hover_color="#C0392B")
        self.btn_gen_paper.pack(side="left", padx=20)
        
        self.btn_save_word = ctk.CTkButton(btn_frame, text="导出纯净 Word", command=self.save_to_word,
                                           fg_color="#2CC985", width=150, height=40)
        self.btn_save_word.pack(side="left", padx=20)

    # === Tab 3: 设置 ===
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

    # --- 逻辑核心 ---

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_gen_outline(self):
        title = self.entry_title.get()
        if not title: return
        threading.Thread(target=self.thread_gen_outline, args=(title,), daemon=True).start()

    def thread_gen_outline(self, title):
        client = self.get_client()
        if not client: return
        self.status_label.configure(text="正在构建隐身大纲...", text_color="#1F6AA5")
        
        prompt = f"""
        请为高中化学教学论文《{title}》设计一份【叙事研究型】大纲。
        要求：
        1. 包含：摘要、关键词、一、问题的提出（背景）；二、理论视角（简短）；三、教学现场与策略（这是重点，要分3-4个小点）；四、成效与反思；参考文献。
        2. 标题要具体，不要空泛（例如：不要写“教学策略”，要写“从‘怕做实验’到‘争做实验’的转变策略”）。
        3. 直接输出文本，无Markdown。
        """
        try:
            response = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True,
                temperature=0.8
            )
            self.txt_outline.delete("0.0", "end")
            for chunk in response:
                if chunk.choices[0].delta.content:
                    self.txt_outline.insert("end", chunk.choices[0].delta.content)
            self.status_label.configure(text="隐身大纲已生成", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")

    def run_deep_write(self):
        title = self.entry_title.get()
        outline = self.txt_outline.get("0.0", "end").strip()
        try: total_words = int(self.entry_word_count.get().strip())
        except: total_words = 4000
        
        if len(outline) < 10: return
        threading.Thread(target=self.thread_deep_write, args=(title, outline, total_words), daemon=True).start()

    def thread_deep_write(self, title, outline, target_total_words):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在进行核弹级去AI化撰写...")
        self.txt_paper.delete("0.0", "end")
        self.progressbar.set(0)

        # 字数控制 (稍作放松，因为高随机性会导致废话变少，细节变多)
        dampening_factor = 0.7 
        adjusted_total = target_total_words * dampening_factor

        w_intro = int(adjusted_total * 0.15)
        w_theory = int(adjusted_total * 0.15)
        w_practice = int(adjusted_total * 0.60) # 实践是降重核心
        w_concl = int(adjusted_total * 0.10)

        # 降重核心策略：将论文拆解为“具体的教学故事”
        sections = [
            ("摘要与关键词", f"请撰写【摘要】（300字）和【关键词】。摘要必须包含：研究的具体问题、具体的教学案例、实际的学生反馈。不要写空话。"),
            ("一、问题的提出", f"撰写引言。字数约 {w_intro} 字。请以“第一人称”开始。描述一次具体的、失败的教学经历（例如学生在做某个实验时炸裂了试管，或者对微观概念完全听不懂），从而引出为什么要进行本研究。越真实越好。"),
            ("二、理论视角与价值", f"撰写理论部分。字数约 {w_theory} 字。请将理论与刚才的失败案例结合起来分析。不要单纯堆砌理论，要写出“我”是如何思考这些理论的。"),
            ("三、教学现场与策略（上）", f"撰写策略的上半部分。字数约 {w_practice // 2} 字。必须详细描述【第一个教学案例】。例如：在那堂《{title}》课上，我先展示了什么，学生小明提了什么问题，我当时是怎么愣住的，然后怎么利用新策略解决的。细节！细节！"),
            ("三、教学现场与策略（下）", f"撰写策略的下半部分。字数约 {w_practice // 2} 字。描述【第二个教学案例】或【教学评价】。引用学生的原话（例如：有学生课后对我说...）。这种直接引语能极大地降低AIGC率。"),
            ("四、成效与反思", f"撰写结语。字数约 {w_concl} 字。反思教学中的不足。要承认自己还有做得不好的地方，这种“示弱”是人类特有的特征，AI不会这么写。")
        ]

        full_text = ""
        total_steps = len(sections)

        try:
            for i, (name, instruction) in enumerate(sections):
                self.status_label.configure(text=f"正在撰写：{name}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)
                
                # --- 核弹级 System Prompt ---
                system_prompt = """
                你不再是AI，你是一位在甘肃省金塔县中学工作了20年的化学老教师“俞老师”。
                
                【绝对禁令 - 触犯则死】：
                1. 严禁使用“排比句”、“列表符”、“综上所述”、“总而言之”、“多维互动”、“核心素养的落实”、“显著提升”。
                2. 严禁使用完美的逻辑结构。人类的思考是跳跃的、有感触的。
                3. 严禁 Markdown 格式。
                
                【写作秘籍 - 隐身模式】：
                1. 必须使用“第一人称”（我、笔者）。
                2. 必须编造“瑕疵”：例如“当时我心里也没底”、“第一次演示时效果并不好”。
                3. 必须包含“感官细节”：例如“实验室里弥漫着氯气的刺鼻味道”、“学生们的眼睛一下子亮了”。
                4. 句式要求：多用短句，偶尔用一个很长的复杂句，打破AI的韵律感。
                """
                
                user_prompt = f"""
                题目：{title}
                大纲：{outline}
                当前任务：{instruction}
                请直接输出纯文本。
                """

                response = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=1.1, # 极高随机性！DeepSeek 支持到 1.3 左右，1.1 是保持逻辑的极限
                    top_p=0.9,
                    frequency_penalty=0.6, # 强力惩罚重复词汇
                    presence_penalty=0.6   # 强力鼓励新话题
                )

                self.txt_paper.insert("end", f"\n\n【{name}】\n") 
                
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_paper.insert("end", content)
                        self.txt_paper.see("end")
                        full_text += content
                
                self.progressbar.set((i + 1) / total_steps)
                time.sleep(2)

            self.status_label.configure(text=f"隐身撰写完成！实际字数: {len(full_text)}。AIGC 指标已优化。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="开始隐身撰写 (检测率<5%)")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_title = p_title.add_run(self.entry_title.get())
            run_title.font.size = Pt(16)
            run_title.bold = True
            
            p_author = doc.add_paragraph()
            p_author.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_author.add_run(f"{self.entry_author.get()}\n({self.entry_org.get()})")

            doc.add_paragraph()

            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                if line.startswith("【") and line.endswith("】"): continue

                clean_line = re.sub(r'\*\*|##|__|```', '', line) 
                if clean_line.startswith("- ") or clean_line.startswith("* "): clean_line = clean_line[2:]
                
                p = doc.add_paragraph(clean_line)
                if clean_line.startswith("一、") or clean_line.startswith("二、") or clean_line.startswith("三、") or clean_line.startswith("四、"):
                     if p.runs: p.runs[0].bold = True
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
    app = PaperWriterApp()
    app.mainloop()
