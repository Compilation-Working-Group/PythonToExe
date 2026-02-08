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
APP_VERSION = "v9.0.0 (Structure Locked + Length Boost)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"期刊论文撰写系统 (结构锁死+字数增强版) - {DEV_NAME}")
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
        self.tab_write = self.tabview.add("2. 深度撰写")
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
        self.entry_org = ctk.CTkEntry(t, placeholder_text="甘肃省金塔县中学, 甘肃金塔 735399")
        self.entry_org.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 字数控制
        ctk.CTkLabel(t, text="目标字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_word_count = ctk.CTkEntry(t, placeholder_text="4500")
        self.entry_word_count.insert(0, "4500") 
        self.entry_word_count.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        hint = ctk.CTkLabel(t, text="提示：系统将把论文拆解为 8-9 个微任务，强制堆叠字数并保证结构完整。", text_color="#1F6AA5", font=("Arial", 10))
        hint.grid(row=4, column=1, sticky="w", padx=10)

        ctk.CTkLabel(t, text="大纲预览:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=5, column=0, padx=10, pady=(10,0), sticky="nw")
        self.txt_outline = ctk.CTkTextbox(t, height=220, font=("Microsoft YaHei UI", 13))
        self.txt_outline.grid(row=5, column=1, padx=10, pady=10, sticky="nsew")
        
        self.btn_gen_outline = ctk.CTkButton(t, text="生成标准结构大纲", command=self.run_gen_outline, fg_color="#1F6AA5")
        self.btn_gen_outline.grid(row=6, column=1, pady=10, sticky="e")

    # === Tab 2: 深度撰写 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        info = "提示：导出为纯文本Word。请耐心等待，为了字数达标，写作过程会比之前慢一倍。"
        ctk.CTkLabel(t, text=info, text_color="red").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="开始深度撰写 (结构+字数)", command=self.run_deep_write, 
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
        self.status_label.configure(text="正在构建范文结构...", text_color="#1F6AA5")
        
        # 强制结构 Prompt
        prompt = f"""
        请为高中化学教学论文《{title}》设计大纲。
        【强制结构要求】：
        1. 摘要与关键词
        2. 一、高中化学...的教学价值（理论支撑）
        3. 二、高中化学...的教学策略（实践路径）
        4. 结语
        5. 参考文献
        注意：请确保“策略”部分至少有4个子标题，以便后续扩展字数。
        直接输出文本，无Markdown。
        """
        try:
            response = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True,
                temperature=0.7
            )
            self.txt_outline.delete("0.0", "end")
            for chunk in response:
                if chunk.choices[0].delta.content:
                    self.txt_outline.insert("end", chunk.choices[0].delta.content)
            self.status_label.configure(text="大纲已生成", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")

    def run_deep_write(self):
        title = self.entry_title.get()
        outline = self.txt_outline.get("0.0", "end").strip()
        try: total_words = int(self.entry_word_count.get().strip())
        except: total_words = 4500
        
        if len(outline) < 10: return
        threading.Thread(target=self.thread_deep_write, args=(title, outline, total_words), daemon=True).start()

    def thread_deep_write(self, title, outline, target_total_words):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在执行多段式写作...")
        self.txt_paper.delete("0.0", "end")
        self.progressbar.set(0)

        # === 核心算法：切香肠战术 (Micro-Chunking) ===
        # 将论文强制拆分为 9 个部分，无论 AI 想怎么偷懒，都必须写满这 9 段。
        # 假设目标 4500 字，每段只需承担 500 字，AI 很容易完成，且总字数必达标。
        
        chunk_target = target_total_words // 8  # 平均每段字数

        sections = [
            # 头部
            ("摘要与关键词", f"请撰写【摘要】（300字）和【关键词】。摘要需包含目的、方法、结果、结论。"),
            
            # 第一部分：引言
            ("一、问题的提出（背景）", f"撰写论文引言。字数约 {chunk_target} 字。从具体的教学痛点切入（如：传统实验的危险性、微观概念的抽象性）。请使用第一人称叙述具体的教学困境。"),
            
            # 第二部分：价值（拆分为两段写，保证字数）
            ("二、教学价值（理论层面）", f"撰写“教学价值”的前半部分。侧重于理论支撑（如核心素养、深度学习）。字数约 {chunk_target} 字。请引用建构主义或UBD理论，但要用自己的话解释。"),
            ("二、教学价值（现实层面）", f"撰写“教学价值”的后半部分。侧重于解决实际问题（如突破时空限制、降低风险）。字数约 {chunk_target} 字。结合高中化学具体知识点（如氯气、浓硫酸）。"),
            
            # 第三部分：策略（核心重点，拆分为三段写！）
            ("三、教学策略 - 课前准备与设计", f"撰写“策略”的第一点。字数约 {chunk_target} 字。描述如何进行教学设计、资源整合。请虚构一个具体的课题（如《原电池》）作为案例贯穿始终。"),
            ("三、教学策略 - 课堂实施与互动", f"撰写“策略”的第二点。字数约 {chunk_target} 字。描述课堂上的具体师生互动。请描写具体的实验现象、学生的惊呼、老师的引导。细节决定成败（降重关键）！"),
            ("三、教学策略 - 课后评价与反馈", f"撰写“策略”的第三点。字数约 {chunk_target} 字。描述如何利用数据反馈调整教学。请写出具体的“学生反馈原话”。"),
            
            # 第四部分
            ("四、成效与反思", f"撰写成效分析和教学反思。字数约 {chunk_target} 字。承认教学中存在的不足（如设备成本高、部分学生不适应），这能极大降低AIGC率。"),
            
            # 尾部
            ("参考文献", f"列出参考文献（8-10条）。格式规范。")
        ]

        full_text = ""
        total_steps = len(sections)

        try:
            for i, (name, instruction) in enumerate(sections):
                self.status_label.configure(text=f"正在撰写 ({i+1}/{total_steps})：{name}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)
                
                # System Prompt: 兼顾降重与结构
                system_prompt = """
                你是一位资深高中化学教师。
                1. 【结构要求】：严格按照用户指定的章节写作，不要自己发明章节。
                2. 【降重秘籍】：
                   - 多讲故事，少讲道理。
                   - 多用“我”、“笔者”、“学生小明”等具体主语。
                   - 严禁使用“综上所述、总而言之、多维互动”等AI词汇。
                3. 【格式】：纯文本输出，无Markdown。
                """
                
                user_prompt = f"""
                题目：{title}
                当前章节标题：{name}
                写作任务：{instruction}
                
                请直接输出该章节的正文内容（不要重复标题）。
                """

                response = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.85, # 稍微降低一点点随机性，保证结构不崩，但依然保持“人味”
                    top_p=0.9
                )

                # 插入章节标记（方便阅读）
                self.txt_paper.insert("end", f"\n\n【{name}】\n") 
                
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_paper.insert("end", content)
                        self.txt_paper.see("end")
                        full_text += content
                
                self.progressbar.set((i + 1) / total_steps)
                time.sleep(1.5) # 稍微休息防封

            self.status_label.configure(text=f"撰写完成！总字数: {len(full_text)}。结构完整，细节丰富。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="开始深度撰写")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 头部
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
                if line.startswith("【") and line.endswith("】"): 
                    # 可以在这里选择保留或删除章节标记，为了排版方便，建议删除或保留作为参考
                    # 这里我们将其作为加粗的段落保留，方便用户定位
                    continue

                # 清洗
                clean_line = re.sub(r'\*\*|##|__|```', '', line) 
                if clean_line.startswith("- ") or clean_line.startswith("* "): clean_line = clean_line[2:]
                
                p = doc.add_paragraph(clean_line)
                
                # 简单加粗标题
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
