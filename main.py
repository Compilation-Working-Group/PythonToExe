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

# --- 配置区域 ---
APP_VERSION = "v3.0.0 (Custom Style Ed.)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"期刊论文定制撰写系统 - {DEV_NAME}专用版")
        self.geometry("1100x850")
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
        
        self.tab_info = self.tabview.add("1. 论文信息设定")
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

    # === Tab 1: 信息设定 (模仿范文格式) ===
    def setup_info_tab(self):
        t = self.tab_info
        t.grid_columnconfigure(1, weight=1)

        # 题目
        ctk.CTkLabel(t, text="论文题目:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_title = ctk.CTkEntry(t, placeholder_text="例如：高中化学虚拟仿真实验教学的价值与策略研究", height=35)
        self.entry_title.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # 作者信息 (用于生成头部格式)
        ctk.CTkLabel(t, text="作者姓名:", font=("Microsoft YaHei UI", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_author = ctk.CTkEntry(t, placeholder_text="俞晋全")
        self.entry_author.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(t, text="单位及邮编:", font=("Microsoft YaHei UI", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_org = ctk.CTkEntry(t, placeholder_text="甘肃省金塔县中学, 甘肃金塔 735399")
        self.entry_org.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 核心论点控制
        ctk.CTkLabel(t, text="核心关键词:", font=("Microsoft YaHei UI", 12)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_keywords = ctk.CTkEntry(t, placeholder_text="高中化学; 仿真实验; 教学设计 (用分号隔开)")
        self.entry_keywords.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 大纲预览区
        ctk.CTkLabel(t, text="自动生成的大纲 (可修改):", font=("Microsoft YaHei UI", 12, "bold")).grid(row=4, column=0, padx=10, pady=(10,0), sticky="nw")
        self.txt_outline = ctk.CTkTextbox(t, height=300, font=("Microsoft YaHei UI", 13))
        self.txt_outline.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        
        self.btn_gen_outline = ctk.CTkButton(t, text="生成标准结构大纲", command=self.run_gen_outline, fg_color="#1F6AA5")
        self.btn_gen_outline.grid(row=5, column=1, pady=10, sticky="e")

    # === Tab 2: 深度撰写 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        info = "提示：系统将严格按照您上传的范文格式撰写。为达到 3000-6000 字，写作过程较长，请耐心等待。"
        ctk.CTkLabel(t, text=info, text_color="gray").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="开始深度撰写 (Pro)", command=self.run_deep_write, 
                                           width=200, height=40, font=("Microsoft YaHei UI", 14, "bold"))
        self.btn_gen_paper.pack(side="left", padx=20)
        
        self.btn_save_word = ctk.CTkButton(btn_frame, text="导出 Word (范文格式)", command=self.save_to_word,
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
        
        # 强制模仿范文结构的 Prompt
        prompt = f"""
        请为高中化学教学论文《{title}》设计大纲。
        【严格格式要求】：
        1. 必须包含：摘要、关键词、一、引言（背景与意义）；二、核心价值（理论支撑）；三、具体策略（实践方法）；四、结语；参考文献。
        2. 正文标题必须使用汉字数字格式：
           一、...
           （一）...
           （二）...
           二、...
           （一）...
        3. “策略”部分至少要有4个子标题，以确保字数充足。
        4. 不需要输出“作者”等信息，只输出正文大纲结构。
        """
        try:
            response = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True
            )
            self.txt_outline.delete("0.0", "end")
            for chunk in response:
                if chunk.choices[0].delta.content:
                    self.txt_outline.insert("end", chunk.choices[0].delta.content)
            self.status_label.configure(text="大纲已生成，请核对结构", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")

    def run_deep_write(self):
        title = self.entry_title.get()
        outline = self.txt_outline.get("0.0", "end").strip()
        if len(outline) < 10: return
        threading.Thread(target=self.thread_deep_write, args=(title, outline), daemon=True).start()

    def thread_deep_write(self, title, outline):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在深度撰写...")
        self.txt_paper.delete("0.0", "end")
        self.progressbar.set(0)

        # 定义分块写作任务，确保字数和深度
        # 模仿范文：摘要 -> 引言 -> 价值(理论) -> 策略(核心) -> 结语
        sections = [
            ("摘要与关键词", "请撰写【摘要】（250-300字，概括研究背景、价值、策略）和【关键词】（3-5个）。风格务实，不要废话。"),
            ("一、引言与背景", "撰写论文的开头部分（不带标题，直接写内容）。分析当前高中化学教学的痛点（如：传统实验危险、微观概念难理解、时空受限），引出本文的研究主题。字数800字。"),
            ("二、核心价值/教学意义", "撰写论文的“价值/意义”部分。请分点阐述（如：降低风险、突破限制、强化微观认知）。一定要结合高中化学具体知识点（如：氯气、浓硫酸、有机合成）。字数1000字。"),
            ("三、教学策略/实践路径（上）", "撰写“教学策略”的前两点。必须结合具体的教学案例（如：‘钠与水反应’、‘原电池’）。详细描述教师如何做、学生如何做、系统如何反馈。这是降重的关键，细节要多！字数1200字。"),
            ("三、教学策略/实践路径（下）", "撰写“教学策略”的后两点。侧重于‘分层任务设计’或‘评价反馈机制’。引用教育学理论（如最近发展区、UBD理论）。字数1200字。"),
            ("四、结语与参考文献", "撰写【结语】（总结全文，展望未来）和【参考文献】（列出5-8条相关文献，格式规范）。")
        ]

        full_text = ""
        total = len(sections)

        try:
            for i, (name, instruction) in enumerate(sections):
                self.status_label.configure(text=f"正在撰写：{name}...", text_color="#1F6AA5")
                self.progressbar.set(i / total)
                
                # 提示词工程：去AI化，增加具体案例
                system_prompt = """
                你是一位拥有20年经验的高中化学高级教师。
                你的写作风格：
                1. 严谨、务实，多用“笔者认为”、“在教学实践中”。
                2. 拒绝AI常用的空洞连接词（如“综上所述、总而言之”少用）。
                3. 必须大量引用高中化学具体教材内容（如必修一、必修二的具体实验）。
                4. 字数必须充足，逻辑必须连贯。
                """
                
                user_prompt = f"""
                论文题目：{title}
                参考大纲：{outline}
                
                【当前任务】：{instruction}
                
                请直接输出正文内容，不要重复标题。
                """

                response = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.75 # 提高随机性以降重
                )

                self.txt_paper.insert("end", f"\n\n【{name}】\n") 
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_paper.insert("end", content)
                        self.txt_paper.see("end")
                        full_text += content
                
                self.progressbar.set((i + 1) / total)
                time.sleep(2)

            self.status_label.configure(text="论文撰写完成！", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="开始深度撰写 (Pro)")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            
            # --- 设置中文字体 ---
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 1. 标题 (黑体，居中，二号)
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_title = p_title.add_run(self.entry_title.get())
            run_title.font.name = u'黑体'
            run_title.font.size = Pt(18)
            run_title._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            
            # 2. 作者信息 (楷体，居中，小四)
            p_author = doc.add_paragraph()
            p_author.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            info_str = f"{self.entry_author.get()}\n({self.entry_org.get()})"
            run_author = p_author.add_run(info_str)
            run_author.font.name = u'楷体'
            run_author.font.size = Pt(12)
            run_author._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')

            # 3. 正文处理
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line or "【" in line: continue # 跳过系统标记
                
                p = doc.add_paragraph()
                # 识别标题
                if line.startswith("摘要") or line.startswith("关键词"):
                    run = p.add_run(line)
                    run.bold = True
                elif line.startswith("一、") or line.startswith("二、") or line.startswith("三、") or line.startswith("四、"):
                    # 一级标题
                    run = p.add_run(line)
                    run.font.name = u'黑体'
                    run.font.size = Pt(14)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                elif line.startswith("（一）") or line.startswith("（二）"):
                    # 二级标题
                    run = p.add_run(line)
                    run.font.name = u'楷体'
                    run.bold = True
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
                else:
                    # 正文
                    p.add_run(line)
                    p.paragraph_format.first_line_indent = Pt(24) # 首行缩进

            doc.save(file_path)
            self.status_label.configure(text=f"已导出范文格式: {os.path.basename(file_path)}", text_color="green")

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
