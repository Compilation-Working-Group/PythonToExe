import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from tkinter import filedialog
import json
import time

# --- 配置区域 ---
APP_VERSION = "v2.0.0 (Pro Journal Ed.)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"AI 期刊论文深度撰写系统 - {DEV_NAME}")
        self.geometry("1000x800")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 默认配置
        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()

        # --- 主选项卡 ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_outline = self.tabview.add("1. 拟定大纲")
        self.tab_write = self.tabview.add("2. 深度撰写")
        self.tab_settings = self.tabview.add("3. 系统设置")

        self.setup_outline_tab()
        self.setup_write_tab()
        self.setup_settings_tab()

        # 状态栏
        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        # 进度条 (新功能)
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 大纲生成 ===
    def setup_outline_tab(self):
        t = self.tab_outline
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(t, text="请输入论文题目:", font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10,0))
        
        self.entry_title = ctk.CTkEntry(t, placeholder_text="例如: 基于“宏微结合”素养的高中化学教学实践研究", height=40, font=("Microsoft YaHei UI", 12))
        self.entry_title.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        self.txt_outline = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14), height=300)
        self.txt_outline.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        
        # 预设提示词
        default_prompt = "（点击下方按钮生成大纲，或在此处手动粘贴大纲...）"
        self.txt_outline.insert("0.0", default_prompt)

        self.btn_gen_outline = ctk.CTkButton(t, text="生成标准期刊论文大纲", command=self.run_gen_outline, height=40, font=("Microsoft YaHei UI", 14, "bold"))
        self.btn_gen_outline.grid(row=3, column=0, pady=10)

    # === Tab 2: 深度撰写 (核心升级) ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        info_text = "提示：系统将采用“分步接力”模式写作，以确保字数达到 3000-6000 字并降低查重率。\n整个过程可能需要 3-5 分钟，请耐心等待。"
        ctk.CTkLabel(t, text=info_text, font=("Microsoft YaHei UI", 12), text_color="gray", justify="left").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        # 两个功能按钮
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="开始深度撰写 (分章节)", command=self.run_deep_write, 
                                           fg_color="#1F6AA5", font=("Microsoft YaHei UI", 14, "bold"), width=200, height=40)
        self.btn_gen_paper.pack(side="left", padx=20)
        
        self.btn_save_word = ctk.CTkButton(btn_frame, text="导出 Word 文档", command=self.save_to_word,
                                           fg_color="#2CC985", hover_color="#229966", width=150, height=40)
        self.btn_save_word.pack(side="left", padx=20)

    # === Tab 3: 设置 ===
    def setup_settings_tab(self):
        t = self.tab_settings
        ctk.CTkLabel(t, text="API Key:").pack(pady=(20, 5))
        self.entry_key = ctk.CTkEntry(t, width=400, show="*")
        self.entry_key.insert(0, self.api_config.get("api_key", ""))
        self.entry_key.pack(pady=5)

        ctk.CTkLabel(t, text="Base URL:").pack(pady=(10, 5))
        self.entry_url = ctk.CTkEntry(t, width=400)
        self.entry_url.insert(0, self.api_config.get("base_url", ""))
        self.entry_url.pack(pady=5)
        
        ctk.CTkLabel(t, text="Model Name:").pack(pady=(10, 5))
        self.entry_model = ctk.CTkEntry(t, width=400)
        self.entry_model.insert(0, self.api_config.get("model", ""))
        self.entry_model.pack(pady=5)

        ctk.CTkButton(t, text="保存配置", command=self.save_config).pack(pady=20)

    # --- 逻辑功能 ---

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请先配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_gen_outline(self):
        title = self.entry_title.get()
        if not title:
            self.status_label.configure(text="请输入题目！", text_color="red")
            return
        threading.Thread(target=self.thread_gen_outline, args=(title,), daemon=True).start()

    def thread_gen_outline(self, title):
        client = self.get_client()
        if not client: return
        self.btn_gen_outline.configure(state="disabled", text="正在规划架构...")
        self.status_label.configure(text="正在生成大纲...", text_color="#1F6AA5")
        
        # 专门针对“降低查重”和“期刊结构”的 Prompt
        prompt = f"""
        请为学术论文题目《{title}》设计一份详细的大纲。
        【要求】：
        1. 结构必须包含：中文摘要（300字左右）、关键词（3-5个）、一、引言；二、理论基础；三、教学实践/研究设计；四、结果与分析；五、结论与反思；六、参考文献。
        2. 正文部分请细化到二级标题（如 2.1, 2.2）。
        3. 这是一个高中化学名师的论文，请体现“素养为本”、“真实情境”、“教学案例”等要素。
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
            self.status_label.configure(text="大纲生成完成，请修改后点击下一步", text_color="green")
            self.tabview.set("1. 拟定大纲")
        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal", text="生成标准期刊论文大纲")

    def run_deep_write(self):
        title = self.entry_title.get()
        outline = self.txt_outline.get("0.0", "end").strip()
        if len(outline) < 20:
            self.status_label.configure(text="请先生成大纲！", text_color="red")
            return
        threading.Thread(target=self.thread_deep_write, args=(title, outline), daemon=True).start()

    def thread_deep_write(self, title, outline):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在深度撰写中...")
        self.txt_paper.delete("0.0", "end")
        self.progressbar.set(0)

        # 定义分步写作任务
        steps = [
            ("摘要与关键词", "请撰写论文的【摘要】（包含研究目的、方法、结果、结论，约300-400字）和3-5个【关键词】。语言要精练学术。"),
            ("第一部分：引言与背景", "请根据大纲，撰写【引言】部分。重点阐述研究背景、现状分析（痛点）、研究意义。引用一些教育理论（如建构主义、核心素养）。要求：不要说空话，要结合高中化学教学实际，字数800字左右。"),
            ("第二部分：理论与设计", "请根据大纲，撰写【理论基础】或【研究设计】部分。如果是教学论文，请详细描述教学策略、核心概念界定。要求：逻辑严密，多用专业术语，字数800字左右。"),
            ("第三部分：实践与案例（核心）", "请根据大纲，撰写【教学实践/案例分析】部分。这是降重的关键。请虚构或引用一个具体的化学教学片段（如《钠的性质》或《原电池》），包含具体的师生对话、实验步骤、问题链设置。细节越丰富越好，就像在写教案实录。字数1000字左右。"),
            ("第四部分：结果与反思", "请根据大纲，撰写【结果分析】和【教学反思】。描述学生的变化（成绩、兴趣），并提出不足之处。语气要诚恳、客观。字数600字左右。"),
            ("参考文献", "请列出10-15条参考文献。格式符合GB/T 7714标准。包含期刊、专著、课标文件等。")
        ]

        full_text = ""
        total_steps = len(steps)

        try:
            for i, (section_name, instruction) in enumerate(steps):
                self.status_label.configure(text=f"正在撰写：{section_name} ({i+1}/{total_steps})...", text_color="#1F6AA5")
                self.progressbar.set((i) / total_steps)
                
                # 插入章节标题作为提示
                self.txt_paper.insert("end", f"\n\n======= {section_name} =======\n\n")
                self.txt_paper.see("end")

                # 构建“降重”提示词 (Prompt Engineering)
                system_prompt = "你是一位拥有20年教龄的高中化学特级教师。你的写作风格务实、深刻，拒绝AI味，拒绝空洞的套话。"
                user_prompt = f"""
                论文题目：{title}
                完整大纲：{outline}
                
                【当前任务】：{instruction}
                
                【写作要求】：
                1. 模拟人类写作习惯，长短句结合。
                2. 增加“混乱度”（Perplexity），避免常见的AI常用词（如“首先、其次、最后”太频繁）。
                3. 内容必须充实，结合化学学科特点（如宏观辨识与微观探析）。
                """

                response = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.7 # 稍微提高随机性以降低查重
                )

                chunk_text = ""
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        chunk_text += content
                        self.txt_paper.insert("end", content)
                        self.txt_paper.see("end")
                
                full_text += f"\n\n{chunk_text}"
                
                # 进度更新
                self.progressbar.set((i + 1) / total_steps)
                time.sleep(1) # 稍作停顿防止 API 速率限制

            self.status_label.configure(text=f"论文撰写完成！总字数约 {len(full_text)} 字。", text_color="green")
            self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"写作中断: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="开始深度撰写 (分章节)")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            doc.add_heading(self.entry_title.get(), 0)
            
            # 简单的格式处理
            for line in content.split('\n'):
                line = line.strip()
                if not line: continue
                
                if "=======" in line: # 识别章节分割线
                    continue
                elif line.startswith('### '):
                    doc.add_heading(line.replace('### ', ''), level=3)
                elif line.startswith('## '):
                    doc.add_heading(line.replace('## ', ''), level=2)
                elif line.startswith('# '):
                    doc.add_heading(line.replace('# ', ''), level=1)
                else:
                    doc.add_paragraph(line)
            
            doc.save(file_path)
            self.status_label.configure(text=f"已保存至: {os.path.basename(file_path)}", text_color="green")

    def load_config(self):
        try:
            with open("config.json", "r") as f:
                self.api_config = json.load(f)
        except: pass

    def save_config(self):
        self.api_config["api_key"] = self.entry_key.get().strip()
        self.api_config["base_url"] = self.entry_url.get().strip()
        self.api_config["model"] = self.entry_model.get().strip()
        with open("config.json", "w") as f:
            json.dump(self.api_config, f)
        self.status_label.configure(text="配置保存成功", text_color="green")

if __name__ == "__main__":
    app = PaperWriterApp()
    app.mainloop()
