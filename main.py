import customtkinter as ctk
from openai import OpenAI
import os
import sys
import re
from tkinter import filedialog, messagebox
from docx import Document

# Onefile 模式资源路径修复
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
    os.chdir(application_path)

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class WritingAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 写作助手 Pro")
        self.geometry("1300x980")
        self.client = None

        self.create_widgets()

    def create_widgets(self):
        # API 设置区（保持不变）
        api_frame = ctk.CTkFrame(self)
        api_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(api_frame, text="API Key:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.key_entry = ctk.CTkEntry(api_frame, width=350, show="*")
        self.key_entry.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(api_frame, text="Base URL:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.url_entry = ctk.CTkEntry(api_frame, width=300)
        self.url_entry.grid(row=0, column=3, padx=5, pady=5)
        self.url_entry.insert(0, "https://api.openai.com/v1")

        ctk.CTkLabel(api_frame, text="模型:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.model_combo = ctk.CTkComboBox(api_frame, width=200, values=[
            "gpt-4o", "gpt-4o-mini", "claude-3-5-sonnet-20241022", "llama3-70b-8192", "grok-beta"
        ])
        self.model_combo.set("gpt-4o-mini")
        self.model_combo.grid(row=0, column=5, padx=5, pady=5)

        ctk.CTkButton(api_frame, text="保存 API 设置", command=self.save_api).grid(row=0, column=6, padx=10, pady=5)

        # 输入区
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(input_frame, text="写作类型:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.type_combo = ctk.CTkComboBox(input_frame, values=[
            "期刊论文", "项目计划", "个人反思", "案例分析", "工作总结", "自定义"
        ], command=self.toggle_custom_prompt)
        self.type_combo.set("期刊论文")
        self.type_combo.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(input_frame, text="题目/主题:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.title_entry = ctk.CTkEntry(input_frame, width=400)
        self.title_entry.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkLabel(input_frame, text="目标字数:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.word_count_entry = ctk.CTkEntry(input_frame, width=110, placeholder_text="6000")
        self.word_count_entry.grid(row=0, column=5, padx=5, pady=5)

        ctk.CTkButton(input_frame, text="生成大纲", command=self.generate_outline).grid(row=0, column=6, padx=10, pady=5)

        # 额外要求区
        extra_frame = ctk.CTkFrame(self)
        extra_frame.pack(pady=8, padx=20, fill="x")
        ctk.CTkLabel(extra_frame, text="额外内容要求（可选）:").pack(anchor="w", padx=10)
        self.extra_content = ctk.CTkTextbox(extra_frame, height=55)
        self.extra_content.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(extra_frame, text="额外格式要求（可选）:").pack(anchor="w", padx=10)
        self.extra_format = ctk.CTkTextbox(extra_frame, height=55)
        self.extra_format.pack(fill="x", padx=10, pady=4)

        # 参考文献
        refs_frame = ctk.CTkFrame(self)
        refs_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(refs_frame, text="附加参考文献或材料（可选）:").pack(anchor="w", padx=10)
        self.refs_text = ctk.CTkTextbox(refs_frame, height=80)
        self.refs_text.pack(fill="x", padx=10, pady=5)

        # 大纲区
        outline_frame = ctk.CTkFrame(self)
        outline_frame.pack(pady=10, padx=20, fill="both", expand=True)
        btn1 = ctk.CTkFrame(outline_frame)
        btn1.pack(fill="x", pady=5)
        ctk.CTkLabel(btn1, text="大纲（可直接编辑）:").pack(side="left", padx=10)
        ctk.CTkButton(btn1, text="清空大纲", command=lambda: self.outline_text.delete("1.0", "end")).pack(side="right", padx=10)
        self.outline_text = ctk.CTkTextbox(outline_frame)
        self.outline_text.pack(fill="both", expand=True, padx=10, pady=5)

        ctk.CTkButton(outline_frame, text="根据大纲生成全文", command=self.generate_full).pack(pady=10)

        # 结果区
        result_frame = ctk.CTkFrame(self)
        result_frame.pack(pady=10, padx=20, fill="both", expand=True)
        btn2 = ctk.CTkFrame(result_frame)
        btn2.pack(fill="x", pady=5)
        ctk.CTkLabel(btn2, text="生成结果:").pack(side="left", padx=10)
        ctk.CTkButton(btn2, text="清空结果", command=lambda: self.result_text.delete("1.0", "end")).pack(side="right", padx=5)
        ctk.CTkButton(btn2, text="导出 Word（纯文本）", command=self.export_word).pack(side="right", padx=5)
        ctk.CTkButton(btn2, text="导出 Markdown", command=self.export_md).pack(side="right", padx=5)
        ctk.CTkButton(btn2, text="导出 TXT", command=self.export_txt).pack(side="right", padx=5)

        self.result_text = ctk.CTkTextbox(result_frame)
        self.result_text.pack(fill="both", expand=True, padx=10, pady=5)

        self.custom_prompt = ctk.CTkTextbox(self, height=100)
        self.custom_prompt.insert("1.0", "在此输入你的详细写作要求...")
        self.toggle_custom_prompt(self.type_combo.get())

    def toggle_custom_prompt(self, choice):
        if choice == "自定义":
            self.custom_prompt.pack(pady=10, padx=20, fill="x")
        else:
            self.custom_prompt.pack_forget()

    def save_api(self):
        api_key = self.key_entry.get().strip()
        base_url = self.url_entry.get().strip() or None
        if not api_key:
            messagebox.showerror("错误", "请填写 API Key")
            return
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.model = self.model_combo.get()
        messagebox.showinfo("成功", f"API 设置保存成功\n模型: {self.model}")

    def generate_outline(self):
        if not self.client: 
            messagebox.showerror("错误", "请先保存 API 设置"); return
        title = self.title_entry.get().strip()
        if not title:
            messagebox.showwarning("提示", "请填写题目/主题"); return

        prompt = self.build_prompt(self.type_combo.get(), title, is_outline=True)
        self.call_api(prompt, self.outline_text)

    def generate_full(self):
        if not self.client: 
            messagebox.showerror("错误", "请先保存 API 设置"); return
        outline = self.outline_text.get("1.0", "end").strip()
        if not outline:
            messagebox.showwarning("提示", "大纲为空"); return

        prompt = self.build_prompt(
            self.type_combo.get(),
            self.title_entry.get().strip(),
            is_outline=False,
            outline=outline,
            refs=self.refs_text.get("1.0", "end").strip(),
            extra_content=self.extra_content.get("1.0", "end").strip(),
            extra_format=self.extra_format.get("1.0", "end").strip(),
            word_count=self.word_count_entry.get().strip(),
            custom=self.custom_prompt.get("1.0", "end").strip() if self.type_combo.get() == "自定义" else None
        )
        self.call_api(prompt, self.result_text, max_tokens=15000)

    def call_api(self, prompt, textbox, max_tokens=8000):
        textbox.delete("1.0", "end")
        textbox.insert("1.0", "正在生成，请稍候...")
        self.update_idletasks()

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=max_tokens
            )
            content = response.choices[0].message.content.strip()
            textbox.delete("1.0", "end")
            textbox.insert("1.0", content)
        except Exception as e:
            messagebox.showerror("生成失败", str(e))

    # ==================== 核心修复：加强题目约束 ====================
    def build_prompt(self, writing_type, title, is_outline, outline=None, refs=None,
                     extra_content=None, extra_format=None, word_count=None, custom=None):
        
        refs_part = f"\n\n附加参考材料：\n{refs}" if refs else ""
        word_part = f"\n全文严格控制在约 {word_count} 字左右。" if word_count and word_count.isdigit() else ""
        content_part = f"\n额外内容要求：{extra_content}" if extra_content else ""
        format_part = f"\n额外格式要求：{extra_format}" if extra_format else ""

        # 核心修复：把真实题目强行放在最前面，并多次强调
        if is_outline:
            return f"""你现在是一位专业的学术/专业写作助手。
请严格围绕题目《{title}》生成详细大纲。
要求：使用中文，层次清晰，用数字编号，每节给出简要描述。"""

        else:
            return f"""你现在是一位专业的学术/专业写作助手。
请**严格按照**以下要求撰写内容：

**题目**：{title}

**写作类型**：{writing_type}

**大纲**（必须严格遵守，不得增减或改变顺序）：
{outline}

{word_part}{content_part}{format_part}{refs_part}

重要指令：
1. 全文必须紧紧围绕题目《{title}》展开，绝对不能跑题。
2. 所有内容都要服务于该题目，不允许写无关内容。
3. 语言正式、专业、逻辑严密。
4. 严格按照上面给出的大纲结构撰写。"""

    # Markdown 清理函数（保持不变）
    def clean_markdown(self, text):
        text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)
        text = re.sub(r'(\*\*|__)(.+?)\1', r'\2', text)
        text = re.sub(r'(\*|_)(.+?)\1', r'\2', text)
        text = re.sub(r'^\s*[-*+]\s+', '', text, flags=re.MULTILINE)
        text = re.sub(r'^\s*\d+\.\s+', '', text, flags=re.MULTILINE)
        text = re.sub(r'`(.*?)`', r'\1', text)
        text = re.sub(r'```[\s\S]*?```', '', text)
        return text.strip()

    def export_word(self):
        raw_text = self.result_text.get("1.0", "end").strip()
        if not raw_text: return
        file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word 文件", "*.docx")])
        if file:
            doc = Document()
            doc.add_heading(self.title_entry.get() or "未命名文档", level=0)
            clean_text = self.clean_markdown(raw_text)
            for para in clean_text.split("\n\n"):
                if para.strip():
                    doc.add_paragraph(para.strip())
            doc.save(file)
            messagebox.showinfo("导出成功", f"已保存纯文本 Word 文件：\n{file}")

    def export_md(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text: return
        file = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown 文件", "*.md")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(f"# {self.title_entry.get() or '未命名文档'}\n\n{text}")
            messagebox.showinfo("成功", f"已保存：{file}")

    def export_txt(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text: return
        file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(text)
            messagebox.showinfo("成功", f"已保存：{file}")

if __name__ == "__main__":
    app = WritingAssistant()
    app.mainloop()
