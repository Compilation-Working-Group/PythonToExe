import customtkinter as ctk
from openai import OpenAI
import os
import sys
import re
from tkinter import filedialog, messagebox
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

# Onefile 模式資源路徑修復（Linux 特別重要）
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
    os.chdir(application_path)

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class WritingAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 寫作助手 Pro")
        self.geometry("1350x1050")
        self.client = None
        self.model = "gpt-4o-mini"

        self.create_widgets()

    def create_widgets(self):
        # API 設置區
        api_frame = ctk.CTkFrame(self)
        api_frame.pack(pady=12, padx=20, fill="x")

        ctk.CTkLabel(api_frame, text="API Key:").grid(row=0, column=0, padx=6, pady=8, sticky="w")
        self.key_entry = ctk.CTkEntry(api_frame, width=360, show="*")
        self.key_entry.grid(row=0, column=1, padx=6, pady=8)

        ctk.CTkLabel(api_frame, text="Base URL:").grid(row=0, column=2, padx=6, pady=8, sticky="w")
        self.url_entry = ctk.CTkEntry(api_frame, width=320)
        self.url_entry.grid(row=0, column=3, padx=6, pady=8)
        self.url_entry.insert(0, "https://api.openai.com/v1")

        ctk.CTkLabel(api_frame, text="模型:").grid(row=0, column=4, padx=6, pady=8, sticky="w")
        self.model_combo = ctk.CTkComboBox(api_frame, width=240,
            values=[
                "gpt-4o", "gpt-4o-mini",
                "claude-3-5-sonnet-20241022",
                "llama3-70b-8192", "grok-beta",
                "deepseek-chat", "deepseek-reasoner"
            ],
            command=self.on_model_change)
        self.model_combo.set("gpt-4o-mini")
        self.model_combo.grid(row=0, column=5, padx=6, pady=8)

        ctk.CTkButton(api_frame, text="保存 API 設置", command=self.save_api).grid(row=0, column=6, padx=12, pady=8)

        ctk.CTkLabel(api_frame, text="（選擇 DeepSeek 會自動切換 Base URL）", 
                     text_color="gray").grid(row=1, column=2, columnspan=4, sticky="w", padx=6, pady=4)

        # 輸入區
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(input_frame, text="寫作類型:").grid(row=0, column=0, padx=6, pady=8, sticky="w")
        self.type_combo = ctk.CTkComboBox(input_frame, values=[
            "期刊論文", "項目計劃", "個人反思", "案例分析", "工作總結", "自定義"
        ], command=self.toggle_custom_prompt)
        self.type_combo.set("期刊論文")
        self.type_combo.grid(row=0, column=1, padx=6, pady=8)

        ctk.CTkLabel(input_frame, text="題目/主題:").grid(row=0, column=2, padx=6, pady=8, sticky="w")
        self.title_entry = ctk.CTkEntry(input_frame, width=440)
        self.title_entry.grid(row=0, column=3, padx=6, pady=8)

        ctk.CTkLabel(input_frame, text="目標字數:").grid(row=0, column=4, padx=6, pady=8, sticky="w")
        self.word_count_entry = ctk.CTkEntry(input_frame, width=130, placeholder_text="約6000字")
        self.word_count_entry.grid(row=0, column=5, padx=6, pady=8)

        ctk.CTkButton(input_frame, text="生成大綱", command=self.generate_outline).grid(row=0, column=6, padx=12, pady=8)

        # 額外要求區
        extra_frame = ctk.CTkFrame(self)
        extra_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(extra_frame, text="額外內容要求（可選）:").pack(anchor="w", padx=10, pady=4)
        self.extra_content = ctk.CTkTextbox(extra_frame, height=65)
        self.extra_content.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(extra_frame, text="額外格式要求（可選）:").pack(anchor="w", padx=10, pady=4)
        self.extra_format = ctk.CTkTextbox(extra_frame, height=65)
        self.extra_format.pack(fill="x", padx=10, pady=4)

        # 參考文獻
        refs_frame = ctk.CTkFrame(self)
        refs_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(refs_frame, text="附加參考文獻或材料（可選）:").pack(anchor="w", padx=10, pady=4)
        self.refs_text = ctk.CTkTextbox(refs_frame, height=90)
        self.refs_text.pack(fill="x", padx=10, pady=4)

        # 大綱區
        outline_frame = ctk.CTkFrame(self)
        outline_frame.pack(pady=10, padx=20, fill="both", expand=True)
        btn_frame1 = ctk.CTkFrame(outline_frame)
        btn_frame1.pack(fill="x", pady=6)
        ctk.CTkLabel(btn_frame1, text="大綱（可直接編輯）:").pack(side="left", padx=10)
        ctk.CTkButton(btn_frame1, text="清空大綱", command=lambda: self.outline_text.delete("1.0", "end")).pack(side="right", padx=10)
        self.outline_text = ctk.CTkTextbox(outline_frame)
        self.outline_text.pack(fill="both", expand=True, padx=10, pady=6)
        ctk.CTkButton(outline_frame, text="根據大綱生成全文", command=self.generate_full).pack(pady=10)

        # 結果區
        result_frame = ctk.CTkFrame(self)
        result_frame.pack(pady=10, padx=20, fill="both", expand=True)
        btn_frame2 = ctk.CTkFrame(result_frame)
        btn_frame2.pack(fill="x", pady=6)
        ctk.CTkLabel(btn_frame2, text="生成結果:").pack(side="left", padx=10)
        ctk.CTkButton(btn_frame2, text="清空結果", command=lambda: self.result_text.delete("1.0", "end")).pack(side="right", padx=6)
        ctk.CTkButton(btn_frame2, text="導出 Word（純文本）", command=self.export_word).pack(side="right", padx=6)
        ctk.CTkButton(btn_frame2, text="導出 Markdown", command=self.export_md).pack(side="right", padx=6)
        ctk.CTkButton(btn_frame2, text="導出 TXT", command=self.export_txt).pack(side="right", padx=6)

        self.result_text = ctk.CTkTextbox(result_frame)
        self.result_text.pack(fill="both", expand=True, padx=10, pady=6)

        # 自定義提示詞區
        self.custom_prompt = ctk.CTkTextbox(self, height=110)
        self.custom_prompt.insert("1.0", "在此輸入你的詳細寫作要求與結構...")
        self.toggle_custom_prompt(self.type_combo.get())

    def on_model_change(self, choice):
        choice_lower = choice.lower()
        if "deepseek" in choice_lower:
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.deepseek.com/v1")
        elif "grok" in choice_lower:
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.x.ai/v1")
        elif "llama" in choice_lower or "mixtral" in choice_lower:
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.groq.com/openai/v1")
        else:
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.openai.com/v1")

    def toggle_custom_prompt(self, choice):
        if choice == "自定義":
            self.custom_prompt.pack(pady=10, padx=20, fill="x")
        else:
            self.custom_prompt.pack_forget()

    def save_api(self):
        api_key = self.key_entry.get().strip()
        base_url = self.url_entry.get().strip() or None
        if not api_key:
            messagebox.showerror("錯誤", "請填寫 API Key")
            return
        try:
            self.client = OpenAI(api_key=api_key, base_url=base_url)
            self.model = self.model_combo.get()
            messagebox.showinfo("成功", f"API 設置保存成功\n模型: {self.model}")
        except Exception as e:
            messagebox.showerror("錯誤", f"API 初始化失敗：{str(e)}")

    def generate_outline(self):
        if not self.client:
            messagebox.showerror("錯誤", "請先保存 API 設置")
            return
        title = self.title_entry.get().strip()
        if not title:
            messagebox.showwarning("提示", "請填寫題目/主題")
            return
        prompt = self.build_prompt(self.type_combo.get(), title, is_outline=True)
        self.call_api(prompt, self.outline_text, max_tokens=2500)

    def generate_full(self):
        if not self.client:
            messagebox.showerror("錯誤", "請先保存 API 設置")
            return
        outline = self.outline_text.get("1.0", "end").strip()
        if not outline:
            messagebox.showwarning("提示", "大綱為空")
            return

        prompt = self.build_prompt(
            writing_type=self.type_combo.get(),
            title=self.title_entry.get().strip(),
            is_outline=False,
            outline=outline,
            refs=self.refs_text.get("1.0", "end").strip(),
            extra_content=self.extra_content.get("1.0", "end").strip(),
            extra_format=self.extra_format.get("1.0", "end").strip(),
            word_count=self.word_count_entry.get().strip(),
            custom=self.custom_prompt.get("1.0", "end").strip() if self.type_combo.get() == "自定義" else None
        )
        self.call_api(prompt, self.result_text, max_tokens=18000)

    def call_api(self, prompt, textbox, max_tokens=8000):
        textbox.delete("1.0", "end")
        textbox.insert("1.0", "正在生成，請稍候...")
        self.update_idletasks()

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.70,
                max_tokens=max_tokens
            )
            content = response.choices[0].message.content.strip()
            textbox.delete("1.0", "end")
            textbox.insert("1.0", content)
        except Exception as e:
            messagebox.showerror("生成失敗", str(e))

    def build_prompt(self, writing_type, title, is_outline, outline=None, refs=None,
                     extra_content=None, extra_format=None, word_count=None, custom=None):
        refs_part = f"\n\n附加參考材料（請適當引用）：\n{refs}" if refs else ""
        word_part = f"\n全文嚴格控制在約 {word_count} 字左右（含標點）。" if word_count and word_count.strip().isdigit() else ""
        content_part = f"\n額外內容要求：{extra_content}" if extra_content else ""
        format_part = f"\n額外格式要求：{extra_format}" if extra_format else ""

        if is_outline:
            return f"""你是一位嚴謹的學術/專業寫作助手。
請嚴格圍繞題目《{title}》生成詳細大綱。
要求：使用中文，層次清晰，使用 1.  1.1  1.2 等編號，每節給出簡要描述。"""

        else:
            return f"""你是一位嚴謹的學術/專業寫作助手。
請**嚴格按照**以下要求撰寫完整內容：

**題目**：{title}
**寫作類型**：{writing_type}

**大綱**（必須嚴格遵守，不得增減或改變順序）：
{outline}

{word_part}
{content_part}
{format_part}
{refs_part}

核心規則：
1. 全文必須緊緊圍繞《{title}》展開，絕對不能跑題。
2. 所有內容都要服務於該題目。
3. 語言正式、專業、邏輯嚴密。
4. 嚴格遵循上面給出的大綱結構。"""

    def clean_markdown(self, text):
        if not text:
            return ""
        lines = text.split('\n')
        cleaned = []
        prev = ""

        for line in lines:
            orig = line
            line = line.strip()

            if not line:
                cleaned.append("")
                continue

            # 去除常見 Markdown
            line = re.sub(r'^#{1,6}\s*', '', line)
            line = re.sub(r'(\*\*|__)(.+?)\1', r'\2', line)
            line = re.sub(r'(\*|_)(.+?)\1', r'\2', line)
            line = re.sub(r'^\s*[-*+]\s+', '', line)
            line = re.sub(r'^\s*\d+\.\s*', '', line)
            line = re.sub(r'`(.*?)`', r'\1', line)
            if line.startswith('```'):
                continue

            # 避免連續相同行（重複標題）
            if line == prev:
                continue

            cleaned.append(orig)  # 保留原始縮進
            prev = line

        return '\n'.join(cleaned)

    def export_word(self):
        raw_text = self.result_text.get("1.0", "end").strip()
        if not raw_text:
            return

        file = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")]
        )
        if not file:
            return

        doc = Document()
        title = self.title_entry.get().strip() or "未命名文檔"

        # 主標題
        doc.add_heading(title, level=1)

        clean_text = self.clean_markdown(raw_text)
        paragraphs = clean_text.split('\n')

        for para in paragraphs:
            para = para.strip()
            if not para:
                continue

            if re.match(r'^\d+\.\s', para):          # 1. 引言
                doc.add_heading(para, level=2)
            elif re.match(r'^\d+\.\d+\s', para):     # 1.1 研究背景
                doc.add_heading(para, level=3)
            elif re.match(r'^\d+\.\d+\.\d+\s', para): # 3.2.1 生活化情境
                doc.add_heading(para, level=4)
            else:
                p = doc.add_paragraph(para)
                p.style = 'Normal'

        doc.save(file)
        messagebox.showinfo("導出成功", f"已保存優化後的 Word 文件：\n{file}")

    def export_md(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text:
            return
        file = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown 文件", "*.md")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(f"# {self.title_entry.get().strip() or '未命名文檔'}\n\n{text}")
            messagebox.showinfo("成功", f"已保存：{file}")

    def export_txt(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text:
            return
        file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(text)
            messagebox.showinfo("成功", f"已保存：{file}")

if __name__ == "__main__":
    app = WritingAssistant()
    app.mainloop()
