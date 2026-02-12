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
        self.geometry("1300x1000")
        self.client = None

        self.create_widgets()

    def create_widgets(self):
        # API 设置区
        api_frame = ctk.CTkFrame(self)
        api_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(api_frame, text="API Key:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.key_entry = ctk.CTkEntry(api_frame, width=350, show="*")
        self.key_entry.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(api_frame, text="Base URL:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.url_entry = ctk.CTkEntry(api_frame, width=320)
        self.url_entry.grid(row=0, column=3, padx=5, pady=5)
        self.url_entry.insert(0, "https://api.openai.com/v1")

        ctk.CTkLabel(api_frame, text="模型:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.model_combo = ctk.CTkComboBox(api_frame, width=220, 
            values=[
                "gpt-4o", "gpt-4o-mini",
                "claude-3-5-sonnet-20241022",
                "llama3-70b-8192", "grok-beta",
                "deepseek-chat", "deepseek-reasoner"
            ],
            command=self.on_model_change)   # ← 新增：模型改变时自动切换 Base URL
        self.model_combo.set("gpt-4o-mini")
        self.model_combo.grid(row=0, column=5, padx=5, pady=5)

        ctk.CTkButton(api_frame, text="保存 API 设置", command=self.save_api).grid(row=0, column=6, padx=10, pady=5)

        # 小提示
        ctk.CTkLabel(api_frame, text="（选择 DeepSeek 模型会自动切换 Base URL）", 
                     text_color="gray").grid(row=1, column=3, columnspan=3, sticky="w", padx=5)

        # 其余界面代码保持不变（输入区、额外要求、大纲、结果区等）
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

        # 大纲区、结果区等（省略中间重复代码，保持和上一个版本完全一致）
        # ...（此处省略大纲、结果、导出等代码，与上一个版本完全相同）

        outline_frame = ctk.CTkFrame(self)
        outline_frame.pack(pady=10, padx=20, fill="both", expand=True)
        btn1 = ctk.CTkFrame(outline_frame)
        btn1.pack(fill="x", pady=5)
        ctk.CTkLabel(btn1, text="大纲（可直接编辑）:").pack(side="left", padx=10)
        ctk.CTkButton(btn1, text="清空大纲", command=lambda: self.outline_text.delete("1.0", "end")).pack(side="right", padx=10)
        self.outline_text = ctk.CTkTextbox(outline_frame)
        self.outline_text.pack(fill="both", expand=True, padx=10, pady=5)
        ctk.CTkButton(outline_frame, text="根据大纲生成全文", command=self.generate_full).pack(pady=10)

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

    # 新增：模型改变时自动切换 Base URL
    def on_model_change(self, choice):
        if "deepseek" in choice.lower():
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.deepseek.com/v1")
        elif "grok" in choice.lower():
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.x.ai/v1")
        elif "llama" in choice.lower() or "mixtral" in choice.lower():
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.groq.com/openai/v1")
        else:
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, "https://api.openai.com/v1")

    # 下面所有方法（toggle_custom_prompt、save_api、generate_outline、generate_full、call_api、build_prompt、clean_markdown、export_xxx）与上一个版本完全一致
    # 为节省篇幅这里不再重复粘贴，你可以直接把上一个版本中对应的函数复制过来替换即可
    # 如果你需要我把完整代码一次性发给你，请直接说“我要完整版”，我马上发出。

if __name__ == "__main__":
    app = WritingAssistant()
    app.mainloop()
