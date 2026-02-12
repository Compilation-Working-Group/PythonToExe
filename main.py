"""
AI 写作助手 - 智能文稿创作平台
支持学术论文、研究报告、工作计划、反思总结、案例分析、工作总结及自定义文稿的智能撰写
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import anthropic
import json
import os
from datetime import datetime

# ── 主题配置 ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── 常量定义 ────────────────────────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".ai_writer_config.json")
APP_VERSION = "v1.0.0"

DOCUMENT_TYPES = [
    ("📄", "学术论文",  "含摘要、引言、方法、结果、讨论、参考文献"),
    ("📊", "研究报告",  "含背景、分析框架、结论与建议"),
    ("📋", "工作计划",  "含目标、阶段步骤、时间线、资源安排"),
    ("🔍", "反思总结",  "含经历回顾、收获、不足与改进方向"),
    ("🔬", "案例分析",  "含案例背景、问题呈现、深度分析、启示"),
    ("📝", "工作总结",  "含工作概述、核心成果、问题与展望"),
    ("✨", "自定义",    "根据您的描述自由定制文稿类型与结构"),
]

OUTLINE_SYSTEM = """你是一位资深写作顾问，擅长为各类专业文稿设计清晰、合理的结构大纲。

请根据用户提供的文稿类型、题目和要求，输出一份层次分明的大纲。

格式规范：
- 一级章节：1. 章节名称（简要说明本章核心内容）
- 二级章节：1.1 小节名称（说明）
- 三级要点：1.1.1 要点（如有必要）
- 每个条目要精炼，括号内说明控制在20字以内

注意：
- 直接输出大纲正文，无需前言或解释
- 学术论文须包含摘要、关键词、引言、正文各节、结论、参考文献
- 其他类型按其行文惯例组织结构
- 大纲条目数量适中，一般10~20条为宜
"""

WRITING_SYSTEM = """你是一位经验丰富的专业写作专家，擅长撰写高质量、内容充实的各类文稿。

请严格依据提供的文稿类型、题目、要求和大纲，撰写完整的正文内容。

写作规范：
- 语言专业、准确、流畅，符合相应文体规范
- 内容充实，论据充分，逻辑严密
- 严格按照大纲结构依次展开，不得遗漏章节
- 每个章节内容饱满，避免空洞
- 学术论文须有理论依据，工作类文稿须结合实际
- 使用 Markdown 格式：# 一级标题，## 二级标题，**加粗**等
- 直接输出正文，无需额外说明
"""


# ── 配置管理器 ──────────────────────────────────────────────────────────────
class ConfigManager:
    def __init__(self):
        self._data = self._load()

    def _load(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {"api_key": "", "model": "claude-sonnet-4-5-20250929", "last_type": "学术论文"}

    def save(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get(self, key, default=""):
        return self._data.get(key, default)

    def set(self, key, value):
        self._data[key] = value
        self.save()


# ── 可滚动文本框组件 ────────────────────────────────────────────────────────
class TextEditor(ctk.CTkFrame):
    def __init__(self, parent, font=None, **kwargs):
        super().__init__(parent, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        _font = font or ctk.CTkFont(size=13)
        self.textbox = ctk.CTkTextbox(self, font=_font, wrap="word", **kwargs)
        self.textbox.grid(row=0, column=0, sticky="nsew")

    def get(self) -> str:
        return self.textbox.get("1.0", "end-1c")

    def set(self, text: str):
        self.textbox.delete("1.0", "end")
        if text:
            self.textbox.insert("1.0", text)

    def append(self, text: str):
        self.textbox.insert("end", text)
        self.textbox.see("end")

    def clear(self):
        self.textbox.delete("1.0", "end")

    def set_readonly(self, readonly: bool):
        state = "disabled" if readonly else "normal"
        self.textbox.configure(state=state)


# ── 侧边栏文档类型按钮 ──────────────────────────────────────────────────────
class DocTypeButton(ctk.CTkButton):
    ACTIVE_COLOR   = ("#2B6CB0", "#1A4F8A")   # 深蓝选中
    INACTIVE_COLOR = ("transparent", "transparent")
    HOVER_COLOR    = ("#EBF4FF", "#1E3A5F")

    def __init__(self, parent, icon, name, desc, command, **kwargs):
        super().__init__(
            parent,
            text=f"  {icon}  {name}",
            anchor="w",
            font=ctk.CTkFont(size=13),
            height=40,
            corner_radius=8,
            fg_color=self.INACTIVE_COLOR,
            hover_color=self.HOVER_COLOR,
            command=command,
            **kwargs,
        )
        self._name = name
        self._desc = desc

    def activate(self):
        self.configure(fg_color=self.ACTIVE_COLOR, font=ctk.CTkFont(size=13, weight="bold"))

    def deactivate(self):
        self.configure(fg_color=self.INACTIVE_COLOR, font=ctk.CTkFont(size=13))


# ── 主应用窗口 ──────────────────────────────────────────────────────────────
class AIWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self._cfg    = ConfigManager()
        self._busy   = False
        self._doc_type = self._cfg.get("last_type", "学术论文")
        self._type_btns: dict[str, DocTypeButton] = {}

        self.title(f"✍️  AI 写作助手  {APP_VERSION}")
        self.geometry("1280x820")
        self.minsize(960, 620)

        self._build_ui()
        self._load_config_values()
        self._select_type(self._doc_type, save=False)

    # ── UI 构建 ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=240, corner_radius=0,
                          fg_color=("#1A2744", "#0F1A33"))
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_columnconfigure(0, weight=1)
        sb.grid_rowconfigure(9, weight=1)   # spacer row

        # ── Logo 区域 ──
        logo_frame = ctk.CTkFrame(sb, fg_color="transparent")
        logo_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(22, 4))

        ctk.CTkLabel(logo_frame, text="✍️", font=ctk.CTkFont(size=28)).pack(side="left")
        title_col = ctk.CTkFrame(logo_frame, fg_color="transparent")
        title_col.pack(side="left", padx=(8, 0))
        ctk.CTkLabel(title_col, text="AI 写作助手",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="white").pack(anchor="w")
        ctk.CTkLabel(title_col, text="智能文稿创作平台",
                     font=ctk.CTkFont(size=10),
                     text_color="#7FA8D4").pack(anchor="w")

        # ── 分隔线 ──
        ctk.CTkLabel(sb, text="─" * 26, font=ctk.CTkFont(size=9),
                     text_color="#2A4070").grid(row=1, column=0, pady=(8, 4))

        ctk.CTkLabel(sb, text="文稿类型",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7FA8D4").grid(row=2, column=0, sticky="w", padx=18, pady=(0, 6))

        # ── 文档类型按钮 ──
        for idx, (icon, name, desc) in enumerate(DOCUMENT_TYPES):
            btn = DocTypeButton(
                sb, icon=icon, name=name, desc=desc,
                command=lambda n=name: self._select_type(n)
            )
            btn.grid(row=3 + idx, column=0, padx=10, pady=2, sticky="ew")
            self._type_btns[name] = btn

        # ── 弹性空间 ──
        ctk.CTkLabel(sb, text="").grid(row=9, column=0, sticky="nsew")

        # ── 设置区域 ──
        ctk.CTkLabel(sb, text="─" * 26, font=ctk.CTkFont(size=9),
                     text_color="#2A4070").grid(row=10, column=0, pady=(0, 6))

        ctk.CTkLabel(sb, text="Anthropic API Key",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7FA8D4").grid(row=11, column=0, sticky="w", padx=18, pady=(0, 4))

        self._api_entry = ctk.CTkEntry(
            sb, placeholder_text="sk-ant-api...", show="*", height=34,
            fg_color=("#0D1B36", "#0A1228"), border_color="#2A4070",
            text_color="white", placeholder_text_color="#4A6FA0"
        )
        self._api_entry.grid(row=12, column=0, padx=10, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(sb, text="模型",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7FA8D4").grid(row=13, column=0, sticky="w", padx=18, pady=(0, 4))

        self._model_var = ctk.StringVar(value="claude-sonnet-4-5-20250929")
        self._model_menu = ctk.CTkOptionMenu(
            sb,
            variable=self._model_var,
            values=[
                "claude-opus-4-5-20251101",
                "claude-sonnet-4-5-20250929",
                "claude-haiku-4-5-20251001",
            ],
            height=34,
            fg_color=("#0D1B36", "#0A1228"),
            button_color=("#2B6CB0", "#1A4F8A"),
        )
        self._model_menu.grid(row=14, column=0, padx=10, pady=(0, 8), sticky="ew")

        save_btn = ctk.CTkButton(
            sb, text="💾  保存设置", height=34,
            fg_color=("#1A4F8A", "#153D6F"),
            hover_color=("#2B6CB0", "#1A4F8A"),
            command=self._save_settings,
        )
        save_btn.grid(row=15, column=0, padx=10, pady=(0, 20), sticky="ew")

    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=(0, 12), pady=12)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        # ── 顶栏：类型标签 + 状态 ──
        topbar = ctk.CTkFrame(main, fg_color="transparent", height=42)
        topbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        topbar.grid_columnconfigure(1, weight=1)
        topbar.grid_propagate(False)

        self._badge = ctk.CTkLabel(
            topbar, text="📄  学术论文",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#2B6CB0", "#1A4F8A"),
            corner_radius=8, padx=14, pady=6,
        )
        self._badge.grid(row=0, column=0, padx=(0, 12))

        self._status_var = tk.StringVar(value="就绪 · 请输入题目后生成大纲")
        status_lbl = ctk.CTkLabel(
            topbar, textvariable=self._status_var,
            font=ctk.CTkFont(size=12), text_color="#7FA8D4",
        )
        status_lbl.grid(row=0, column=1, sticky="w")

        # ── 输入区 ──
        input_card = ctk.CTkFrame(main, corner_radius=10)
        input_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        input_card.grid_columnconfigure(1, weight=2)
        input_card.grid_columnconfigure(3, weight=3)

        ctk.CTkLabel(input_card, text="题目 / 主题",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#A8C8F0").grid(
            row=0, column=0, padx=(16, 8), pady=14, sticky="w"
        )
        self._title_entry = ctk.CTkEntry(
            input_card,
            placeholder_text="输入文稿题目或主题...",
            height=38, font=ctk.CTkFont(size=13),
        )
        self._title_entry.grid(row=0, column=1, padx=(0, 16), pady=14, sticky="ew")

        ctk.CTkLabel(input_card, text="附加要求",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#A8C8F0").grid(
            row=0, column=2, padx=(0, 8), pady=14, sticky="w"
        )
        self._req_entry = ctk.CTkEntry(
            input_card,
            placeholder_text="字数限制、风格偏好、特定内容要求等（可选）...",
            height=38, font=ctk.CTkFont(size=13),
        )
        self._req_entry.grid(row=0, column=3, padx=(0, 16), pady=14, sticky="ew")

        # ── 标签页 ──
        self._tabs = ctk.CTkTabview(main, corner_radius=10)
        self._tabs.grid(row=2, column=0, sticky="nsew")

        self._build_outline_tab(self._tabs.add("📋  大纲编辑"))
        self._build_output_tab(self._tabs.add("📄  正文输出"))

        # ── 进度条 ──
        self._progress = ctk.CTkProgressBar(main, mode="indeterminate", height=4)
        self._progress.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        self._progress.set(0)

    def _build_outline_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(tab, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(4, 8))

        self._btn_gen_outline = ctk.CTkButton(
            toolbar, text="🔮  生成大纲",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38, width=140,
            command=self._on_gen_outline,
        )
        self._btn_gen_outline.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="🗑  清空",
            font=ctk.CTkFont(size=12), height=38, width=72,
            fg_color="transparent", border_width=1,
            command=lambda: self._outline_editor.clear(),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="✍️  开始撰写",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38, width=140,
            fg_color=("#276749", "#1A4731"),
            hover_color=("#2F855A", "#22543D"),
            command=self._on_gen_text,
        ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            toolbar,
            text="💡 大纲生成后可直接编辑，修改完成后点击「开始撰写」",
            font=ctk.CTkFont(size=12), text_color="#7FA8D4",
        ).pack(side="left")

        self._outline_editor = TextEditor(
            tab,
            font=ctk.CTkFont(size=13, family="Consolas"),
        )
        self._outline_editor.grid(row=1, column=0, sticky="nsew")

    def _build_output_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(tab, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(4, 8))

        self._btn_gen_text = ctk.CTkButton(
            toolbar, text="✍️  开始撰写",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38, width=140,
            fg_color=("#276749", "#1A4731"),
            hover_color=("#2F855A", "#22543D"),
            command=self._on_gen_text,
        )
        self._btn_gen_text.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="📋  复制",
            font=ctk.CTkFont(size=12), height=38, width=72,
            fg_color="transparent", border_width=1,
            command=self._copy_output,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            toolbar, text="💾  保存",
            font=ctk.CTkFont(size=12), height=38, width=72,
            fg_color="transparent", border_width=1,
            command=self._save_output,
        ).pack(side="left", padx=(0, 12))

        self._wc_var = tk.StringVar(value="字数：0")
        ctk.CTkLabel(
            toolbar, textvariable=self._wc_var,
            font=ctk.CTkFont(size=12), text_color="#7FA8D4",
        ).pack(side="left")

        self._output_editor = TextEditor(tab, font=ctk.CTkFont(size=13))
        self._output_editor.grid(row=1, column=0, sticky="nsew")

    # ── 事件处理 ────────────────────────────────────────────────────────────
    def _select_type(self, name: str, save: bool = True):
        self._doc_type = name
        for n, btn in self._type_btns.items():
            btn.activate() if n == name else btn.deactivate()
        icon = next((i for i, n, _ in DOCUMENT_TYPES if n == name), "✨")
        self._badge.configure(text=f"{icon}  {name}")
        if save:
            self._cfg.set("last_type", name)

    def _load_config_values(self):
        self._api_entry.insert(0, self._cfg.get("api_key", ""))
        saved_model = self._cfg.get("model", "claude-sonnet-4-5-20250929")
        self._model_var.set(saved_model)

    def _save_settings(self):
        self._cfg.set("api_key", self._api_entry.get().strip())
        self._cfg.set("model", self._model_var.get())
        self._set_status("✅  设置已保存", "#68D391")

    def _get_client(self):
        key = self._api_entry.get().strip()
        if not key:
            messagebox.showerror("缺少 API Key", "请在左侧设置中输入 Anthropic API Key！")
            return None
        return anthropic.Anthropic(api_key=key)

    def _set_status(self, text: str, color: str = "#7FA8D4"):
        self._status_var.set(text)
        # 动态找到 status label 更新颜色（通过引用已存储的 widget）
        # 简化处理：直接更新 status_var，颜色通过已配置的 label 显示

    def _set_busy(self, busy: bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        self._btn_gen_outline.configure(state=state)
        self._btn_gen_text.configure(state=state)
        if busy:
            self._progress.start()
        else:
            self._progress.stop()
            self._progress.set(0)

    # ── 生成大纲 ────────────────────────────────────────────────────────────
    def _on_gen_outline(self):
        if self._busy:
            return
        title = self._title_entry.get().strip()
        if not title:
            messagebox.showwarning("提示", "请先输入文稿题目或主题！")
            return
        client = self._get_client()
        if not client:
            return

        self._set_busy(True)
        self._set_status("⏳  正在生成大纲...")
        self._outline_editor.clear()
        self._tabs.set("📋  大纲编辑")

        doc_type = self._doc_type
        req      = self._req_entry.get().strip()
        model    = self._model_var.get()

        prompt = f"文稿类型：{doc_type}\n题目：{title}"
        if req:
            prompt += f"\n特殊要求：{req}"

        def run():
            try:
                with client.messages.stream(
                    model=model,
                    max_tokens=2048,
                    system=OUTLINE_SYSTEM,
                    messages=[{"role": "user", "content": prompt}],
                ) as stream:
                    for chunk in stream.text_stream:
                        self.after(0, lambda c=chunk: self._outline_editor.append(c))
                self.after(0, lambda: self._set_status("✅  大纲生成完成 · 可直接编辑后点击「开始撰写」"))
            except Exception as exc:
                self.after(0, lambda e=exc: messagebox.showerror("生成失败", str(e)))
                self.after(0, lambda: self._set_status("❌  大纲生成失败"))
            finally:
                self.after(0, lambda: self._set_busy(False))

        threading.Thread(target=run, daemon=True).start()

    # ── 生成正文 ────────────────────────────────────────────────────────────
    def _on_gen_text(self):
        if self._busy:
            return
        title   = self._title_entry.get().strip()
        outline = self._outline_editor.get().strip()

        if not title:
            messagebox.showwarning("提示", "请先输入文稿题目或主题！")
            return
        if not outline:
            messagebox.showwarning("提示", "请先生成或填写大纲内容！")
            return

        client = self._get_client()
        if not client:
            return

        self._set_busy(True)
        self._set_status("⏳  正在撰写正文，请稍候...")
        self._output_editor.clear()
        self._wc_var.set("字数：0")
        self._tabs.set("📄  正文输出")

        doc_type = self._doc_type
        req      = self._req_entry.get().strip()
        model    = self._model_var.get()

        prompt = f"文稿类型：{doc_type}\n题目：{title}\n大纲：\n{outline}"
        if req:
            prompt += f"\n特殊要求：{req}"

        def run():
            char_count = 0
            try:
                with client.messages.stream(
                    model=model,
                    max_tokens=8192,
                    system=WRITING_SYSTEM,
                    messages=[{"role": "user", "content": prompt}],
                ) as stream:
                    for chunk in stream.text_stream:
                        char_count += len(chunk)
                        self.after(0, lambda c=chunk: self._output_editor.append(c))
                        self.after(0, lambda n=char_count: self._wc_var.set(f"字数：{n}"))
                self.after(0, lambda: self._set_status(
                    f"✅  撰写完成 · 共 {char_count} 字"))
            except Exception as exc:
                self.after(0, lambda e=exc: messagebox.showerror("生成失败", str(e)))
                self.after(0, lambda: self._set_status("❌  撰写失败"))
            finally:
                self.after(0, lambda: self._set_busy(False))

        threading.Thread(target=run, daemon=True).start()

    # ── 复制 / 保存 ─────────────────────────────────────────────────────────
    def _copy_output(self):
        text = self._output_editor.get()
        if not text:
            messagebox.showinfo("提示", "暂无可复制的内容。")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self._set_status("✅  已复制到剪贴板")

    def _save_output(self):
        text = self._output_editor.get()
        if not text:
            messagebox.showinfo("提示", "暂无可保存的内容。")
            return
        title      = self._title_entry.get().strip() or "文稿"
        timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_fn = f"{title}_{timestamp}"

        fp = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[
                ("Markdown 文件 (*.md)",  "*.md"),
                ("纯文本文件 (*.txt)",    "*.txt"),
                ("所有文件",              "*.*"),
            ],
            initialfile=default_fn,
            title="保存文稿",
        )
        if fp:
            try:
                with open(fp, "w", encoding="utf-8") as f:
                    f.write(text)
                self._set_status(f"✅  已保存：{os.path.basename(fp)}")
            except Exception as exc:
                messagebox.showerror("保存失败", str(exc))


# ── 入口 ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AIWriterApp()
    app.mainloop()
