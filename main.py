import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import re
import threading
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement

# --- 全局配置与默认值 ---
APP_NAME = "公文自动排版助手"
APP_VERSION = "v1.0.0"
AUTHOR_INFO = "开发者：Python开发者\n基于 GB/T 9704-2012 标准"

DEFAULT_CONFIG = {
    "margins": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
    "line_spacing": 28,  # 磅值
    "fonts": {
        "title": "方正小标宋简体", # 注意：电脑需安装此字体，否则Word会回退
        "h1": "黑体",
        "h2": "楷体_GB2312",
        "h3": "仿宋_GB2312",
        "body": "仿宋_GB2312"
    },
    "sizes": {
        "title": 22, # 二号
        "h1": 16,    # 三号
        "h2": 16,
        "h3": 16,
        "body": 16
    }
}

class GongWenFormatterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("900x700")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.config = self.load_config()
        self.file_list = []

        self.setup_ui()

    def load_config(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return DEFAULT_CONFIG
        return DEFAULT_CONFIG

    def save_config(self):
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", "配置已保存！")

    def setup_ui(self):
        # 侧边导航
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text=APP_NAME, font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        
        self.btn_home = ctk.CTkButton(self.sidebar, text="排版工作台", command=lambda: self.show_frame("home"))
        self.btn_home.pack(pady=10, padx=10)
        self.btn_settings = ctk.CTkButton(self.sidebar, text="参数设置", command=lambda: self.show_frame("settings"))
        self.btn_settings.pack(pady=10, padx=10)
        self.btn_about = ctk.CTkButton(self.sidebar, text="使用说明", command=lambda: self.show_frame("about"))
        self.btn_about.pack(pady=10, padx=10)

        # 主内容区
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.frames = {}
        self.create_home_frame()
        self.create_settings_frame()
        self.create_about_frame()

        self.show_frame("home")

    def create_home_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["home"] = f
        
        # 按钮区
        btn_box = ctk.CTkFrame(f, fg_color="transparent")
        btn_box.pack(fill="x", pady=10)
        
        ctk.CTkButton(btn_box, text="📂 上传文档 (支持多选)", command=self.upload_files, width=200).pack(side="left", padx=10)
        ctk.CTkButton(btn_box, text="▶ 开始一键排版", command=self.start_processing, width=200, fg_color="green").pack(side="left", padx=10)
        self.btn_export = ctk.CTkButton(btn_box, text="💾 导出结果", command=self.export_files, width=200, state="disabled")
        self.btn_export.pack(side="left", padx=10)

        # 列表区
        self.file_listbox = ctk.CTkTextbox(f, height=400)
        self.file_listbox.pack(fill="both", expand=True, pady=10)
        self.file_listbox.insert("0.0", "请上传 .docx 文档...\n")
        self.file_listbox.configure(state="disabled")

        # 进度条
        self.progressbar = ctk.CTkProgressBar(f)
        self.progressbar.pack(fill="x", pady=10)
        self.progressbar.set(0)
        
        self.status_label = ctk.CTkLabel(f, text="就绪")
        self.status_label.pack()

    def create_settings_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["settings"] = f
        
        ctk.CTkLabel(f, text="排版参数设置 (单位: cm / 磅)", font=("Arial", 20)).pack(pady=20)
        
        # 简单的参数输入示例
        self.entries = {}
        settings = [
            ("上边距 (cm)", "top", self.config["margins"]["top"]),
            ("下边距 (cm)", "bottom", self.config["margins"]["bottom"]),
            ("左边距 (cm)", "left", self.config["margins"]["left"]),
            ("右边距 (cm)", "right", self.config["margins"]["right"]),
            ("行间距 (磅)", "line_spacing", self.config["line_spacing"])
        ]

        for label_text, key, val in settings:
            row = ctk.CTkFrame(f, fg_color="transparent")
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text=label_text, width=100).pack(side="left")
            entry = ctk.CTkEntry(row)
            entry.insert(0, str(val))
            entry.pack(side="left", fill="x", expand=True)
            self.entries[key] = entry

        ctk.CTkButton(f, text="保存设置", command=self.update_config).pack(pady=20)

    def create_about_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["about"] = f
        
        info = f"""{APP_NAME}
版本：{APP_VERSION}
{AUTHOR_INFO}

【使用说明】
1. 点击“上传文档”，选择一个或多个 Word (.docx) 文件。
2. 点击“开始一键排版”，程序将自动处理。
3. 处理完成后，点击“导出结果”选择保存文件夹。

【排版规则】
- 自动识别“一、”、“（一）”、“1.”等层级。
- 自动设置国标版心（上3.7 下3.5 左2.8 右2.6）。
- 自动设置仿宋、黑体、楷体等公文专用字体。
- 自动设置固定行距。

注意：请确保电脑安装了“方正小标宋简体”、“仿宋_GB2312”、“楷体_GB2312”等字体，否则显示可能不正确。
"""
        lbl = ctk.CTkTextbox(f, font=("Arial", 14), wrap="word")
        lbl.insert("0.0", info)
        lbl.configure(state="disabled")
        lbl.pack(fill="both", expand=True)

    def show_frame(self, name):
        for frame in self.frames.values():
            frame.grid_forget()
        self.frames[name].grid(row=0, column=0, sticky="nsew")

    def update_config(self):
        try:
            self.config["margins"]["top"] = float(self.entries["top"].get())
            self.config["margins"]["bottom"] = float(self.entries["bottom"].get())
            self.config["margins"]["left"] = float(self.entries["left"].get())
            self.config["margins"]["right"] = float(self.entries["right"].get())
            self.config["line_spacing"] = float(self.entries["line_spacing"].get())
            self.save_config()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")

    def upload_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word Document", "*.docx")])
        if files:
            self.file_list = list(files)
            self.log(f"已加载 {len(files)} 个文件。")
            self.btn_export.configure(state="disabled")

    def log(self, text):
        self.file_listbox.configure(state="normal")
        self.file_listbox.delete("0.0", "end")
        for f in self.file_list:
            self.file_listbox.insert("end", f"{os.path.basename(f)}\n")
        self.file_listbox.insert("end", f"\n>>> {text}\n")
        self.file_listbox.configure(state="disabled")

    def start_processing(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先上传文件")
            return
        
        self.processed_docs = []
        threading.Thread(target=self.process_thread, daemon=True).start()

    def process_thread(self):
        total = len(self.file_list)
        for index, file_path in enumerate(self.file_list):
            self.status_label.configure(text=f"正在处理: {os.path.basename(file_path)}...")
            self.progressbar.set((index) / total)
            
            try:
                doc = self.format_document(file_path)
                self.processed_docs.append((file_path, doc))
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
            
            self.progressbar.set((index + 1) / total)
        
        self.status_label.configure(text="处理完成！请点击导出。")
        self.btn_export.configure(state="normal")

    def export_files(self):
        save_dir = filedialog.askdirectory()
        if not save_dir: return
        
        for original_path, doc in self.processed_docs:
            filename = os.path.basename(original_path)
            # 添加 "_排版后" 后缀，或者直接覆盖，这里选择保留原名但在新文件夹
            save_path = os.path.join(save_dir, filename)
            doc.save(save_path)
        
        messagebox.showinfo("完成", f"所有文件已导出至 {save_dir}")
        os.startfile(save_dir) if os.name == 'nt' else None

    # --- 核心排版逻辑 ---
    def format_document(self, file_path):
        doc = Document(file_path)
        cfg = self.config

        # 1. 页面设置
        section = doc.sections[0]
        section.top_margin = Cm(cfg["margins"]["top"])
        section.bottom_margin = Cm(cfg["margins"]["bottom"])
        section.left_margin = Cm(cfg["margins"]["left"])
        section.right_margin = Cm(cfg["margins"]["right"])
        
        # 尝试设置文档网格 (python-docx对此支持有限，通过行距模拟)
        # 2. 样式处理
        self.set_default_style(doc)

        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

            # 设置固定行距
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph.paragraph_format.line_spacing = Pt(cfg["line_spacing"])

            # 标题识别与字体设置
            # 标题 (简单假设第一段是标题，实际可能需要更复杂的逻辑)
            if paragraph == doc.paragraphs[0] and len(text) < 30: 
                self.set_font(paragraph, cfg["fonts"]["title"], cfg["sizes"]["title"], bold=False)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                continue

            # 一级标题 (一、)
            if re.match(r"^[一二三四五六七八九十]+、", text):
                self.set_font(paragraph, cfg["fonts"]["h1"], cfg["sizes"]["h1"], bold=False) # 黑体本身不需要加粗
                continue

            # 二级标题 ( (一) )
            if re.match(r"^（[一二三四五六七八九十]+）", text):
                self.set_font(paragraph, cfg["fonts"]["h2"], cfg["sizes"]["h2"], bold=False)
                continue

            # 三级标题 ( 1. )
            if re.match(r"^\d+\.", text):
                self.set_font(paragraph, cfg["fonts"]["h3"], cfg["sizes"]["h3"], bold=True) # 仿宋加粗
                continue

            # 正文
            self.set_font(paragraph, cfg["fonts"]["body"], cfg["sizes"]["body"])
            paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["body"] * 2) # 首行缩进2字符
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        # 表格处理
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self.set_font(p, "仿宋_GB2312", 14) # 表格内容通常小一号

        # 页码处理 (python-docx 插入页码非常复杂，通常需要底层XML操作)
        # 这里使用一种简化的 Footer 插入方式
        self.add_page_number(doc.sections[0].footer.paragraphs[0])

        return doc

    def set_font(self, paragraph, font_name, font_size, bold=False):
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    def set_default_style(self, doc):
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman' # 西文
        style.font.size = Pt(16)
        style._element.rPr.rFonts.set(qn('w:eastAsia'), self.config["fonts"]["body"])

    def add_page_number(self, paragraph):
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        # 简单设置页码字体
        run.font.name = "宋体"
        run.font.size = Pt(14)

if __name__ == "__main__":
    app = GongWenFormatterApp()
    app.mainloop()
