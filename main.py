import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import re
import threading
import time
import traceback
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement

# --- 全局配置 ---
APP_NAME = "公文自动排版助手"
APP_VERSION = "v1.0.4 (Debug Mode)"
AUTHOR_INFO = "开发者：Python开发者\n基于 GB/T 9704-2012 标准"

DEFAULT_CONFIG = {
    "margins": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
    "line_spacing": 28,
    "fonts": {
        "title": "方正小标宋简体",
        "h1": "黑体",
        "h2": "楷体_GB2312",
        "h3": "仿宋_GB2312",
        "body": "仿宋_GB2312"
    },
    "sizes": {
        "title": 22,
        "h1": 16,
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
        self.processed_docs = [] 

        self.setup_ui()

    def load_config(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return DEFAULT_CONFIG

    def save_config(self):
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "配置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")

    def setup_ui(self):
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

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        self.frames = {}
        self.create_home_frame()
        self.create_settings_frame()
        self.create_about_frame()
        self.show_frame("home")

    def create_home_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["home"] = f
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        
        btn_box = ctk.CTkFrame(f, fg_color="transparent")
        btn_box.grid(row=0, column=0, sticky="ew", pady=10)
        
        ctk.CTkButton(btn_box, text="📂 1. 上传文档", command=self.upload_files, width=180).pack(side="left", padx=10)
        self.btn_process = ctk.CTkButton(btn_box, text="▶ 2. 开始排版", command=self.start_processing, width=180, fg_color="green", state="disabled")
        self.btn_process.pack(side="left", padx=10)
        self.btn_export = ctk.CTkButton(btn_box, text="💾 3. 导出结果", command=self.export_files, width=180, state="disabled")
        self.btn_export.pack(side="left", padx=10)

        self.log_box = ctk.CTkTextbox(f)
        self.log_box.grid(row=1, column=0, sticky="nsew", pady=10)
        self.log_box.insert("0.0", ">>> 欢迎使用！请先上传 Word 文档。\n")
        self.log_box.configure(state="disabled")

        self.progressbar = ctk.CTkProgressBar(f)
        self.progressbar.grid(row=2, column=0, sticky="ew", pady=10)
        self.progressbar.set(0)

    def create_settings_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["settings"] = f
        ctk.CTkLabel(f, text="排版参数设置", font=("Arial", 20)).pack(pady=20)
        
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
            ctk.CTkLabel(row, text=label_text, width=120).pack(side="left")
            entry = ctk.CTkEntry(row)
            entry.insert(0, str(val))
            entry.pack(side="left", fill="x", expand=True)
            self.entries[key] = entry

        ctk.CTkButton(f, text="保存设置", command=self.update_config).pack(pady=20)

    def create_about_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["about"] = f
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(0, weight=1)
        
        info = f"{APP_NAME}\n{APP_VERSION}\n{AUTHOR_INFO}\n\n【常见问题】\n如果排版无反应，可能是因为：\n1. 文档被其他程序占用了。\n2. 文档格式损坏（请另存为新docx重试）。\n3. 系统缺失必要字体。"
        lbl = ctk.CTkTextbox(f, font=("Arial", 14), wrap="word")
        lbl.insert("0.0", info)
        lbl.configure(state="disabled")
        lbl.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_frame(self, name):
        for frame in self.frames.values(): frame.grid_forget()
        self.frames[name].grid(row=0, column=0, sticky="nsew")

    def log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def update_config(self):
        try:
            self.config["margins"]["top"] = float(self.entries["top"].get())
            self.config["margins"]["bottom"] = float(self.entries["bottom"].get())
            self.config["margins"]["left"] = float(self.entries["left"].get())
            self.config["margins"]["right"] = float(self.entries["right"].get())
            self.config["line_spacing"] = float(self.entries["line_spacing"].get())
            self.save_config()
        except ValueError:
            messagebox.showerror("错误", "输入必须是数字")

    def upload_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word Document", "*.docx")])
        if files:
            self.file_list = list(files)
            self.processed_docs = [] 
            self.log(f"已加载 {len(files)} 个文件。请点击“开始排版”。")
            self.btn_process.configure(state="normal")
            self.btn_export.configure(state="disabled")

    def start_processing(self):
        # 1. 立即给用户反馈
        self.log(">>> 正在启动排版引擎...")
        self.btn_process.configure(state="disabled")
        self.btn_upload.configure(state="disabled")
        
        # 2. 清空旧数据
        self.processed_docs = []
        
        # 3. 启动线程
        t = threading.Thread(target=self.process_thread, daemon=True)
        t.start()

    # --- 线程安全更新 UI 的辅助方法 ---
    def safe_log(self, text):
        self.log(text)

    def safe_progress(self, val):
        self.progressbar.set(val)

    def safe_finish(self, count):
        self.on_process_finish(count)

    def process_thread(self):
        try:
            total = len(self.file_list)
            success_count = 0
            
            # 这里的 log 确认线程确实跑起来了
            self.after(0, self.safe_log, f"准备处理 {total} 个文件...")

            for index, file_path in enumerate(self.file_list):
                filename = os.path.basename(file_path)
                self.after(0, self.safe_log, f"正在分析: {filename} ...")
                self.after(0, self.safe_progress, index / total)
                
                try:
                    # 核心排版逻辑
                    doc = self.format_document(file_path)
                    self.processed_docs.append((file_path, doc))
                    success_count += 1
                    self.after(0, self.safe_log, f"✅ {filename} 排版成功")
                except Exception as e:
                    # 捕获单个文件的错误
                    error_msg = str(e)
                    print(f"Error processing {filename}: {error_msg}")
                    self.after(0, self.safe_log, f"❌ {filename} 失败: {error_msg}")
                
                self.after(0, self.safe_progress, (index + 1) / total)
                time.sleep(0.1) # 防止界面卡死

            self.after(0, self.safe_finish, success_count)

        except Exception as e:
            # 捕获整个线程的崩溃
            error_msg = f"排版引擎发生严重错误: {str(e)}"
            self.after(0, self.safe_log, error_msg)
            self.after(0, lambda: messagebox.showerror("错误", error_msg))
            self.after(0, self.safe_finish, 0)

    def on_process_finish(self, count):
        self.btn_process.configure(state="normal")
        self.btn_upload.configure(state="normal")
        if count > 0:
            self.btn_export.configure(state="normal")
            messagebox.showinfo("完成", f"已完成 {count} 个文档的排版！\n请点击“导出结果”保存文件。")
        else:
            self.log("⚠️ 没有生成任何结果，请检查文件是否损坏。")

    def export_files(self):
        if not self.processed_docs: return
        save_dir = filedialog.askdirectory(title="选择导出文件夹")
        if not save_dir: return
        
        count = 0
        for original_path, doc in self.processed_docs:
            try:
                base_name = os.path.basename(original_path)
                name, ext = os.path.splitext(base_name)
                new_name = f"{name}_排版后{ext}"
                save_path = os.path.join(save_dir, new_name)
                doc.save(save_path)
                self.log(f"已导出: {save_path}")
                count += 1
            except Exception as e:
                self.log(f"导出失败 {base_name}: {e}")

        messagebox.showinfo("导出完成", f"成功导出 {count} 个文件。\n路径: {save_dir}")
        if os.name == 'nt':
            try: os.startfile(save_dir)
            except: pass

    # --- 核心排版逻辑 ---
    def format_document(self, file_path):
        # 增加容错：检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError("找不到文件")

        doc = Document(file_path)
        cfg = self.config

        # 1. 页面设置
        for section in doc.sections:
            section.top_margin = Cm(cfg["margins"]["top"])
            section.bottom_margin = Cm(cfg["margins"]["bottom"])
            section.left_margin = Cm(cfg["margins"]["left"])
            section.right_margin = Cm(cfg["margins"]["right"])
            section.page_width = Cm(21)
            section.page_height = Cm(29.7)

        # 2. 字体设置 (增加 try-except 防止 lxml 报错)
        try:
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(cfg["sizes"]["body"])
            style._element.rPr.rFonts.set(qn('w:eastAsia'), cfg["fonts"]["body"])
        except Exception as e:
            print(f"样式初始化警告: {e}")

        # 3. 遍历排版
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text: continue

            try:
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                paragraph.paragraph_format.line_spacing = Pt(cfg["line_spacing"])
            except: pass
            
            # 标题识别
            if paragraph == doc.paragraphs[0] and len(text) < 50:
                self.set_font(paragraph, cfg["fonts"]["title"], cfg["sizes"]["title"], bold=False)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                paragraph.paragraph_format.space_after = Pt(cfg["line_spacing"])
                continue

            if re.match(r"^[一二三四五六七八九十]+、", text):
                self.set_font(paragraph, cfg["fonts"]["h1"], cfg["sizes"]["h1"], bold=False)
                paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["h1"] * 2)
                continue

            if re.match(r"^（[一二三四五六七八九十]+）", text):
                self.set_font(paragraph, cfg["fonts"]["h2"], cfg["sizes"]["h2"], bold=False)
                paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["h2"] * 2)
                continue

            if re.match(r"^\d+\.", text):
                self.set_font(paragraph, cfg["fonts"]["h3"], cfg["sizes"]["h3"], bold=True)
                paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["h3"] * 2)
                continue

            self.set_font(paragraph, cfg["fonts"]["body"], cfg["sizes"]["body"])
            paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["body"] * 2)
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        # 4. 表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self.set_font(p, "仿宋_GB2312", 14) 

        # 5. 页码
        try:
            footer = doc.sections[0].footer
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            self.add_page_number(p)
        except: pass

        return doc

    def set_font(self, paragraph, font_name, font_size, bold=False):
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.bold = bold
            try:
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            except: pass

    def add_page_number(self, paragraph):
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
        run.font.name = "宋体"; run.font.size = Pt(14)

if __name__ == "__main__":
    app = GongWenFormatterApp()
    app.mainloop()
