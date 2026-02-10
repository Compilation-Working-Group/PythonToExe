import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import re
import time
import traceback
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement

# --- Config ---
APP_NAME = "公文自动排版助手"
APP_VERSION = "v2.1.0 (Smart Structure)"
AUTHOR_INFO = "开发者：Python开发者\n基于 GB/T 9704-2012 标准"

DEFAULT_CONFIG = {
    "margins": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
    "line_spacing": 28, 
    "fonts": {
        "title": "方正小标宋简体",
        "subtitle": "楷体_GB2312",
        "h1": "黑体",
        "h2": "楷体_GB2312",
        "h3": "仿宋_GB2312",
        "body": "仿宋_GB2312"
    },
    "sizes": {
        "title": 22,    # 二号
        "subtitle": 16, # 三号
        "h1": 16,       # 三号
        "h2": 16,
        "h3": 16,
        "body": 16
    }
}

class GongWenFormatterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1000x750")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.config = self.load_config()
        self.file_list = []
        self.processed_docs = [] 
        self.process_queue = []
        self.setup_ui()

    def load_config(self):
        if os.path.exists("config.json"):
            try: return json.load(open("config.json", "r", encoding="utf-8"))
            except: pass
        return DEFAULT_CONFIG

    def save_config(self):
        try:
            json.dump(self.config, open("config.json", "w", encoding="utf-8"), ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e: messagebox.showerror("错误", str(e))

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.sidebar = ctk.CTkFrame(self, width=180, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text=APP_NAME, font=("Arial", 18, "bold")).pack(pady=20)
        
        btns = [("排版工作台", "home"), ("参数设置", "settings"), ("使用说明", "about")]
        for text, frame in btns:
            ctk.CTkButton(self.sidebar, text=text, command=lambda f=frame: self.show_frame(f)).pack(pady=10, padx=10)

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
        
        self.btn_upload = ctk.CTkButton(btn_box, text="📂 1. 上传文档", command=self.upload_files, width=180)
        self.btn_upload.pack(side="left", padx=10)
        self.btn_process = ctk.CTkButton(btn_box, text="▶ 2. 开始排版", command=self.start_processing, width=180, fg_color="green", state="disabled")
        self.btn_process.pack(side="left", padx=10)
        self.btn_export = ctk.CTkButton(btn_box, text="💾 3. 导出结果", command=self.export_files, width=180, state="disabled")
        self.btn_export.pack(side="left", padx=10)

        self.log_box = ctk.CTkTextbox(f)
        self.log_box.grid(row=1, column=0, sticky="nsew", pady=10)
        self.log_box.insert("0.0", ">>> 欢迎使用！请上传文档。\n")
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
            ("上边距 (cm)", "top", 3.7), ("下边距 (cm)", "bottom", 3.5),
            ("左边距 (cm)", "left", 2.8), ("右边距 (cm)", "right", 2.6),
            ("行间距 (磅)", "line_spacing", 28)
        ]
        for txt, key, val in settings:
            row = ctk.CTkFrame(f, fg_color="transparent")
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text=txt, width=120).pack(side="left")
            e = ctk.CTkEntry(row); e.insert(0, str(self.config["margins"].get(key, val) if key != "line_spacing" else self.config["line_spacing"]))
            e.pack(side="left", fill="x", expand=True)
            self.entries[key] = e
        ctk.CTkButton(f, text="保存设置", command=self.update_config).pack(pady=20)

    def create_about_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["about"] = f
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(0, weight=1)
        info = f"{APP_NAME} {APP_VERSION}\n\n改进说明：\n1. 标题识别不再受《》符号干扰。\n2. 修复了作者信息和称谓的对齐问题。\n3. 自动清除文本中的 'SAFE' 干扰字符。"
        lbl = ctk.CTkTextbox(f, font=("Arial", 14), wrap="word")
        lbl.insert("0.0", info)
        lbl.configure(state="disabled")
        lbl.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_frame(self, name):
        for f in self.frames.values(): f.grid_forget()
        self.frames[name].grid(row=0, column=0, sticky="nsew")

    def log(self, text):
        print(f"[LOG] {text}")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def update_config(self):
        try:
            for k, e in self.entries.items():
                val = float(e.get())
                if k == "line_spacing": self.config[k] = val
                else: self.config["margins"][k] = val
            self.save_config()
        except: messagebox.showerror("错误", "请输入数字")

    def upload_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word Document", "*.docx")])
        if files:
            self.file_list = list(files)
            self.processed_docs = []
            self.log(f"已加载 {len(files)} 个文件。")
            self.btn_process.configure(state="normal")
            self.btn_export.configure(state="disabled")

    def start_processing(self):
        self.btn_process.configure(state="disabled")
        self.btn_upload.configure(state="disabled")
        self.processed_docs = []
        self.process_queue = list(enumerate(self.file_list))
        self.total_files = len(self.file_list)
        self.success_count = 0
        self.update()
        self.after(100, self.process_next)

    def process_next(self):
        if not self.process_queue:
            self.finish_process()
            return
        idx, path = self.process_queue.pop(0)
        self.progressbar.set(idx / self.total_files)
        self.log(f"正在处理: {os.path.basename(path)} ...")
        self.update()
        try:
            doc = self.format_doc(path)
            self.processed_docs.append((path, doc))
            self.success_count += 1
            self.log("✅ 成功")
        except Exception as e:
            self.log(f"❌ 失败: {e}")
            traceback.print_exc()
        self.after(50, self.process_next)

    def finish_process(self):
        self.progressbar.set(1.0)
        self.btn_process.configure(state="normal")
        self.btn_upload.configure(state="normal")
        if self.success_count > 0:
            self.btn_export.configure(state="normal")
            messagebox.showinfo("完成", f"已处理 {self.success_count} 个文件")
        else:
            messagebox.showwarning("失败", "无文件成功处理")

    def export_files(self):
        d = filedialog.askdirectory()
        if not d: return
        count = 0
        for p, doc in self.processed_docs:
            try:
                name = os.path.splitext(os.path.basename(p))[0] + "_排版后.docx"
                doc.save(os.path.join(d, name))
                count += 1
            except Exception as e: self.log(f"导出错: {e}")
        messagebox.showinfo("完成", f"已导出 {count} 个文件到 {d}")
        if os.name == 'nt': os.startfile(d)

    # --- CORE FORMATTING LOGIC ---
    def format_doc(self, path):
        if not os.path.exists(path): raise Exception("文件丢失")
        doc = Document(path)
        cfg = self.config

        # 1. Page Setup
        try:
            sect = doc.sections[0]
            sect.top_margin = Cm(cfg["margins"]["top"])
            sect.bottom_margin = Cm(cfg["margins"]["bottom"])
            sect.left_margin = Cm(cfg["margins"]["left"])
            sect.right_margin = Cm(cfg["margins"]["right"])
            sect.page_width = Cm(21); sect.page_height = Cm(29.7)
        except: pass

        # 2. Structure Analysis & Formatting
        # We need to identify: Title, Subtitle, Author/Unit, Salutation (Start of Body), Headings, Body
        
        # State flags
        has_title = False
        body_started = False
        
        for i, p in enumerate(doc.paragraphs):
            # Clean "SAFE" artifact
            if "SAFE" in p.text: p.text = p.text.replace("SAFE", "")
            
            txt = p.text.strip()
            if not txt: continue

            # Reset format
            try:
                p.paragraph_format.first_line_indent = None
                p.paragraph_format.left_indent = None
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
            except: pass

            # Apply Base Line Spacing & Grid
            try:
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                p.paragraph_format.line_spacing = Pt(cfg["line_spacing"])
                self.set_grid_xml(p)
            except: pass

            # --- Logic ---
            
            # A. Explicit Body Start (Salutation)
            if re.match(r"^(尊敬的|各位|亲爱的|大家好)", txt):
                body_started = True
                self.style_body(p, cfg) # Salutation is part of body, left align, indent 2
                continue

            # B. If body already started, detect Headings or Body
            if body_started:
                if re.match(r"^[一二三四五六七八九十]+、", txt):
                    self.style_h1(p, cfg)
                elif re.match(r"^（[一二三四五六七八九十]+）", txt):
                    self.style_h2(p, cfg)
                elif re.match(r"^\d+\.", txt): # Matches "1." or "1、" if normalized
                    self.style_h3(p, cfg)
                else:
                    self.style_body(p, cfg)
                continue

            # C. Header Area (Before Body)
            
            # 1. Title (First significant line)
            if not has_title:
                # Rule: Short enough, no ending punctuation, or starts with 《
                if len(txt) < 50 and not txt.startswith("——") and not txt.startswith("--"):
                    self.style_title(p, cfg)
                    has_title = True
                    continue
            
            # 2. Subtitle (Starts with dash)
            if txt.startswith("——") or txt.startswith("--") or (txt.startswith("（") and txt.endswith("）") and len(txt)<30):
                self.style_subtitle(p, cfg)
                continue

            # 3. Author/Unit (Short, centered, after title, before body)
            if len(txt) < 25 and has_title and not body_started:
                self.style_subtitle(p, cfg) # Use subtitle style (KaiTi, centered)
                continue
            
            # 4. Abstract/Keywords
            if txt.startswith("摘要") or txt.startswith("关键词"):
                self.style_body(p, cfg) # Treat as body text style (FangSong) but maybe bold label?
                continue

            # Fallback: Treat as body
            body_started = True
            self.style_body(p, cfg)

        # Tables
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        if "SAFE" in p.text: p.text = p.text.replace("SAFE", "")
                        self.set_font(p, "仿宋_GB2312", 14)
                        self.set_grid_xml(p)

        # Page Number
        try:
            ftr = doc.sections[0].footer
            p = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
            self.add_page_num(p)
        except: pass

        return doc

    # --- Styling Helpers ---
    def style_title(self, p, cfg):
        self.set_font(p, cfg["fonts"]["title"], cfg["sizes"]["title"])
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.paragraph_format.space_after = Pt(cfg["line_spacing"]) # Gap after title
        self.set_indent_xml(p, 0)

    def style_subtitle(self, p, cfg):
        self.set_font(p, cfg["fonts"]["subtitle"], cfg["sizes"]["subtitle"])
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.set_indent_xml(p, 0)

    def style_h1(self, p, cfg):
        self.set_font(p, cfg["fonts"]["h1"], cfg["sizes"]["h1"])
        self.set_indent_xml(p, 2) # H1 usually has indent in some standards, or 0. Using 2 based on your feedback.

    def style_h2(self, p, cfg):
        self.set_font(p, cfg["fonts"]["h2"], cfg["sizes"]["h2"])
        self.set_indent_xml(p, 2)

    def style_h3(self, p, cfg):
        self.set_font(p, cfg["fonts"]["h3"], cfg["sizes"]["h3"], bold=True)
        self.set_indent_xml(p, 2)

    def style_body(self, p, cfg):
        self.set_font(p, cfg["fonts"]["body"], cfg["sizes"]["body"])
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        self.set_indent_xml(p, 2) # Standard 2 char indent

    # --- XML Helpers ---
    def set_font(self, p, name, size, bold=False):
        try:
            for r in p.runs:
                r.font.name = name
                r.font.size = Pt(size)
                r.bold = bold
                r._element.rPr.rFonts.set(qn('w:eastAsia'), name)
        except: pass

    def set_indent_xml(self, p, chars):
        try:
            pPr = p._p.get_or_add_pPr()
            ind = pPr.get_or_add_ind()
            if chars == 0:
                if 'w:firstLineChars' in ind.attrib: del ind.attrib['w:firstLineChars']
            else:
                ind.set(qn('w:firstLineChars'), str(int(chars * 100)))
        except: pass

    def set_grid_xml(self, p):
        try:
            pPr = p._p.get_or_add_pPr()
            snap = pPr.find(qn('w:snapToGrid'))
            if not snap: snap = OxmlElement('w:snapToGrid'); pPr.append(snap)
            snap.set(qn('w:val'), '1')
        except: pass

    def add_page_num(self, p):
        try:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            r = p.add_run()
            # Simple page num field
            f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
            txt = OxmlElement('w:instrText'); txt.set(qn('xml:space'), 'preserve'); txt.text = "PAGE"
            f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
            r._r.append(f1); r._r.append(txt); r._r.append(f2)
            r.font.name = "宋体"; r.font.size = Pt(14)
        except: pass

if __name__ == "__main__":
    app = GongWenFormatterApp()
    app.mainloop()
