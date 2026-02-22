import customtkinter as ctk
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog
import threading
import json
import os
from openai import OpenAI
from docx import Document

# 设置 CustomTkinter 的全局主题和颜色
ctk.set_appearance_mode("System")  # 跟随系统深色/浅色模式
ctk.set_default_color_theme("blue") # 主题色

CONFIG_FILE = "docwriter_config.json"

class ModernAIDocWriter:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepSeek 智能写作 Pro版 v3.1 (超长文本支持)")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)
        
        self.is_generating = False
        self.stop_flag = False
        
        self.load_config()
        self.create_ui()

    def load_config(self):
        """加载本地保存的配置文件（如 API Key）"""
        self.config = {"api_key": ""}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            except:
                pass

    def save_config(self, api_key):
        """保存配置到本地"""
        self.config["api_key"] = api_key
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f)
        except:
            pass

    def create_ui(self):
        # 整体网格布局
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # ==================== 左侧侧边栏 ====================
        self.sidebar = ctk.CTkFrame(self.root, width=290, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(7, weight=1) 

        # Logo / 标题
        self.logo_label = ctk.CTkLabel(self.sidebar, text="✨ AI 写作 Pro", font=ctk.CTkFont(family="微软雅黑", size=24, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 20))

        # 1. API Key 输入框
        self.api_key_entry = ctk.CTkEntry(self.sidebar, placeholder_text="输入 DeepSeek API Key", show="*")
        self.api_key_entry.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="ew")
        if self.config.get("api_key"):
            self.api_key_entry.insert(0, self.config["api_key"])

        # 2. 需求描述输入
        self.topic_label = ctk.CTkLabel(self.sidebar, text="🎯 具体写作需求：", anchor="w", font=ctk.CTkFont(weight="bold"))
        self.topic_label.grid(row=2, column=0, padx=20, pady=(5, 0), sticky="ew")
        self.topic_textbox = ctk.CTkTextbox(self.sidebar, height=100)
        self.topic_textbox.grid(row=3, column=0, padx=20, pady=(5, 15), sticky="ew")
        self.topic_textbox.insert("1.0", "例如：写一份关于新能源汽车市场下半年的调研报告，侧重于电池技术的突破...")

        # 3. 语气与篇幅设置 (双列布局)
        self.settings_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.settings_frame.grid(row=4, column=0, padx=20, pady=5, sticky="ew")
        self.settings_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(self.settings_frame, text="语气风格:").grid(row=0, column=0, sticky="w")
        self.tone_var = ctk.StringVar(value="专业严谨")
        self.tone_menu = ctk.CTkOptionMenu(self.settings_frame, values=["专业严谨", "幽默风趣", "热情洋溢", "平易近人"], variable=self.tone_var, width=110)
        self.tone_menu.grid(row=1, column=0, sticky="w", pady=5)

        # 【核心优化】：将下拉菜单更换为 ComboBox（组合框），支持手动输入
        ctk.CTkLabel(self.settings_frame, text="字数(可点进去手填):").grid(row=0, column=1, sticky="w", padx=(5,0))
        self.length_var = ctk.StringVar(value="详细(约2000字)")
        self.length_menu = ctk.CTkComboBox(
            self.settings_frame, 
            values=["简短(约500字)", "适中(约1000字)", "详细(约2000字)", "长篇(约5000字)", "超长篇(约8000字)"], 
            variable=self.length_var, 
            width=135
        )
        self.length_menu.grid(row=1, column=1, sticky="w", padx=(5,0), pady=5)

        # 4. 文档生成按钮区
        self.doc_types = ["📝 学术论文", "📊 研究报告", "📅 工作计划", "💡 总结反思", "📢 演讲稿件", "📧 商业邮件"]
        
        self.btn_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.btn_frame.grid(row=5, column=0, padx=20, pady=15, sticky="ew")
        self.btn_frame.grid_columnconfigure((0, 1), weight=1)

        for i, doc in enumerate(self.doc_types):
            btn = ctk.CTkButton(self.btn_frame, text=doc, command=lambda d=doc: self.start_generation(d), fg_color="#2b6b84", hover_color="#1f5368")
            btn.grid(row=i//2, column=i%2, padx=3, pady=5, sticky="ew")

        # 停止按钮 (默认隐藏)
        self.stop_btn = ctk.CTkButton(self.sidebar, text="🛑 停止生成", fg_color="#c0392b", hover_color="#a53125", command=self.stop_generation)
        
        # 外观模式切换
        self.appearance_mode_menu = ctk.CTkOptionMenu(self.sidebar, values=["System", "Dark", "Light"], command=self.change_appearance)
        self.appearance_mode_menu.grid(row=8, column=0, padx=20, pady=(10, 20), sticky="ew")

        # ==================== 右侧编辑与导出区 ====================
        self.main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.text_area = ctk.CTkTextbox(self.main_frame, font=ctk.CTkFont(family="微软雅黑", size=14), wrap="word")
        self.text_area.grid(row=0, column=0, columnspan=3, sticky="nsew", pady=(0, 15))

        self.clear_btn = ctk.CTkButton(self.main_frame, text="🗑️ 清空面板", fg_color="gray", command=self.clear_text, width=120)
        self.clear_btn.grid(row=1, column=0, sticky="w")

        self.export_md_btn = ctk.CTkButton(self.main_frame, text="💾 导出为 Markdown", command=self.export_md, width=150)
        self.export_md_btn.grid(row=1, column=1, sticky="e", padx=(0, 10))

        self.export_word_btn = ctk.CTkButton(self.main_frame, text="📄 导出为 Word", command=self.export_word, fg_color="#27ae60", hover_color="#219653", width=150)
        self.export_word_btn.grid(row=1, column=2, sticky="e")

    def change_appearance(self, new_mode):
        ctk.set_appearance_mode(new_mode)

    def start_generation(self, doc_type):
        if self.is_generating:
            return

        api_key = self.api_key_entry.get().strip()
        topic = self.topic_textbox.get("1.0", "end").strip()

        if not api_key:
            messagebox.showerror("错误", "请先输入 DeepSeek API Key！")
            return
        if not topic or topic.startswith("例如："):
            messagebox.showerror("错误", "请具体描述一下你的文档需求！")
            return

        if len(self.text_area.get("1.0", "end").strip()) > 0:
            if not messagebox.askyesno("确认", "编辑器已有内容，是否清空并重新生成？"):
                return

        self.save_config(api_key)
        self.is_generating = True
        self.stop_flag = False
        
        self.stop_btn.grid(row=6, column=0, padx=20, pady=10, sticky="ew")
        
        self.text_area.delete("1.0", "end")
        self.text_area.insert("end", f"🚀 正在连接 DeepSeek 大模型，构思【{doc_type}】...\n\n")

        tone = self.tone_var.get()
        length = self.length_var.get() # 这里能直接获取到用户手打的任意自定义字数

        threading.Thread(target=self.call_deepseek, args=(api_key, topic, doc_type, tone, length), daemon=True).start()

    def call_deepseek(self, api_key, topic, doc_type, tone, length):
        try:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            sys_prompt = "你是一个顶级文档写作专家，精通各类公文、学术、职场和商业文档的撰写，排版结构完美。"
            
            # 【核心优化】：针对长文本专门强化的 Prompt 提示词工程
            user_prompt = f"""请帮我撰写一份【{doc_type}】。
- 核心主题/需求：{topic}
- 语气风格：{tone}
- 篇幅字数要求：严格遵循【{length}】的长度标准！
  *特别注意*：如果是长篇或超长篇，请务必通过【增加多维度的深度分析】、【提供丰富的具体案例】、【详实的数据与步骤拆解】等方式来实质性扩充篇幅！切忌车轱辘话来回凑字数，坚决不要草草收尾。
- 排版格式：使用清晰的 Markdown 格式输出，包含层级标题（#、##）。不要输出任何寒暄废话，直接给我正文内容。"""

            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                stream=True,
                max_tokens=8192 # 【核心优化】：解锁单次生成的最大 Token 限制，支持几万字的巨长文本不被截断
            )

            self.root.after(0, self.text_area.delete, "1.0", "end")

            for chunk in response:
                if self.stop_flag:
                    self.root.after(0, self.append_text, "\n\n[⚠️ 生成已被用户手动中断]")
                    break
                    
                delta = chunk.choices[0].delta.content
                if delta:
                    self.root.after(0, self.append_text, delta)

        except Exception as e:
            self.root.after(0, self.append_text, f"\n\n❌ 生成发生错误：\n{str(e)}")
        finally:
            self.root.after(0, self.finish_generation)

    def stop_generation(self):
        self.stop_flag = True

    def finish_generation(self):
        self.is_generating = False
        self.stop_btn.grid_forget()

    def append_text(self, text):
        self.text_area.insert("end", text)
        self.text_area.see("end")

    def clear_text(self):
        if messagebox.askyesno("确认", "确定要清空编辑器内容吗？"):
            self.text_area.delete("1.0", "end")

    def export_md(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown 文件", "*.md")], title="导出为 Markdown")
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.text_area.get("1.0", "end"))
            messagebox.showinfo("成功", "Markdown 文件导出成功！")

   def export_word(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word 文档", "*.docx")], title="导出为公文排版 Word")
        if not file_path: return
        
        try:
            doc = Document()
            
            # === 1. 公文格式页面设置 (A4标准) ===
            sections = doc.sections
            for section in sections:
                section.page_height = Mm(297)
                section.page_width = Mm(210)
                section.top_margin = Mm(37)
                section.bottom_margin = Mm(35)
                section.left_margin = Mm(28)
                section.right_margin = Mm(26)
            
            # === 2. 全局正文样式 (仿宋, 三号字(16磅), 1.5倍行距) ===
            style = doc.styles['Normal']
            style.font.name = '仿宋'
            style._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
            style.font.size = Pt(16)
            style.paragraph_format.line_spacing = 1.5
            style.paragraph_format.first_line_indent = Pt(32) # 首行缩进2字符(三号字=16磅，两个字=32磅)

            content = self.text_area.get("1.0", "end").strip()
            
            # === 3. 逐行解析与去 Markdown 符号处理 ===
            for line in content.split('\n'):
                line = line.strip()
                if not line:
                    continue
                
                # 解析标题级数
                heading_level = 0
                match = re.match(r'^(#+)\s*(.*)', line)
                if match:
                    heading_level = len(match.group(1))
                    line = match.group(2) # 剥离前面的 # 号
                
                # 清除行首无用的无序列表符号 (- 或 *)
                line = re.sub(r'^[\-\*]\s+', '', line)
                
                if heading_level == 1:
                    # 一级标题：黑体，三号，居中，不加粗
                    p = doc.add_paragraph()
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    p.paragraph_format.first_line_indent = 0 # 居中不缩进
                    
                    line_clean = line.replace('*', '').replace('#', '') # 暴力清除残留符号
                    run = p.add_run(line_clean)
                    run.font.name = '黑体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                    run.font.size = Pt(16)
                    run.font.bold = False
                    
                elif heading_level == 2:
                    # 二级标题：楷体，三号，加粗，不缩进 (更醒目)
                    p = doc.add_paragraph()
                    p.paragraph_format.first_line_indent = 0
                    
                    line_clean = line.replace('*', '').replace('#', '')
                    run = p.add_run(line_clean)
                    run.font.name = '楷体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
                    run.font.size = Pt(16)
                    run.font.bold = True
                    
                else:
                    # 正文或三级及以下标题：处理内联的 **粗体**
                    p = doc.add_paragraph()
                    
                    if heading_level >= 3:
                        p.paragraph_format.first_line_indent = Pt(32) # 缩进2字符
                        run = p.add_run(line.replace('*', '').replace('#', ''))
                        run.font.name = '仿宋'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                        run.font.size = Pt(16)
                        run.font.bold = True
                    else:
                        # 普通正文拆分处理加粗
                        parts = re.split(r'(\*\*.*?\*\*)', line)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                run = p.add_run(part[2:-2]) # 去除**符号
                                run.font.bold = True
                            else:
                                # 清理其余单独残留的Markdown乱码符号
                                clean_part = part.replace('*', '').replace('#', '')
                                if clean_part:
                                    run = p.add_run(clean_part)
                            
                            # 统一设定正文字体
                            run.font.name = '仿宋'
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
            
            doc.save(file_path)
            messagebox.showinfo("成功", f"✅ 公文级 Word 已成功导出，无残留符号！\n路径:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出 Word 失败:\n{str(e)}")
