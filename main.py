import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox
import json
import time
import re

# --- 配置区域 ---
APP_VERSION = "v22.0.0 (Smart Header Filter + Safe Cleaning)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 文体风格定义 ===
STYLE_GUIDE = {
    "期刊论文": {
        "desc": "参照《虚拟仿真》、《热重分析》等范文。学术严谨，理实结合。",
        "outline_prompt": "请设计一份标准的教育期刊论文大纲。必须包含：摘要、关键词、一、问题的提出；二、核心概念/理论；三、教学策略/模型建构（核心）；四、成效与反思；参考文献。",
        "writing_prompt": "语气要学术、客观。策略部分必须结合具体的化学知识点（如氯气、氧化还原）。多用数据和案例支撑。",
    },
    "教学反思": {
        "desc": "参照《二轮复习反思》。第一人称，深度剖析。",
        "outline_prompt": "请设计一份深度教学反思大纲。建议结构：一、教学初衷；二、课堂实录与问题；三、原因深度剖析；四、改进措施。",
        "writing_prompt": "使用第一人称‘我’。拒绝套话，重点描写课堂上真实的遗憾、突发状况和学生的真实反应。剖析要深刻。",
    },
    "教学案例": {
        "desc": "叙事风格，还原课堂现场。",
        "outline_prompt": "请设计一份教学案例大纲。建议结构：一、案例背景；二、情境描述（片段）；三、案例分析；四、教学启示。",
        "writing_prompt": "采用‘叙事研究’风格。像写故事一样描述课堂冲突、师生对话和实验现象。",
    },
    "工作计划": {
        "desc": "行政公文风格，条理清晰。",
        "outline_prompt": "请设计一份工作计划大纲。包含：指导思想、工作目标、主要措施、行事历。",
        "writing_prompt": "语言简练，多用‘一要...二要...’的句式。措施要具体，多用数据。",
    },
    "工作总结": {
        "desc": "汇报风格，数据详实。",
        "outline_prompt": "请设计一份工作总结大纲。包含：工作概况、主要成绩、存在不足、未来展望。",
        "writing_prompt": "用数据说话（平均分、获奖数）。既要展示亮点，也要诚恳分析不足。",
    },
    "自由定制": {
        "desc": "根据指令自动生成。",
        "outline_prompt": "请根据用户的具体指令设计最合理的大纲结构。",
        "writing_prompt": "严格遵循用户的特殊要求。",
    }
}

class MasterWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能写作系统 - {APP_VERSION}")
        self.geometry("1300x900")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()
        self.stop_event = threading.Event()

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_write = self.tabview.add("写作工作台")
        self.tab_settings = self.tabview.add("系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(5, weight=1) 

        # --- 顶部控制区 ---
        ctrl_frame = ctk.CTkFrame(t, fg_color="transparent")
        ctrl_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(ctrl_frame, text="文体类型:", font=("bold", 14)).pack(side="left", padx=5)
        self.combo_mode = ctk.CTkComboBox(ctrl_frame, values=list(STYLE_GUIDE.keys()), width=180, command=self.on_mode_change)
        self.combo_mode.set("期刊论文")
        self.combo_mode.pack(side="left", padx=5)
        
        ctk.CTkLabel(ctrl_frame, text="目标字数:", font=("bold", 14)).pack(side="left", padx=(20, 5))
        self.entry_words = ctk.CTkEntry(ctrl_frame, width=100)
        self.entry_words.insert(0, "3000")
        self.entry_words.pack(side="left", padx=5)

        ctk.CTkLabel(t, text="文章标题:", font=("bold", 12)).grid(row=1, column=0, padx=10, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=600)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(t, text="具体指令:", font=("bold", 12)).grid(row=2, column=0, padx=10, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=50, font=("Arial", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # --- 核心双面板区 ---
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=2) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧：大纲
        outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        outline_frame.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(outline_frame, text="Step 1: 生成并修改大纲", text_color="#1F6AA5", font=("bold", 13)).pack(side="left")
        
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 12)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_o_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_o_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_o_frame, text="生成/重置大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5)
        ctk.CTkButton(btn_o_frame, text="清空", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=60).pack(side="right", padx=5)

        # 右侧：正文
        content_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        content_frame.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(content_frame, text="Step 2: 按大纲撰写全文", text_color="#2CC985", font=("bold", 13)).pack(side="left")
        self.status_label = ctk.CTkLabel(content_frame, text="就绪", text_color="gray")
        self.status_label.pack(side="right")

        self.txt_content = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 14))
        self.txt_content.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        btn_w_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_w_frame.grid(row=2, column=1, sticky="ew")
        self.btn_run_write = ctk.CTkButton(btn_w_frame, text="开始撰写全文", command=self.run_full_write, fg_color="#2CC985", font=("bold", 14))
        self.btn_run_write.pack(side="left", padx=5)
        self.btn_stop = ctk.CTkButton(btn_w_frame, text="🔴 停止", command=self.stop_writing, fg_color="#C0392B", width=80)
        self.btn_stop.pack(side="left", padx=5)
        self.btn_clear_all = ctk.CTkButton(btn_w_frame, text="🧹 清空", command=self.clear_all, fg_color="gray", width=80)
        self.btn_clear_all.pack(side="right", padx=5)
        self.btn_export = ctk.CTkButton(btn_w_frame, text="导出 Word", command=self.save_to_word, width=120)
        self.btn_export.pack(side="right", padx=5)

        self.progressbar = ctk.CTkProgressBar(t, mode="determinate", height=2)
        self.progressbar.grid(row=6, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self.progressbar.set(0)

        self.on_mode_change("期刊论文")

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

    # --- 逻辑控制 ---

    def on_mode_change(self, choice):
        if choice == "期刊论文":
            self.entry_topic.delete(0, "end")
            self.entry_topic.insert(0, "高中化学虚拟仿真实验教学的价值与策略研究")
            self.txt_instructions.delete("0.0", "end")
            self.txt_instructions.insert("0.0", "参照《氯气》和《热重》范文风格。内容要扎实，多举例。")
            self.entry_words.delete(0, "end")
            self.entry_words.insert(0, "3000")
        elif choice == "教学反思":
            self.entry_topic.delete(0, "end")
            self.entry_topic.insert(0, "高三化学二轮复习课后的深刻反思")
            self.entry_words.delete(0, "end")
            self.entry_words.insert(0, "2000")
        self.txt_outline.delete("0.0", "end")
        self.txt_outline.insert("0.0", f"（请点击“生成大纲”按钮，AI将为您规划【{choice}】的结构...）")

    def stop_writing(self):
        self.stop_event.set()
        self.status_label.configure(text="已停止", text_color="red")

    def clear_all(self):
        self.txt_outline.delete("0.0", "end")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)
        self.status_label.configure(text="已清空")

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    # --- 生成大纲 ---
    def run_gen_outline(self):
        self.stop_event.clear()
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        if not topic:
            self.status_label.configure(text="请输入标题！", text_color="red")
            return
        threading.Thread(target=self.thread_outline, args=(mode, topic, instr), daemon=True).start()

    def thread_outline(self, mode, topic, instr):
        client = self.get_client()
        if not client: return
        self.btn_gen_outline.configure(state="disabled")
        self.status_label.configure(text="正在规划结构...", text_color="#1F6AA5")
        
        style_cfg = STYLE_GUIDE.get(mode, STYLE_GUIDE["自由定制"])
        
        prompt = f"""
        任务：为《{topic}》写一份【{mode}】的详细大纲。
        【参考风格】：{style_cfg['desc']}
        【结构建议】：{style_cfg['outline_prompt']}
        【用户指令】：{instr}
        【要求】：
        1. 必须包含一级标题（如一、二、三）和二级标题（如（一）（二））。
        2. 不要包含Markdown符号。
        3. 直接输出大纲，不要废话。
        """
        try:
            resp = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True
            )
            self.txt_outline.delete("0.0", "end")
            for chunk in resp:
                if self.stop_event.is_set(): break
                if chunk.choices[0].delta.content:
                    c = chunk.choices[0].delta.content
                    self.txt_outline.insert("end", c)
                    self.txt_outline.see("end")
            self.status_label.configure(text="大纲已生成，请手动修改。", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal")

    # --- 撰写全文 (关键修复逻辑) ---
    def run_full_write(self):
        self.stop_event.clear()
        
        # 1. 获取大纲
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 5:
            self.status_label.configure(text="请先生成或输入大纲", text_color="red")
            return
            
        lines = [l.strip() for l in outline_raw.split('\n') if l.strip()]
        
        # === 关键修复：智能滤除标题行 ===
        # 如果大纲第一行和文章标题高度相似（超过50%匹配），则认为是标题行，删除它
        if len(lines) > 0:
            first_line = lines[0]
            topic = self.entry_topic.get().strip()
            # 简单模糊匹配：如果标题核心词在第一行里
            if len(topic) > 2 and topic[:4] in first_line:
                print(f"检测到标题行，已自动过滤: {first_line}")
                lines = lines[1:]

        # 2. 智能切分大纲
        tasks = []
        current_task = []
        for line in lines:
            is_header = False
            # 兼容多种大纲格式
            if re.match(r'^[一二三四五六七八九十]+、', line): is_header = True
            if "摘要" in line or "参考文献" in line: is_header = True
            
            if is_header:
                if current_task: tasks.append(current_task)
                current_task = [line]
            else:
                current_task.append(line)
        if current_task: tasks.append(current_task)

        # 确保有任务
        if not tasks:
            self.status_label.configure(text="大纲格式无法识别，请确保包含'一、'或'摘要'", text_color="red")
            return

        topic = self.entry_topic.get()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        try: total_words = int(self.entry_words.get())
        except: total_words = 3000
        
        threading.Thread(target=self.thread_write, args=(tasks, mode, topic, instr, total_words), daemon=True).start()

    def thread_write(self, tasks, mode, topic, instr, total_words):
        client = self.get_client()
        if not client: return

        self.btn_run_write.configure(state="disabled")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)
        
        style_cfg = STYLE_GUIDE.get(mode, STYLE_GUIDE["自由定制"])
        
        # 字数计算
        core_tasks = [t for t in tasks if "摘要" not in t[0] and "参考文献" not in t[0]]
        core_count = len(core_tasks) if len(core_tasks) > 0 else 1
        
        reserved_words = 0
        if any("摘要" in t[0] for t in tasks): reserved_words += 300
        
        available_words = total_words - reserved_words
        if available_words < 500: available_words = 500
        avg_core_words = available_words // core_count

        last_paragraph = "（文章刚开始，暂无上文）"

        try:
            for i, task_lines in enumerate(tasks):
                if self.stop_event.is_set(): break
                
                header = task_lines[0]
                sub_points = "\n".join(task_lines[1:])
                
                # 动态参数
                current_limit = avg_core_words
                prompt_suffix = ""
                
                if "摘要" in header: 
                    current_limit = 300
                    prompt_suffix = "【特殊要求】：必须在摘要下方另起一行，列出3-5个【关键词】。"
                elif "参考文献" in header: 
                    current_limit = 0
                elif any(x in header for x in ["一、", "引言", "结语"]): 
                    current_limit = int(avg_core_words * 0.6)
                else:
                    current_limit = int(avg_core_words * 1.2)
                
                self.status_label.configure(text=f"撰写: {header}...", text_color="#1F6AA5")
                self.progressbar.set(i / len(tasks))

                self.txt_content.insert("end", f"\n\n【{header}】\n")
                self.txt_content.see("end")

                sys_prompt = f"""
                你是一位资深教育专家。
                文体：{mode}
                风格要求：{style_cfg['writing_prompt']}
                
                【写作铁律】：
                1. 严禁复述章节标题！(系统已自动插入)。
                2. 严禁Markdown格式。
                3. 内容务实，拒绝空洞套话。
                4. {prompt_suffix}
                """
                
                user_prompt = f"""
                题目：{topic}
                当前章节：{header}
                要点：{sub_points}
                
                上下文：...{last_paragraph[-150:]}
                
                字数：约 {current_limit} 字。
                请直接输出正文。
                """

                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[{"role":"system","content":sys_prompt}, {"role":"user","content":user_prompt}],
                    temperature=0.7
                )
                
                raw = resp.choices[0].message.content
                clean_text = raw.strip()
                
                # === 关键修复：温柔清洗算法 ===
                # 如果是摘要，只去掉"摘要："前缀，保留内容
                if "摘要" in header:
                    clean_text = re.sub(r'^【?摘要】?[:：]?\s*', '', clean_text)
                else:
                    # 其他章节，如果第一行和标题重复，则删除第一行
                    lines_content = clean_text.split('\n')
                    if len(lines_content) > 0:
                        # 提取标题核心词（去掉一、二、）
                        header_core = re.sub(r'^[一二三四五六七八九十]+、', '', header)
                        if header_core in lines_content[0]:
                            clean_text = "\n".join(lines_content[1:]).strip()

                self.txt_content.insert("end", clean_text)
                self.txt_content.see("end")
                
                if len(clean_text) > 50: last_paragraph = clean_text
                time.sleep(0.5)

            if not self.stop_event.is_set():
                self.status_label.configure(text="撰写完成！", text_color="green")
                self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"API错误: {str(e)}", text_color="red")
        finally:
            self.btn_run_write.configure(state="normal")

    def save_to_word(self):
        content = self.txt_content.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_t = p_title.add_run(self.entry_topic.get())
            run_t.font.name = u'黑体'
            run_t._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run_t.font.size = Pt(18)
            run_t.bold = True
            
            p_auth = doc.add_paragraph()
            p_auth.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_a = p_auth.add_run(f"{DEV_NAME}\n({DEV_ORG})")
            run_a.font.name = u'楷体'
            run_a._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
            run_a.font.size = Pt(12)
            
            doc.add_paragraph() 

            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue

                if line.startswith("【") and line.endswith("】"):
                    header = line.replace("【", "").replace("】", "")
                    
                    if "摘要" in header or "关键词" in header:
                        p = doc.add_paragraph()
                        run = p.add_run(header)
                        run.bold = True
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    elif re.match(r'^[一二三四五六七八九十]+、', header):
                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(12)
                        run = p.add_run(header)
                        run.bold = True
                        run.font.size = Pt(14)
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    else:
                        p = doc.add_paragraph(header)
                        p.runs[0].bold = True
                else:
                    p = doc.add_paragraph(line)
                    p.paragraph_format.first_line_indent = Pt(24) 
                    p.paragraph_format.line_spacing = 1.25

            doc.save(file_path)
            self.status_label.configure(text=f"已导出: {os.path.basename(file_path)}", text_color="green")

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
    app = MasterWriterApp()
    app.mainloop()
