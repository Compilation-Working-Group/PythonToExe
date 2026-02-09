import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox
import json
import time
import re

# --- 配置区域 ---
APP_VERSION = "v18.0.0 (Journal Standard + Strict Word Control)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室" # 根据您上传的文件自动定制
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 动态预设库 (已根据您上传的范文优化) ===
PRESET_CONFIGS = {
    "期刊论文 (标准版)": {
        "topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "instruction": "严格参照《化学教育》或《中学化学教学参考》的风格。\n1. 必须包含：摘要、关键词、一、问题的提出；二、核心概念；三、实践策略；四、成效反思；参考文献。\n2. 策略部分必须结合具体案例（如氯气）。",
        "words": "3000",
        "structure_mode": "journal" 
    },
    "教学反思": {
        "topic": "高三化学二轮复习课后的深刻反思",
        "instruction": "第一人称。剖析真实问题（如学生对复杂情境应用吃力）。\n结构：教学初衷 -> 课堂实录 -> 原因剖析 -> 改进措施。",
        "words": "1500",
        "structure_mode": "general"
    },
    "教学案例": {
        "topic": "《钠与水反应》教学案例分析",
        "instruction": "叙事风格。还原师生对话。重点描写课堂冲突和意外生成。",
        "words": "2500",
        "structure_mode": "general"
    },
    "工作计划": {
        "topic": "2026年春季学期高二化学备课组工作计划",
        "instruction": "行政公文风。条理清晰，多用数据。",
        "words": "2000",
        "structure_mode": "general"
    },
    "自由定制": {
        "topic": "（在此输入主题）",
        "instruction": "请详细描述要求。",
        "words": "1000",
        "structure_mode": "general"
    }
}

class JournalWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"期刊论文精准撰写系统 - {DEV_NAME}")
        self.geometry("1280x900")
        
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

        self.status_label = ctk.CTkLabel(self, text="就绪 - 请先生成大纲", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(6, weight=1)

        # 第一行：基础信息
        ctk.CTkLabel(t, text="文体类型:", font=("bold", 12)).grid(row=0, column=0, padx=10, sticky="e")
        self.combo_mode = ctk.CTkComboBox(t, values=list(PRESET_CONFIGS.keys()), width=250, command=self.on_mode_change)
        self.combo_mode.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(t, text="论文题目:", font=("bold", 12)).grid(row=1, column=0, padx=10, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 第二行：作者信息 (用于范文格式)
        ctk.CTkLabel(t, text="作者/单位:", font=("bold", 12)).grid(row=2, column=0, padx=10, sticky="e")
        info_frame = ctk.CTkFrame(t, fg_color="transparent")
        info_frame.grid(row=2, column=1, sticky="w")
        self.entry_author = ctk.CTkEntry(info_frame, width=150, placeholder_text="作者名")
        self.entry_author.insert(0, DEV_NAME)
        self.entry_author.pack(side="left", padx=(10, 5))
        self.entry_org = ctk.CTkEntry(info_frame, width=340, placeholder_text="单位信息")
        self.entry_org.insert(0, "甘肃省金塔县中学")
        self.entry_org.pack(side="left", padx=5)

        # 第三行：控制参数
        ctk.CTkLabel(t, text="指令要求:", font=("bold", 12)).grid(row=3, column=0, padx=10, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=50, font=("Arial", 12))
        self.txt_instructions.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(t, text="严格控字:", font=("bold", 12)).grid(row=4, column=0, padx=10, sticky="e")
        self.entry_words = ctk.CTkEntry(t, width=150)
        self.entry_words.grid(row=4, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(t, text="(系统将根据此字数自动分配各章节篇幅，误差控制在±15%)", text_color="gray", font=("Arial", 10)).grid(row=4, column=1, padx=170, sticky="w")

        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=5, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 双面板
        self.paned = ctk.CTkFrame(t, fg_color="transparent")
        self.paned.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned.grid_columnconfigure(0, weight=1)
        self.paned.grid_columnconfigure(1, weight=3) # 右侧宽一些
        self.paned.grid_rowconfigure(1, weight=1)

        # 左侧：大纲
        ctk.CTkLabel(self.paned, text="Step 1: 结构大纲 (可手动调整)", font=("bold", 12), text_color="#1F6AA5").grid(row=0, column=0, sticky="w", padx=5)
        self.txt_outline = ctk.CTkTextbox(self.paned, font=("Arial", 13))
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_f1 = ctk.CTkFrame(self.paned, fg_color="transparent")
        btn_f1.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_f1, text="生成标准大纲", command=self.run_gen_outline, width=120)
        self.btn_gen_outline.pack(side="left", padx=5)

        # 右侧：正文
        ctk.CTkLabel(self.paned, text="Step 2: 正文预览 (自动清洗格式)", font=("bold", 12), text_color="#2CC985").grid(row=0, column=1, sticky="w", padx=5)
        self.txt_content = ctk.CTkTextbox(self.paned, font=("Arial", 14))
        self.txt_content.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        btn_f2 = ctk.CTkFrame(self.paned, fg_color="transparent")
        btn_f2.grid(row=2, column=1, sticky="ew")
        self.btn_write = ctk.CTkButton(btn_f2, text="按大纲精准撰写", command=self.run_full_write, fg_color="#2CC985", font=("bold", 13))
        self.btn_write.pack(side="left", padx=5)
        self.btn_stop = ctk.CTkButton(btn_f2, text="停止", command=self.stop_writing, fg_color="#C0392B", width=60)
        self.btn_stop.pack(side="left", padx=5)
        self.btn_export = ctk.CTkButton(btn_f2, text="导出期刊格式Word", command=self.save_to_word, width=150)
        self.btn_export.pack(side="right", padx=5)

        self.on_mode_change("期刊论文 (标准版)")

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

    # --- 逻辑 ---

    def on_mode_change(self, choice):
        preset = PRESET_CONFIGS.get(choice, PRESET_CONFIGS["期刊论文 (标准版)"])
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, preset["topic"])
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", preset["instruction"])
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, preset["words"])

    def stop_writing(self):
        self.stop_event.set()
        self.status_label.configure(text="已停止", text_color="red")

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_gen_outline(self):
        self.stop_event.clear()
        threading.Thread(target=self.thread_outline, daemon=True).start()

    def thread_outline(self):
        client = self.get_client()
        if not client: return
        
        mode = self.combo_mode.get()
        topic = self.entry_topic.get()
        instr = self.txt_instructions.get("0.0", "end").strip()

        self.btn_gen_outline.configure(state="disabled")
        self.status_label.configure(text="正在生成结构大纲...", text_color="#1F6AA5")

        # 针对期刊论文的强制结构 Prompt
        structure_req = ""
        if "期刊论文" in mode:
            structure_req = """
            【期刊论文强制结构】：
            摘要
            关键词
            一、问题的提出 (研究背景与意义)
            二、(理论框架或核心概念)
            三、(教学策略或实践路径，需包含3个小点)
            四、成效与反思
            参考文献
            """
        else:
            structure_req = "结构：按照标准公文或教学文档结构，列出一级标题（一、二...）。"

        prompt = f"""
        任务：为《{topic}》设计大纲。
        文体：{mode}
        用户指令：{instr}
        
        {structure_req}
        
        要求：
        1. 直接输出标题列表，不要Markdown。
        2. 不要任何解释性文字。
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
                    self.txt_outline.insert("end", chunk.choices[0].delta.content)
            self.status_label.configure(text="大纲已生成，请检查", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal")

    def run_full_write(self):
        self.stop_event.clear()
        threading.Thread(target=self.thread_write, daemon=True).start()

    def thread_write(self):
        client = self.get_client()
        if not client: return

        # 1. 解析大纲：只提取“一级标题”
        # 以前我们提取所有标题，导致字数爆炸。现在我们只提取 "一、" "二、" 这种大块。
        outline_text = self.txt_outline.get("0.0", "end").strip()
        lines = [l.strip() for l in outline_text.split('\n') if l.strip()]
        
        # 智能分组：将大纲分为若干个 "写作任务块"
        # 摘要、关键词算一块；每个 "一、" 算一块；参考文献算一块。
        tasks = []
        current_task = []
        
        for line in lines:
            # 识别新板块的标志
            is_new_block = False
            if line.startswith("一、") or line.startswith("二、") or line.startswith("三、") or line.startswith("四、") or line.startswith("五、"):
                is_new_block = True
            elif "摘要" in line or "参考文献" in line:
                is_new_block = True
            
            if is_new_block:
                if current_task: tasks.append(current_task)
                current_task = [line]
            else:
                current_task.append(line)
        if current_task: tasks.append(current_task)

        # 2. 字数分配算法
        try: total_words = int(self.entry_words.get())
        except: total_words = 3000
        
        # 预留摘要(300)、结语(300)、参考文献(0)
        # 剩下的字数分给中间的正文核心板块
        core_words = total_words - 600
        if core_words < 500: core_words = 500
        
        # 计算核心板块数量
        core_tasks_count = 0
        for t in tasks:
            header = t[0]
            if any(x in header for x in ["一、", "二、", "三、", "四、", "五、"]):
                core_tasks_count += 1
        
        avg_core_words = int(core_words / (core_tasks_count if core_tasks_count > 0 else 1))

        self.btn_write.configure(state="disabled")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)

        try:
            for i, task_lines in enumerate(tasks):
                if self.stop_event.is_set(): break
                
                header = task_lines[0] # 该块的主标题
                sub_points = "\n".join(task_lines[1:]) # 该块下的小点
                
                # 确定本块字数
                target_len = 300 # 默认
                if "摘要" in header: target_len = 300
                elif "参考文献" in header: target_len = 0
                elif any(x in header for x in ["一、", "二、", "三、", "四、"]): target_len = avg_words
                
                # 重点章节加量 (通常 "三、策略" 是重点)
                if "三、" in header or "策略" in header or "实践" in header:
                    target_len = int(target_len * 1.5)

                self.status_label.configure(text=f"正在撰写: {header} (目标 {target_len} 字)...", text_color="#1F6AA5")
                self.progressbar.set((i) / len(tasks))

                # 插入标题 (仅插入主标题)
                self.txt_content.insert("end", f"\n【{header}】\n")
                self.txt_content.see("end")

                # Prompt
                sys_prompt = f"""
                你是一位高中化学名师。当前任务：撰写章节【{header}】。
                文体：{self.combo_mode.get()}
                
                【绝对指令】：
                1. 严禁复述标题。直接写正文。
                2. 严禁Markdown。
                3. 本章节目标字数：{target_len} 字左右（误差不超过20%）。请严格控制篇幅，既不要太短，也不要写成几千字的长文。
                4. 严格遵守用户指令：{self.txt_instructions.get("0.0", "end").strip()}
                """
                
                user_prompt = f"""
                题目：{self.entry_topic.get()}
                当前章节：{header}
                该章节包含的要点（请将这些要点融合成连贯的文章，不要列条目）：
                {sub_points}
                
                写作提示：
                - 如果是“摘要”，请写成一段流畅的短文。
                - 如果是“正文”，请结合具体化学案例（如氯气、钠），多用数据。
                - 必须控制字数在 {target_len} 左右！
                """

                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[{"role":"system","content":sys_prompt}, {"role":"user","content":user_prompt}],
                    temperature=0.7
                )
                
                raw = resp.choices[0].message.content
                
                # 清洗：去掉开头可能重复的标题
                clean = raw.strip()
                # 简单清洗逻辑：如果前20个字里包含了标题的核心词，就去掉那一行
                header_core = re.sub(r'[一二三四五、\d\.]', '', header).strip()
                if len(clean) > len(header_core) and header_core in clean[:len(header_core)+10]:
                    # 尝试按换行符切分，丢弃第一行
                    parts = clean.split('\n', 1)
                    if len(parts) > 1: clean = parts[1].strip()

                self.txt_content.insert("end", clean + "\n")
                self.txt_content.see("end")
                time.sleep(1)

            self.progressbar.set(1)
            self.status_label.configure(text="撰写完成！", text_color="green")

        except Exception as e:
            self.status_label.configure(text=str(e), text_color="red")
        finally:
            self.btn_write.configure(state="normal")

    def save_to_word(self):
        content = self.txt_content.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            # 字体设置
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 1. 标题 (黑体, 二号, 居中)
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_t = p_title.add_run(self.entry_topic.get())
            run_t.font.name = u'黑体'
            run_t._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run_t.font.size = Pt(18)
            run_t.bold = True
            
            # 2. 作者与单位 (楷体, 小四, 居中)
            p_info = doc.add_paragraph()
            p_info.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            info_text = f"{self.entry_author.get()}\n({self.entry_org.get()})"
            run_i = p_info.add_run(info_text)
            run_i.font.name = u'楷体'
            run_i._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
            run_i.font.size = Pt(12)
            
            doc.add_paragraph() # 空行

            # 3. 正文处理
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                
                # 识别标题 【XXX】
                if line.startswith("【") and line.endswith("】"):
                    header_text = line.replace("【", "").replace("】", "")
                    
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(12)
                    run = p.add_run(header_text)
                    
                    # 摘要和关键词特殊处理
                    if "摘要" in header_text or "关键词" in header_text:
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                        run.bold = True
                    # 一级标题 (一、)
                    elif re.match(r'^[一二三四五六七八九十]+、', header_text):
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                        run.font.size = Pt(14)
                        run.bold = True
                    # 其他标题
                    else:
                        run.bold = True
                else:
                    # 正文段落
                    p = doc.add_paragraph(line)
                    p.paragraph_format.first_line_indent = Pt(24) # 首行缩进2字符
                    p.paragraph_format.line_spacing = 1.5 # 1.5倍行距

            doc.save(file_path)
            self.status_label.configure(text=f"已导出范文格式: {os.path.basename(file_path)}", text_color="green")

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
    app = JournalWriterApp()
    app.mainloop()
