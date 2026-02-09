import customtkinter as ctk
import threading
from openai import OpenAI
import os
import sys
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox
import json
import time
import re

# --- Linux 显示环境自动修正 ---
if sys.platform.startswith('linux'):
    try:
        import tkinter
        root = tkinter.Tk()
        root.destroy()
    except:
        if os.environ.get('DISPLAY','') == '':
            os.environ.__setitem__('DISPLAY', ':0')

# --- 配置区域 ---
APP_VERSION = "v27.0.0 (Dynamic Structure Adaptive)"
DEV_NAME = "俞晋全"
DEV_ORG = "甘肃省金塔县中学，甘肃金塔 735300"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 文体风格定义 (已移除死板的结构模板，改为风格指导) ===
STYLE_GUIDE = {
    "期刊论文": {
        "desc": "学术严谨，理实结合，适合发表。",
        "default_topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "default_words": "3000",
        "default_instruction": "要求：\n1. 结合具体的化学教学案例（如氯气、钠、电化学等）。\n2. 数据详实，逻辑严密。\n3. 适合《化学教育》或《中学化学教学参考》风格。",
        "writing_prompt": "语气学术、客观、务实。严禁堆砌空洞理论，必须用具体的化学知识点和教学片段来支撑观点。",
    },
    "教学反思": {
        "desc": "第一人称，深度剖析，真诚走心。",
        "default_topic": "高三化学二轮复习课后的深刻反思",
        "default_words": "2000",
        "default_instruction": "要求：\n1. 必须使用第一人称‘我’。\n2. 重点复盘课堂上的‘遗憾点’和‘生成性问题’。\n3. 剖析原因要深刻，不要找客观理由。",
        "writing_prompt": "使用第一人称。文风要诚恳、犀利。多描写课堂上的真实细节（如学生的错题、冷场的瞬间）。",
    },
    "教学案例": {
        "desc": "叙事风格，还原现场，生动具体。",
        "default_topic": "《钠与水反应》教学案例分析",
        "default_words": "2500",
        "default_instruction": "要求：\n1. 采用‘教育叙事’手法，像讲故事一样记录课堂。\n2. 还原师生对话，描写实验现象。\n3. 突出‘意外’与‘机智化解’。",
        "writing_prompt": "采用叙事风格。大量使用对话描写、动作描写。还原真实的课堂冲突和教学灵感。",
    },
    "工作计划": {
        "desc": "行政公文，条理清晰，数据导向。",
        "default_topic": "2026年春季学期高二化学备课组工作计划",
        "default_words": "2000",
        "default_instruction": "要求：\n1. 语言简练，干脆利落。\n2. 包含具体的行事历（周次安排）。\n3. 目标要量化（如平均分、及格率）。",
        "writing_prompt": "行政公文风格。多用‘一要...二要...’句式。内容必须具体可执行，包含时间节点。",
    },
    "工作总结": {
        "desc": "汇报风格，亮点突出，分析透彻。",
        "default_topic": "2025年度个人教学工作总结",
        "default_words": "3000",
        "default_instruction": "要求：\n1. 用数据说话。\n2. 既要展示成绩，也要诚恳分析不足。\n3. 结构严谨，层次分明。",
        "writing_prompt": "汇报风格。多用数据对比。对成绩要总结经验，对不足要分析原因并提出对策。",
    },
    "自由定制": {
        "desc": "完全根据指令生成。",
        "default_topic": "（在此输入题目）",
        "default_words": "1000",
        "default_instruction": "请详细描述您的要求...",
        "writing_prompt": "严格遵循用户的特殊指令，风格不限。",
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

        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=2) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        outline_frame.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(outline_frame, text="Step 1: 生成智能大纲 (可修改)", text_color="#1F6AA5", font=("bold", 13)).pack(side="left")
        
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 12)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_o_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_o_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_o_frame, text="根据题目生成大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5)
        ctk.CTkButton(btn_o_frame, text="清空", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=60).pack(side="right", padx=5)

        content_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        content_frame.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(content_frame, text="Step 2: 撰写预览 (实时流式)", text_color="#2CC985", font=("bold", 13)).pack(side="left")
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

    def on_mode_change(self, choice):
        config = STYLE_GUIDE.get(choice, STYLE_GUIDE["自由定制"])
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, config.get("default_topic", ""))
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", config.get("default_instruction", ""))
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, config.get("default_words", "3000"))
        
        self.txt_outline.delete("0.0", "end")
        self.txt_outline.insert("0.0", f"（已切换至【{choice}】模式，请点击“根据题目生成大纲”...）")

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

    # --- 生成大纲 (核心修改：语义分析) ---
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
        self.status_label.configure(text="正在分析题目并构建大纲...", text_color="#1F6AA5")
        
        # 获取风格，但不获取死板结构
        style_cfg = STYLE_GUIDE.get(mode, STYLE_GUIDE["自由定制"])
        
        # 动态 Prompt
        prompt = f"""
        你是一位资深的高中化学教育专家。
        
        【任务】：请为《{topic}》这篇文章设计一份独特的、逻辑严密的大纲。
        【文体类型】：{mode}
        【风格参考】：{style_cfg['desc']}
        【用户要求】：{instr}
        
        【重要指令 - 拒绝千篇一律】：
        1. 请不要套用通用模板（如“一、背景；二、意义”）。
        2. 请深入分析题目《{topic}》的具体的含义。
           - 如果题目是实验研究，大纲要包含“实验设计、数据分析”。
           - 如果题目是复习策略，大纲要包含“学情诊断、模型建构”。
           - 如果题目是工作总结，大纲要包含具体的“成绩数据、反思改进”。
        3. 必须包含一级标题（如一、二、三）和二级标题（如（一）（二））。
        4. 不要包含Markdown符号。
        5. 直接输出大纲内容，不要前言后语。
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
            self.status_label.configure(text="大纲已生成，请检查并修改。", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal")

    # --- 撰写全文 ---
    def run_full_write(self):
        self.stop_event.clear()
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 5:
            self.status_label.configure(text="请先生成或输入大纲", text_color="red")
            return
            
        lines = [l.strip() for l in outline_raw.split('\n') if l.strip()]
        
        # 智能滤除标题行
        if len(lines) > 0:
            first_line = lines[0]
            topic = self.entry_topic.get().strip()
            if len(topic) > 2 and topic[:4] in first_line:
                lines = lines[1:]

        tasks = []
        current_task = []
        for line in lines:
            is_header = False
            # 兼容多种标题格式
            if re.match(r'^[一二三四五六七八九十]+、', line): is_header = True
            if re.match(r'^第[一二三四五六七八九十]+部分', line): is_header = True
            if "摘要" in line or "参考文献" in line: is_header = True
            
            if is_header:
                if current_task: tasks.append(current_task)
                current_task = [line]
            else:
                current_task.append(line)
        if current_task: tasks.append(current_task)

        if not tasks:
            self.status_label.configure(text="大纲格式无法识别（需包含'一、'或'摘要'）", text_color="red")
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
                5. 严格执行字数限制，不要大幅超写或少写。
                """
                
                user_prompt = f"""
                题目：{topic}
                当前章节：{header}
                包含要点：
                {sub_points}
                
                上下文：...{last_paragraph[-150:]}
                
                字数：约 {current_limit} 字。
                请直接输出正文。
                """

                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[{"role":"system","content":sys_prompt}, {"role":"user","content":user_prompt}],
                    temperature=0.7,
                    stream=True
                )
                
                current_section_text = ""
                
                # 强力标题清洗
                def get_core(text): return re.sub(r'[^\u4e00-\u9fa50-9]', '', text)

                for chunk in resp:
                    if self.stop_event.is_set(): break
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        temp_text = current_section_text + content
                        
                        if "摘要" in header:
                            if len(temp_text) < 10 and ("摘" in temp_text or "要" in temp_text):
                                current_section_text += content
                                continue 
                            clean_chunk = re.sub(r'^【?摘要】?[:：]?\s*', '', content)
                            self.txt_content.insert("end", clean_chunk)
                        else:
                            header_core = get_core(header)
                            if len(temp_text) < 50 and (header_core in get_core(temp_text)):
                                current_section_text += content
                            else:
                                if len(current_section_text) > 0 and len(current_section_text) < 50:
                                    if header_core in get_core(current_section_text):
                                        parts = current_section_text.split('\n', 1)
                                        if len(parts) > 1:
                                            self.txt_content.insert("end", parts[1] + content)
                                        else:
                                            self.txt_content.insert("end", content)
                                    else:
                                        self.txt_content.insert("end", current_section_text + content)
                                    current_section_text = "SAFE" 
                                else:
                                    self.txt_content.insert("end", content)
                        
                        self.txt_content.see("end")
                        if len(temp_text) > 50: last_paragraph = temp_text

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
