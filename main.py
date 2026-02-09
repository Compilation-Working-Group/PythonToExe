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
APP_VERSION = "v19.0.0 (Yu Style Custom Edition)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 核心：基于您上传文档的深度定制模版 ===
# 这里把您的三类文稿结构“写死”在代码里，确保 AI 不乱发挥
YU_TEMPLATES = {
    "教研论文 (参照《虚拟仿真/真实情境》)": [
        {"title": "摘要与关键词", "prompt": "请模仿《高中化学虚拟仿真实验教学的价值与策略研究》的摘要风格。写一段300字左右的摘要，概括研究背景、方法（如案例分析）、核心策略及成效。接着列出3-5个关键词。"},
        {"title": "一、问题的提出", "prompt": "请模仿范文的第一部分。从高中化学教学的实际痛点切入（如传统实验的危险性、微观概念的抽象性）。引用一两个具体的教学场景作为引子。"},
        {"title": "二、核心概念与教学价值", "prompt": "请模仿范文的第二部分。阐述本研究主题（如虚拟仿真/真实情境）对化学核心素养培养的具体价值。要结合具体知识点（如氧化还原、离子反应）进行分析，不要空谈理论。"},
        {"title": "三、教学策略与实践", "prompt": "【这是全文重点，占总字数40%】。请模仿范文第三部分。提出3个具体的教学策略（策略一、策略二、策略三）。\n要求：每个策略必须结合一个具体的化学教学案例（如氯气制备、钠的性质），详细描述教学过程、师生互动和设计意图。"},
        {"title": "四、成效与反思", "prompt": "请模仿范文第四部分。写出教学成效（最好有对比数据，如及格率提升）和存在的不足（技术瓶颈、学生依赖等）。"},
        {"title": "参考文献", "prompt": "列出5-8条参考文献，格式符合GB/T 7714。"}
    ],
    "解题指导 (参照《热重图像分析》)": [
        {"title": "摘要", "prompt": "简述本类题型在高考中的地位及解题模型建构的重要性。"},
        {"title": "一、模型一：依据质量/温度变化的分析", "prompt": "请设计一个典型例题（关于晶体热分解），然后给出【解析】和【解题建模】（总结规律）。"},
        {"title": "二、模型二：依据残留率/损失率的分析", "prompt": "请设计一个关于固体残留率的例题，给出【解析】和【解题建模】。"},
        {"title": "三、模型三：氧化还原型图像分析", "prompt": "针对难点（氧化还原热重），设计例题，并重点讲解如何判断氧化剂/还原剂介入。给出【解题建模】。"},
        {"title": "四、结语", "prompt": "总结此类题型的备考建议。"}
    ],
    "深度反思 (参照《二轮复习反思》)": [
        {"title": "引言", "prompt": "简述本次教学/复习的背景、目标以及整体的课堂反馈。"},
        {"title": "一、教学初衷与设计思路", "prompt": "包含三个小点：(一)核心目标定位；(二)基于学情的考量；(三)期望达成的素养。请用第一人称‘我’，写出备课时的真实想法。"},
        {"title": "二、课堂教学的实际效果", "prompt": "包含三个小点：(一)学生反馈分析；(二)预设与生成的对比；(三)亮点与收获。要写出课堂上真实的落差感。"},
        {"title": "三、暴露出的不足与根源", "prompt": "包含三个小点：(一)知识整合的断层；(二)策略实施的乏力；(三)对临界生的疏漏。这是反思的核心，要深刻、犀利地剖析自己。"},
        {"title": "四、改进思路与具体措施", "prompt": "包含三个小点：(一)重构专题；(二)细化分层；(三)优化讲练。措施要具体可行。"}
    ],
    "自由定制": [] # 特殊处理
}

class YuWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"俞晋全名师工作室专用写作系统 - {APP_VERSION}")
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
        
        self.tab_write = self.tabview.add("文稿撰写")
        self.tab_settings = self.tabview.add("设置")

        self.setup_write_tab()
        self.setup_settings_tab()

    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(5, weight=1) 

        # 1. 文体选择
        ctk.CTkLabel(t, text="文体模版:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.combo_mode = ctk.CTkComboBox(t, values=list(YU_TEMPLATES.keys()), width=300, command=self.on_mode_change)
        self.combo_mode.set("教研论文 (参照《虚拟仿真/真实情境》)")
        self.combo_mode.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # 2. 标题
        ctk.CTkLabel(t, text="文章标题:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 3. 指令
        ctk.CTkLabel(t, text="补充指令:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=60, font=("Microsoft YaHei UI", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.txt_instructions.insert("0.0", "例如：重点结合《氯气》或者《钠》的教学案例；字数严格控制在3000字以内。")

        # 4. 字数估算显示
        ctk.CTkLabel(t, text="预估字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.lbl_word_count = ctk.CTkLabel(t, text="约 3500 字 (由模版自动控制)", text_color="gray", anchor="w")
        self.lbl_word_count.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 5. 主工作区
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned_frame.grid_columnconfigure(0, weight=1) 
        self.paned_frame.grid_columnconfigure(1, weight=3) 
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧：结构预览
        ctk.CTkLabel(self.paned_frame, text="结构预览 (AI将按此顺序撰写)", text_color="#1F6AA5", font=("bold", 12)).grid(row=0, column=0, sticky="w", padx=5)
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 12)) 
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        # 右侧：正文
        ctk.CTkLabel(self.paned_frame, text="正文生成区", text_color="#2CC985", font=("bold", 12)).grid(row=0, column=1, sticky="w", padx=5)
        self.txt_content = ctk.CTkTextbox(self.paned_frame, font=("Microsoft YaHei UI", 14))
        self.txt_content.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        # 按钮区
        btn_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        
        self.btn_run = ctk.CTkButton(btn_frame, text="开始撰写", command=self.run_writing, fg_color="#1F6AA5", font=("bold", 14), width=150)
        self.btn_run.pack(side="left", padx=5)
        
        self.btn_stop = ctk.CTkButton(btn_frame, text="停止", command=self.stop_writing, fg_color="#C0392B", width=80)
        self.btn_stop.pack(side="left", padx=5)

        self.btn_clear = ctk.CTkButton(btn_frame, text="清空", command=self.clear_all, fg_color="gray", width=80)
        self.btn_clear.pack(side="left", padx=5)
        
        self.btn_export = ctk.CTkButton(btn_frame, text="导出为Word (自动排版)", command=self.save_to_word, width=200, fg_color="#2CC985")
        self.btn_export.pack(side="right", padx=5)

        # 初始化加载
        self.on_mode_change("教研论文 (参照《虚拟仿真/真实情境》)")

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
        # 自动填充标题示例
        if "教研论文" in choice:
            self.entry_topic.delete(0, "end")
            self.entry_topic.insert(0, "高中化学教学中真实情境创设的实践研究")
            self.lbl_word_count.configure(text="约 3500-4500 字")
        elif "解题指导" in choice:
            self.entry_topic.delete(0, "end")
            self.entry_topic.insert(0, "高中化学工艺流程题的解题模型建构")
            self.lbl_word_count.configure(text="约 2500-3000 字")
        elif "深度反思" in choice:
            self.entry_topic.delete(0, "end")
            self.entry_topic.insert(0, "高三化学一轮复习教学的深度反思")
            self.lbl_word_count.configure(text="约 2000-2500 字")
        else:
            self.entry_topic.delete(0, "end")
            self.lbl_word_count.configure(text="根据指令自动生成")

        # 预览结构
        self.txt_outline.delete("0.0", "end")
        template = YU_TEMPLATES.get(choice, [])
        if template:
            for item in template:
                self.txt_outline.insert("end", item["title"] + "\n")
        else:
            self.txt_outline.insert("end", "（自由定制模式：点击开始后由AI自动规划）")

    def run_writing(self):
        self.stop_event.clear()
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        
        if not topic:
            messagebox.showerror("错误", "请输入文章标题")
            return

        threading.Thread(target=self.thread_write_process, args=(mode, topic, instr), daemon=True).start()

    def thread_write_process(self, mode, topic, instr):
        client = self.get_client()
        if not client: return

        self.btn_run.configure(state="disabled", text="撰写中...")
        self.txt_content.delete("0.0", "end")
        
        # 1. 获取写作任务列表
        template = YU_TEMPLATES.get(mode, [])
        
        # 如果是自由定制，先生成大纲
        if not template: 
            self.txt_content.insert("end", "正在规划大纲...\n")
            # (此处省略自由大纲生成逻辑，为保持简洁，重点放在定制模版上)
            # 简单回落逻辑：
            template = [
                {"title": "一、背景", "prompt": "写背景"},
                {"title": "二、内容", "prompt": "写主要内容"},
                {"title": "三、总结", "prompt": "写总结"}
            ]

        # 2. 逐章节撰写 (严格控制字数的核心)
        total_steps = len(template)
        
        try:
            for i, section in enumerate(template):
                if self.stop_event.is_set(): break
                
                title = section["title"]
                prompt_req = section["prompt"]
                
                # 插入显眼的标题
                self.txt_content.insert("end", f"\n\n【{title}】\n")
                self.txt_content.see("end")
                
                # 构建“俞晋全风格”的 Prompt
                system_prompt = f"""
                你就是俞晋全老师，一位经验丰富的高中化学名师。
                你的文风特点：
                1. 务实：不喜欢空洞的理论堆砌，喜欢用教学中的真实案例（如学生哪里错了、实验哪里失败了）来说明问题。
                2. 专业：对化学知识点（如氯气、钠、氧化还原、电化学）信手拈来。
                3. 结构：条理极其清晰，喜欢用“（一）...（二）...”这种层级。
                
                【当前任务】：撰写《{topic}》的【{title}】部分。
                【严格限制】：
                1. 本部分字数控制在 400-600 字之间（摘要除外）。
                2. 严禁复述标题。
                3. 严禁 Markdown。
                4. 用户额外指令：{instr}
                """
                
                user_prompt = f"""
                请根据以下要求撰写本部分：
                {prompt_req}
                
                请直接输出正文。
                """

                # 请求 AI
                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.75 # 保持一定创造性，但不过分发散
                )
                
                raw_text = resp.choices[0].message.content
                
                # 3. 智能清洗：去除 AI 可能重复输出的标题
                # 如果第一行包含了标题中的关键词，就删掉第一行
                lines = raw_text.strip().split('\n')
                clean_title = re.sub(r'[一二三四、（）]', '', title) # 去掉序号
                if len(lines) > 0 and clean_title in lines[0]:
                    raw_text = "\n".join(lines[1:])
                
                self.txt_content.insert("end", raw_text.strip())
                self.txt_content.see("end")
                
                time.sleep(1) # 缓冲

            if not self.stop_event.is_set():
                messagebox.showinfo("完成", "文稿撰写完毕！请点击导出 Word。")

        except Exception as e:
            messagebox.showerror("API 错误", str(e))
        finally:
            self.btn_run.configure(state="normal", text="开始撰写")

    def save_to_word(self):
        content = self.txt_content.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            
            # --- 样式设置 (完全复刻您的范文) ---
            style = doc.styles['Normal']
            style.font.name = u'Times New Roman'
            style.font.size = Pt(12)
            style._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 1. 标题 (黑体三号居中)
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_t = p_title.add_run(self.entry_topic.get())
            run_t.font.name = u'黑体'
            run_t._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run_t.font.size = Pt(16)
            run_t.bold = True
            
            # 2. 作者信息
            p_author = doc.add_paragraph()
            p_author.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_a = p_author.add_run(f"{DEV_NAME}\n({DEV_ORG}，甘肃 金塔 735399)")
            run_a.font.name = u'楷体'
            run_a._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
            run_a.font.size = Pt(10.5) # 五号
            
            doc.add_paragraph() # 空行

            # 3. 正文解析与排版
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                
                # 识别系统插入的标记 【Title】
                if line.startswith("【") and line.endswith("】"):
                    header = line.replace("【", "").replace("】", "")
                    
                    # 判断是一级标题（如“一、”）还是其他
                    if "摘要" in header or "参考文献" in header:
                        p = doc.add_paragraph()
                        run = p.add_run(header)
                        run.bold = True
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    elif re.match(r'^[一二三四五六七八九十]+、', header):
                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(12) # 段前间距
                        run = p.add_run(header)
                        run.bold = True
                        run.font.size = Pt(14) # 四号
                        run.font.name = u'黑体'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                    else:
                        # 兜底
                        p = doc.add_paragraph(header)
                        p.runs[0].bold = True
                else:
                    # 正文内容
                    # 检查是否包含（一）（二）这种二级标题，如果有，加粗
                    p = doc.add_paragraph()
                    if re.match(r'^（[一二三四五六七八九十]+）', line):
                        run = p.add_run(line)
                        run.bold = True
                        run.font.name = u'楷体_GB2312'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体_GB2312')
                    else:
                        p.add_run(line)
                        p.paragraph_format.first_line_indent = Pt(24) # 首行缩进
                        p.paragraph_format.line_spacing = 1.25

            doc.save(file_path)
            messagebox.showinfo("成功", f"已导出符合期刊格式的文档：\n{os.path.basename(file_path)}")

    def stop_writing(self):
        self.stop_event.set()
    def clear_all(self):
        self.txt_content.delete("0.0", "end")
    def get_client(self):
        return OpenAI(api_key=self.api_config.get("api_key"), base_url=self.api_config.get("base_url"))
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
    app = YuWriterApp()
    app.mainloop()
