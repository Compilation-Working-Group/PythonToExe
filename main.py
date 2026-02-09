import customtkinter as ctk
import threading
from openai import OpenAI
import os
from datetime import datetime

# 设置外观模式
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class DeepSeekWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 窗口基础设置
        self.title("AI 智能写作助手 - Powered by DeepSeek")
        self.geometry("1000x700")
        
        # 初始化变量
        self.api_key = ctk.StringVar(value="")  # 可以在此处预填 key，或在界面输入
        self.doc_type = ctk.StringVar(value="期刊论文")
        self.generated_outline = ""
        self.full_content = ""
        self.client = None

        # 布局容器
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧边栏 (设置区) ===
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="写作助手 Pro", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.api_label = ctk.CTkLabel(self.sidebar_frame, text="DeepSeek API Key:", anchor="w")
        self.api_label.grid(row=1, column=0, padx=20, pady=(10, 0))
        self.api_entry = ctk.CTkEntry(self.sidebar_frame, textvariable=self.api_key, show="*")
        self.api_entry.grid(row=2, column=0, padx=20, pady=(0, 10))

        self.type_label = ctk.CTkLabel(self.sidebar_frame, text="文稿类型:", anchor="w")
        self.type_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.type_menu = ctk.CTkOptionMenu(self.sidebar_frame, variable=self.doc_type,
                                           values=["期刊论文", "教学计划", "教学反思", "案例分析", "年度总结", "自定义"])
        self.type_menu.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="n")

        # === 右侧主内容区 (Tabview) ===
        self.tabview = ctk.CTkTabview(self, width=750)
        self.tabview.grid(row=0, column=1, padx=(10, 20), pady=(10, 20), sticky="nsew")
        
        self.tab_setup = self.tabview.add("1. 题目与要求")
        self.tab_outline = self.tabview.add("2. 大纲修订")
        self.tab_write = self.tabview.add("3. 正文生成")

        self.setup_ui_tab1()
        self.setup_ui_tab2()
        self.setup_ui_tab3()

    def setup_ui_tab1(self):
        """Tab 1: 输入题目和具体要求"""
        self.tab_setup.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.tab_setup, text="文稿标题/主题:", anchor="w").grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        self.title_entry = ctk.CTkEntry(self.tab_setup, placeholder_text="例如：高中化学探究式教学在电化学单元的应用")
        self.title_entry.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="ew")

        ctk.CTkLabel(self.tab_setup, text="具体指令或特殊要求 (可选):", anchor="w").grid(row=2, column=0, padx=10, pady=(10, 0), sticky="w")
        self.requirements_text = ctk.CTkTextbox(self.tab_setup, height=200)
        self.requirements_text.grid(row=3, column=0, padx=10, pady=(5, 10), sticky="nsew")
        self.requirements_text.insert("0.0", "例如：重点分析学生在原电池理解上的常见误区，结合具体课堂案例，字数约3000字。")

        self.btn_gen_outline = ctk.CTkButton(self.tab_setup, text="生成大纲 >>", command=self.start_generate_outline)
        self.btn_gen_outline.grid(row=4, column=0, padx=10, pady=20)
        
        self.status_label_1 = ctk.CTkLabel(self.tab_setup, text="")
        self.status_label_1.grid(row=5, column=0)

    def setup_ui_tab2(self):
        """Tab 2: 显示并允许用户修改大纲"""
        self.tab_outline.grid_columnconfigure(0, weight=1)
        self.tab_outline.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.tab_outline, text="请检查并修改生成的论文大纲:", anchor="w").grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        self.outline_editor = ctk.CTkTextbox(self.tab_outline)
        self.outline_editor.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew")

        self.btn_gen_full = ctk.CTkButton(self.tab_outline, text="确认大纲并撰写正文 >>", fg_color="green", command=self.start_generate_content)
        self.btn_gen_full.grid(row=2, column=0, padx=10, pady=20)

        self.status_label_2 = ctk.CTkLabel(self.tab_outline, text="")
        self.status_label_2.grid(row=3, column=0)

    def setup_ui_tab3(self):
        """Tab 3: 显示最终正文"""
        self.tab_write.grid_columnconfigure(0, weight=1)
        self.tab_write.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.tab_write, text="生成的完整文稿:", anchor="w").grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        self.content_display = ctk.CTkTextbox(self.tab_write)
        self.content_display.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew")

        self.btn_save = ctk.CTkButton(self.tab_write, text="保存为 Markdown 文件", command=self.save_to_file)
        self.btn_save.grid(row=2, column=0, padx=10, pady=20)

    # === 逻辑处理部分 ===

    def init_client(self):
        key = self.api_key.get().strip()
        if not key:
            return False
        # DeepSeek API 配置
        self.client = OpenAI(api_key=key, base_url="https://api.deepseek.com")
        return True

    def start_generate_outline(self):
        if not self.init_client():
            self.status_label_1.configure(text="错误: 请先输入 API Key", text_color="red")
            return
        
        title = self.title_entry.get()
        reqs = self.requirements_text.get("0.0", "end")
        dtype = self.doc_type.get()
        
        if not title:
            self.status_label_1.configure(text="错误: 标题不能为空", text_color="red")
            return

        self.status_label_1.configure(text="DeepSeek 正在思考并生成大纲...", text_color="blue")
        self.btn_gen_outline.configure(state="disabled")
        
        # 开启线程避免界面卡死
        threading.Thread(target=self.run_outline_api, args=(title, dtype, reqs)).start()

    def run_outline_api(self, title, dtype, reqs):
        try:
            prompt = f"""
            你是一个专业的学术写作助手。请为一篇主题为“{title}”的【{dtype}】撰写一份详细的大纲。
            用户额外要求：{reqs}
            
            请只输出大纲结构，层级分明（如：一、二、1. 2. 等），不要输出多余的寒暄语。
            """
            
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个严谨的学术助手。"},
                    {"role": "user", "content": prompt},
                ],
                stream=False
            )
            
            result = response.choices[0].message.content
            
            # 回到主线程更新 UI
            self.outline_editor.insert("0.0", result)
            self.tabview.set("2. 大纲修订") # 自动跳转
            self.status_label_1.configure(text="大纲生成完毕！", text_color="green")
            
        except Exception as e:
            self.status_label_1.configure(text=f"API 请求失败: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal")

    def start_generate_content(self):
        outline = self.outline_editor.get("0.0", "end").strip()
        if not outline:
            self.status_label_2.configure(text="错误: 大纲不能为空", text_color="red")
            return

        self.status_label_2.configure(text="DeepSeek 正在根据大纲撰写全文，这可能需要一两分钟...", text_color="blue")
        self.btn_gen_full.configure(state="disabled")
        
        threading.Thread(target=self.run_content_api, args=(outline,)).start()

    def run_content_api(self, outline):
        try:
            title = self.title_entry.get()
            dtype = self.doc_type.get()
            reqs = self.requirements_text.get("0.0", "end")
            
            prompt = f"""
            请根据以下大纲，撰写一篇完整的【{dtype}】。
            题目：{title}
            额外要求：{reqs}
            
            大纲如下：
            {outline}
            
            要求：内容详实，逻辑严密，语言专业，符合{dtype}的规范。请直接输出正文，使用 Markdown 格式。
            """
            
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个资深的领域专家和专业写作者。"},
                    {"role": "user", "content": prompt},
                ],
                stream=False
            )
            
            result = response.choices[0].message.content
            
            self.content_display.insert("0.0", result)
            self.tabview.set("3. 正文生成") # 自动跳转
            self.status_label_2.configure(text="全文撰写完毕！", text_color="green")

        except Exception as e:
            self.status_label_2.configure(text=f"API 请求失败: {str(e)}", text_color="red")
        finally:
            self.btn_gen_full.configure(state="normal")

    def save_to_file(self):
        content = self.content_display.get("0.0", "end")
        title = self.title_entry.get().strip() or "output"
        filename = f"{title}_{datetime.now().strftime('%Y%m%d')}.md"
        
        try:
            with open(filename, "w", encoding="utf-8") as f:
                f.write(content)
            self.btn_save.configure(text=f"已保存为 {filename}", fg_color="gray")
        except Exception as e:
            self.btn_save.configure(text="保存失败", fg_color="red")

if __name__ == "__main__":
    app = DeepSeekWriterApp()
    app.mainloop()
