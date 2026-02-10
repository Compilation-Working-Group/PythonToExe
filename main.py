import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import json
import re
from datetime import datetime
import pyperclip
from openai import OpenAI

# --- 扩展功能库 ---
from duckduckgo_search import DDGS
import pypdf
from docx import Document
import pandas as pd
try:
    from pptx import Presentation
except ImportError:
    Presentation = None

# --- 配置区域 ---
APP_NAME = "DeepSeek Pro"
APP_VERSION = "v2.1.0 (History & Multi-Attach)"
DEV_INFO = "开发者：Yu Jinquan | 核心：DeepSeek-V3/R1"

DEFAULT_CONFIG = {
    "api_key": "",
    "model": "deepseek-chat",
    "use_search": False,
    "is_r1": False,
    "system_prompt": "你是一个乐于助人的AI助手。输出代码时请使用Markdown格式。"
}

# 颜色配置
COLOR_USER_BUBBLE = "#95EC69"
COLOR_USER_TEXT = "#000000"
COLOR_AI_BUBBLE = "#FFFFFF"
COLOR_AI_BUBBLE_DARK = "#2B2B2B"
COLOR_BG = ("#F2F2F2", "#1a1a1a")

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ChatBubble(ctk.CTkFrame):
    """ 增强版气泡：支持悬停显示复制按钮 """
    def __init__(self, master, role, text, is_reasoning=False, timestamp=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.role = role
        self.raw_text = text # 保存原始文本用于复制
        
        # 布局
        self.grid_columnconfigure(0 if role == "user" else 1, weight=1)
        self.grid_columnconfigure(1 if role == "user" else 0, weight=0)
        
        # 样式
        if role == "user":
            bubble_color = COLOR_USER_BUBBLE
            text_color = COLOR_USER_TEXT
            anchor = "e"
        else:
            bubble_color = (COLOR_AI_BUBBLE, COLOR_AI_BUBBLE_DARK)
            text_color = ("black", "white")
            anchor = "w"

        if is_reasoning:
            bubble_color = ("#F0F0F0", "#333333")
            text_color = "gray"
            display_text = f"🧠 深度思考:\n{text}"
        else:
            display_text = text

        # 气泡容器
        self.bubble_inner = ctk.CTkFrame(self, fg_color=bubble_color, corner_radius=12)
        self.bubble_inner.grid(row=0, column=1 if role == "user" else 0, padx=10, pady=5, sticky=anchor)

        # 绑定悬停事件
        self.bubble_inner.bind("<Enter>", self.on_enter)
        self.bubble_inner.bind("<Leave>", self.on_leave)

        # 渲染内容
        self.render_content(self.bubble_inner, display_text, text_color)

        # 时间戳
        if timestamp:
            ts_lbl = ctk.CTkLabel(self, text=timestamp, font=("Arial", 10), text_color="gray")
            ts_lbl.grid(row=1, column=1 if role == "user" else 0, padx=15, sticky=anchor)

        # 复制按钮 (初始隐藏)
        self.btn_copy = ctk.CTkButton(self.bubble_inner, text="📋", width=24, height=24, 
                                      fg_color="transparent", hover_color=("gray80", "gray40"),
                                      text_color=("gray50", "gray80"),
                                      font=("Arial", 14),
                                      command=self.copy_content)
        # 不pack，悬停时pack

    def on_enter(self, event):
        # 鼠标进入显示复制按钮
        self.btn_copy.place(relx=1.0, rely=1.0, anchor="se", x=-2, y=-2)

    def on_leave(self, event):
        # 鼠标离开隐藏
        self.btn_copy.place_forget()

    def copy_content(self):
        try:
            pyperclip.copy(self.raw_text)
            self.btn_copy.configure(text="✅")
            self.after(1000, lambda: self.btn_copy.configure(text="📋"))
        except Exception as e:
            print(f"Copy failed: {e}")

    def render_content(self, parent, text, text_color):
        parts = re.split(r'(```[\s\S]*?```)', text)
        for part in parts:
            if part.startswith("```") and part.endswith("```"):
                # 代码块
                code = part.strip("`")
                if '\n' in code:
                    code = code.split('\n', 1)[1] # 去掉第一行语言标记
                
                f = ctk.CTkFrame(parent, fg_color="#1E1E1E", corner_radius=5)
                f.pack(fill="x", padx=8, pady=5)
                
                t = ctk.CTkTextbox(f, font=("Consolas", 12), text_color="#D4D4D4", fg_color="transparent", 
                                   height=min(len(code.split('\n'))*20 + 20, 300), wrap="none")
                t.insert("0.0", code)
                t.configure(state="disabled")
                t.pack(fill="x", padx=5, pady=5)
                
                # 代码块自带显式复制按钮
                ctk.CTkButton(f, text="复制代码", height=20, width=60, font=("Arial", 10),
                              fg_color="#333333", hover_color="#444444",
                              command=lambda c=code: pyperclip.copy(c)).pack(anchor="ne", padx=5, pady=2)
            else:
                if part.strip():
                    ctk.CTkLabel(parent, text=part, text_color=text_color, justify="left", 
                                 font=("Microsoft YaHei UI", 14), wraplength=550).pack(fill="x", padx=10, pady=5)

class DeepSeekApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1200x850")
        
        self.config = self.load_json("config.json", DEFAULT_CONFIG)
        self.history_data = self.load_json("history.json", []) # 加载历史记录
        
        self.client = None
        self.is_running = False 
        
        # 附件列表：[{'name': 'a.txt', 'content': '...'}, ...]
        self.attachments = [] 

        self.setup_ui()
        self.restore_history() # 恢复界面
        
        if self.config["api_key"]:
            self.init_client()

    def load_json(self, path, default):
        if os.path.exists(path):
            try: return json.load(open(path, "r", encoding="utf-8"))
            except: pass
        return default

    def save_config(self):
        json.dump(self.config, open("config.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    def save_history(self):
        # 保存结构化数据
        json.dump(self.history_data, open("history.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    def init_client(self):
        if not self.config["api_key"]: return
        self.client = OpenAI(api_key=self.config["api_key"], base_url="https://api.deepseek.com")

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 侧边栏 ===
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(10, weight=1)

        ctk.CTkLabel(self.sidebar, text="DeepSeek Pro", font=("Arial", 20, "bold")).pack(pady=(30, 5))
        ctk.CTkLabel(self.sidebar, text=APP_VERSION, font=("Arial", 10), text_color="gray").pack(pady=(0, 20))

        # 设置开关
        frame_set = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        frame_set.pack(fill="x", padx=10)
        
        self.r1_var = ctk.BooleanVar(value=self.config.get("is_r1", False))
        ctk.CTkSwitch(frame_set, text="深度思考 (R1)", variable=self.r1_var, command=self.update_settings).pack(pady=10, anchor="w")
        
        self.search_var = ctk.BooleanVar(value=self.config["use_search"])
        ctk.CTkSwitch(frame_set, text="联网搜索", variable=self.search_var, command=self.update_settings).pack(pady=10, anchor="w")

        # Key
        ctk.CTkLabel(self.sidebar, text="API Key:", anchor="w").pack(padx=15, pady=(20, 0), fill="x")
        self.entry_key = ctk.CTkEntry(self.sidebar, show="*", placeholder_text="sk-...")
        self.entry_key.insert(0, self.config["api_key"])
        self.entry_key.pack(padx=15, pady=5, fill="x")
        ctk.CTkButton(self.sidebar, text="保存 / 重连", height=30, command=self.save_key).pack(padx=15, pady=5, fill="x")

        # 底部
        ctk.CTkButton(self.sidebar, text="🗑️ 清空所有记录", fg_color="#C0392B", hover_color="#E74C3C", command=self.clear_all_history).pack(side="bottom", padx=15, pady=10, fill="x")
        ctk.CTkLabel(self.sidebar, text=DEV_INFO, font=("Arial", 10), text_color="gray50").pack(side="bottom", pady=5)

        # === 主区域 ===
        self.main_area = ctk.CTkFrame(self, fg_color=COLOR_BG)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        self.main_area.grid_rowconfigure(0, weight=1)
        self.main_area.grid_columnconfigure(0, weight=1)

        # 聊天列表
        self.chat_scroll = ctk.CTkScrollableFrame(self.main_area, fg_color="transparent")
        self.chat_scroll.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # 输入区域容器
        input_container = ctk.CTkFrame(self.main_area, fg_color=("white", "#2B2B2B"), height=160)
        input_container.grid(row=1, column=0, sticky="ew", padx=15, pady=15)
        input_container.grid_columnconfigure(0, weight=1)

        # 工具栏 (附件)
        tool_frame = ctk.CTkFrame(input_container, fg_color="transparent", height=30)
        tool_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        
        self.btn_attach = ctk.CTkButton(tool_frame, text="📎 添加附件", width=80, height=24, 
                                        fg_color="transparent", border_width=1, 
                                        text_color=("gray20", "gray80"), command=self.upload_files)
        self.btn_attach.pack(side="left", padx=5)
        
        # 附件列表显示Label
        self.lbl_files = ctk.CTkLabel(tool_frame, text="", font=("Arial", 11), text_color="gray")
        self.lbl_files.pack(side="left", padx=5)
        
        self.btn_clear_files = ctk.CTkButton(tool_frame, text="清空附件", width=60, height=20, 
                                             fg_color="transparent", text_color="red", 
                                             font=("Arial", 11), command=self.clear_attachments)
        # 初始隐藏

        # 输入框
        self.entry_msg = ctk.CTkTextbox(input_container, height=80, font=("Microsoft YaHei UI", 14), fg_color="transparent", border_width=0)
        self.entry_msg.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.entry_msg.bind("<Return>", self.on_enter_press)

        # 发送按钮
        btn_action_frame = ctk.CTkFrame(input_container, fg_color="transparent")
        btn_action_frame.grid(row=1, column=1, sticky="s", padx=10, pady=10)
        
        self.btn_send = ctk.CTkButton(btn_action_frame, text="发送", width=80, command=self.send_message)
        self.btn_send.pack(side="bottom")
        
        self.btn_stop = ctk.CTkButton(btn_action_frame, text="⏹️", width=30, fg_color="#C0392B", command=self.stop_generation)

    # --- 逻辑 ---

    def update_settings(self):
        self.config["use_search"] = self.search_var.get()
        self.config["is_r1"] = self.r1_var.get()
        self.config["model"] = "deepseek-reasoner" if self.r1_var.get() else "deepseek-chat"
        self.save_config()

    def save_key(self):
        key = self.entry_key.get().strip()
        self.config["api_key"] = key
        self.save_config()
        self.init_client()
        messagebox.showinfo("OK", "Key Saved")

    def upload_files(self):
        # 允许所有文件
        filetypes = [("All Files", "*.*")]
        # 支持多选
        filepaths = filedialog.askopenfilenames(filetypes=filetypes)
        if not filepaths: return
        
        count = 0
        for path in filepaths:
            try:
                name = os.path.basename(path)
                content = self.extract_text(path)
                
                # 大文件截断保护
                if len(content) > 30000:
                    content = content[:30000] + f"\n\n[System: File '{name}' truncated due to length limit.]"
                
                self.attachments.append({"name": name, "content": content})
                count += 1
            except Exception as e:
                print(f"Error reading {path}: {e}")
        
        if count > 0:
            self.update_file_label()

    def update_file_label(self):
        if not self.attachments:
            self.lbl_files.configure(text="")
            self.btn_clear_files.pack_forget()
            return
            
        names = [f["name"] for f in self.attachments]
        display_text = " | ".join(names)
        if len(display_text) > 50: display_text = display_text[:47] + "..."
        
        self.lbl_files.configure(text=f"已添加 {len(names)} 个文件: {display_text}")
        self.btn_clear_files.pack(side="left", padx=5)

    def clear_attachments(self):
        self.attachments = []
        self.update_file_label()

    def extract_text(self, filepath):
        """ 增强版文件读取器：尝试读取所有格式 """
        ext = os.path.splitext(filepath)[1].lower()
        try:
            if ext == '.pdf':
                reader = pypdf.PdfReader(filepath)
                return "\n".join([p.extract_text() or "" for p in reader.pages])
            elif ext == '.docx':
                doc = Document(filepath)
                return "\n".join([p.text for p in doc.paragraphs])
            elif ext == '.pptx' and Presentation:
                prs = Presentation(filepath)
                txt = []
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"): txt.append(shape.text)
                return "\n".join(txt)
            elif ext in ['.xlsx', '.xls', '.csv']:
                df = pd.read_excel(filepath) if 'xls' in ext else pd.read_csv(filepath)
                return df.to_string()
            else:
                # 尝试作为纯文本读取 (含代码文件)
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()
        except Exception:
            # 如果无法读取内容（如图片/二进制），返回文件名占位符
            return f"[System: User uploaded a file named '{os.path.basename(filepath)}'. This file type cannot be parsed as text, but be aware of its existence.]"

    def restore_history(self):
        """ 启动时恢复历史记录到界面 """
        for item in self.history_data:
            # item结构: {'role': 'user'/'ai', 'content': '...', 'reasoning': '...'}
            role = item.get('role')
            content = item.get('content', '')
            reasoning = item.get('reasoning', '')
            ts = item.get('timestamp', '')

            if role == 'user':
                self.add_bubble_ui('user', content, timestamp=ts)
            else:
                if reasoning:
                    self.add_bubble_ui('ai', reasoning, is_reasoning=True, timestamp=ts)
                if content:
                    self.add_bubble_ui('ai', content, is_reasoning=False, timestamp=ts)
                    
        # 滚动到底部
        self.chat_scroll.update_idletasks()
        try: self.chat_scroll._parent_canvas.yview_moveto(1.0)
        except: pass

    def clear_all_history(self):
        if messagebox.askyesno("确认", "确定要删除所有聊天记录吗？"):
            self.history_data = []
            self.save_history()
            for widget in self.chat_scroll.winfo_children(): widget.destroy()
            self.add_system_msg("记录已清空")

    def add_system_msg(self, text):
        ctk.CTkLabel(self.chat_scroll, text=text, font=("Arial", 10), text_color="gray").pack(pady=5)

    def add_bubble_ui(self, role, text, is_reasoning=False, timestamp=None):
        if not timestamp: timestamp = datetime.now().strftime("%m-%d %H:%M")
        bubble = ChatBubble(self.chat_scroll, role, text, is_reasoning, timestamp)
        bubble.pack(fill="x", pady=5)
        return bubble

    def on_enter_press(self, event):
        if not event.state & 0x0001: 
            self.send_message()
            return "break"

    def stop_generation(self):
        self.is_running = False
        self.btn_stop.pack_forget()
        self.btn_send.configure(state="normal", text="发送")

    def perform_search(self, query):
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=3))
                if results: return "\n".join([f"Source: {r['title']}\n{r['body']}" for r in results])
        except: pass
        return ""

    def send_message(self):
        text = self.entry_msg.get("0.0", "end").strip()
        if not text and not self.attachments: return
        if not self.client: return messagebox.showerror("Error", "No API Key")

        # 1. 准备上下文
        user_display_text = text
        full_prompt = ""
        
        # 处理附件
        if self.attachments:
            file_texts = [f"--- File: {f['name']} ---\n{f['content']}\n" for f in self.attachments]
            full_prompt += "\n".join(file_texts) + "\n"
            user_display_text += f"\n[已发送 {len(self.attachments)} 个文件]"
            self.clear_attachments() # UI清空

        full_prompt += text

        # 2. UI 更新
        self.entry_msg.delete("0.0", "end")
        ts = datetime.now().strftime("%m-%d %H:%M")
        self.add_bubble_ui("user", user_display_text, timestamp=ts)
        
        # 存入历史数据 (注意：历史数据存完整的Prompt以便上下文记忆，但UI只显示简洁的)
        # 改进策略：为了节省token，历史记录里附件内容可能需要裁剪，这里暂存完整版
        self.history_data.append({"role": "user", "content": full_prompt, "timestamp": ts})
        self.save_history()

        # 3. 启动线程
        self.is_running = True
        self.btn_send.configure(state="disabled", text="...")
        self.btn_stop.pack(side="bottom")
        
        threading.Thread(target=self.process_stream, args=(full_prompt,), daemon=True).start()

    def process_stream(self, prompt):
        # 联网搜索
        if self.search_var.get():
            self.after(0, lambda: self.add_system_msg("🔍 正在搜索..."))
            s = self.perform_search(prompt[-100:]) # 仅搜索最后一部分文本
            if s: prompt = f"参考资料:\n{s}\n\n问题:\n{prompt}"

        # 构建API消息列表 (仅取最近10轮以节省Token)
        api_messages = [{"role": "system", "content": self.config["system_prompt"]}]
        # 转换 history_data 格式给 API
        for h in self.history_data[-10:]: 
            # 过滤掉 reasoning 字段，只发 content
            api_messages.append({"role": "user" if h["role"]=="user" else "assistant", "content": h["content"]})
        
        # 如果刚才追加了搜索结果，替换最后一条
        api_messages[-1]["content"] = prompt

        try:
            response = self.client.chat.completions.create(
                model=self.config["model"],
                messages=api_messages,
                stream=True
            )
            
            ai_text = ""
            r1_text = ""
            
            # 临时流式容器
            self.current_r1_box = None
            self.current_ai_box = None
            
            def append_ui(widget, txt):
                widget.configure(state="normal")
                widget.insert("end", txt)
                widget.see("end")
                widget.configure(state="disabled")

            def init_box(is_r1):
                f = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
                f.pack(fill="x", pady=2)
                t = ctk.CTkTextbox(f, font=("Microsoft YaHei UI", 14), height=60, wrap="word", fg_color=("white", "#2B2B2B"))
                t.pack(fill="x", padx=10)
                if is_r1: 
                    t.configure(text_color="gray", font=("Arial", 12))
                    t.insert("0.0", "🧠 思考中...\n")
                return f, t

            for chunk in response:
                if not self.is_running: break
                delta = chunk.choices[0].delta
                
                # R1 思考
                if hasattr(delta, 'reasoning_content') and delta.reasoning_content:
                    c = delta.reasoning_content
                    r1_text += c
                    if not self.current_r1_box:
                        ready = threading.Event()
                        def _make():
                            self.fr1, self.current_r1_box = init_box(True)
                            ready.set()
                        self.after(0, _make)
                        ready.wait()
                    self.after(0, lambda x=c: append_ui(self.current_r1_box, x))

                # 正文
                if hasattr(delta, 'content') and delta.content:
                    c = delta.content
                    ai_text += c
                    if not self.current_ai_box:
                        ready = threading.Event()
                        def _make():
                            self.fai, self.current_ai_box = init_box(False)
                            ready.set()
                        self.after(0, _make)
                        ready.wait()
                    self.after(0, lambda x=c: append_ui(self.current_ai_box, x))

            # 结束：替换为正式气泡并保存
            def finalize():
                if self.current_r1_box: self.fr1.destroy()
                if self.current_ai_box: self.fai.destroy()
                
                ts = datetime.now().strftime("%m-%d %H:%M")
                if r1_text: self.add_bubble_ui("ai", r1_text, True, ts)
                if ai_text: self.add_bubble_ui("ai", ai_text, False, ts)
                
                # 保存到历史
                self.history_data.append({
                    "role": "ai", 
                    "content": ai_text, 
                    "reasoning": r1_text, 
                    "timestamp": ts
                })
                self.save_history()
                
                self.is_running = False
                self.btn_stop.pack_forget()
                self.btn_send.configure(state="normal", text="发送")

            self.after(0, finalize)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("API Error", str(e)))
            self.after(0, lambda: self.btn_send.configure(state="normal", text="发送"))

if __name__ == "__main__":
    app = DeepSeekApp()
    app.mainloop()
