import os
import sys
import asyncio
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import docx
import edge_tts
from openai import OpenAI
import imageio_ffmpeg
import re

# 默认配置
DEFAULT_DEEPSEEK_URL = "https://api.deepseek.com"

# --- 完整的 Edge-TTS 免费中文语音库 ---
VOICE_MAP = {
    "晓晓 (女声 - 活泼/默认)": "zh-CN-XiaoxiaoNeural",
    "晓伊 (女声 - 可爱/儿童)": "zh-CN-XiaoyiNeural",
    "云希 (男声 - 沉稳/影视)": "zh-CN-YunxiNeural",
    "云健 (男声 - 体育/解说)": "zh-CN-YunjianNeural",
    "云扬 (男声 - 新闻/播音)": "zh-CN-YunyangNeural",
    "云夏 (男声 - 少年)": "zh-CN-YunxiaNeural",
    "辽宁小北 (东北话 - 女声)": "zh-CN-Liaoning-XiaobeiNeural",
    "陕西小妮 (陕西话 - 女声)": "zh-CN-Shaanxi-XiaoniNeural",
    "香港晓佳 (粤语 - 女声1)": "zh-HK-HiuGaaiNeural",
    "香港晓曼 (粤语 - 女声2)": "zh-HK-HiuMaanNeural",
    "香港云龙 (粤语 - 男声)": "zh-HK-WanLungNeural",
    "台湾晓臻 (台湾腔 - 女声1)": "zh-TW-HsiaoChenNeural",
    "台湾晓雨 (台湾腔 - 女声2)": "zh-TW-HsiaoYuNeural",
    "台湾云哲 (台湾腔 - 男声)": "zh-TW-YunJheNeural",
    "英语 (女声 - Aria)": "en-US-AriaNeural",
    "英语 (男声 - Guy)": "en-US-GuyNeural"
}

class TTSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepSeek 智能语音合成助手 - 作者: Yu JinQuan")
        
        # 稍微加高了窗口，以容纳新的参数面板
        window_width = 980
        window_height = 760
        self.center_window(window_width, window_height)
        self.root.minsize(900, 600)
        
        self.is_playing = False
        self.is_generating = False 
        self.temp_audio_file = "temp_preview.mp3"
        self.loop = asyncio.new_event_loop()
        
        self.selected_voice_key = ttk.StringVar(value="晓晓 (女声 - 活泼/默认)")
        
        # 初始化参数变量
        self.rate_var = tk.DoubleVar(value=0)
        self.volume_var = tk.DoubleVar(value=0)
        self.pitch_var = tk.DoubleVar(value=0)
        
        threading.Thread(target=self.start_loop, daemon=True).start()
        self.create_ui()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def start_loop(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def create_ui(self):
        # 1. 顶部操作区
        frame_top = ttk.Labelframe(self.root, text="文件与编辑", padding=15, bootstyle="info")
        frame_top.pack(side=TOP, fill=X, padx=15, pady=(15, 5))
        
        ttk.Button(frame_top, text="📂 导入文本/Word", command=self.import_file, bootstyle="primary-outline").pack(side=LEFT, padx=5)
        ttk.Button(frame_top, text="🗑️ 清空内容", command=self.clear_text, bootstyle="danger-outline").pack(side=LEFT, padx=5)
        
        ttk.Frame(frame_top, width=30).pack(side=LEFT) 
        ttk.Label(frame_top, text="选中多音字后点击 ->", foreground="gray").pack(side=LEFT)
        ttk.Button(frame_top, text="📝 修正选中字读音", command=self.fix_pronunciation, bootstyle="warning").pack(side=LEFT, padx=5)

        # === 以下按倒序(BOTTOM)加载 ===

        # 2. 状态栏 (最底)
        frame_status = ttk.Frame(self.root, padding=5)
        frame_status.pack(side=BOTTOM, fill=X)
        self.status_label = ttk.Label(frame_status, text="状态: 就绪", bootstyle="secondary")
        self.status_label.pack(side=LEFT, padx=10)
        ttk.Label(frame_status, text="Author: Yu JinQuan", bootstyle="secondary").pack(side=RIGHT, padx=10)

        # 3. 语音控制与导出 (倒数第二)
        frame_bottom = ttk.Labelframe(self.root, text="语音控制与导出", padding=15, bootstyle="primary")
        frame_bottom.pack(side=BOTTOM, fill=X, padx=15, pady=(5, 10))
        
        ttk.Label(frame_bottom, text="发音人:").pack(side=LEFT, padx=(5, 5))
        voice_combo = ttk.Combobox(frame_bottom, textvariable=self.selected_voice_key, values=list(VOICE_MAP.keys()), state="readonly", width=25, bootstyle="primary")
        voice_combo.pack(side=LEFT, padx=5)

        ttk.Separator(frame_bottom, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=15)

        ttk.Button(frame_bottom, text="▶️ 生成并播放", command=self.play_audio, bootstyle="success").pack(side=LEFT, padx=5)
        ttk.Button(frame_bottom, text="⏹️ 停止", command=self.stop_audio, bootstyle="danger").pack(side=LEFT, padx=5)
        
        ttk.Separator(frame_bottom, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=15)
        
        ttk.Button(frame_bottom, text="💾 导出 MP3", command=lambda: self.export_audio("mp3"), bootstyle="info").pack(side=LEFT, padx=5)
        ttk.Button(frame_bottom, text="🎵 导出 WAV", command=lambda: self.export_audio("wav"), bootstyle="info").pack(side=LEFT, padx=5)

        # 4. 新增：高级参数调节区 (倒数第三)
        frame_params = ttk.Labelframe(self.root, text="高级语音参数", padding=10, bootstyle="warning")
        frame_params.pack(side=BOTTOM, fill=X, padx=15, pady=5)
        
        # 使用 Grid 布局使滑块对齐更美观
        # --- 语速 ---
        ttk.Label(frame_params, text="语速调节:").grid(row=0, column=0, padx=(10, 5), pady=5, sticky="e")
        scale_rate = ttk.Scale(frame_params, from_=-50, to=50, variable=self.rate_var, command=self.update_param_labels, bootstyle="primary")
        scale_rate.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.lbl_rate = ttk.Label(frame_params, text="0%", width=5, font=("Arial", 10, "bold"))
        self.lbl_rate.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        # --- 音量 ---
        ttk.Label(frame_params, text="音量调节:").grid(row=0, column=3, padx=(20, 5), pady=5, sticky="e")
        scale_vol = ttk.Scale(frame_params, from_=-50, to=50, variable=self.volume_var, command=self.update_param_labels, bootstyle="success")
        scale_vol.grid(row=0, column=4, padx=5, pady=5, sticky="ew")
        self.lbl_vol = ttk.Label(frame_params, text="0%", width=5, font=("Arial", 10, "bold"))
        self.lbl_vol.grid(row=0, column=5, padx=5, pady=5, sticky="w")
        
        # --- 音调 ---
        ttk.Label(frame_params, text="音调调节:").grid(row=0, column=6, padx=(20, 5), pady=5, sticky="e")
        scale_pitch = ttk.Scale(frame_params, from_=-50, to=50, variable=self.pitch_var, command=self.update_param_labels, bootstyle="warning")
        scale_pitch.grid(row=0, column=7, padx=5, pady=5, sticky="ew")
        self.lbl_pitch = ttk.Label(frame_params, text="0Hz", width=6, font=("Arial", 10, "bold"))
        self.lbl_pitch.grid(row=0, column=8, padx=5, pady=5, sticky="w")
        
        # --- 重置按钮 ---
        ttk.Button(frame_params, text="🔄 重置参数", command=self.reset_params, bootstyle="secondary-outline").grid(row=0, column=9, padx=(20, 10), pady=5)

        # 配置列权重，让滑块占据剩余空间
        frame_params.columnconfigure(1, weight=1)
        frame_params.columnconfigure(4, weight=1)
        frame_params.columnconfigure(7, weight=1)

        # 5. AI 润色区 (倒数第四)
        frame_ai = ttk.Labelframe(self.root, text="DeepSeek AI 智能处理", padding=15, bootstyle="success")
        frame_ai.pack(side=BOTTOM, fill=X, padx=15, pady=5)
        ttk.Label(frame_ai, text="提示: 借助大模型将生硬的文本改写为更自然、流畅的口语化播音文案。").pack(side=LEFT, padx=5)
        ttk.Button(frame_ai, text="✨ 开始智能润色", command=self.run_deepseek_polish, bootstyle="success-outline").pack(side=RIGHT, padx=5)

        # 6. 中间文本区 (置顶，填充剩余空间)
        frame_text = ttk.Frame(self.root, padding=2)
        frame_text.pack(side=TOP, expand=True, fill=BOTH, padx=15, pady=10)
        self.text_area = scrolledtext.ScrolledText(frame_text, font=("Microsoft YaHei", 12), wrap=tk.WORD, bd=1, relief=tk.SOLID)
        self.text_area.pack(expand=True, fill=BOTH)

        # === 右键菜单 ===
        self.context_menu = tk.Menu(self.root, tearoff=0, font=("Microsoft YaHei", 10))
        self.context_menu.add_command(label="剪切", command=self.cut_text)
        self.context_menu.add_command(label="复制", command=self.copy_text)
        self.context_menu.add_command(label="粘贴", command=self.paste_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="全选", command=self.select_all_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📝 修正选中字读音", command=self.fix_pronunciation)

        self.text_area.bind("<ButtonRelease-3>", self.show_context_menu)
        if sys.platform == "darwin":
            self.text_area.bind("<ButtonRelease-2>", self.show_context_menu)

    # --- 新增：参数调节联动功能 ---
    def update_param_labels(self, *args):
        # 获取滑块的整数值
        r = int(self.rate_var.get())
        v = int(self.volume_var.get())
        p = int(self.pitch_var.get())
        
        # 格式化显示文本 (带有正负号)
        self.lbl_rate.config(text=f"{r:+d}%" if r else "0%")
        self.lbl_vol.config(text=f"{v:+d}%" if v else "0%")
        self.lbl_pitch.config(text=f"{p:+d}Hz" if p else "0Hz")

    def reset_params(self):
        self.rate_var.set(0)
        self.volume_var.set(0)
        self.pitch_var.set(0)
        self.update_param_labels()

    # --- 右键菜单功能 ---
    def show_context_menu(self, event):
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def cut_text(self):
        self.text_area.event_generate("<<Cut>>")

    def copy_text(self):
        self.text_area.event_generate("<<Copy>>")

    def paste_text(self):
        self.text_area.event_generate("<<Paste>>")

    def select_all_text(self):
        self.text_area.tag_add(tk.SEL, "1.0", tk.END)
        self.text_area.mark_set(tk.INSERT, "1.0")
        self.text_area.see(tk.INSERT)
        return 'break'

    # --- 逻辑功能区 ---
    def update_status(self, text):
        self.status_label.config(text=f"状态: {text}")
        self.root.update_idletasks()

    def fix_pronunciation(self):
        try:
            selection = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            messagebox.showwarning("提示", "请先在文本框中选中需要修正读音的汉字！")
            return

        if not selection.strip():
            return

        hint = f"请输入 [{selection}] 的【同音字】\n例如选了“单”，这里输入发音相同的“善”"
        homophone = simpledialog.askstring("同音字替换", hint)
        
        if homophone:
            replacement = f"[{selection}|{homophone.strip()}]"
            self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.text_area.insert(tk.INSERT, replacement)
            self.update_status(f"已设置同音字: {selection} -> {homophone}")

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text/Word", "*.txt *.docx")])
        if not file_path: return
        try:
            content = ""
            if file_path.lower().endswith(".txt"):
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
            elif file_path.lower().endswith(".docx"):
                doc = docx.Document(file_path)
                content = "\n".join([para.text for para in doc.paragraphs])
            self.text_area.delete("1.0", tk.END)
            self.text_area.insert(tk.END, content)
            self.update_status(f"已加载: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def clear_text(self):
        self.text_area.delete("1.0", tk.END)
        self.stop_audio()
        self.update_status("内容已清空")

    def run_deepseek_polish(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("提示", "请先输入需要润色的文本")
            return
        
        api_key = os.getenv("DEEPSEEK_API_KEY")
        if not api_key:
            api_key = simpledialog.askstring("API Key", "请输入 DeepSeek API Key:", show="*")
            if not api_key: return
            os.environ["DEEPSEEK_API_KEY"] = api_key 

        threading.Thread(target=self._deepseek_thread, args=(text, api_key)).start()

    def _deepseek_thread(self, text, api_key):
        self.update_status("正在连接 DeepSeek AI...")
        try:
            client = OpenAI(api_key=api_key, base_url=DEFAULT_DEEPSEEK_URL)
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个专业的配音文案润色专家。请将用户输入的文本修改为适合朗读的口语化文案。直接输出结果。"},
                    {"role": "user", "content": text},
                ],
                stream=False
            )
            polished = response.choices[0].message.content
            self.root.after(0, lambda: self.text_area.delete("1.0", tk.END))
            self.root.after(0, lambda: self.text_area.insert(tk.END, polished))
            self.root.after(0, lambda: self.update_status("润色完成"))
            self.root.after(0, lambda: messagebox.showinfo("完成", "DeepSeek 润色已完成！"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("API 错误", f"请求失败: {str(e)}"))
            self.root.after(0, lambda: self.update_status("润色失败"))

    # --- 升级版语音合成核心 (融入三大参数) ---
    async def _generate_audio_task(self, text, output_file):
        selected_name = self.selected_voice_key.get()
        voice_id = VOICE_MAP.get(selected_name, "zh-CN-XiaoxiaoNeural")
        
        processed_text = re.sub(r'\[.*?\|(.*?)\]', r'\1', text)
        
        # 提取参数值并格式化为 Edge-TTS 接受的字符串标准
        r = int(self.rate_var.get())
        v = int(self.volume_var.get())
        p = int(self.pitch_var.get())
        
        rate_str = f"{r:+d}%"
        vol_str = f"{v:+d}%"
        pitch_str = f"{p:+d}Hz"
        
        # 将参数一并发送给引擎
        communicate = edge_tts.Communicate(
            text=processed_text, 
            voice=voice_id,
            rate=rate_str,
            volume=vol_str,
            pitch=pitch_str
        )
        await communicate.save(output_file)

    def play_audio(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text: return
        self.stop_audio()
        self.is_generating = True
        self.update_status(f"正在合成 ({self.selected_voice_key.get()})...")
        
        def run_gen():
            try:
                future = asyncio.run_coroutine_threadsafe(
                    self._generate_audio_task(text, self.temp_audio_file), self.loop
                )
                future.result() 
                if not self.is_generating: return
                self.root.after(0, self._play_sound)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("合成错误", str(e)))
                self.root.after(0, lambda: self.update_status("合成出错"))

        threading.Thread(target=run_gen).start()

    def _play_sound(self):
        try:
            import pygame
            pygame.mixer.init()
            pygame.mixer.music.load(self.temp_audio_file)
            pygame.mixer.music.play()
            self.is_playing = True
            self.is_generating = False
            self.update_status("正在播放...")
        except Exception as e:
            messagebox.showerror("播放错误", str(e))

    def stop_audio(self):
        self.is_generating = False 
        try:
            import pygame
            pygame.mixer.init()
            if pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
                pygame.mixer.music.unload()
        except:
            pass
        self.is_playing = False
        self.update_status("已停止")

    def export_audio(self, fmt):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text: return

        ext = ".mp3" if fmt == "mp3" else ".wav"
        save_path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[(f"{fmt.upper()} File", f"*{ext}")])
        if not save_path: return

        self.update_status(f"正在导出为 {fmt}...")

        def run_export():
            try:
                temp_mp3 = "temp_export.mp3"
                future = asyncio.run_coroutine_threadsafe(
                    self._generate_audio_task(text, temp_mp3), self.loop
                )
                future.result()

                if fmt == "mp3":
                    import shutil
                    shutil.move(temp_mp3, save_path)
                    
                elif fmt == "wav":
                    self.root.after(0, lambda: self.update_status("正在转换格式 (FFmpeg)..."))
                    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
                    cmd = [
                        ffmpeg_exe, "-y",
                        "-i", temp_mp3,
                        "-acodec", "pcm_s16le",
                        "-ar", "44100", 
                        "-ac", "2", 
                        save_path
                    ]
                    subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    if os.path.exists(temp_mp3):
                        os.remove(temp_mp3)

                self.root.after(0, lambda: messagebox.showinfo("成功", f"导出成功！\n保存路径: {save_path}"))
                self.root.after(0, lambda: self.update_status("导出完成"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("导出失败", f"错误详情:\n{str(e)}"))
                self.root.after(0, lambda: self.update_status("导出失败"))

        threading.Thread(target=run_export).start()

if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = TTSApp(root)
    root.mainloop()
