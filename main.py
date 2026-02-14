import os
import sys
import asyncio
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
from tkinter import ttk
import docx
import edge_tts
from openai import OpenAI
import imageio_ffmpeg
import re  # 新增正则表达式库，用于处理同音字替换

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
        
        window_width = 950
        window_height = 700
        self.center_window(window_width, window_height)
        self.root.minsize(800, 500)
        
        self.is_playing = False
        self.is_generating = False 
        self.temp_audio_file = "temp_preview.mp3"
        self.loop = asyncio.new_event_loop()
        
        self.selected_voice_key = tk.StringVar(value="晓晓 (女声 - 活泼/默认)")
        
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
        frame_top = tk.LabelFrame(self.root, text="文件与编辑", padx=10, pady=5)
        frame_top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))
        
        tk.Button(frame_top, text="📂 导入文本/Word", command=self.import_file).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_top, text="🗑️ 清空内容", command=self.clear_text, bg="#ffebee").pack(side=tk.LEFT, padx=5)
        
        # 恢复并优化的多音字修正按钮
        tk.Frame(frame_top, width=20).pack(side=tk.LEFT) 
        tk.Label(frame_top, text="选中多音字后点击 ->", fg="gray").pack(side=tk.LEFT)
        tk.Button(frame_top, text="📝 修正选中字读音 (同音字法)", command=self.fix_pronunciation, bg="#fff3e0").pack(side=tk.LEFT, padx=5)

        # 2. 底部控制区 (倒序)
        frame_status = tk.Frame(self.root, bd=1, relief=tk.SUNKEN, bg="#f0f0f0")
        frame_status.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = tk.Label(frame_status, text="状态: 就绪", anchor=tk.W, bg="#f0f0f0")
        self.status_label.pack(side=tk.LEFT, padx=5)
        tk.Label(frame_status, text="Author: Yu JinQuan", anchor=tk.E, bg="#f0f0f0", fg="#666").pack(side=tk.RIGHT, padx=10)

        frame_bottom = tk.LabelFrame(self.root, text="语音控制与导出", padx=10, pady=5)
        frame_bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5, 10))
        
        tk.Label(frame_bottom, text="选择语音:").pack(side=tk.LEFT, padx=(5, 0))
        voice_combo = ttk.Combobox(frame_bottom, textvariable=self.selected_voice_key, values=list(VOICE_MAP.keys()), state="readonly", width=25)
        voice_combo.pack(side=tk.LEFT, padx=5)

        tk.Frame(frame_bottom, width=2, bg="#ccc").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        tk.Button(frame_bottom, text="▶️ 生成并播放", command=self.play_audio, bg="#e8f5e9", width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="⏹️ 停止", command=self.stop_audio, bg="#ffcdd2", width=8).pack(side=tk.LEFT, padx=5)
        
        tk.Frame(frame_bottom, width=2, bg="#ccc").pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        tk.Button(frame_bottom, text="💾 导出 MP3", command=lambda: self.export_audio("mp3")).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="🎵 导出 WAV", command=lambda: self.export_audio("wav")).pack(side=tk.LEFT, padx=5)

        # 3. AI 润色区
        frame_ai = tk.LabelFrame(self.root, text="DeepSeek AI 润色", padx=10, pady=5)
        frame_ai.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        tk.Label(frame_ai, text="提示: 将文本改写为更自然的口语风格").pack(side=tk.LEFT)
        tk.Button(frame_ai, text="✨ 开始智能润色", command=self.run_deepseek_polish, bg="#e3f2fd", fg="#0d47a1").pack(side=tk.RIGHT, padx=5)

        # 4. 中间文本区
        self.text_area = scrolledtext.ScrolledText(self.root, font=("Microsoft YaHei", 12), wrap=tk.WORD)
        self.text_area.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=10, pady=5)

    def update_status(self, text):
        self.status_label.config(text=f"状态: {text}")
        self.root.update_idletasks()

    # --- 恢复同音字替换逻辑 ---
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
            # 格式化为 [原字|同音字]，例如 [单|善]
            replacement = f"[{selection}|{homophone.strip()}]"
            self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.text_area.insert(tk.INSERT, replacement)
            self.update_status(f"已设置同音字: {selection} -> {homophone}")

    # --- 文件操作 ---
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

    # --- DeepSeek ---
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

    # --- 语音合成核心 ---
    async def _generate_audio_task(self, text, output_file):
        selected_name = self.selected_voice_key.get()
        voice_id = VOICE_MAP.get(selected_name, "zh-CN-XiaoxiaoNeural")
        
        # 核心逻辑：利用正则表达式，将 [原字|同音字] 提取出同音字发给引擎
        # 比如：引擎收到的不再是 [单|善]老师，而是 善老师
        processed_text = re.sub(r'\[.*?\|(.*?)\]', r'\1', text)
        
        # 使用最纯净的文本进行发音，完美兼容所有声音模型
        communicate = edge_tts.Communicate(processed_text, voice_id)
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
    root = tk.Tk()
    app = TTSApp(root)
    root.mainloop()
