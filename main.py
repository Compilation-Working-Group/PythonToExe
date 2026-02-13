import os
import sys
import asyncio
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import docx
import edge_tts
from openai import OpenAI

# 默认配置
DEFAULT_DEEPSEEK_URL = "https://api.deepseek.com"

class TTSApp:
    def __init__(self, root):
        self.root = root
        # 1. 修改标题，增加作者信息
        self.root.title("DeepSeek 智能语音合成助手 - 作者: Yu JinQuan")
        
        # 设置合理的初始大小
        window_width = 850
        window_height = 600
        self.center_window(window_width, window_height)
        self.root.minsize(800, 500)
        
        # 变量初始化
        self.is_playing = False
        self.is_generating = False 
        self.temp_audio_file = "temp_preview.mp3"
        self.loop = asyncio.new_event_loop()
        
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
        # === 布局核心策略：三明治布局 ===

        # 1. 顶部区域
        frame_top = tk.LabelFrame(self.root, text="文件操作", padx=10, pady=5)
        frame_top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))
        
        tk.Button(frame_top, text="📂 导入文本/Word", command=self.import_file).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_top, text="🗑️ 清空内容", command=self.clear_text, bg="#ffebee").pack(side=tk.LEFT, padx=5)

        # 2. 底部区域 (倒序添加)
        
        # 2.1 状态栏与作者信息 (最底部)
        frame_status = tk.Frame(self.root, bd=1, relief=tk.SUNKEN, bg="#f0f0f0")
        frame_status.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = tk.Label(frame_status, text="状态: 就绪", anchor=tk.W, bg="#f0f0f0")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # 增加作者信息标签
        tk.Label(frame_status, text="Author: Yu JinQuan", anchor=tk.E, bg="#f0f0f0", fg="#666").pack(side=tk.RIGHT, padx=10)

        # 2.2 语音合成与导出区
        frame_bottom = tk.LabelFrame(self.root, text="语音合成与导出", padx=10, pady=5)
        frame_bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5, 10))
        
        tk.Button(frame_bottom, text="▶️ 生成并播放", command=self.play_audio, bg="#e8f5e9", width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="⏹️ 停止 / 重置", command=self.stop_audio, bg="#ffcdd2", width=12).pack(side=tk.LEFT, padx=5)
        tk.Frame(frame_bottom, width=2, bg="#ccc").pack(side=tk.LEFT, fill=tk.Y, padx=15)
        
        tk.Button(frame_bottom, text="💾 导出 MP3", command=lambda: self.export_audio("mp3")).pack(side=tk.LEFT, padx=5)
        # 修改为导出 WAV
        tk.Button(frame_bottom, text="🎵 导出 WAV", command=lambda: self.export_audio("wav")).pack(side=tk.LEFT, padx=5)

        # 2.3 AI 润色区
        frame_ai = tk.LabelFrame(self.root, text="DeepSeek AI 润色", padx=10, pady=5)
        frame_ai.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        
        tk.Label(frame_ai, text="提示: 将文本改写为更自然的口语风格").pack(side=tk.LEFT)
        tk.Button(frame_ai, text="✨ 开始智能润色", command=self.run_deepseek_polish, bg="#e3f2fd", fg="#0d47a1").pack(side=tk.RIGHT, padx=5)

        # 3. 中间文本区
        self.text_area = scrolledtext.ScrolledText(self.root, font=("Microsoft YaHei", 12), wrap=tk.WORD)
        self.text_area.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=10, pady=5)

    def update_status(self, text):
        self.status_label.config(text=f"状态: {text}")
        self.root.update_idletasks()

    # --- 功能函数 ---
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
                    {"role": "system", "content": "你是一个专业的配音文案润色专家。请将用户输入的文本修改为适合朗读的口语化文案，去除生硬的书面语，增加自然的连接词。请直接输出润色后的结果，不要包含任何解释性语言。"},
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

    async def _generate_audio_task(self, text, output_file):
        voice = "zh-CN-XiaoxiaoNeural"
        communicate = edge_tts.Communicate(text, voice)
        await communicate.save(output_file)

    def play_audio(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text: return
        
        self.stop_audio()
        self.is_generating = True
        self.update_status("正在合成语音 (Edge-TTS)...")
        
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
            messagebox.showerror("播放错误", f"无法播放音频: {e}")

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

        # 设置后缀名
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
                    self.root.after(0, lambda: self.update_status("正在转换格式 (MoviePy)..."))
                    from moviepy.editor import AudioFileClip
                    
                    # 使用 moviepy 将 MP3 转换为 WAV
                    audio = AudioFileClip(temp_mp3)
                    # pcm_s16le 是最通用的 WAV 编码
                    audio.write_audiofile(save_path, codec='pcm_s16le', logger=None)
                    audio.close()
                    os.remove(temp_mp3)

                self.root.after(0, lambda: messagebox.showinfo("成功", f"导出成功！\n保存路径: {save_path}"))
                self.root.after(0, lambda: self.update_status("导出完成"))
            
            except ImportError:
                 self.root.after(0, lambda: messagebox.showerror("组件缺失", "导出音频需要 moviepy 库。"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("导出失败", f"错误详情:\n{str(e)}"))
                self.root.after(0, lambda: self.update_status("导出失败"))

        threading.Thread(target=run_export).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = TTSApp(root)
    root.mainloop()
