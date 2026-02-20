import os
import json
import re
import xml.sax.saxutils as saxutils
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import azure.cognitiveservices.speech as speechsdk

# 尝试在 Windows 系统中开启高 DPI 感知，防止界面模糊
if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

# 尝试导入音频播放模块
try:
    import pygame
    pygame.mixer.init()
    AUDIO_SUPPORTED = True
except ImportError:
    AUDIO_SUPPORTED = False
    print("未安装 pygame，试听功能将被禁用。")

# 尝试导入 Word 读取模块
try:
    import docx
    DOCX_SUPPORTED = True
except ImportError:
    DOCX_SUPPORTED = False
    print("未安装 python-docx，Word 导入功能将受限。")

# ================= 配置与全局变量 =================
CONFIG_FILE = "tts_config.json"

VOICES = {
    # ---- 🇨🇳 大陆普通话 (女声) ----
    "晓晓 (标准女声 - 温暖亲切，推荐讲解)": "zh-CN-XiaoxiaoNeural",
    "晓伊 (标准女声 - 轻松自然，知性优雅)": "zh-CN-XiaoyiNeural",
    "晓辰 (标准女声 - 休闲随性，适合生活)": "zh-CN-XiaochenNeural",
    "晓涵 (标准女声 - 情感丰富，适合朗读)": "zh-CN-XiaohanNeural",
    "晓睿 (成熟女声 - 沉稳专业，适合新闻)": "zh-CN-XiaoruiNeural",
    "晓双 (儿童女声 - 可爱童音，适合故事)": "zh-CN-XiaoshuangNeural",
    "晓萱 (标准女声 - 柔和优美，从容不迫)": "zh-CN-XiaoxuanNeural",
    "晓墨 (知性女声 - 角色扮演，清晰有力)": "zh-CN-XiaomoNeural",
    "晓秋 (温柔女声 - 阅读旁白，唯美抒情)": "zh-CN-XiaoqiuNeural",
    "晓柔 (活泼女声 - 甜美可爱，撒娇感)": "zh-CN-XiaorouNeural",
    "晓甄 (成熟女声 - 严厉干练，适合批评)": "zh-CN-XiaozhenNeural",

    # ---- 🇨🇳 大陆普通话 (男声) ----
    "云希 (年轻男声 - 活泼阳光，推荐测试)": "zh-CN-YunxiNeural",
    "云健 (成熟男声 - 稳重影视，适合纪录片)": "zh-CN-YunjianNeural",
    "云扬 (标准男声 - 新闻播报，字正腔圆)": "zh-CN-YunyangNeural",
    "云泽 (成熟男声 - 磁性叙事，抓人耳朵)": "zh-CN-YunzeNeural",
    "云枫 (年轻男声 - 爽朗热情，活力四射)": "zh-CN-YunfengNeural",
    "云皓 (年轻男声 - 轻松愉悦，语速轻快)": "zh-CN-YunhaoNeural",
    "云野 (成熟男声 - 浑厚有力，深沉沧桑)": "zh-CN-YunyeNeural",

    # ---- 🌶️ 地方口音/方言 ----
    "辽宁晓北 (东北话女声 - 幽默豪爽)": "zh-CN-liaoning-XiaobeiNeural",
    "陕西晓妮 (陕西话女声 - 纯正自然)": "zh-CN-shaanxi-XiaoniNeural",
    "四川云希 (四川话男声 - 亲切接地气)": "zh-CN-sichuan-YunxiNeural",

    # ---- 🍵 港台地区 ----
    "台湾晓臻 (甜美女生 - 台湾腔国语)": "zh-TW-HsiaoChenNeural",
    "台湾云哲 (温和男生 - 台湾腔国语)": "zh-TW-YunJheNeural",
    "香港晓曼 (标准粤语女声 - 自然流畅)": "zh-HK-HiuMaanNeural",
    "香港云龙 (标准粤语男声 - 经典港剧音)": "zh-HK-WanLungNeural",

    # ---- 🇺🇸 英语 - 美国 (English US) ----
    "Jenny (美国女声 - 友好清晰，适合讲解)": "en-US-JennyNeural",
    "Aria (美国女声 - 情感丰富，自然流畅)": "en-US-AriaNeural",
    "Guy (美国男声 - 专业沉稳，适合纪录片)": "en-US-GuyNeural",
    "Davis (美国男声 - 活泼热情，适合对话)": "en-US-DavisNeural",
    "Jane (美国女声 - 温和专业播报)": "en-US-JaneNeural",
    "Jason (美国男声 - 成熟稳重有力)": "en-US-JasonNeural",
    "Sara (美国女声 - 年轻活力美音)": "en-US-SaraNeural",
    "Tony (美国男声 - 清晰有力播报)": "en-US-TonyNeural",
    "Amber (美国女声 - 青春洋溢美音)": "en-US-AmberNeural",

    # ---- 🇬🇧 英语 - 英国 (English UK) ----
    "Sonia (英国女声 - 优雅纯正英音)": "en-GB-SoniaNeural",
    "Ryan (英国男声 - 专业英式播报)": "en-GB-RyanNeural",
    "Libby (英国女声 - 轻松自然英音)": "en-GB-LibbyNeural",
    "Oliver (英国男声 - 活力年轻英音)": "en-GB-OliverNeural",

    # ---- 🇦🇺 英语 - 澳洲/加拿大 ----
    "Natasha (澳洲女声 - 地道澳音)": "en-AU-NatashaNeural",
    "William (澳洲男声 - 自然清晰澳音)": "en-AU-WilliamNeural",
    "Clara (加拿大女声 - 温和自然)": "en-CA-ClaraNeural",
    "Liam (加拿大男声 - 专业清晰)": "en-CA-LiamNeural"
}

PLACEHOLDER_TEXT = """【微课语音生成专业版 - 使用指南】
1. 首次使用：请在【右侧控制面板】填写您的 Azure API 密钥和区域代码。
2. 文本输入：点击此处直接输入内容，或使用上方“导入”按钮读取本地的 TXT/Word 文档。
3. 读音修正：选中生僻字或多音字（如：重），右键点击“修正读音”，输入拼音（如 zhong4）。
4. 试听导出：在右侧面板调节语速/音调，点击“试听”，满意后选择导出 MP3 或 无损 WAV。
5. 撤销/重做：支持系统级快捷键 Ctrl+Z, Ctrl+Y，也可使用鼠标右键菜单。
（鼠标点击此处开始输入，本提示将自动消失...）"""

is_paused = False
is_playing = False
temp_preview_file = "temp_preview_audio.mp3"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(key, region):
    config = {"speech_key": key, "service_region": region}
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f)
    except Exception:
        pass

def process_custom_pinyin(raw_text):
    parts = re.split(r'(\[.*?\|.*?\])', raw_text)
    ssml_result = ""
    for part in parts:
        if part.startswith('[') and part.endswith(']') and '|' in part:
            inner = part[1:-1]
            try:
                word, pinyin = inner.split('|', 1)
                esc_word = saxutils.escape(word)
                formatted_pinyin = re.sub(r'([a-zA-Z])(\d)', r'\1 \2', pinyin)
                formatted_pinyin = re.sub(r'\s+', ' ', formatted_pinyin).strip()
                ssml_result += f'<phoneme alphabet="sapi" ph="{formatted_pinyin}">{esc_word}</phoneme>'
            except ValueError:
                ssml_result += saxutils.escape(part)
        else:
            ssml_result += saxutils.escape(part)
    return ssml_result

def generate_ssml(text, voice_name, rate, pitch, volume):
    rate_str = f"{rate}%" if rate <= 0 else f"+{rate}%"
    pitch_str = f"{pitch}%" if pitch <= 0 else f"+{pitch}%"
    processed_text = process_custom_pinyin(text)
    lang_code = voice_name[:5] 
    
    ssml = f"""<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="{lang_code}">
        <voice name="{voice_name}">
            <prosody rate="{rate_str}" pitch="{pitch_str}" volume="{volume}">
                {processed_text}
            </prosody>
        </voice>
    </speak>"""
    return ssml

def text_to_speech_file(text, file_path, voice_name, rate, pitch, volume):
    speech_key = entry_key.get().strip()
    service_region = entry_region.get().strip()
    
    if not speech_key or not service_region:
        return False, "请在右侧面板填写 Azure API 密钥和区域代码！"
        
    save_config(speech_key, service_region)
    
    try:
        speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)
        if file_path.lower().endswith('.wav'):
            speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Riff24Khz16BitMonoPcm)
        else:
            speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio16Khz128KBitRateMonoMp3)
            
        audio_config = speechsdk.audio.AudioOutputConfig(filename=file_path)
        speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
        
        ssml_string = generate_ssml(text, voice_name, rate, pitch, volume)
        result = speech_synthesizer.speak_ssml_async(ssml_string).get()
        
        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            return True, "合成成功！"
        elif result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = result.cancellation_details
            error_msg = f"合成被拒绝: {cancellation_details.reason}"
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                error_msg += f"\n详细原因: {cancellation_details.error_details}\n(提示: 请检查密钥/区域是否填写正确)"
            return False, error_msg
        else:
            return False, f"未知错误: {result.reason}"
    except Exception as e:
        return False, f"发生异常: {str(e)}"

# ================= 占位符与清空控制 =================
def remove_placeholder(event=None):
    if getattr(text_input, "is_placeholder", False):
        text_input.delete("1.0", tk.END)
        text_input.config(fg="black")
        text_input.is_placeholder = False
        text_input.edit_reset()

def add_placeholder(event=None):
    content = text_input.get("1.0", tk.END).strip()
    if not content:
        text_input.insert("1.0", PLACEHOLDER_TEXT)
        text_input.config(fg="gray")
        text_input.is_placeholder = True
        text_input.edit_reset()

def check_empty_input():
    if getattr(text_input, "is_placeholder", False) or not text_input.get("1.0", tk.END).strip():
        messagebox.showwarning("提示", "内容为空，请先输入或导入需要转换的文本！")
        return True
    return False

# ================= 终极键鼠融合接管 =================
def perform_action(action):
    text_input.focus_set()
    def _do_action():
        try:
            if action == "copy": text_input.event_generate("<<Copy>>")
            elif action == "cut": text_input.event_generate("<<Cut>>")
            elif action == "paste": text_input.event_generate("<<Paste>>")
            elif action == "select_all": text_input.tag_add("sel", "1.0", "end")
            elif action == "undo": text_input.event_generate("<<Undo>>")
            elif action == "redo": text_input.event_generate("<<Redo>>")
        except tk.TclError:
            pass 
    text_input.after(20, _do_action)

def on_paste_intercept(event):
    remove_placeholder()
    try:
        if text_input.tag_ranges("sel"):
            text_input.delete(tk.SEL_FIRST, tk.SEL_LAST)
    except tk.TclError:
        pass
    return None 

def on_select_all_intercept(event):
    text_input.tag_add("sel", "1.0", "end")
    return "break" 

def popup_context_menu(event):
    text_input.focus_set()
    remove_placeholder()
    try:
        if not text_input.tag_ranges("sel"):
            text_input.mark_set("insert", f"@{event.x},{event.y}")
    except tk.TclError:
        pass
    context_menu.tk_popup(event.x_root, event.y_root)

# ================= UI 交互、重置与作者信息 =================
def reset_params():
    rate_scale.set(0)
    pitch_scale.set(0)
    volume_scale.set(100)
    status_label.config(text="⚙️ 参数已重置为默认", bootstyle=SUCCESS)

def show_about():
    about_text = (
        "微课语音生成专业版 (ChemTTS Pro)\n"
        "==========================\n\n"
        "👨‍🏫 作者：俞金泉 (Yu)\n"
        "🏫 单位：金塔县中学\n"
        "🧪 职务：化学教研组长 / 高中化学名师工作室主持人\n"
        "🎓 班级：高二(1)班班主任\n\n"
        "💡 专为一线教学、微课制作与新高考教案定制开发。\n"
        "✅ 支持双语混合、SSML注音修正、无损 WAV 导出。\n\n"
        "✨ 祝老师们工作顺利，桃李满天下！"
    )
    messagebox.showinfo("关于软件与作者", about_text)

def check_playback_status():
    global is_playing, is_paused
    if not AUDIO_SUPPORTED or not is_playing: return
        
    if not pygame.mixer.music.get_busy() and not is_paused:
        is_playing = False
        status_label.config(text="试听已结束", bootstyle=SUCCESS)
        btn_pause.config(text="⏸ 暂停")
    else:
        root.after(500, check_playback_status)

def stop_playback():
    global is_paused, is_playing
    if AUDIO_SUPPORTED:
        pygame.mixer.music.stop()
        try: pygame.mixer.music.unload()
        except AttributeError: pass
    is_paused = False
    is_playing = False
    btn_pause.config(text="⏸ 暂停")
    status_label.config(text="已停止播放", bootstyle=SECONDARY)

def on_preview():
    if check_empty_input(): return
    global is_playing, is_paused
    if not AUDIO_SUPPORTED:
        messagebox.showerror("错误", "未找到 pygame 模块，无法试听。")
        return

    text = text_input.get("1.0", tk.END).strip()
    stop_playback()
    status_label.config(text="正在呼叫 Azure 生成试听音频...", bootstyle=INFO)
    root.update()

    selected_voice = VOICES[voice_combo.get()]
    success, msg = text_to_speech_file(text, temp_preview_file, selected_voice, rate_scale.get(), pitch_scale.get(), volume_scale.get())
    
    if success:
        status_label.config(text="正在播放试听...", bootstyle=SUCCESS)
        pygame.mixer.music.load(temp_preview_file)
        pygame.mixer.music.play()
        is_playing = True
        is_paused = False
        check_playback_status()
    else:
        status_label.config(text="试听生成失败", bootstyle=DANGER)
        messagebox.showerror("生成失败", msg)

def on_toggle_pause():
    global is_paused, is_playing
    if not AUDIO_SUPPORTED or not is_playing: return
        
    if is_paused:
        pygame.mixer.music.unpause()
        btn_pause.config(text="⏸ 暂停")
        is_paused = False
        status_label.config(text="正在播放试听...", bootstyle=SUCCESS)
    else:
        pygame.mixer.music.pause()
        btn_pause.config(text="▶ 继续")
        is_paused = True
        status_label.config(text="试听已暂停", bootstyle=WARNING)

def on_convert(audio_format="mp3"):
    if check_empty_input(): return
    text = text_input.get("1.0", tk.END).strip()
    stop_playback()
    
    if audio_format == "wav":
        save_path = filedialog.asksaveasfilename(
            title="保存无损 WAV 音频", defaultextension=".wav",
            filetypes=[("WAV 无损音频", "*.wav"), ("所有文件", "*.*")], initialfile="化学微课语音_01.wav"
        )
    else:
        save_path = filedialog.asksaveasfilename(
            title="保存 MP3 音频", defaultextension=".mp3",
            filetypes=[("MP3 音频", "*.mp3"), ("所有文件", "*.*")], initialfile="化学微课语音_01.mp3"
        )
        
    if not save_path: return
        
    status_label.config(text=f"正在导出 {audio_format.upper()} 文件，请稍候...", bootstyle=INFO)
    root.update()
    
    selected_voice = VOICES[voice_combo.get()]
    success, msg = text_to_speech_file(text, save_path, selected_voice, rate_scale.get(), pitch_scale.get(), volume_scale.get())
    
    if success:
        status_label.config(text=f"导出成功！保存在: {save_path}", bootstyle=SUCCESS)
        messagebox.showinfo("成功", f"语音合成成功！文件位于:\n{save_path}")
    else:
        status_label.config(text="合成失败", bootstyle=DANGER)
        messagebox.showerror("生成失败", msg)

def on_import_file():
    file_path = filedialog.askopenfilename(
        title="导入文档",
        filetypes=[
            ("支持的文档 (TXT/Word)", ("*.txt", "*.docx")), 
            ("文本文件", "*.txt"), 
            ("Word文档", "*.docx"), 
            ("所有文件", "*.*")
        ]
    )
    if not file_path: return

    try:
        content = ""
        if file_path.lower().endswith('.docx'):
            if not DOCX_SUPPORTED:
                messagebox.showerror("缺少库", "未安装 python-docx 库。\n请在终端运行: pip3 install python-docx")
                return
            doc = docx.Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        else:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
        
        remove_placeholder()
        text_input.insert(tk.END, content + "\n")
        status_label.config(text=f"成功导入: {os.path.basename(file_path)}", bootstyle=SUCCESS)
    except Exception as e:
        messagebox.showerror("读取失败", f"无法读取该文件: {str(e)}")

def on_export_txt():
    if check_empty_input(): return
    content = text_input.get("1.0", tk.END).strip()
        
    file_path = filedialog.asksaveasfilename(
        title="保存文稿", defaultextension=".txt",
        filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")], initialfile="微课文稿_备份.txt"
    )
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            status_label.config(text=f"文稿已保存至: {os.path.basename(file_path)}", bootstyle=SUCCESS)
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存: {str(e)}")

def on_correct_pinyin():
    if getattr(text_input, "is_placeholder", False):
        messagebox.showinfo("提示", "请先输入或导入课件文本！")
        return

    try:
        selected_text = text_input.selection_get()
    except tk.TclError:
        messagebox.showinfo("提示", "请先用鼠标选中需要注音的汉字！")
        return

    pinyin = simpledialog.askstring("修正读音", f"请输入【{selected_text}】的拼音和数字声调\n(例如：zhong4)：")
    if pinyin:
        pinyin = pinyin.strip()
        try:
            start_idx = text_input.index(tk.SEL_FIRST)
            end_idx = text_input.index(tk.SEL_LAST)
            text_input.delete(start_idx, end_idx)
            text_input.insert(start_idx, f"[{selected_text}|{pinyin}]")
        except tk.TclError:
            pass

def on_clear():
    text_input.delete("1.0", tk.END)
    add_placeholder()
    root.focus()
    status_label.config(text="已清空", bootstyle=SECONDARY)


# ================= 界面设计部分 (全新分栏布局) =================
# 放大初始窗口，适合宽屏展现
root = tb.Window(title="微课语音生成专业版 (多平台/版权所有)", themename="litera", size=(1000, 680))
root.minsize(850, 600)

saved_config = load_config()

# ---- 顶层主容器 (包含左右分栏) ----
main_frame = tb.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

# ================= 左侧：工作区 =================
left_panel = tb.Frame(main_frame)
left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))

# 工具栏
toolbar_frame = tb.Frame(left_panel)
toolbar_frame.pack(fill=tk.X, pady=(0, 8))

btn_import = tb.Button(toolbar_frame, text="📂 导入(TXT/Word)", command=on_import_file, bootstyle=(INFO, OUTLINE))
btn_import.pack(side=tk.LEFT, padx=(0, 10))

btn_export = tb.Button(toolbar_frame, text="💾 保存为 TXT", command=on_export_txt, bootstyle=(WARNING, OUTLINE))
btn_export.pack(side=tk.LEFT, padx=(0, 10))

btn_clear = tb.Button(toolbar_frame, text="🗑 清空", command=on_clear, bootstyle=(SECONDARY, OUTLINE))
btn_clear.pack(side=tk.LEFT)

btn_pinyin = tb.Button(toolbar_frame, text="✍ 修正选中字读音", command=on_correct_pinyin, bootstyle=(PRIMARY, OUTLINE))
btn_pinyin.pack(side=tk.RIGHT)

# 文本输入框 (自动填满左侧剩余空间)
text_input = tk.Text(left_panel, font=("微软雅黑", 12), wrap=tk.WORD, undo=True, maxundo=-1, relief=tk.FLAT, bg="#F8F9FA", padx=10, pady=10)
text_input.pack(fill=tk.BOTH, expand=True)

text_input.bind("<FocusIn>", remove_placeholder)
text_input.bind("<FocusOut>", add_placeholder)
text_input.bind("<<Paste>>", on_paste_intercept) 
text_input.bind("<Control-a>", on_select_all_intercept) 
text_input.bind("<Control-A>", on_select_all_intercept)
add_placeholder()

context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="✍ 修正选中字读音", command=on_correct_pinyin)
context_menu.add_separator()
context_menu.add_command(label="↶ 撤销 (Undo)", command=lambda: perform_action("undo"))
context_menu.add_command(label="↷ 重做 (Redo)", command=lambda: perform_action("redo"))
context_menu.add_separator()
context_menu.add_command(label="✂ 剪切 (Cut)", command=lambda: perform_action("cut"))
context_menu.add_command(label="📋 复制 (Copy)", command=lambda: perform_action("copy"))
context_menu.add_command(label="📝 粘贴 (Paste)", command=lambda: perform_action("paste"))
context_menu.add_separator()
context_menu.add_command(label="☑ 全选 (Select All)", command=lambda: perform_action("select_all"))
context_menu.add_command(label="🗑 清空内容", command=on_clear)

text_input.bind("<Button-3>", popup_context_menu)
text_input.bind("<Button-2>", popup_context_menu)

# ================= 右侧：控制区 =================
# 设定一个大概的宽度，保证控制面板不会被文本框挤压
right_panel = tb.Frame(main_frame)
right_panel.pack(side=tk.RIGHT, fill=tk.Y)

# 1. API 配置面板
api_frame = tb.Labelframe(right_panel, text=" ⚙️ Azure 接口配置 ", padding=12, bootstyle=INFO)
api_frame.pack(fill=tk.X, pady=(0, 15))

tb.Label(api_frame, text="API 密钥:", font=("微软雅黑", 9)).pack(anchor="w", pady=(0, 2))
entry_key = tb.Entry(api_frame, show="*")
entry_key.pack(fill=tk.X, pady=(0, 10))
entry_key.insert(0, saved_config.get("speech_key", ""))

tb.Label(api_frame, text="区域 (Region):", font=("微软雅黑", 9)).pack(anchor="w", pady=(0, 2))
entry_region = tb.Entry(api_frame)
entry_region.pack(fill=tk.X)
entry_region.insert(0, saved_config.get("service_region", ""))

# 2. 语音及参数面板
voice_frame = tb.Labelframe(right_panel, text=" 🗣️ 语音与参数 ", padding=12, bootstyle=PRIMARY)
voice_frame.pack(fill=tk.X, pady=(0, 15))

tb.Label(voice_frame, text="发音人:", font=("微软雅黑", 9, "bold")).pack(anchor="w", pady=(0, 5))
voice_combo = tb.Combobox(voice_frame, values=list(VOICES.keys()), state="readonly", bootstyle=PRIMARY)
voice_combo.pack(fill=tk.X, pady=(0, 10))
voice_combo.current(0)

tb.Label(voice_frame, text="语速:").pack(anchor="w")
rate_scale = tb.Scale(voice_frame, from_=-50, to=50, orient=tk.HORIZONTAL, bootstyle=INFO)
rate_scale.set(0)
rate_scale.pack(fill=tk.X, pady=(0, 8))

tb.Label(voice_frame, text="音调:").pack(anchor="w")
pitch_scale = tb.Scale(voice_frame, from_=-50, to=50, orient=tk.HORIZONTAL, bootstyle=WARNING)
pitch_scale.set(0)
pitch_scale.pack(fill=tk.X, pady=(0, 8))

tb.Label(voice_frame, text="音量:").pack(anchor="w")
volume_scale = tb.Scale(voice_frame, from_=0, to=100, orient=tk.HORIZONTAL, bootstyle=SUCCESS)
volume_scale.set(100)
volume_scale.pack(fill=tk.X, pady=(0, 10))

btn_reset = tb.Button(voice_frame, text="↺ 恢复默认参数", command=reset_params, bootstyle=(SECONDARY, OUTLINE))
btn_reset.pack(fill=tk.X)

# 3. 试听控制面板
preview_frame = tb.Labelframe(right_panel, text=" 🎧 试听控制 ", padding=12, bootstyle=WARNING)
preview_frame.pack(fill=tk.X, pady=(0, 15))

# 使用网格将试听按钮放一行
btn_play = tb.Button(preview_frame, text="🔊 试听", command=on_preview, bootstyle=WARNING)
btn_play.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

btn_pause = tb.Button(preview_frame, text="⏸ 暂停", command=on_toggle_pause, bootstyle=(SECONDARY, OUTLINE))
btn_pause.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

btn_stop = tb.Button(preview_frame, text="⏹ 停止", command=stop_playback, bootstyle=(DANGER, OUTLINE))
btn_stop.pack(side=tk.LEFT, expand=True, fill=tk.X)

# 4. 导出操作面板
export_frame = tb.Labelframe(right_panel, text=" 💾 生成音频 ", padding=12, bootstyle=SUCCESS)
export_frame.pack(fill=tk.X)

convert_btn_mp3 = tb.Button(export_frame, text="🎵 导出 MP3 (推荐)", command=lambda: on_convert("mp3"), bootstyle=SUCCESS)
convert_btn_mp3.pack(fill=tk.X, pady=(0, 8))

convert_btn_wav = tb.Button(export_frame, text="🎚️ 导出 WAV (无损)", command=lambda: on_convert("wav"), bootstyle=PRIMARY)
convert_btn_wav.pack(fill=tk.X)


# ================= 底部：状态栏 =================
bottom_bar = tb.Frame(root)
bottom_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(0, 10))

status_label = tb.Label(bottom_bar, text="准备就绪", font=("微软雅黑", 9), bootstyle=SECONDARY)
status_label.pack(side=tk.LEFT)

author_label = tb.Label(bottom_bar, text="© 俞晋全 | 金塔县中学高中化学名师工作室", font=("微软雅黑", 9), bootstyle=SECONDARY, cursor="hand2")
author_label.pack(side=tk.RIGHT)
author_label.bind("<Button-1>", lambda e: show_about())

def on_closing():
    stop_playback()
    if os.path.exists(temp_preview_file):
        try: os.remove(temp_preview_file)
        except OSError: pass
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.focus()
root.mainloop()
