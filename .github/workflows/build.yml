import os
import sys
import json
import re
import datetime
from dataclasses import dataclass
from typing import Optional, List, Dict

from PySide6 import QtCore, QtWidgets
from docx import Document

APP_NAME = "PaperWriter"
APP_ORG = "YuJinQuanLab"
SETTINGS_FILE = "settings.json"

# DeepSeek defaults (OpenAI-compatible) [5](https://www.cnblogs.com/bibaodi/p/18925124)
DEFAULT_API_BASE = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"


# ---------------------------
# Linux preflight & Qt env fixups
# ---------------------------
def linux_preflight_check():
    """
    Check critical runtime libs before creating QApplication to avoid Qt core dump.
    Qt 6.5+ needs libxcb-cursor0 to load xcb platform plugin. [6](https://blog.csdn.net/m0_46144825/article/details/122462453)[7](https://fcitx-im.org/wiki/Input_method_related_environment_variables/zh-hans)
    """
    if not sys.platform.startswith("linux"):
        return

    import ctypes
    required = [
        ("libxcb-cursor.so.0", "sudo apt update && sudo apt install -y libxcb-cursor0"),
    ]
    missing = []
    for so_name, hint in required:
        try:
            ctypes.CDLL(so_name)
        except OSError:
            missing.append((so_name, hint))

    if missing:
        msg_lines = [
            "启动失败：系统缺少 Qt(xcb) 运行所需动态库。",
            "",
            "缺少：",
        ]
        for so_name, _ in missing:
            msg_lines.append(f"  - {so_name}")
        msg_lines += [
            "",
            "在 Ubuntu/Debian 上可执行：",
        ]
        hints = sorted(set(hint for _, hint in missing))
        for h in hints:
            msg_lines.append(f"  {h}")
        msg_lines += [
            "",
            "说明：Qt 6.5+ 加载 xcb 平台插件需要该库，否则会出现：",
            "  'From 6.5.0, xcb-cursor0 or libxcb-cursor0 is needed...'",
        ]
        print("\n".join(msg_lines), file=sys.stderr)
        sys.exit(1)


def _find_pyside_qt_plugins_dir() -> Optional[str]:
    try:
        import PySide6
        base = os.path.dirname(PySide6.__file__)
        qt_plugins = os.path.join(base, "Qt", "plugins")
        return qt_plugins if os.path.isdir(qt_plugins) else None
    except Exception:
        return None


def _choose_qt_im_module(qt_plugins_dir: Optional[str]) -> Optional[str]:
    """
    Prefer fcitx if its plugin exists, otherwise ibus.
    Reason: PySide6 commonly includes ibus plugin but not fcitx plugin,
    causing Chinese input failure on fcitx5 systems. [2](https://stackoverflow.com/questions/51778203/permission-error-when-trying-to-use-pyinstaller)[3](https://blog.csdn.net/qq_48379015/article/details/153739176)
    """
    if not qt_plugins_dir:
        return None
    pic_dir = os.path.join(qt_plugins_dir, "platforminputcontexts")
    if not os.path.isdir(pic_dir):
        return None

    files = set(os.listdir(pic_dir))
    has_fcitx = any("fcitx" in f and f.endswith(".so") for f in files)
    has_ibus = any("ibus" in f and f.endswith(".so") for f in files)

    if has_fcitx:
        return "fcitx"
    if has_ibus:
        return "ibus"
    return None


def qt_env_fixups():
    """
    Make Qt find its plugins when packaged (PyInstaller).
    Also set QT_IM_MODULE based on available input context plugins.
    """
    qt_plugins = _find_pyside_qt_plugins_dir()
    if qt_plugins:
        os.environ.setdefault("QT_PLUGIN_PATH", qt_plugins)
        os.environ.setdefault("QT_QPA_PLATFORM_PLUGIN_PATH", os.path.join(qt_plugins, "platforms"))

        # Pick input method module if user didn't set it
        if not os.environ.get("QT_IM_MODULE"):
            chosen = _choose_qt_im_module(qt_plugins)
            if chosen:
                os.environ["QT_IM_MODULE"] = chosen

    # Respect user's env; set a safe default only if unset
    if sys.platform.startswith("linux"):
        os.environ.setdefault("XMODIFIERS", os.environ.get("XMODIFIERS", "@im=fcitx"))


# ---------------------------
# Utilities
# ---------------------------
def user_config_dir() -> str:
    base = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.AppConfigLocation)
    path = os.path.join(base, APP_ORG, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def load_settings() -> dict:
    path = os.path.join(user_config_dir(), SETTINGS_FILE)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_settings(data: dict):
    path = os.path.join(user_config_dir(), SETTINGS_FILE)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def now_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clamp_text(s: str, max_len: int = 12000) -> str:
    s = s.strip()
    return s if len(s) <= max_len else s[:max_len] + "\n...[截断]..."


def parse_kv_instructions(text: str) -> Dict[str, str]:
    d = {}
    if not text.strip():
        return d
    parts = re.split(r"[;\n，,；]+", text.strip())
    for p in parts:
        if "=" in p:
            k, v = p.split("=", 1)
            k, v = k.strip(), v.strip()
            if k and v:
                d[k] = v
    return d


# ---------------------------
# LLM Client (OpenAI-compatible; DeepSeek supported)
# ---------------------------
@dataclass
class LLMConfig:
    api_base: str = DEFAULT_API_BASE
    api_key: str = ""
    model: str = DEFAULT_MODEL
    temperature: float = 0.6
    max_tokens: int = 1800


class OpenAICompatClient:
    """
    OpenAI-compatible Chat Completions client.
    DeepSeek: base_url can be https://api.deepseek.com (or /v1 for compatibility),
    endpoint is /chat/completions, models deepseek-chat / deepseek-reasoner. [5](https://www.cnblogs.com/bibaodi/p/18925124)
    """
    def __init__(self, cfg: LLMConfig, logger):
        self.cfg = cfg
        self.logger = logger

    def is_ready(self) -> bool:
        return bool(self.cfg.api_base.strip()) and bool(self.cfg.api_key.strip()) and bool(self.cfg.model.strip())

    def _build_url(self) -> str:
        base = self.cfg.api_base.rstrip("/")
        # If user already uses /v1, do not duplicate it.
        if base.endswith("/v1"):
            return base + "/chat/completions"

        # DeepSeek canonical endpoint uses /chat/completions under base_url. [5](https://www.cnblogs.com/bibaodi/p/18925124)
        if "api.deepseek.com" in base:
            return base + "/chat/completions"

        # Default OpenAI-compatible pattern
        return base + "/v1/chat/completions"

    def chat(self, messages: List[Dict[str, str]]) -> str:
        import urllib.request

        url = self._build_url()
        payload = {
            "model": self.cfg.model,
            "messages": messages,
            "temperature": self.cfg.temperature,
            "max_tokens": self.cfg.max_tokens,
            "stream": False,
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("Authorization", f"Bearer {self.cfg.api_key}")

        self.logger(f"[{now_str()}] 调用 LLM: {url} | model={self.cfg.model}")

        with urllib.request.urlopen(req, timeout=90) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
        obj = json.loads(raw)
        return obj["choices"][0]["message"]["content"].strip()


# ---------------------------
# Prompt Templates & Generators
# ---------------------------
GENRES = ["论文", "计划", "反思", "案例", "总结", "自定义"]

DEFAULT_OUTLINE_TEMPLATES = {
    "论文": [
        "题目",
        "摘要",
        "关键词",
        "1 引言",
        "2 研究方法",
        "3 结果",
        "4 讨论",
        "5 结论与展望",
        "参考文献（占位符）",
    ],
    "计划": [
        "标题",
        "一、背景与目标",
        "二、现状分析",
        "三、实施步骤",
        "四、时间安排（里程碑）",
        "五、风险与应对",
        "六、评估指标",
    ],
    "反思": [
        "标题",
        "一、事件/课堂概述",
        "二、目标与预期",
        "三、实际发生了什么（证据）",
        "四、问题诊断（原因分析）",
        "五、改进策略（可操作）",
        "六、后续跟进",
    ],
    "案例": [
        "标题",
        "一、背景",
        "二、问题描述",
        "三、关键决策/行动",
        "四、过程与结果",
        "五、经验与启示",
        "六、可迁移做法",
    ],
    "总结": [
        "标题",
        "一、总体回顾",
        "二、关键成果",
        "三、问题与不足",
        "四、经验提炼",
        "五、下一步计划",
    ],
    "自定义": [
        "标题",
        "一、背景",
        "二、主体内容（按需细化）",
        "三、结语/行动项",
    ],
}


def outline_to_markdown(lines: List[str]) -> str:
    md = []
    for s in lines:
        s = s.strip()
        if not s:
            continue
        md.append(f"## {s}")
    return "\n".join(md).strip() + "\n"


def split_outline_markdown(md: str) -> List[str]:
    lines = []
    for line in md.splitlines():
        line = line.strip()
        if line.startswith("#"):
            title = line.lstrip("#").strip()
            if title:
                lines.append(title)
    if not lines:
        lines = [x.strip() for x in md.splitlines() if x.strip()]
    return lines


def rule_based_outline(title: str, genre: str, instructions: str) -> str:
    base = DEFAULT_OUTLINE_TEMPLATES.get(genre, DEFAULT_OUTLINE_TEMPLATES["自定义"])
    lines = base.copy()
    if lines and lines[0] in ("题目", "标题"):
        lines[0] = f"{lines[0]}：{title}" if title else lines[0]
    else:
        if title:
            lines.insert(0, f"标题：{title}")

    kv = parse_kv_instructions(instructions)
    if genre == "论文":
        structure = (kv.get("结构", "") or "").upper()
        if "IMRAD" in structure:
            lines = [
                f"题目：{title}" if title else "题目",
                "摘要",
                "关键词",
                "1 引言（Introduction）",
                "2 方法（Methods）",
                "3 结果（Results）",
                "4 讨论（Discussion）",
                "5 结论（Conclusion）",
                "参考文献（占位符）",
            ]
    return outline_to_markdown(lines)


def rule_based_draft(title: str, genre: str, outline_md: str, instructions: str) -> str:
    sections = split_outline_markdown(outline_md)
    kv = parse_kv_instructions(instructions)
    lang = kv.get("语言", "中文")
    word_count = kv.get("字数", "")
    style = kv.get("风格", "清晰、结构化")

    intro = f"# {title or '未命名文稿'}\n\n"
    intro += f"> 体裁：{genre}｜语言：{lang}｜目标字数：{word_count or '未指定'}｜风格：{style}\n\n"
    if instructions.strip():
        intro += f"**写作要求/指令：** {instructions.strip()}\n\n"

    body = []
    for sec in sections:
        body.append(f"## {sec}\n")
        if genre == "论文":
            if "摘要" in sec:
                body.append("（150~300字概括研究背景、方法、主要发现与结论。）\n")
            elif "关键词" in sec:
                body.append("关键词：___；___；___；___\n")
            elif "引言" in sec:
                body.append("（交代研究背景、问题、意义与研究目标，指出研究空白。）\n")
            elif "方法" in sec:
                body.append("（描述研究设计、对象/材料、变量、步骤、数据处理方法。）\n")
            elif "结果" in sec:
                body.append("（按研究问题呈现结果，可用小节：3.1、3.2…）\n")
            elif "讨论" in sec:
                body.append("（解释结果、与已有研究对比、指出局限与启示。）\n")
            elif "结论" in sec:
                body.append("（总结主要结论、贡献、应用价值与未来工作。）\n")
            elif "参考文献" in sec:
                body.append("（参考文献占位符：\n- [1] 作者. 题目. 期刊, 年, 卷(期): 页码.\n- [2] ...）\n")
            else:
                body.append("（按该小节主题展开论述，建议每节2~4段。）\n")
        else:
            body.append("（围绕该小节主题写1~3段，尽量给出事实、证据与可执行建议。）\n")
        body.append("\n")

    return intro + "".join(body)


class WriterEngine:
    def __init__(self, llm: Optional[OpenAICompatClient], logger):
        self.llm = llm
        self.logger = logger

    def gen_outline(self, title: str, genre: str, instructions: str) -> str:
        if self.llm and self.llm.is_ready():
            sys_prompt = (
                "你是一个专业写作助手。用户会给出题目、体裁与要求。"
                "任务：生成可编辑的中文Markdown大纲。"
                "要求：层级清晰，章节标题具体，不要写正文，不要写多余解释。"
            )
            user_prompt = (
                f"题目：{title}\n"
                f"体裁：{genre}\n"
                f"写作要求：{instructions}\n\n"
                "输出Markdown大纲（使用 #/##/### 标题），适合直接扩写为完整文稿。"
            )
            try:
                out = self.llm.chat([
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt},
                ])
                return clamp_text(out, 12000)
            except Exception as e:
                self.logger(f"[{now_str()}] LLM 调用失败，回退离线模板：{e}")
                return rule_based_outline(title, genre, instructions)
        return rule_based_outline(title, genre, instructions)

    def write_full(self, title: str, genre: str, outline_md: str, instructions: str) -> str:
        if self.llm and self.llm.is_ready():
            sys_prompt = (
                "你是一个专业写作助手。你将根据用户的大纲生成完整文稿。"
                "要求：结构与大纲一致，语言为中文（除非指令指定），表达严谨、连贯。"
                "如需引用，用[1][2]做占位符，不要编造真实DOI或作者。"
            )
            user_prompt = (
                f"题目：{title}\n"
                f"体裁：{genre}\n"
                f"写作要求：{instructions}\n\n"
                f"大纲（Markdown）：\n{outline_md}\n\n"
                "生成完整文稿（Markdown），保持标题层级与大纲一致。"
            )
            try:
                out = self.llm.chat([
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt},
                ])
                return clamp_text(out, 45000)
            except Exception as e:
                self.logger(f"[{now_str()}] LLM 调用失败，回退离线模板：{e}")
                return rule_based_draft(title, genre, outline_md, instructions)
        return rule_based_draft(title, genre, outline_md, instructions)


# ---------------------------
# Exporters
# ---------------------------
def export_markdown(path: str, text: str):
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def export_docx(path: str, markdown_text: str):
    doc = Document()
    for line in markdown_text.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith(">"):
            doc.add_paragraph(line.lstrip(">").strip(), style="Intense Quote")
        else:
            doc.add_paragraph(line)
    doc.save(path)


# ---------------------------
# UI Components
# ---------------------------
class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, settings=None):
        super().__init__(parent)
        self.setWindowTitle("设置 - DeepSeek/OpenAI 兼容接口（可选）")
        self.setModal(True)
        self.resize(640, 300)

        self.api_base = QtWidgets.QLineEdit()
        self.api_key = QtWidgets.QLineEdit()
        self.api_key.setEchoMode(QtWidgets.QLineEdit.Password)
        self.model = QtWidgets.QLineEdit()
        self.temperature = QtWidgets.QDoubleSpinBox()
        self.temperature.setRange(0.0, 2.0)
        self.temperature.setSingleStep(0.1)
        self.max_tokens = QtWidgets.QSpinBox()
        self.max_tokens.setRange(128, 8000)

        form = QtWidgets.QFormLayout()
        form.addRow("API Base（DeepSeek: https://api.deepseek.com）", self.api_base)
        form.addRow("API Key", self.api_key)
        form.addRow("Model（DeepSeek: deepseek-chat / deepseek-reasoner）", self.model)
        form.addRow("Temperature", self.temperature)
        form.addRow("Max tokens", self.max_tokens)

        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addLayout(form)

        note = QtWidgets.QLabel(
            "说明：不填写也能使用（将采用离线模板生成）。\n"
            "DeepSeek API 与 OpenAI 兼容：base_url 可用 https://api.deepseek.com（或 /v1）。"
        )
        note.setStyleSheet("color:#666;")
        layout.addWidget(note)
        layout.addWidget(btns)

        if settings:
            self.api_base.setText(settings.get("api_base", DEFAULT_API_BASE))
            self.api_key.setText(settings.get("api_key", ""))
            self.model.setText(settings.get("model", DEFAULT_MODEL))
            self.temperature.setValue(float(settings.get("temperature", 0.6)))
            self.max_tokens.setValue(int(settings.get("max_tokens", 1800)))
        else:
            self.api_base.setText(DEFAULT_API_BASE)
            self.model.setText(DEFAULT_MODEL)
            self.temperature.setValue(0.6)
            self.max_tokens.setValue(1800)

    def get_data(self) -> dict:
        return {
            "api_base": self.api_base.text().strip() or DEFAULT_API_BASE,
            "api_key": self.api_key.text().strip(),
            "model": self.model.text().strip() or DEFAULT_MODEL,
            "temperature": float(self.temperature.value()),
            "max_tokens": int(self.max_tokens.value()),
        }


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} - 期刊论文/通用文稿写作")
        self.resize(1200, 760)

        self.settings = load_settings()
        self.llm_cfg = LLMConfig(
            api_base=self.settings.get("api_base", DEFAULT_API_BASE),
            api_key=self.settings.get("api_key", ""),
            model=self.settings.get("model", DEFAULT_MODEL),
            temperature=float(self.settings.get("temperature", 0.6)),
            max_tokens=int(self.settings.get("max_tokens", 1800)),
        )

        self.log_box = QtWidgets.QPlainTextEdit()
        self.log_box.setReadOnly(True)

        self.title_edit = QtWidgets.QLineEdit()
        self.title_edit.setPlaceholderText("输入题目（例如：基于探究式学习的高中化学实验教学效果研究）")

        self.genre_combo = QtWidgets.QComboBox()
        self.genre_combo.addItems(GENRES)

        self.instruction_edit = QtWidgets.QPlainTextEdit()
        self.instruction_edit.setPlaceholderText(
            "输入明确指令（可选）。例如：结构=IMRaD; 语言=中文; 字数=2000; 风格=严谨\n"
            "也可以写自然语言要求。"
        )

        self.outline_edit = QtWidgets.QPlainTextEdit()
        self.outline_edit.setPlaceholderText("这里会生成 Markdown 大纲，你可以随意修改。")

        self.draft_edit = QtWidgets.QPlainTextEdit()
        self.draft_edit.setPlaceholderText("点击“撰写全文”后，这里显示完整文稿（Markdown）。")

        self.btn_outline = QtWidgets.QPushButton("① 生成大纲")
        self.btn_write = QtWidgets.QPushButton("② 撰写全文")
        self.btn_export_md = QtWidgets.QPushButton("导出 .md")
        self.btn_export_docx = QtWidgets.QPushButton("导出 .docx")
        self.btn_settings = QtWidgets.QPushButton("设置（API 可选）")
        self.btn_clear = QtWidgets.QPushButton("清空")

        self.btn_outline.clicked.connect(self.on_generate_outline)
        self.btn_write.clicked.connect(self.on_write_full)
        self.btn_export_md.clicked.connect(self.on_export_md)
        self.btn_export_docx.clicked.connect(self.on_export_docx)
        self.btn_settings.clicked.connect(self.on_settings)
        self.btn_clear.clicked.connect(self.on_clear)

        top_bar = QtWidgets.QHBoxLayout()
        top_bar.addWidget(QtWidgets.QLabel("题目："))
        top_bar.addWidget(self.title_edit, 1)
        top_bar.addWidget(QtWidgets.QLabel("体裁："))
        top_bar.addWidget(self.genre_combo)
        top_bar.addWidget(self.btn_outline)
        top_bar.addWidget(self.btn_write)
        top_bar.addWidget(self.btn_export_md)
        top_bar.addWidget(self.btn_export_docx)
        top_bar.addWidget(self.btn_settings)
        top_bar.addWidget(self.btn_clear)

        left = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.addWidget(QtWidgets.QLabel("写作指令/要求（可选）"))
        left_layout.addWidget(self.instruction_edit, 1)
        left_layout.addWidget(QtWidgets.QLabel("大纲（Markdown，可编辑）"))
        left_layout.addWidget(self.outline_edit, 2)

        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.addWidget(QtWidgets.QLabel("全文（Markdown）"))
        right_layout.addWidget(self.draft_edit, 4)
        right_layout.addWidget(QtWidgets.QLabel("日志"))
        right_layout.addWidget(self.log_box, 1)

        splitter = QtWidgets.QSplitter()
        splitter.setOrientation(QtCore.Qt.Horizontal)
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        central = QtWidgets.QWidget()
        central_layout = QtWidgets.QVBoxLayout(central)
        central_layout.addLayout(top_bar)
        central_layout.addWidget(splitter, 1)
        self.setCentralWidget(central)

        self.apply_style()
        self.logger(f"[{now_str()}] 启动完成。未配置 API 也可使用离线模板。")
        if sys.platform.startswith("linux"):
            self.logger(f"[{now_str()}] QT_IM_MODULE={os.environ.get('QT_IM_MODULE','(unset)')}")

    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow { background: #f6f7fb; }
            QLabel { color: #222; font-weight: 600; }
            QLineEdit, QPlainTextEdit {
                background: #ffffff;
                border: 1px solid #d7dbe7;
                border-radius: 8px;
                padding: 8px;
                font-family: "Segoe UI", "Microsoft YaHei", "PingFang SC";
                font-size: 13px;
            }
            QPushButton {
                background: #2b6ff3;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 8px 12px;
                font-weight: 700;
            }
            QPushButton:hover { background: #215fda; }
            QPushButton:pressed { background: #184db5; }
            QComboBox {
                background: #ffffff;
                border: 1px solid #d7dbe7;
                border-radius: 10px;
                padding: 6px 10px;
            }
        """)

    def logger(self, msg: str):
        self.log_box.appendPlainText(msg)

    def build_engine(self) -> WriterEngine:
        client = OpenAICompatClient(self.llm_cfg, self.logger)
        return WriterEngine(client, self.logger)

    def collect_inputs(self):
        title = self.title_edit.text().strip()
        genre = self.genre_combo.currentText().strip()
        instructions = self.instruction_edit.toPlainText().strip()
        return title, genre, instructions

    def on_generate_outline(self):
        title, genre, instructions = self.collect_inputs()
        if not title:
            QtWidgets.QMessageBox.warning(self, "提示", "请先输入题目。")
            return
        self.logger(f"[{now_str()}] 开始生成大纲：{genre} | {title}")
        engine = self.build_engine()

        self.btn_outline.setEnabled(False)
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            outline = engine.gen_outline(title, genre, instructions)
            self.outline_edit.setPlainText(outline)
            self.logger(f"[{now_str()}] 大纲生成完成。")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()
            self.btn_outline.setEnabled(True)

    def on_write_full(self):
        title, genre, instructions = self.collect_inputs()
        outline_md = self.outline_edit.toPlainText().strip()

        if not title:
            QtWidgets.QMessageBox.warning(self, "提示", "请先输入题目。")
            return
        if not outline_md:
            QtWidgets.QMessageBox.warning(self, "提示", "请先生成或填写大纲。")
            return

        self.logger(f"[{now_str()}] 开始撰写全文：{genre} | {title}")
        engine = self.build_engine()

        self.btn_write.setEnabled(False)
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            draft = engine.write_full(title, genre, outline_md, instructions)
            self.draft_edit.setPlainText(draft)
            self.logger(f"[{now_str()}] 全文生成完成。")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()
            self.btn_write.setEnabled(True)

    def on_export_md(self):
        text = self.draft_edit.toPlainText().strip() or self.outline_edit.toPlainText().strip()
        if not text:
            QtWidgets.QMessageBox.information(self, "提示", "没有可导出的内容（请先生成大纲或全文）。")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出 Markdown", f"{APP_NAME}.md", "Markdown (*.md)")
        if not path:
            return
        export_markdown(path, text)
        self.logger(f"[{now_str()}] 已导出 Markdown：{path}")

    def on_export_docx(self):
        text = self.draft_edit.toPlainText().strip()
        if not text:
            QtWidgets.QMessageBox.information(self, "提示", "请先生成全文，再导出 Word。")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出 Word", f"{APP_NAME}.docx", "Word (*.docx)")
        if not path:
            return
        export_docx(path, text)
        self.logger(f"[{now_str()}] 已导出 Word：{path}")

    def on_settings(self):
        dlg = SettingsDialog(self, settings={
            "api_base": self.llm_cfg.api_base,
            "api_key": self.llm_cfg.api_key,
            "model": self.llm_cfg.model,
            "temperature": self.llm_cfg.temperature,
            "max_tokens": self.llm_cfg.max_tokens,
        })
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            data = dlg.get_data()
            self.llm_cfg = LLMConfig(**data)
            save_settings(data)
            ready = OpenAICompatClient(self.llm_cfg, self.logger).is_ready()
            self.logger(f"[{now_str()}] 设置已保存。API {'已启用' if ready else '未启用（将使用离线模板）'}。")

    def on_clear(self):
        self.title_edit.clear()
        self.instruction_edit.clear()
        self.outline_edit.clear()
        self.draft_edit.clear()
        self.logger(f"[{now_str()}] 已清空。")


# ---------------------------
# CLI + self-test (for CI)
# ---------------------------
def run_cli_if_requested():
    import argparse
    parser = argparse.ArgumentParser(description="PaperWriter CLI")
    parser.add_argument("--title", type=str, default="")
    parser.add_argument("--genre", type=str, default="论文", choices=GENRES)
    parser.add_argument("--instructions", type=str, default="")
    parser.add_argument("--outline-only", action="store_true")
    parser.add_argument("--selftest", action="store_true", help="Run a minimal startup test and exit.")
    args, _ = parser.parse_known_args()

    if args.selftest:
        print("SELFTEST_OK")
        sys.exit(0)

    if args.title:
        settings = load_settings()
        cfg = LLMConfig(
            api_base=settings.get("api_base", DEFAULT_API_BASE),
            api_key=settings.get("api_key", ""),
            model=settings.get("model", DEFAULT_MODEL),
            temperature=float(settings.get("temperature", 0.6)),
            max_tokens=int(settings.get("max_tokens", 1800)),
        )

        def _log(msg):
            print(msg, file=sys.stderr)

        client = OpenAICompatClient(cfg, _log)
        engine = WriterEngine(client, _log)
        outline = engine.gen_outline(args.title, args.genre, args.instructions)
        if args.outline_only:
            print(outline)
        else:
            draft = engine.write_full(args.title, args.genre, outline, args.instructions)
            print(draft)
        sys.exit(0)


def main():
    run_cli_if_requested()

    linux_preflight_check()
    qt_env_fixups()

    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
