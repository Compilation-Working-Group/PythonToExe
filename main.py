import sys
import json
import os
import re
import requests
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox,
    QFileDialog, QDialog, QFormLayout, QMessageBox, QMenuBar, QAction,
    QTabWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

CONFIG_FILE = "config.json"
DRAFT_FILE = "draft.json"


# ===================== 工具函数：读写 JSON =====================

def load_json(path, default=None):
    if default is None:
        default = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return default
    return default


def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass


def load_config():
    return load_json(CONFIG_FILE, {})


def save_config(cfg):
    save_json(CONFIG_FILE, cfg)


def load_draft():
    return load_json(DRAFT_FILE, {})


def save_draft(title, doc_type, outline, fulltext):
    data = {
        "title": title,
        "doc_type": doc_type,
        "outline": outline,
        "fulltext": fulltext,
    }
    save_json(DRAFT_FILE, data)


# ===================== 文稿模板管理 =====================

class TemplateManager:
    def __init__(self):
        self.templates = {
            "期刊论文": {
                "outline": (
                    "请为以下期刊论文生成详细大纲，包含：\n"
                    "一、摘要\n"
                    "二、引言\n"
                    "三、研究方法\n"
                    "四、研究结果\n"
                    "五、讨论\n"
                    "六、结论与展望\n"
                    "七、参考文献（仅列出结构）。\n"
                    "使用如下分级格式：一、（一）1.（1）。"
                ),
                "full": (
                    "请根据给定大纲撰写一篇完整的中文期刊论文，包含：摘要、引言、方法、结果、讨论、结论。\n"
                    "要求：\n"
                    "1. 文风学术、规范，逻辑清晰。\n"
                    "2. 各部分内容充实，有论证、有分析。\n"
                    "3. 标题层级与大纲保持一致。\n"
                    "4. 参考文献部分可给出示例格式。"
                ),
            },
            "工作计划": {
                "outline": (
                    "请为以下工作计划生成大纲，建议结构：\n"
                    "一、指导思想\n"
                    "二、工作目标\n"
                    "三、主要措施\n"
                    "四、时间安排\n"
                    "五、保障措施。\n"
                    "使用公文式分级标题。"
                ),
                "full": (
                    "请根据大纲撰写一篇完整的工作计划，要求：\n"
                    "1. 文风正式、务实，条理清晰。\n"
                    "2. 目标明确，措施具体可操作。\n"
                    "3. 适合用于学校/单位正式文件。"
                ),
            },
            "教学反思": {
                "outline": (
                    "请为以下教学反思生成大纲，建议结构：\n"
                    "一、教学背景\n"
                    "二、教学过程回顾\n"
                    "三、存在问题\n"
                    "四、改进措施\n"
                    "五、反思与提升。\n"
                    "使用公文式分级标题。"
                ),
                "full": (
                    "请根据大纲撰写一篇完整的教学反思，要求：\n"
                    "1. 结合具体教学情境，有细节、有案例。\n"
                    "2. 反思要真诚、深入，避免空泛。\n"
                    "3. 改进措施要具体可行。"
                ),
            },
            "案例分析": {
                "outline": (
                    "请为以下案例分析生成大纲，建议结构：\n"
                    "一、案例背景\n"
                    "二、案例经过\n"
                    "三、问题分析\n"
                    "四、对策与建议\n"
                    "五、启示与总结。\n"
                    "使用公文式分级标题。"
                ),
                "full": (
                    "请根据大纲撰写一篇完整的案例分析，要求：\n"
                    "1. 案例描述清晰具体。\n"
                    "2. 分析有理论支撑或实践依据。\n"
                    "3. 对策建议具有可操作性。"
                ),
            },
            "工作总结": {
                "outline": (
                    "请为以下工作总结生成大纲，建议结构：\n"
                    "一、基本情况\n"
                    "二、主要工作与成效\n"
                    "三、存在问题\n"
                    "四、改进方向\n"
                    "五、下一步工作思路。\n"
                    "使用公文式分级标题。"
                ),
                "full": (
                    "请根据大纲撰写一篇完整的工作总结，要求：\n"
                    "1. 实事求是，有数据或事例支撑。\n"
                    "2. 结构清晰，重点突出。\n"
                    "3. 既总结成绩，也分析问题与改进方向。"
                ),
            },
            "自定义文稿": {
                "outline": (
                    "请根据标题和文稿类型，自主设计一个结构合理、层级清晰的大纲，"
                    "使用公文式分级标题。"
                ),
                "full": (
                    "请根据大纲撰写一篇完整的正式中文文稿，文风可根据类型自适应，"
                    "但要求结构清晰、内容充实、逻辑严谨。"
                ),
            },
        }

    def get_outline_prompt(self, doc_type):
        return self.templates.get(doc_type, self.templates["自定义文稿"])["outline"]

    def get_full_prompt(self, doc_type):
        return self.templates.get(doc_type, self.templates["自定义文稿"])["full"]


# ===================== DeepSeek 流式写作线程 =====================

class WritingThread(QThread):
    text_chunk = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, api_key, base_url, model, system_prompt, user_prompt):
        super().__init__()
        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.system_prompt = system_prompt
        self.user_prompt = user_prompt
        self._running = True

    def stop(self):
        self._running = False

    def run(self):
        try:
            url = f"{self.base_url}/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": self.user_prompt}
                ],
                "stream": True
            }

            with requests.post(url, headers=headers, json=payload, stream=True) as resp:
                resp.raise_for_status()
                for line in resp.iter_lines():
                    if not self._running:
                        break
                    if line:
                        try:
                            data = json.loads(line.decode("utf-8").replace("data: ", ""))
                            delta = data["choices"][0]["delta"].get("content", "")
                            if delta:
                                self.text_chunk.emit(delta)
                        except:
                            continue

            self.finished.emit()

        except Exception as e:
            self.error.emit(str(e))


# ===================== 设置窗口 =====================

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("DeepSeek API 设置")
        self.resize(400, 220)

        self.config = load_config()

        layout = QFormLayout()

        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.api_key_edit.setText(self.config.get("api_key", ""))

        self.base_url_edit = QLineEdit()
        self.base_url_edit.setText(self.config.get("base_url", "https://api.deepseek.com"))

        self.model_edit = QLineEdit()
        self.model_edit.setText(self.config.get("model", "deepseek-chat"))

        layout.addRow("API Key：", self.api_key_edit)
        layout.addRow("Base URL：", self.base_url_edit)
        layout.addRow("Model 名称：", self.model_edit)

        btns = QHBoxLayout()
        btn_ok = QPushButton("保存")
        btn_cancel = QPushButton("取消")
        btn_ok.clicked.connect(self.save_and_close)
        btn_cancel.clicked.connect(self.reject)
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)

        layout.addRow(btns)
        self.setLayout(layout)

    def save_and_close(self):
        self.config["api_key"] = self.api_key_edit.text().strip()
        self.config["base_url"] = self.base_url_edit.text().strip()
        self.config["model"] = self.model_edit.text().strip()
        save_config(self.config)
        self.accept()


# ===================== 主界面 =====================

class WriterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("智能文稿撰写助手（论文 / 公文 / 总结等）")
        self.resize(1000, 720)

        self.config = load_config()
        self.templates = TemplateManager()

        self.init_ui()
        self.load_draft_if_any()

    # ---------- UI ----------

    def init_ui(self):
        main_layout = QVBoxLayout()

        # 菜单栏
        menubar = QMenuBar(self)
        menu_settings = menubar.addMenu("设置")
        act_settings = QAction("DeepSeek API 设置", self)
        act_settings.triggered.connect(self.open_settings)
        menu_settings.addAction(act_settings)

        menu_file = menubar.addMenu("文件")
        act_save_draft = QAction("保存草稿", self)
        act_save_draft.triggered.connect(self.manual_save_draft)
        menu_file.addAction(act_save_draft)

        main_layout.setMenuBar(menubar)

        # 顶部：类型 + 标题
        top = QHBoxLayout()
        top.addWidget(QLabel("文稿类型："))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["期刊论文", "工作计划", "教学反思", "案例分析", "工作总结", "自定义文稿"])
        top.addWidget(self.type_combo)

        top.addWidget(QLabel("标题："))
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("请输入标题")
        top.addWidget(self.title_edit, 1)

        main_layout.addLayout(top)

        # 按钮区
        btn_row1 = QHBoxLayout()
        self.btn_outline = QPushButton("生成大纲")
        self.btn_full = QPushButton("撰写全文（流式）")
        self.btn_stop = QPushButton("终止撰写")
        self.btn_clear = QPushButton("清空内容")

        self.btn_outline.clicked.connect(self.generate_outline)
        self.btn_full.clicked.connect(self.start_writing)
        self.btn_stop.clicked.connect(self.stop_writing)
        self.btn_clear.clicked.connect(self.clear_text)

        btn_row1.addWidget(self.btn_outline)
        btn_row1.addWidget(self.btn_full)
        btn_row1.addWidget(self.btn_stop)
        btn_row1.addWidget(self.btn_clear)

        main_layout.addLayout(btn_row1)

        btn_row2 = QHBoxLayout()
        self.btn_abstract = QPushButton("生成摘要")
        self.btn_refs = QPushButton("生成参考文献")
        self.btn_export_docx = QPushButton("导出 Word")
        self.btn_export_md = QPushButton("导出 Markdown")
        self.btn_export_txt = QPushButton("导出 TXT")

        self.btn_abstract.clicked.connect(self.generate_abstract)
        self.btn_refs.clicked.connect(self.generate_references)
        self.btn_export_docx.clicked.connect(self.export_word)
        self.btn_export_md.clicked.connect(self.export_markdown)
        self.btn_export_txt.clicked.connect(self.export_txt)

        btn_row2.addWidget(self.btn_abstract)
        btn_row2.addWidget(self.btn_refs)
        btn_row2.addWidget(self.btn_export_docx)
        btn_row2.addWidget(self.btn_export_md)
        btn_row2.addWidget(self.btn_export_txt)

        main_layout.addLayout(btn_row2)

        # 中部：Tab（大纲 / 正文）
        self.tabs = QTabWidget()
        self.outline_edit = QTextEdit()
        self.fulltext_edit = QTextEdit()

        self.outline_edit.setPlaceholderText("这里是大纲，可手动修改……")
        self.fulltext_edit.setPlaceholderText("这里是正文，可手动修改……")

        self.tabs.addTab(self.outline_edit, "大纲")
        self.tabs.addTab(self.fulltext_edit, "正文")

        main_layout.addWidget(self.tabs, 1)

        self.setLayout(main_layout)

        # 应用美化样式
        self.apply_style()

    # ---------- 美化界面 ----------

    def apply_style(self):
        self.setStyleSheet("""
            QWidget {
                font-family: 'Microsoft YaHei';
                font-size: 14px;
            }
            QPushButton {
                background-color: #4A90E2;
                color: white;
                border-radius: 6px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background-color: #357ABD;
            }
            QPushButton:pressed {
                background-color: #2C5A93;
            }
            QLineEdit, QTextEdit, QComboBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                padding: 4px;
            }
            QTabWidget::pane {
                border: 1px solid #CCCCCC;
            }
            QTabBar::tab {
                padding: 8px 16px;
            }
            QTabBar::tab:selected {
                background-color: #E6F0FA;
            }
        """)

    # ---------- DeepSeek 调用 ----------

    def call_deepseek(self, system_prompt, user_prompt, temperature=0.7):
        cfg = load_config()
        api_key = cfg.get("api_key", "")
        base_url = cfg.get("base_url", "").rstrip("/")
        model = cfg.get("model", "")

        if not api_key or not base_url or not model:
            QMessageBox.warning(self, "缺少配置", "请先在“设置”中填写 API Key、Base URL 和 Model。")
            return None

        url = f"{base_url}/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": temperature
        }

        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=120)
            resp.raise_for_status()
            return resp.json()["choices"][0]["message"]["content"]
        except Exception as e:
            QMessageBox.critical(self, "调用失败", str(e))
            return None

    # ---------- 生成大纲 ----------

    def generate_outline(self):
        title = self.title_edit.text().strip()
        doc_type = self.type_combo.currentText()
        if not title:
            QMessageBox.warning(self, "缺少标题", "请先输入标题")
            return

        system_prompt = (
            "你是一名专业中文写作助手，请根据标题和文稿类型生成结构严谨、层级清晰的大纲。\n"
            "使用如下分级格式：\n"
            "一、……\n"
            "（一）……\n"
            "1. ……\n"
            "（1）……\n"
        )
        type_prompt = self.templates.get_outline_prompt(doc_type)
        user_prompt = f"文稿类型：{doc_type}\n标题：{title}\n{type_prompt}"

        content = self.call_deepseek(system_prompt, user_prompt, temperature=0.5)
        if content:
            self.outline_edit.setPlainText(content)
            self.tabs.setCurrentWidget(self.outline_edit)
            self.auto_save()

    # ---------- 流式撰写正文 ----------

    def start_writing(self):
        title = self.title_edit.text().strip()
        outline = self.outline_edit.toPlainText().strip()
        doc_type = self.type_combo.currentText()

        if not title or not outline:
            QMessageBox.warning(self, "缺少内容", "请先输入标题并生成大纲")
            return

        system_prompt = (
            "你是一名专业中文写作者，请根据给定大纲流式撰写完整文稿。\n"
            "要求：\n"
            "1. 文风正式、规范，适合正式发表或存档。\n"
            "2. 严格按照大纲结构展开，标题层级保持一致。\n"
            "3. 内容充实，有分析、有论证，避免空话套话。"
        )
        type_prompt = self.templates.get_full_prompt(doc_type)
        user_prompt = (
            f"文稿类型：{doc_type}\n"
            f"标题：{title}\n"
            f"{type_prompt}\n"
            "以下是大纲：\n"
            f"{outline}\n"
            "请开始撰写正文。"
        )

