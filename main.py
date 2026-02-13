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
from PyQt5.QtCore import Qt
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

CONFIG_FILE = "config.json"
DRAFT_FILE = "draft.json"


# ===================== 配置读写 =====================

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


# ===================== 文稿模板 =====================

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
                    "七、参考文献（仅列出结构，不必具体文献）。\n"
                    "使用如下分级格式：一、（一）1.（1）。"
                ),
                "full": (
                    "请根据给定大纲撰写一篇完整的中文期刊论文，包含：摘要、引言、方法、结果、讨论、结论。"
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


# ===================== 主程序 =====================

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
        self.btn_full = QPushButton("撰写全文")
        self.btn_abstract = QPushButton("生成摘要")
        self.btn_refs = QPushButton("生成参考文献示例")

        self.btn_outline.clicked.connect(self.generate_outline)
        self.btn_full.clicked.connect(self.generate_fulltext)
        self.btn_abstract.clicked.connect(self.generate_abstract)
        self.btn_refs.clicked.connect(self.generate_references)

        btn_row1.addWidget(self.btn_outline)
        btn_row1.addWidget(self.btn_full)
        btn_row1.addWidget(self.btn_abstract)
        btn_row1.addWidget(self.btn_refs)

        main_layout.addLayout(btn_row1)

        btn_row2 = QHBoxLayout()
        self.btn_export_docx = QPushButton("导出 Word（公文格式）")
        self.btn_export_md = QPushButton("导出 Markdown")
        self.btn_export_txt = QPushButton("导出 TXT")

        self.btn_export_docx.clicked.connect(self.export_word)
        self.btn_export_md.clicked.connect(self.export_markdown)
        self.btn_export_txt.clicked.connect(self.export_txt)

        btn_row2.addWidget(self.btn_export_docx)
        btn_row2.addWidget(self.btn_export_md)
        btn_row2.addWidget(self.btn_export_txt)

        main_layout.addLayout(btn_row2)

        # 中部：Tab（大纲 / 正文）
        self.tabs = QTabWidget()
        self.outline_edit = QTextEdit()
        self.fulltext_edit = QTextEdit()

        self.outline_edit.setPlaceholderText("这里是大纲，可手动修改……")
        self.fulltext_edit.setPlaceholderText("这里是全文，可手动修改……")

        self.tabs.addTab(self.outline_edit, "大纲")
        self.tabs.addTab(self.fulltext_edit, "正文")

        main_layout.addWidget(self.tabs, 1)

        self.setLayout(main_layout)

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
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
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
            data = resp.json()
            return data["choices"][0]["message"]["content"]
        except Exception as e:
            QMessageBox.critical(self, "调用失败", f"调用 DeepSeek 失败：\n{e}")
            return None

    # ---------- 业务：大纲 / 正文 / 摘要 / 参考文献 ----------

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

    def generate_fulltext(self):
        title = self.title_edit.text().strip()
        doc_type = self.type_combo.currentText()
        outline = self.outline_edit.toPlainText().strip()

        if not title:
            QMessageBox.warning(self, "缺少标题", "请先输入标题")
            return
        if not outline:
            QMessageBox.warning(self, "缺少大纲", "请先生成或编写大纲")
            return

        system_prompt = (
            "你是一名专业中文写作者，请根据给定大纲撰写完整文稿。\n"
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
            "------------------\n"
            f"{outline}\n"
            "------------------\n"
            "请据此撰写完整文稿。"
        )

        content = self.call_deepseek(system_prompt, user_prompt, temperature=0.7)
        if content:
            self.fulltext_edit.setPlainText(content)
            self.tabs.setCurrentWidget(self.fulltext_edit)
            self.auto_save()

    def generate_abstract(self):
        title = self.title_edit.text().strip()
        fulltext = self.fulltext_edit.toPlainText().strip()
        if not title or not fulltext:
            QMessageBox.warning(self, "缺少内容", "请先生成或撰写正文，再生成摘要。")
            return

        system_prompt = (
            "你是一名学术写作助手，请根据给定正文生成一个中文摘要。\n"
            "要求：\n"
            "1. 200～300 字左右。\n"
            "2. 概括研究/内容的目的、方法、结果、结论（若适用）。\n"
            "3. 文风简洁、准确。"
        )
        user_prompt = f"标题：{title}\n正文如下：\n{fulltext}"

        content = self.call_deepseek(system_prompt, user_prompt, temperature=0.4)
        if content:
            # 将摘要插入正文最前面
            new_text = f"【摘要】\n{content.strip()}\n\n{fulltext}"
            self.fulltext_edit.setPlainText(new_text)
            self.tabs.setCurrentWidget(self.fulltext_edit)
            self.auto_save()

    def generate_references(self):
        title = self.title_edit.text().strip()
        doc_type = self.type_combo.currentText()
        fulltext = self.fulltext_edit.toPlainText().strip()
        if not title or not fulltext:
            QMessageBox.warning(self, "缺少内容", "请先生成或撰写正文，再生成参考文献示例。")
            return

        system_prompt = (
            "你是一名学术写作助手，请根据标题和正文内容，生成若干条示例参考文献，"
            "使用中文常见期刊/图书/网络文献格式，注意：\n"
            "1. 可以是虚构但要格式规范。\n"
            "2. 不需要太多，一般 5～10 条即可。\n"
            "3. 每条独立成行。"
        )
        user_prompt = f"文稿类型：{doc_type}\n标题：{title}\n正文如下：\n{fulltext}"

        content = self.call_deepseek(system_prompt, user_prompt, temperature=0.6)
        if content:
            new_text = self.fulltext_edit.toPlainText().rstrip()
            new_text += "\n\n【参考文献】\n" + content.strip()
            self.fulltext_edit.setPlainText(new_text)
            self.tabs.setCurrentWidget(self.fulltext_edit)
            self.auto_save()

    # ---------- 导出：Word / Markdown / TXT ----------

    def detect_heading_level(self, line: str) -> int:
        line = line.lstrip()
        if re.match(r"^[一二三四五六七八九十]+、", line):
            return 1
        if re.match(r"^（[一二三四五六七八九十]+）", line):
            return 2
        if re.match(r"^\d+\.", line):
            return 3
        if re.match(r"^（\d+）", line):
            return 4
        return 0

    def export_word(self):
        text = self.fulltext_edit.toPlainText().strip()
        title = self.title_edit.text().strip()
        if not text:
            QMessageBox.warning(self, "无内容", "正文为空，无法导出。")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出 Word（公文格式）", f"{title or '文稿'}.docx", "Word 文档 (*.docx)"
        )
        if not file_path:
            return

        try:
            doc = Document()

            # 正文默认样式
            style = doc.styles["Normal"]
            style.font.name = "宋体"
            style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            style.font.size = Pt(12)

            # 标题（小标宋，居中）
            if title:
                p = doc.add_paragraph()
                p.alignment = 1
                run = p.add_run(title)
                run.bold = True
                run.font.size = Pt(16)
                run.font.name = "方正小标宋简体"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "方正小标宋简体")

            for line in text.splitlines():
                line = line.rstrip()
                if not line:
                    doc.add_paragraph("")
                    continue

                level = self.detect_heading_level(line)

                if level == 1:
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    run.bold = True
                    run.font.name = "黑体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
                    run.font.size = Pt(15)
                elif level == 2:
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    run.bold = True
                    run.font.name = "黑体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
                    run.font.size = Pt(14)
                elif level == 3:
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    run.bold = True
                    run.font.name = "楷体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "楷体")
                    run.font.size = Pt(13)
                elif level == 4:
                    p = doc.add_paragraph()
                    run = p.add_run(line)
                    run.bold = True
                    run.font.name = "楷体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "楷体")
                    run.font.size = Pt(12)
                else:
                    p = doc.add_paragraph()
                    p.paragraph_format.first_line_indent = Pt(24)
                    p.paragraph_format.line_spacing = 1.5
                    run = p.add_run(line)
                    run.font.name = "宋体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                    run.font.size = Pt(12)

            doc.save(file_path)
            QMessageBox.information(self, "导出成功", f"已导出：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出 Word 失败：\n{e}")

    def export_markdown(self):
        text = self.fulltext_edit.toPlainText().strip()
        title = self.title_edit.text().strip()
        if not text:
            QMessageBox.warning(self, "无内容", "正文为空，无法导出。")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出 Markdown", f"{title or '文稿'}.md", "Markdown 文件 (*.md)"
        )
        if not file_path:
            return

        lines = []
        if title:
            lines.append(f"# {title}\n")

        for line in text.splitlines():
            l = line.strip()
            if not l:
                lines.append("")
                continue
            level = self.detect_heading_level(l)
            if level == 1:
                lines.append(f"## {l}")
            elif level == 2:
                lines.append(f"### {l}")
            elif level == 3:
                lines.append(f"#### {l}")
            elif level == 4:
                lines.append(f"##### {l}")
            else:
                lines.append(l)

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            QMessageBox.information(self, "导出成功", f"已导出：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出 Markdown 失败：\n{e}")

    def export_txt(self):
        text = self.fulltext_edit.toPlainText().strip()
        title = self.title_edit.text().strip()
        if not text:
            QMessageBox.warning(self, "无内容", "正文为空，无法导出。")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出 TXT", f"{title or '文稿'}.txt", "文本文件 (*.txt)"
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                if title:
                    f.write(title + "\n\n")
                f.write(text)
            QMessageBox.information(self, "导出成功", f"已导出：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出 TXT 失败：\n{e}")

    # ---------- 草稿自动保存 / 载入 ----------

    def auto_save(self):
        save_draft(
            self.title_edit.text().strip(),
            self.type_combo.currentText(),
            self.outline_edit.toPlainText(),
            self.fulltext_edit.toPlainText(),
        )

    def manual_save_draft(self):
        self.auto_save()
        QMessageBox.information(self, "草稿已保存", "当前标题、大纲、正文已保存到本地草稿。")

    def load_draft_if_any(self):
        draft = load_draft()
        if not draft:
            return
        self.title_edit.setText(draft.get("title", ""))
        doc_type = draft.get("doc_type", "期刊论文")
        idx = self.type_combo.findText(doc_type)
        if idx >= 0:
            self.type_combo.setCurrentIndex(idx)
        self.outline_edit.setPlainText(draft.get("outline", ""))
        self.fulltext_edit.setPlainText(draft.get("fulltext", ""))

    def closeEvent(self, event):
        self.auto_save()
        event.accept()

    # ---------- 设置 ----------

    def open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec_()


def main():
    app = QApplication(sys.argv)
    win = WriterApp()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
