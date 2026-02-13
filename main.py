import sys
import os
import json
import requests
from requests.exceptions import RequestException
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox,
    QFileDialog, QMessageBox, QDialog, QFormLayout
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QObject, QTimer
)
from PyQt6.QtGui import QFont
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ===================== 配置文件路径 =====================
CONFIG_PATH = "config.json"
# ======================================================

# ===================== 流式API调用线程 =====================
class StreamAPICaller(QThread):
    """流式API调用线程（避免界面卡死）"""
    new_content = pyqtSignal(str)  # 新内容信号
    finished_signal = pyqtSignal(bool, str)  # 完成信号（是否成功，错误信息）
    stopped = False  # 终止标记

    def __init__(self, api_key, prompt):
        super().__init__()
        self.api_key = api_key
        self.prompt = prompt
        self.session = requests.Session()
        self.request = None

    def run(self):
        """线程执行函数：流式调用DeepSeek API"""
        self.stopped = False
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": self.prompt}],
            "temperature": 0.2,
            "stream": True  # 开启流式输出
        }

        try:
            # 发起流式请求
            self.request = self.session.post(
                "https://api.deepseek.com/v1/chat/completions",
                json=data,
                headers=headers,
                stream=True,
                timeout=90
            )
            self.request.raise_for_status()

            # 逐行解析流式响应
            for line in self.request.iter_lines():
                if self.stopped:  # 检测终止信号
                    self.finished_signal.emit(False, "已终止撰写")
                    return
                if line:
                    line = line.decode('utf-8').strip()
                    if line.startswith('data: '):
                        line = line[6:]
                        if line == '[DONE]':
                            break
                        try:
                            json_data = json.loads(line)
                            if 'choices' in json_data and len(json_data['choices']) > 0:
                                delta = json_data['choices'][0].get('delta', {})
                                content = delta.get('content', '')
                                if content:
                                    self.new_content.emit(content)  # 发送新内容
                        except json.JSONDecodeError:
                            continue

            self.finished_signal.emit(True, "")
        except RequestException as e:
            error_msg = f"API调用失败：{str(e)}"
            if "401" in str(e):
                error_msg = "API调用失败：401未授权（Key无效/过期）"
            elif "403" in str(e):
                error_msg = "API调用失败：403禁止访问（余额不足）"
            elif "429" in str(e):
                error_msg = "API调用失败：429请求频繁（请稍后再试）"
            self.finished_signal.emit(False, error_msg)
        except Exception as e:
            self.finished_signal.emit(False, f"未知错误：{str(e)}")
        finally:
            # 关闭请求
            if self.request:
                self.request.close()

    def stop(self):
        """终止API调用"""
        self.stopped = True
        if self.request:
            self.request.close()

# ===================== 配置管理 =====================
class ConfigManager:
    """配置文件管理：保存/加载API Key"""
    @staticmethod
    def load_config():
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载配置失败: {e}")
                return {"deepseek_api_key": ""}
        return {"deepseek_api_key": ""}

    @staticmethod
    def save_api_key(api_key):
        config = {"deepseek_api_key": api_key.strip()}
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(None, "错误", f"保存配置失败: {str(e)}")

# ===================== API设置弹窗 =====================
class APISettingDialog(QDialog):
    """API Key 设置弹窗（适配中文输入）"""
    def __init__(self, current_key):
        super().__init__()
        self.setWindowTitle("API 设置")
        self.setFixedSize(500, 180)
        self.api_key = current_key
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        # API Key 输入框（强制启用中文输入）
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("请输入 DeepSeek API Key（支持中文粘贴）")
        self.key_input.setText(self.api_key)
        self.key_input.setEchoMode(QLineEdit.EchoMode.Password)
        # 修复Linux中文输入核心：启用输入法
        self.key_input.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.key_input.setAttribute(Qt.WidgetAttribute.WA_KeyCompression, False)
        form_layout.addRow("DeepSeek API Key：", self.key_input)

        # 保存按钮
        self.save_btn = QPushButton("✅ 保存并应用")
        self.save_btn.clicked.connect(self.save_key)
        form_layout.addRow("", self.save_btn)

        layout.addLayout(form_layout)
        self.setLayout(layout)

    def save_key(self):
        key = self.key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "提示", "API Key 不能为空")
            return
        ConfigManager.save_api_key(key)
        QMessageBox.information(self, "成功", "API Key 已保存，下次启动自动加载！")
        self.accept()

# ===================== 主窗口 =====================
class PaperWriter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager.load_config()
        self.DEEPSEEK_API_KEY = self.config.get("deepseek_api_key", "")
        self.stream_thread = None  # 流式调用线程
        self.setWindowTitle("智能公文/论文撰写工具 | 流式输出 | 多平台兼容")
        self.setMinimumSize(950, 780)
        self.init_ui()
        self.init_signal_slots()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # ========== 顶部：API 设置 + 状态 ==========
        top_layout = QHBoxLayout()
        self.api_status_label = QLabel()
        self.update_api_status()
        self.setting_btn = QPushButton("⚙️ API 设置")
        top_layout.addWidget(self.api_status_label)
        top_layout.addStretch()
        top_layout.addWidget(self.setting_btn)
        layout.addLayout(top_layout)

        # ========== 文稿类型 ==========
        type_layout = QHBoxLayout()
        type_label = QLabel("文稿类型：")
        self.type_combo = QComboBox()
        # 修复中文输入/显示
        self.type_combo.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.type_combo.addItems([
            "期刊论文", "工作计划", "工作总结", "学习反思", "教学案例", "汇报材料", "自定义"
        ])
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.type_combo)
        layout.addLayout(type_layout)

        # ========== 题目输入（修复中文输入） ==========
        title_layout = QHBoxLayout()
        title_label = QLabel("题目/要求：")
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("输入完整题目或详细要求，例如：2026年度部门工作总结")
        # 核心：启用输入法 + 禁用按键压缩（Linux中文输入关键）
        self.title_input.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.title_input.setAttribute(Qt.WidgetAttribute.WA_KeyCompression, False)
        type_layout.addWidget(title_label)
        title_layout.addWidget(self.title_input)
        layout.addLayout(title_layout)

        # ========== 大纲操作按钮组 ==========
        outline_btn_layout = QHBoxLayout()
        self.outline_btn = QPushButton("📌 生成标准公文大纲")
        self.stop_outline_btn = QPushButton("🛑 终止生成")
        self.stop_outline_btn.setEnabled(False)  # 默认禁用
        outline_btn_layout.addWidget(self.outline_btn)
        outline_btn_layout.addWidget(self.stop_outline_btn)
        layout.addLayout(outline_btn_layout)

        # ========== 大纲编辑区（修复中文输入） ==========
        layout.addWidget(QLabel("📝 大纲（纯文本公文层级，可直接修改）："))
        self.outline_edit = QTextEdit()
        self.outline_edit.setPlaceholderText("大纲格式：一、 →（一）→1. →（1），禁止使用Markdown")
        self.outline_edit.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.outline_edit.setAttribute(Qt.WidgetAttribute.WA_KeyCompression, False)
        layout.addWidget(self.outline_edit)

        # ========== 全文操作按钮组 ==========
        fulltext_btn_layout = QHBoxLayout()
        self.write_btn = QPushButton("🚀 按公文格式撰写完整文稿")
        self.stop_write_btn = QPushButton("🛑 终止撰写")
        self.stop_write_btn.setEnabled(False)  # 默认禁用
        fulltext_btn_layout.addWidget(self.write_btn)
        fulltext_btn_layout.addWidget(self.stop_write_btn)
        layout.addLayout(fulltext_btn_layout)

        # ========== 文稿展示 ==========
        layout.addWidget(QLabel("📄 完整文稿（纯文本无格式）："))
        self.result_text = QTextEdit()
        self.result_text.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.result_text.setAttribute(Qt.WidgetAttribute.WA_KeyCompression, False)
        layout.addWidget(self.result_text)

        # ========== 导出 + 清空按钮组 ==========
        action_btn_layout = QHBoxLayout()
        self.export_btn = QPushButton("📄 导出【国家标准公文格式】Word文档")
        self.clear_btn = QPushButton("🗑️ 清空所有内容")
        action_btn_layout.addWidget(self.export_btn)
        action_btn_layout.addWidget(self.clear_btn)
        layout.addLayout(action_btn_layout)

    def init_signal_slots(self):
        """初始化信号槽"""
        # 按钮点击事件
        self.setting_btn.clicked.connect(self.open_api_setting)
        self.outline_btn.clicked.connect(self.generate_outline)
        self.stop_outline_btn.clicked.connect(self.stop_outline_generation)
        self.write_btn.clicked.connect(self.generate_full_text)
        self.stop_write_btn.clicked.connect(self.stop_fulltext_generation)
        self.clear_btn.clicked.connect(self.clear_all_content)
        self.export_btn.clicked.connect(self.export_word)

    def update_api_status(self):
        """更新API状态显示"""
        if self.DEEPSEEK_API_KEY:
            self.api_status_label.setText("✅ API Key 已配置")
            self.api_status_label.setStyleSheet("color:green;")
        else:
            self.api_status_label.setText("❌ 未设置 API Key，请先配置")
            self.api_status_label.setStyleSheet("color:red;")

    def open_api_setting(self):
        """打开API设置弹窗"""
        dialog = APISettingDialog(self.DEEPSEEK_API_KEY)
        if dialog.exec():
            self.config = ConfigManager.load_config()
            self.DEEPSEEK_API_KEY = self.config.get("deepseek_api_key", "")
            self.update_api_status()

    def check_api_key(self):
        """检查API是否配置"""
        if not self.DEEPSEEK_API_KEY:
            QMessageBox.critical(self, "错误", "请先点击右上角【API 设置】配置 DeepSeek Key！")
            return False
        return True

    def clear_all_content(self):
        """清空所有输入/输出内容"""
        reply = QMessageBox.question(
            self, "确认", "是否清空所有内容？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.title_input.clear()
            self.outline_edit.clear()
            self.result_text.clear()

    def start_stream_thread(self, prompt, is_outline=True):
        """启动流式调用线程"""
        # 停止已有线程
        if self.stream_thread and self.stream_thread.isRunning():
            self.stream_thread.stop()
            self.stream_thread.wait()

        # 初始化UI状态
        if is_outline:
            self.outline_edit.clear()
            self.outline_btn.setEnabled(False)
            self.stop_outline_btn.setEnabled(True)
        else:
            self.result_text.clear()
            self.write_btn.setEnabled(False)
            self.stop_write_btn.setEnabled(True)

        # 创建并启动线程
        self.stream_thread = StreamAPICaller(self.DEEPSEEK_API_KEY, prompt)
        self.stream_thread.new_content.connect(lambda content: self.append_content(content, is_outline))
        self.stream_thread.finished_signal.connect(lambda success, msg: self.stream_finished(success, msg, is_outline))
        self.stream_thread.start()

    def append_content(self, content, is_outline):
        """追加流式内容到对应编辑框"""
        if is_outline:
            current = self.outline_edit.toPlainText()
            self.outline_edit.setPlainText(current + content)
            # 滚动到末尾
            self.outline_edit.verticalScrollBar().setValue(self.outline_edit.verticalScrollBar().maximum())
        else:
            current = self.result_text.toPlainText()
            self.result_text.setPlainText(current + content)
            self.result_text.verticalScrollBar().setValue(self.result_text.verticalScrollBar().maximum())

    def stream_finished(self, success, error_msg, is_outline):
        """流式调用完成后的处理"""
        # 恢复按钮状态
        if is_outline:
            self.outline_btn.setEnabled(True)
            self.stop_outline_btn.setEnabled(False)
        else:
            self.write_btn.setEnabled(True)
            self.stop_write_btn.setEnabled(False)

        # 显示错误信息
        if not success and error_msg:
            QMessageBox.critical(self, "错误", error_msg)

    def stop_outline_generation(self):
        """终止大纲生成"""
        if self.stream_thread and self.stream_thread.isRunning():
            self.stream_thread.stop()

    def stop_fulltext_generation(self):
        """终止全文撰写"""
        if self.stream_thread and self.stream_thread.isRunning():
            self.stream_thread.stop()

    def generate_outline(self):
        """生成大纲（流式）"""
        if not self.check_api_key():
            return
        doc_type = self.type_combo.currentText()
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "提示", "请输入题目或要求")
            return
        
        prompt = f"""
        你是专业公文写作助手，请为【{doc_type}】生成大纲。
        题目：{title}
        要求：
        1. 纯文本，绝对禁止任何Markdown、符号、表格、代码
        2. 严格使用国家标准公文层级：一、 →（一）→1. →（1）
        3. 结构清晰，可直接用于正式文稿
        只输出大纲，不要多余解释。
        """
        self.start_stream_thread(prompt, is_outline=True)

    def generate_full_text(self):
        """生成全文（流式）"""
        if not self.check_api_key():
            return
        doc_type = self.type_combo.currentText()
        title = self.title_input.text().strip()
        outline = self.outline_edit.toPlainText().strip()
        if not title or not outline:
            QMessageBox.warning(self, "提示", "请先生成并完善大纲")
            return
        
        prompt = f"""
        你是专业公文撰稿人，请按【{doc_type}】正式文体写作。
        题目：{title}
        大纲：{outline}
        要求：
        1. 纯文本，无任何Markdown、格式符、特殊符号
        2. 严格使用公文层级：一、 （一） 1. （1）
        3. 语言正式、逻辑严谨、内容完整
        4. 直接输出正文，不要前言、说明、解释
        """
        self.start_stream_thread(prompt, is_outline=False)

    def export_word(self):
        """导出国家标准公文格式Word"""
        title = self.title_input.text().strip()
        content = self.result_text.toPlainText().strip()
        if not title or not content:
            QMessageBox.warning(self, "提示", "请先生成完整文稿")
            return
        
        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出Word", f"{title}.docx", "Word文档 (*.docx)"
        )
        if not save_path:
            return
        
        try:
            doc = Document()
            # A4公文页面设置
            section = doc.sections[0]
            section.page_height = Cm(29.7)
            section.page_width = Cm(21.0)
            section.left_margin = Cm(2.8)
            section.right_margin = Cm(2.6)
            section.top_margin = Cm(3.7)
            section.bottom_margin = Cm(3.5)

            # 公文标题：二号小标宋体、居中
            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_p.add_run(title)
            title_run.font.size = Pt(22)
            title_run.font.bold = True
            title_run.font.name = "SimHei" if os.name == "posix" else "小标宋体"
            title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '小标宋体')
            doc.add_paragraph()

            # 正文按公文层级自动排版
            lines = content.splitlines()
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                p = doc.add_paragraph()
                run = p.add_run(line)
                run.font.size = Pt(16)  # 三号字

                # 适配Linux字体
                linux_font_map = {
                    "黑体": "SimHei",
                    "楷体_GB2312": "KaiTi",
                    "仿宋_GB2312": "FangSong"
                }

                # 一级标题：一、 黑体
                if line.startswith(("一、","二、","三、","四、","五、")):
                    font_name = linux_font_map["黑体"] if os.name == "posix" else "黑体"
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                    run.font.bold = True
                    p.paragraph_format.first_line_indent = Cm(0)
                # 二级标题：（一） 楷体
                elif line.startswith(("（一）","（二）","（三）")):
                    font_name = linux_font_map["楷体_GB2312"] if os.name == "posix" else "楷体_GB2312"
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '楷体_GB2312')
                    p.paragraph_format.first_line_indent = Cm(0)
                # 三级标题：1.  加粗
                elif line.startswith(("1.","2.","3.")):
                    run.font.bold = True
                    p.paragraph_format.first_line_indent = Cm(0)
                # 正文：仿宋_GB2312 + 首行缩进
                else:
                    font_name = linux_font_map["仿宋_GB2312"] if os.name == "posix" else "仿宋_GB2312"
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
                    p.paragraph_format.first_line_indent = Cm(0.74)
                p.paragraph_format.line_spacing = 1.25

            doc.save(save_path)
            QMessageBox.information(self, "成功", "已按【国家标准公文格式】导出Word！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

# ===================== 主程序入口 =====================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 适配Linux系统中文显示和输入
    if os.name == "posix":
        # 设置系统中文字体
        font = QFont("Noto Sans CJK SC")
        app.setFont(font)
        # 启用输入法支持
        app.setAttribute(Qt.ApplicationAttribute.AA_EnableInputMethods, True)

    window = PaperWriter()
    window.show()
    sys.exit(app.exec())
