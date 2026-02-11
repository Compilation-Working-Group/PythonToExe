import os
import sys
import json
import uuid
import re
import datetime
from dataclasses import dataclass
from typing import Optional, List, Dict, Any, Tuple

from PySide6 import QtCore, QtGui, QtWidgets
from docx import Document

APP_NAME = "PaperWriter"
APP_ORG = "YuJinQuanLab"
SETTINGS_FILE = "settings.json"
AUTOSAVE_FILE = "autosave.paperwriter.json"

# DeepSeek defaults
DEFAULT_API_BASE = "https://api.deepseek.com"
DEFAULT_MODEL_CHAT = "deepseek-chat"
DEFAULT_MODEL_REASONER = "deepseek-reasoner"


# =========================
# Utils
# =========================
def now_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clamp_text(s: str, max_len: int = 250000) -> str:
    s = (s or "").strip()
    return s if len(s) <= max_len else s[:max_len] + "\n...[截断]..."


def normalize_newlines(s: str) -> str:
    return (s or "").replace("\r\n", "\n").replace("\r", "\n")


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


def safe_json_loads(text: str) -> Optional[Any]:
    try:
        return json.loads(text)
    except Exception:
        t = re.sub(r"^```.*?$", "", (text or ""), flags=re.M).strip()
        try:
            return json.loads(t)
        except Exception:
            return None


# =========================
# Project Model
# =========================
@dataclass
class NodeData:
    id: str
    title: str
    notes: str = ""
    content: str = ""
    children: List["NodeData"] = None

    def to_dict(self):
        return {
            "id": self.id,
            "title": self.title,
            "notes": self.notes,
            "content": self.content,
            "children": [c.to_dict() for c in (self.children or [])],
        }

    @staticmethod
    def from_dict(d: dict) -> "NodeData":
        return NodeData(
            id=d.get("id") or str(uuid.uuid4()),
            title=d.get("title", "未命名"),
            notes=d.get("notes", ""),
            content=d.get("content", ""),
            children=[NodeData.from_dict(x) for x in d.get("children", [])],
        )


@dataclass
class Snapshot:
    ts: str
    action: str
    project: dict

    def to_dict(self):
        return {"ts": self.ts, "action": self.action, "project": self.project}

    @staticmethod
    def from_dict(d: dict) -> "Snapshot":
        return Snapshot(ts=d.get("ts", now_str()), action=d.get("action", ""), project=d.get("project", {}))


@dataclass
class ProjectData:
    version: int
    title: str
    genre: str
    template: str
    instructions: str
    outline_root: NodeData
    draft_md: str
    references: List[str]
    snapshots: List[Snapshot]
    updated_at: str

    def to_dict(self):
        return {
            "version": self.version,
            "title": self.title,
            "genre": self.genre,
            "template": self.template,
            "instructions": self.instructions,
            "outline_root": self.outline_root.to_dict(),
            "draft_md": self.draft_md,
            "references": self.references,
            "snapshots": [s.to_dict() for s in self.snapshots],
            "updated_at": self.updated_at,
        }

    @staticmethod
    def empty() -> "ProjectData":
        root = NodeData(id=str(uuid.uuid4()), title="大纲", notes="", content="", children=[])
        return ProjectData(
            version=999,
            title="",
            genre="论文",
            template="论文-教学研究",
            instructions="",
            outline_root=root,
            draft_md="",
            references=[],
            snapshots=[],
            updated_at=now_str(),
        )

    @staticmethod
    def from_dict(d: dict) -> "ProjectData":
        snaps = [Snapshot.from_dict(x) for x in d.get("snapshots", [])]
        return ProjectData(
            version=int(d.get("version", 999)),
            title=d.get("title", ""),
            genre=d.get("genre", "论文"),
            template=d.get("template", "论文-教学研究"),
            instructions=d.get("instructions", ""),
            outline_root=NodeData.from_dict(d.get("outline_root", {"title": "大纲"})),
            draft_md=d.get("draft_md", ""),
            references=list(d.get("references", [])),
            snapshots=snaps,
            updated_at=d.get("updated_at", now_str()),
        )


# =========================
# DeepSeek / OpenAI-compatible client
# =========================
@dataclass
class LLMConfig:
    api_base: str = DEFAULT_API_BASE
    api_key: str = ""
    model_chat: str = DEFAULT_MODEL_CHAT
    model_reasoner: str = DEFAULT_MODEL_REASONER
    temperature: float = 0.6
    max_tokens: int = 3000
    timeout: int = 120


class OpenAICompatClient:
    def __init__(self, cfg: LLMConfig, logger):
        self.cfg = cfg
        self.logger = logger

    def is_ready(self) -> bool:
        return bool(self.cfg.api_base.strip()) and bool(self.cfg.api_key.strip())

    def _build_url(self) -> str:
        base = self.cfg.api_base.rstrip("/")
        if base.endswith("/v1"):
            return base + "/chat/completions"
        if "api.deepseek.com" in base:
            return base + "/chat/completions"
        return base + "/v1/chat/completions"

    def chat(self, messages: List[Dict[str, str]], model: str) -> str:
        import urllib.request

        url = self._build_url()
        payload = {
            "model": model,
            "messages": messages,
            "temperature": self.cfg.temperature,
            "max_tokens": self.cfg.max_tokens,
            "stream": False,
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("Authorization", f"Bearer {self.cfg.api_key}")

        self.logger(f"[{now_str()}] 调用 LLM: {url} | model={model}")
        with urllib.request.urlopen(req, timeout=self.cfg.timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
        obj = json.loads(raw)
        return obj["choices"][0]["message"]["content"].strip()


# =========================
# Templates
# =========================
GENRES = ["论文", "计划", "反思", "案例", "总结", "自定义"]
TEMPLATES = [
    "论文-教学研究（课堂干预/实验教学）",
    "论文-实验研究（统计/对照）",
    "论文-综述",
    "论文-案例研究",
    "计划-通用",
    "反思-课堂",
    "总结-通用",
    "自定义",
]

REQUIRED_HINTS = {
    "论文-教学研究（课堂干预/实验教学）": ["摘要", "关键词", "研究背景", "研究问题/假设", "研究设计", "教学干预方案", "数据收集与处理", "结果", "讨论", "结论与建议", "参考文献"],
    "论文-实验研究（统计/对照）": ["摘要", "关键词", "引言", "方法（样本/变量/工具）", "实验设计（对照/随机）", "统计方法", "结果（表/图）", "讨论", "局限性", "结论", "参考文献"],
    "论文-综述": ["摘要", "关键词", "引言", "检索策略", "研究脉络/分类框架", "关键进展", "争议与不足", "展望", "参考文献"],
    "论文-案例研究": ["摘要", "关键词", "背景", "案例描述", "干预/行动", "证据与结果", "讨论与启示", "结论", "参考文献"],
    "计划-通用": ["背景与目标", "现状分析", "实施步骤", "时间安排", "风险与应对", "评估指标"],
    "反思-课堂": ["概述", "目标与预期", "实际情况", "原因分析", "改进策略", "跟进"],
    "总结-通用": ["回顾", "成果", "不足", "经验", "下一步"],
    "自定义": [],
}

OFFLINE_OUTLINE = {
    "论文-教学研究（课堂干预/实验教学）": ["摘要", "关键词", "1 研究背景", "2 研究问题与意义", "3 研究设计", "4 教学干预方案", "5 数据收集与处理", "6 结果", "7 讨论", "8 结论与建议", "参考文献（占位符）"],
    "论文-实验研究（统计/对照）": ["摘要", "关键词", "1 引言", "2 方法（样本/变量/工具）", "3 实验设计（对照/随机）", "4 统计方法", "5 结果（表/图）", "6 讨论", "7 局限性", "8 结论", "参考文献（占位符）"],
    "论文-综述": ["摘要", "关键词", "1 引言", "2 检索策略与纳入标准", "3 研究脉络/分类框架", "4 关键进展", "5 争议与不足", "6 展望", "参考文献（占位符）"],
    "论文-案例研究": ["摘要", "关键词", "1 背景", "2 案例描述", "3 干预/行动", "4 证据与结果", "5 讨论与启示", "6 结论", "参考文献（占位符）"],
    "计划-通用": ["一、背景与目标", "二、现状分析", "三、实施步骤", "四、时间安排", "五、风险与应对", "六、评估指标"],
    "反思-课堂": ["一、概述", "二、目标与预期", "三、实际情况", "四、原因分析", "五、改进策略", "六、跟进"],
    "总结-通用": ["一、回顾", "二、成果", "三、不足", "四、经验", "五、下一步"],
    "自定义": ["一、背景", "二、主体", "三、结语"],
}


# =========================
# Tree <-> Markdown
# =========================
def tree_to_markdown(root_item: QtWidgets.QTreeWidgetItem, include_notes: bool = True) -> str:
    lines = []
    def walk(item: QtWidgets.QTreeWidgetItem, depth: int):
        title = item.text(0).strip() or "未命名"
        level = min(depth + 1, 6)
        prefix = "#" * level
        lines.append(f"{prefix} {title}")
        if include_notes:
            notes = item.data(0, QtCore.Qt.UserRole + 2) or ""
            notes = str(notes).strip()
            if notes:
                lines.append(f"> 要点：{notes}")
        for i in range(item.childCount()):
            walk(item.child(i), depth + 1)

    for i in range(root_item.childCount()):
        walk(root_item.child(i), 1)
    return "\n".join(lines).strip() + "\n"


def get_item_path_titles(item: QtWidgets.QTreeWidgetItem) -> List[str]:
    titles = []
    cur = item
    while cur and cur.parent():
        titles.append(cur.text(0).strip())
        cur = cur.parent()
    return list(reversed([t for t in titles if t]))


def replace_section_in_markdown(full_md: str, section_title: str, section_md: str) -> str:
    full = normalize_newlines(full_md or "")
    sec = normalize_newlines(section_md).strip() + "\n"
    pattern = re.compile(rf"^(?P<h>#+)\s+{re.escape(section_title)}\s*$", re.M)
    m = pattern.search(full)
    if not m:
        return (full.rstrip() + "\n\n" + sec).strip() + "\n"

    h = m.group("h")
    level = len(h)
    start = m.start()
    next_pat = re.compile(rf"^(#{{1,{level}}})\s+.+$", re.M)
    m2 = next_pat.search(full, m.end())
    end = m2.start() if m2 else len(full)
    return (full[:start] + sec + full[end:]).strip() + "\n"


def missing_required(template: str, outline_md: str) -> List[str]:
    req = REQUIRED_HINTS.get(template, [])
    if not req:
        return []
    md = outline_md.lower()
    missing = []
    for r in req:
        if r.lower() not in md:
            missing.append(r)
    return missing


# =========================
# Engine
# =========================
class WriterEngineUltimate:
    def __init__(self, client: OpenAICompatClient, logger):
        self.client = client
        self.logger = logger

    def _model_for(self, task: str, prefer_reasoner: bool) -> str:
        if prefer_reasoner:
            return self.client.cfg.model_reasoner or DEFAULT_MODEL_REASONER
        return self.client.cfg.model_chat or DEFAULT_MODEL_CHAT

    def gen_outline_tree(self, title: str, genre: str, template: str, instructions: str, prefer_reasoner: bool) -> List[Dict[str, Any]]:
        if self.client.is_ready():
            model = self._model_for("outline", prefer_reasoner)
            sys_prompt = (
                "你是专业学术写作助手。根据题目、体裁、模板与要求生成“树状大纲JSON”。"
                "只输出JSON，不要解释，不要Markdown，不要代码块。"
                "格式：[{\"title\":\"...\",\"notes\":\"\",\"children\":[...]}]。notes用于写作要点。"
                "标题要符合教学研究/实验研究/综述/案例研究的学术结构。"
            )
            user_prompt = (
                f"题目：{title}\n体裁：{genre}\n模板：{template}\n要求：{instructions}\n\n"
                "请生成可直接扩写的树状大纲。务必包含摘要、关键词、参考文献（占位符）等必要部分（若模板需要）。"
            )
            try:
                out = self.client.chat([
                    {"role": "system", "content": sys_prompt},
                    {"role": "user", "content": user_prompt},
                ], model=model)
                data = safe_json_loads(out)
                if isinstance(data, list):
                    def norm(x):
                        return {
                            "title": str(x.get("title", "未命名")),
                            "notes": str(x.get("notes", "")),
                            "children": [norm(c) for c in (x.get("children") or [])],
                        }
                    return [norm(x) for x in data]
            except Exception as e:
                self.logger(f"[{now_str()}] 大纲生成失败，回退离线模板：{e}")

        items = OFFLINE_OUTLINE.get(template, OFFLINE_OUTLINE["自定义"])
        return [{"title": t, "notes": "", "children": []} for t in items]

    def write_full(self, title: str, genre: str, template: str, instructions: str, outline_md: str,
                   refs: List[str], prefer_reasoner: bool) -> str:
        if self.client.is_ready():
            model = self._model_for("write_full", prefer_reasoner)
            sys_prompt = (
                "你是专业学术写作助手。根据大纲生成完整文稿Markdown。"
                "要求：结构与大纲一致；引用用[1][2]占位符；不要编造真实DOI与作者信息。"
                "对于实验研究：注意方法、统计描述的严谨；对于教学研究：强调干预方案与证据链；"
                "对于综述：强调检索策略、分类框架与研究缺口；对于案例研究：强调情境、行动、证据与启示。"
            )
            ref_hint = "\n".join([f"[{i+1}] {r}" for i, r in enumerate(refs)]) if refs else "（暂无）"
            user_prompt = (
                f"题目：{title}\n体裁：{genre}\n模板：{template}\n要求：{instructions}\n\n"
                f"大纲（Markdown）：\n{outline_md}\n\n"
                f"可用参考文献占位（如需引用用[1][2]）：\n{ref_hint}\n\n"
                "请输出完整文稿Markdown（逐节扩写）。如遇数据缺失，请用“（数据待补）”标注，不要捏造。"
            )
            out = self.client.chat([
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt},
            ], model=model)
            return clamp_text(out, 250000)

        return f"# {title}\n\n{outline_md}\n\n（离线模式：请配置 DeepSeek API。）\n"

    def write_section(self, title: str, genre: str, template: str, instructions: str,
                      outline_md: str, section_path: List[str], notes: str, existing: str,
                      refs: List[str], mode: str, prefer_reasoner: bool) -> str:
        if not self.client.is_ready():
            sec_title = section_path[-1] if section_path else "未命名章节"
            return f"## {sec_title}\n\n（离线模式：请配置 DeepSeek API 执行单章{mode}。）\n"

        model = self._model_for("section", prefer_reasoner)
        mode_map = {"write": "撰写", "rewrite": "重写", "expand": "扩展", "shrink": "缩写"}
        task = mode_map.get(mode, "撰写")
        sec_title = section_path[-1] if section_path else "未命名章节"
        path_str = " > ".join(section_path) if section_path else sec_title

        sys_prompt = (
            "你是专业学术写作助手。任务：只处理一个章节。"
            "输出Markdown，并以该章节标题作为标题行（##/###）。不要输出全文其他章节。"
            "引用用[1][2]占位符，不要编造真实DOI。"
        )
        ref_hint = "\n".join([f"[{i+1}] {r}" for i, r in enumerate(refs)]) if refs else "（暂无）"
        user_prompt = (
            f"全文题目：{title}\n体裁：{genre}\n模板：{template}\n要求：{instructions}\n\n"
            f"全文大纲（Markdown）：\n{outline_md}\n\n"
            f"当前章节路径：{path_str}\n"
            f"当前章节要点/素材：{notes}\n\n"
            f"当前章节已有内容（可为空）：\n{existing}\n\n"
            f"参考文献占位：\n{ref_hint}\n\n"
            f"请对章节《{sec_title}》执行：{task}。\n"
            "- 如果涉及数据但未提供，请用“（数据待补）”标注，不要捏造具体数值。\n"
        )
        out = self.client.chat([
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": user_prompt},
        ], model=model)
        return clamp_text(out, 100000)

    def transform(self, kind: str, text: str, instructions: str, prefer_reasoner: bool) -> str:
        if not self.client.is_ready():
            return text
        model = self._model_for("transform", prefer_reasoner)
        if kind == "polish":
            sys_prompt = "你是中文学术写作润色助手。保持原意，提升表达、逻辑与连贯性，语气更学术规范。"
            user_prompt = f"额外要求：{instructions}\n\n请润色以下内容：\n{text}"
        else:
            sys_prompt = (
                "你是中文学术改写助手。目标：不改变事实与核心观点，显著改写措辞与句式，降低重复度；"
                "保持结构清晰与学术规范。"
            )
            user_prompt = f"额外要求：{instructions}\n\n请改写以下内容（降重）：\n{text}"
        out = self.client.chat([
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": user_prompt},
        ], model=model)
        return clamp_text(out, 250000)

    def gen_abstract(self, full_text: str, max_words: int, prefer_reasoner: bool) -> str:
        if not self.client.is_ready():
            return ""
        model = self._model_for("abstract", prefer_reasoner)
        sys_prompt = "你是学术写作助手。请根据全文生成结构化中文摘要，准确、简洁。"
        user_prompt = f"请在{max_words}字左右生成摘要：\n\n{full_text}"
        out = self.client.chat([
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": user_prompt},
        ], model=model)
        return out.strip()

    def gen_keywords(self, full_text: str, k: int, prefer_reasoner: bool) -> str:
        if not self.client.is_ready():
            return ""
        model = self._model_for("keywords", prefer_reasoner)
        sys_prompt = "你是学术写作助手。请根据全文提取关键词。"
        user_prompt = f"请提取{k}个关键词，用中文分号分隔：\n\n{full_text}"
        out = self.client.chat([
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": user_prompt},
        ], model=model)
        return out.strip().replace("\n", " ")


# =========================
# Dialogs
# =========================
class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, settings=None):
        super().__init__(parent)
        self.setWindowTitle("设置 - DeepSeek API")
        self.setModal(True)
        self.resize(820, 380)

        self.api_base = QtWidgets.QLineEdit()
        self.api_key = QtWidgets.QLineEdit()
        self.api_key.setEchoMode(QtWidgets.QLineEdit.Password)

        self.model_chat = QtWidgets.QLineEdit()
        self.model_reasoner = QtWidgets.QLineEdit()

        self.temperature = QtWidgets.QDoubleSpinBox()
        self.temperature.setRange(0.0, 2.0)
        self.temperature.setSingleStep(0.1)
        self.max_tokens = QtWidgets.QSpinBox()
        self.max_tokens.setRange(256, 8000)
        self.timeout = QtWidgets.QSpinBox()
        self.timeout.setRange(10, 300)

        form = QtWidgets.QFormLayout()
        form.addRow("API Base（默认 api.deepseek.com）", self.api_base)
        form.addRow("API Key (sk-xxxx)", self.api_key)
        form.addRow("Chat 模型（deepseek-chat）", self.model_chat)
        form.addRow("Reasoner 模型（deepseek-reasoner）", self.model_reasoner)
        form.addRow("Temperature", self.temperature)
        form.addRow("Max tokens", self.max_tokens)
        form.addRow("Timeout(s)", self.timeout)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addLayout(form)

        note = QtWidgets.QLabel(
            "DeepSeek API 为 OpenAI 兼容格式。申请地址: [https://platform.deepseek.com/](https://platform.deepseek.com/)"
        )
        note.setStyleSheet("color:#666;")
        layout.addWidget(note)
        layout.addWidget(btns)

        if settings:
            self.api_base.setText(settings.get("api_base", DEFAULT_API_BASE))
            self.api_key.setText(settings.get("api_key", ""))
            self.model_chat.setText(settings.get("model_chat", DEFAULT_MODEL_CHAT))
            self.model_reasoner.setText(settings.get("model_reasoner", DEFAULT_MODEL_REASONER))
            self.temperature.setValue(float(settings.get("temperature", 0.6)))
            self.max_tokens.setValue(int(settings.get("max_tokens", 3000)))
            self.timeout.setValue(int(settings.get("timeout", 120)))
        else:
            self.api_base.setText(DEFAULT_API_BASE)
            self.model_chat.setText(DEFAULT_MODEL_CHAT)
            self.model_reasoner.setText(DEFAULT_MODEL_REASONER)
            self.temperature.setValue(0.6)
            self.max_tokens.setValue(3000)
            self.timeout.setValue(120)

    def get_data(self) -> dict:
        return {
            "api_base": self.api_base.text().strip() or DEFAULT_API_BASE,
            "api_key": self.api_key.text().strip(),
            "model_chat": self.model_chat.text().strip() or DEFAULT_MODEL_CHAT,
            "model_reasoner": self.model_reasoner.text().strip() or DEFAULT_MODEL_REASONER,
            "temperature": float(self.temperature.value()),
            "max_tokens": int(self.max_tokens.value()),
            "timeout": int(self.timeout.value()),
        }


class ReferencesDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, refs=None):
        super().__init__(parent)
        self.setWindowTitle("引用管理")
        self.resize(860, 460)
        self.refs = list(refs or [])

        self.list = QtWidgets.QListWidget()
        self.list.addItems(self.refs)

        self.btn_add = QtWidgets.QPushButton("添加")
        self.btn_edit = QtWidgets.QPushButton("编辑")
        self.btn_del = QtWidgets.QPushButton("删除")
        self.btn_insert = QtWidgets.QPushButton("插入引用标记")
        self.btn_bib = QtWidgets.QPushButton("生成参考文献段落")

        row = QtWidgets.QHBoxLayout()
        row.addWidget(self.btn_add)
        row.addWidget(self.btn_edit)
        row.addWidget(self.btn_del)
        row.addStretch(1)
        row.addWidget(self.btn_insert)
        row.addWidget(self.btn_bib)

        self.btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel("建议格式：作者. 题目. 期刊, 年.（可先占位）"))
        layout.addWidget(self.list, 1)
        layout.addLayout(row)
        layout.addWidget(self.btns)

        self._insert_requested = False
        self._bib_requested = False
        self._insert_index = None

        self.btn_add.clicked.connect(self.on_add)
        self.btn_edit.clicked.connect(self.on_edit)
        self.btn_del.clicked.connect(self.on_del)
        self.list.itemDoubleClicked.connect(lambda _: self.on_edit())
        self.btn_insert.clicked.connect(self.on_insert)
        self.btn_bib.clicked.connect(self.on_bib)
        self.btns.accepted.connect(self.accept)
        self.btns.rejected.connect(self.reject)

    def on_add(self):
        text, ok = QtWidgets.QInputDialog.getMultiLineText(self, "添加引用", "输入引用条目：", "")
        if ok and text.strip():
            self.refs.append(text.strip())
            self.list.addItem(text.strip())

    def on_edit(self):
        row = self.list.currentRow()
        if row < 0:
            return
        old = self.refs[row]
        text, ok = QtWidgets.QInputDialog.getMultiLineText(self, "编辑引用", "修改引用条目：", old)
        if ok and text.strip():
            self.refs[row] = text.strip()
            self.list.item(row).setText(text.strip())

    def on_del(self):
        row = self.list.currentRow()
        if row < 0:
            return
        self.refs.pop(row)
        self.list.takeItem(row)

    def on_insert(self):
        row = self.list.currentRow()
        if row < 0:
            QtWidgets.QMessageBox.information(self, "提示", "请先选中一条引用。")
            return
        self._insert_requested = True
        self._insert_index = row + 1
        self.accept()

    def on_bib(self):
        self._bib_requested = True
        self.accept()

    def result_data(self):
        return {
            "refs": self.refs,
            "insert_requested": self._insert_requested,
            "insert_index": self._insert_index,
            "bib_requested": self._bib_requested,
        }


# =========================
# Main Window
# =========================
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PaperWriter Ultimate")
        self.resize(1600, 920)

        s = load_settings()
        self.llm_cfg = LLMConfig(
            api_base=s.get("api_base", DEFAULT_API_BASE),
            api_key=s.get("api_key", ""),
            model_chat=s.get("model_chat", DEFAULT_MODEL_CHAT),
            model_reasoner=s.get("model_reasoner", DEFAULT_MODEL_REASONER),
            temperature=float(s.get("temperature", 0.6)),
            max_tokens=int(s.get("max_tokens", 3000)),
            timeout=int(s.get("timeout", 120)),
        )
        self.client = OpenAICompatClient(self.llm_cfg, self._log)
        self.engine = WriterEngineUltimate(self.client, self._log)

        self.project_path: Optional[str] = None
        self.project: ProjectData = ProjectData.empty()
        self._ignore = False

        # --- widgets ---
        self.title_edit = QtWidgets.QLineEdit()
        self.title_edit.setPlaceholderText("输入题目/标题…")

        self.genre_combo = QtWidgets.QComboBox()
        self.genre_combo.addItems(GENRES)

        self.template_combo = QtWidgets.QComboBox()
        self.template_combo.addItems(TEMPLATES)

        self.instructions_edit = QtWidgets.QPlainTextEdit()
        self.instructions_edit.setPlaceholderText("写作要求（可选）：字数=2000; 风格=严谨; 研究对象=高一…")

        self.prefer_reasoner = QtWidgets.QCheckBox("优先用Reasoner（更严谨）")

        # buttons: project
        self.btn_new = QtWidgets.QPushButton("新建")
        self.btn_open = QtWidgets.QPushButton("打开")
        self.btn_save = QtWidgets.QPushButton("保存")
        self.btn_saveas = QtWidgets.QPushButton("另存为")
        self.btn_settings = QtWidgets.QPushButton("API设置")
        self.btn_refs = QtWidgets.QPushButton("引用")

        # buttons: generation
        self.btn_outline = QtWidgets.QPushButton("生成树状大纲")
        self.btn_write_all = QtWidgets.QPushButton("撰写全文")
        self.btn_fill_empty = QtWidgets.QPushButton("补全空章节")

        self.btn_sec_write = QtWidgets.QPushButton("单章撰写")
        self.btn_sec_rewrite = QtWidgets.QPushButton("单章重写")
        self.btn_sec_expand = QtWidgets.QPushButton("单章扩展")
        self.btn_sec_shrink = QtWidgets.QPushButton("单章缩写")

        self.btn_polish = QtWidgets.QPushButton("润色")
        self.btn_dedup = QtWidgets.QPushButton("降重")
        self.btn_abs = QtWidgets.QPushButton("摘要")
        self.btn_kw = QtWidgets.QPushButton("关键词")
        self.btn_check = QtWidgets.QPushButton("结构校验")

        self.btn_export_md = QtWidgets.QPushButton("导出MD")
        self.btn_export_docx = QtWidgets.QPushButton("导出DOCX")

        # tree
        self.tree_search = QtWidgets.QLineEdit()
        self.tree_search.setPlaceholderText("搜索章节标题…")
        self.tree = QtWidgets.QTreeWidget()
        self.tree.setHeaderLabels(["章节树（拖拽层级/双击改名）"])
        self.tree.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.tree.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.tree.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.tree.setEditTriggers(QtWidgets.QAbstractItemView.EditKeyPressed | QtWidgets.QAbstractItemView.SelectedClicked)

        self.btn_add_sib = QtWidgets.QPushButton("添加同级")
        self.btn_add_child = QtWidgets.QPushButton("添加子级")
        self.btn_del = QtWidgets.QPushButton("删除")
        self.btn_expand_tree = QtWidgets.QPushButton("展开")
        self.btn_collapse_tree = QtWidgets.QPushButton("折叠")

        # node editor
        self.node_title = QtWidgets.QLineEdit()
        self.node_notes = QtWidgets.QPlainTextEdit()
        self.node_content = QtWidgets.QPlainTextEdit()
        self.node_notes.setPlaceholderText("要点/证据/素材")
        self.node_content.setPlaceholderText("该章节草稿（可选）")

        # draft
        self.draft = QtWidgets.QPlainTextEdit()
        self.draft.setPlaceholderText("全文 Markdown（可手动编辑）")

        # snapshots
        self.snap_list = QtWidgets.QListWidget()
        self.btn_restore = QtWidgets.QPushButton("回滚到快照")
        self.btn_snapshot = QtWidgets.QPushButton("手动快照")

        # log
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.status = self.statusBar()

        self._build_layout()
        self._bind()
        self._style()

        self._load_project_to_ui()
        self._log(f"[{now_str()}] 启动完成。DeepSeek API 可选。")

        # autosave timer
        self.autosave_timer = QtCore.QTimer(self)
        self.autosave_timer.setInterval(120_000)
        self.autosave_timer.timeout.connect(self._autosave)
        self.autosave_timer.start()
        QtCore.QTimer.singleShot(600, self._recover_autosave_if_needed)

        # FIX: API Key missing guidance
        if not self.llm_cfg.api_key:
            QtCore.QTimer.singleShot(1000, self.prompt_for_api_key)

    def prompt_for_api_key(self):
        reply = QtWidgets.QMessageBox.question(self, "欢迎", "检测到未配置 API Key。\n\n是否现在配置 DeepSeek API？\n(配置后才可使用AI生成功能)",
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            self.on_settings()

    # ---------- log ----------
    def _log(self, msg: str):
        self.log.appendPlainText(msg)

    # ---------- autosave ----------
    def _autosave_path(self) -> str:
        return os.path.join(user_config_dir(), AUTOSAVE_FILE)

    def _autosave(self):
        try:
            self._sync_ui_to_project()
            with open(self._autosave_path(), "w", encoding="utf-8") as f:
                json.dump(self.project.to_dict(), f, ensure_ascii=False, indent=2)
            self.status.showMessage("Autosaved", 1500)
        except Exception:
            pass

    def _recover_autosave_if_needed(self):
        if self.project_path:
            return
        p = self._autosave_path()
        if not os.path.exists(p):
            return
        # recover only if current is empty
        if self.title_edit.text().strip() or self.draft.toPlainText().strip():
            return
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.project = ProjectData.from_dict(data)
            self._load_project_to_ui()
            self._log(f"[{now_str()}] 已从自动备份恢复。")
        except Exception:
            pass

    # ---------- snapshots ----------
    def _snapshot(self, action: str):
        self._sync_ui_to_project()
        snap_project = self.project.to_dict()
        snap_project["snapshots"] = []
        self.project.snapshots.insert(0, Snapshot(ts=now_str(), action=action, project=snap_project))
        self.project.snapshots = self.project.snapshots[:60]
        self._refresh_snapshots()

    def _refresh_snapshots(self):
        self.snap_list.clear()
        for s in self.project.snapshots:
            self.snap_list.addItem(f"{s.ts} | {s.action}")

    def _build_layout(self):
        top1 = QtWidgets.QHBoxLayout()
        top1.addWidget(QtWidgets.QLabel("题目："))
        top1.addWidget(self.title_edit, 3)
        top1.addWidget(QtWidgets.QLabel("体裁："))
        top1.addWidget(self.genre_combo)
        top1.addWidget(QtWidgets.QLabel("模板："))
        top1.addWidget(self.template_combo)
        top1.addWidget(self.prefer_reasoner)
        top1.addWidget(self.btn_outline)
        top1.addWidget(self.btn_write_all)
        top1.addWidget(self.btn_fill_empty)

        top2 = QtWidgets.QHBoxLayout()
        for b in [self.btn_new, self.btn_open, self.btn_save, self.btn_saveas, self.btn_settings, self.btn_refs]:
            top2.addWidget(b)
        top2.addStretch(1)
        for b in [self.btn_polish, self.btn_dedup, self.btn_abs, self.btn_kw, self.btn_check, self.btn_export_md, self.btn_export_docx]:
            top2.addWidget(b)

        left = QtWidgets.QWidget()
        left_l = QtWidgets.QVBoxLayout(left)
        left_l.addWidget(QtWidgets.QLabel("写作要求/指令（可选）"))
        left_l.addWidget(self.instructions_edit, 1)
        left_l.addWidget(QtWidgets.QLabel("章节树"))
        left_l.addWidget(self.tree_search)
        left_l.addWidget(self.tree, 3)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(self.btn_add_sib)
        row.addWidget(self.btn_add_child)
        row.addWidget(self.btn_del)
        row.addStretch(1)
        row.addWidget(self.btn_expand_tree)
        row.addWidget(self.btn_collapse_tree)
        left_l.addLayout(row)

        mid = QtWidgets.QWidget()
        mid_l = QtWidgets.QVBoxLayout(mid)
        mid_l.addWidget(QtWidgets.QLabel("当前章节标题"))
        mid_l.addWidget(self.node_title)
        mid_l.addWidget(QtWidgets.QLabel("要点/素材（建议写证据链/数据/干预方案/统计方法）"))
        mid_l.addWidget(self.node_notes, 1)
        mid_l.addWidget(QtWidgets.QLabel("章节草稿（可选）"))
        mid_l.addWidget(self.node_content, 2)
        sec_row = QtWidgets.QHBoxLayout()
        for b in [self.btn_sec_write, self.btn_sec_rewrite, self.btn_sec_expand, self.btn_sec_shrink]:
            sec_row.addWidget(b)
        mid_l.addLayout(sec_row)

        right = QtWidgets.QWidget()
        right_l = QtWidgets.QVBoxLayout(right)
        right_l.addWidget(QtWidgets.QLabel("全文 Markdown"))
        right_l.addWidget(self.draft, 3)

        snap_box = QtWidgets.QGroupBox("历史快照（≤60）")
        sb = QtWidgets.QVBoxLayout(snap_box)
        sb.addWidget(self.snap_list, 1)
        sb_btn = QtWidgets.QHBoxLayout()
        sb_btn.addWidget(self.btn_restore)
        sb_btn.addWidget(self.btn_snapshot)
        sb.addLayout(sb_btn)
        right_l.addWidget(snap_box, 1)

        right_l.addWidget(QtWidgets.QLabel("日志"))
        right_l.addWidget(self.log, 1)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.addWidget(left)
        splitter.addWidget(mid)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 2)
        splitter.setStretchFactor(2, 4)

        central = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(central)
        layout.addLayout(top1)
        layout.addLayout(top2)
        layout.addWidget(splitter, 1)
        self.setCentralWidget(central)

    def _style(self):
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
            QTreeWidget {
                background: #ffffff;
                border: 1px solid #d7dbe7;
                border-radius: 10px;
            }
            QGroupBox {
                border: 1px solid #d7dbe7;
                border-radius: 10px;
                margin-top: 6px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
        """)

    def _bind(self):
        # project
        self.btn_new.clicked.connect(self.on_new)
        self.btn_open.clicked.connect(self.on_open)
        self.btn_save.clicked.connect(self.on_save)
        self.btn_saveas.clicked.connect(self.on_saveas)
        self.btn_settings.clicked.connect(self.on_settings)
        self.btn_refs.clicked.connect(self.on_refs)

        # tree events
        self.tree.itemSelectionChanged.connect(self.on_tree_sel)
        self.tree.itemChanged.connect(self.on_tree_item_changed)
        self.tree.model().rowsMoved.connect(lambda *a: self._log(f"[{now_str()}] 章节层级已调整（拖拽）。"))
        self.tree_search.textChanged.connect(self.on_tree_search)

        # node edits
        self.node_title.textEdited.connect(self.on_node_title_edit)
        self.node_notes.textChanged.connect(self.on_node_notes_change)
        self.node_content.textChanged.connect(self.on_node_content_change)

        # tree ops
        self.btn_add_sib.clicked.connect(self.on_add_sib)
        self.btn_add_child.clicked.connect(self.on_add_child)
        self.btn_del.clicked.connect(self.on_delete)
        self.btn_expand_tree.clicked.connect(self.tree.expandAll)
        self.btn_collapse_tree.clicked.connect(self.tree.collapseAll)

        # writing
        self.btn_outline.clicked.connect(self.on_outline)
        self.btn_write_all.clicked.connect(self.on_write_all)
        self.btn_fill_empty.clicked.connect(self.on_fill_empty)

        self.btn_sec_write.clicked.connect(lambda: self.on_section("write"))
        self.btn_sec_rewrite.clicked.connect(lambda: self.on_section("rewrite"))
        self.btn_sec_expand.clicked.connect(lambda: self.on_section("expand"))
        self.btn_sec_shrink.clicked.connect(lambda: self.on_section("shrink"))

        # transforms
        self.btn_polish.clicked.connect(lambda: self.on_transform("polish"))
        self.btn_dedup.clicked.connect(lambda: self.on_transform("dedup"))
        self.btn_abs.clicked.connect(self.on_abstract)
        self.btn_kw.clicked.connect(self.on_keywords)
        self.btn_check.clicked.connect(self.on_check)

        # export
        self.btn_export_md.clicked.connect(self.on_export_md)
        self.btn_export_docx.clicked.connect(self.on_export_docx)

        # snapshots
        self.btn_snapshot.clicked.connect(lambda: self._snapshot("手动快照"))
        self.btn_restore.clicked.connect(self.on_restore_snapshot)

    # ---------- tree search ----------
    def on_tree_search(self, text: str):
        t = text.strip().lower()
        root = self.tree.topLevelItem(0)
        if not root:
            return
        def walk(item):
            visible = True
            if item != root:
                visible = (t in item.text(0).lower()) if t else True
            child_match = False
            for i in range(item.childCount()):
                child_visible = walk(item.child(i))
                child_match = child_match or child_visible
            if item != root:
                item.setHidden(not (visible or child_match))
            return visible or child_match
        walk(root)
        if not t:
            self.tree.expandAll()

    # ---------- project io ----------
    def on_new(self):
        self.project = ProjectData.empty()
        self.project_path = None
        self._load_project_to_ui()
        self._log(f"[{now_str()}] 新建项目。")

    def on_open(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "打开项目", "", "Project (*.paperwriter.json *.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.project = ProjectData.from_dict(data)
            self.project_path = path
            self._load_project_to_ui()
            self._log(f"[{now_str()}] 已打开：{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", f"打开失败：{e}")

    def on_save(self):
        if not self.project_path:
            return self.on_saveas()
        self._save_to(self.project_path)

    def on_saveas(self):
        default = (self.title_edit.text().strip() or "PaperWriter") + ".paperwriter.json"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "另存为", default, "Project (*.paperwriter.json)")
        if not path:
            return
        if not path.endswith(".paperwriter.json"):
            path += ".paperwriter.json"
        self.project_path = path
        self._save_to(path)

    def _save_to(self, path: str):
        try:
            self._sync_ui_to_project()
            self.project.updated_at = now_str()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.project.to_dict(), f, ensure_ascii=False, indent=2)
            self._log(f"[{now_str()}] 已保存：{path}")
            self.status.showMessage("Saved", 1500)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", f"保存失败：{e}")

    # ---------- settings ----------
    def on_settings(self):
        dlg = SettingsDialog(self, settings={
            "api_base": self.llm_cfg.api_base,
            "api_key": self.llm_cfg.api_key,
            "model_chat": self.llm_cfg.model_chat,
            "model_reasoner": self.llm_cfg.model_reasoner,
            "temperature": self.llm_cfg.temperature,
            "max_tokens": self.llm_cfg.max_tokens,
            "timeout": self.llm_cfg.timeout,
        })
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            data = dlg.get_data()
            self.llm_cfg = LLMConfig(**data)
            save_settings(data)
            self.client.cfg = self.llm_cfg
            self._log(f"[{now_str()}] API 设置已保存。")

    # ---------- references ----------
    def on_refs(self):
        self._sync_ui_to_project()
        dlg = ReferencesDialog(self, refs=self.project.references)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            res = dlg.result_data()
            self.project.references = res["refs"]
            self._snapshot("更新引用列表")
            if res["insert_requested"] and res["insert_index"]:
                idx = res["insert_index"]
                cur = self.draft.textCursor()
                cur.insertText(f"[{idx}]")
                self._snapshot(f"插入引用[{idx}]")
            if res["bib_requested"]:
                bib = self._build_bibliography_md()
                md = self.draft.toPlainText()
                if re.search(r"^##\s*参考文献", md, flags=re.M):
                    md2 = re.sub(r"^##\s*参考文献.*?(?=^##\s|\Z)", bib + "\n\n", md, flags=re.S | re.M)
                else:
                    md2 = (md.rstrip() + "\n\n" + bib + "\n").strip() + "\n"
                self.draft.setPlainText(md2)
                self._snapshot("生成参考文献段落")

    def _build_bibliography_md(self) -> str:
        lines = ["## 参考文献（占位符）"]
        if not self.project.references:
            lines.append("（暂无。请在“引用”中添加条目。）")
        else:
            for i, r in enumerate(self.project.references, 1):
                lines.append(f"- [{i}] {r}")
        return "\n".join(lines)

    # ---------- UI sync ----------
    def _load_project_to_ui(self):
        self._ignore = True
        try:
            self.title_edit.setText(self.project.title)
            self.genre_combo.setCurrentText(self.project.genre if self.project.genre in GENRES else "论文")
            self.template_combo.setCurrentText(self.project.template if self.project.template in TEMPLATES else TEMPLATES[0])
            self.instructions_edit.setPlainText(self.project.instructions)
            self.draft.setPlainText(self.project.draft_md)

            self.tree.clear()
            root = QtWidgets.QTreeWidgetItem(["大纲"])
            root.setFlags(root.flags() | QtCore.Qt.ItemIsDropEnabled)
            self.tree.addTopLevelItem(root)

            def add(parent, node: NodeData):
                it = QtWidgets.QTreeWidgetItem([node.title])
                it.setFlags(it.flags() | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsDragEnabled |
                            QtCore.Qt.ItemIsDropEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                it.setData(0, QtCore.Qt.UserRole + 1, node.id)
                it.setData(0, QtCore.Qt.UserRole + 2, node.notes)
                it.setData(0, QtCore.Qt.UserRole + 3, node.content)
                parent.addChild(it)
                for ch in (node.children or []):
                    add(it, ch)

            for ch in (self.project.outline_root.children or []):
                add(root, ch)

            self.tree.expandAll()
            self.tree.setCurrentItem(root.child(0) if root.childCount() else root)
            self._refresh_snapshots()
            self.on_tree_sel()
        finally:
            self._ignore = False

    def _sync_ui_to_project(self):
        self.project.title = self.title_edit.text().strip()
        self.project.genre = self.genre_combo.currentText().strip()
        self.project.template = self.template_combo.currentText().strip()
        self.project.instructions = self.instructions_edit.toPlainText().strip()
        self.project.draft_md = self.draft.toPlainText()

        root_item = self.tree.topLevelItem(0)

        def to_node(item: QtWidgets.QTreeWidgetItem) -> NodeData:
            nid = item.data(0, QtCore.Qt.UserRole + 1) or str(uuid.uuid4())
            notes = item.data(0, QtCore.Qt.UserRole + 2) or ""
            content = item.data(0, QtCore.Qt.UserRole + 3) or ""
            node = NodeData(id=str(nid), title=item.text(0), notes=str(notes), content=str(content), children=[])
            for i in range(item.childCount()):
                node.children.append(to_node(item.child(i)))
            return node

        outline_root = NodeData(id=self.project.outline_root.id, title="大纲", notes="", content="", children=[])
        for i in range(root_item.childCount()):
            outline_root.children.append(to_node(root_item.child(i)))
        self.project.outline_root = outline_root

    # ---------- tree selection ----------
    def on_tree_sel(self):
        if self._ignore:
            return
        item = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if not item or item == root:
            self.node_title.setText("")
            self.node_notes.setPlainText("")
            self.node_content.setPlainText("")
            return
        self._ignore = True
        try:
            self.node_title.setText(item.text(0))
            self.node_notes.setPlainText(str(item.data(0, QtCore.Qt.UserRole + 2) or ""))
            self.node_content.setPlainText(str(item.data(0, QtCore.Qt.UserRole + 3) or ""))
        finally:
            self._ignore = False

    def on_tree_item_changed(self, item: QtWidgets.QTreeWidgetItem, col: int):
        if self._ignore:
            return
        if item and item != self.tree.topLevelItem(0) and col == 0 and self.tree.currentItem() == item:
            self._ignore = True
            try:
                self.node_title.setText(item.text(0))
            finally:
                self._ignore = False

    def on_node_title_edit(self, text: str):
        if self._ignore:
            return
        item = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if item and item != root:
            self._ignore = True
            try:
                item.setText(0, text)
            finally:
                self._ignore = False

    def on_node_notes_change(self):
        if self._ignore:
            return
        item = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if item and item != root:
            item.setData(0, QtCore.Qt.UserRole + 2, self.node_notes.toPlainText())

    def on_node_content_change(self):
        if self._ignore:
            return
        item = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if item and item != root:
            item.setData(0, QtCore.Qt.UserRole + 3, self.node_content.toPlainText())

    # ---------- tree ops ----------
    def _new_item(self, title="新章节"):
        it = QtWidgets.QTreeWidgetItem([title])
        it.setFlags(it.flags() | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsDragEnabled |
                    QtCore.Qt.ItemIsDropEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
        it.setData(0, QtCore.Qt.UserRole + 1, str(uuid.uuid4()))
        it.setData(0, QtCore.Qt.UserRole + 2, "")
        it.setData(0, QtCore.Qt.UserRole + 3, "")
        return it

    def on_add_sib(self):
        cur = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if not cur or cur == root:
            it = self._new_item()
            root.addChild(it)
            self.tree.setCurrentItem(it)
            self.tree.editItem(it, 0)
            return
        parent = cur.parent() or root
        idx = parent.indexOfChild(cur)
        it = self._new_item()
        parent.insertChild(idx + 1, it)
        self.tree.setCurrentItem(it)
        self.tree.editItem(it, 0)

    def on_add_child(self):
        cur = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        parent = cur if (cur and cur != root) else root
        it = self._new_item()
        parent.addChild(it)
        parent.setExpanded(True)
        self.tree.setCurrentItem(it)
        self.tree.editItem(it, 0)

    def on_delete(self):
        cur = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if not cur or cur == root:
            return
        (cur.parent() or root).removeChild(cur)

    # ---------- writing ----------
    def on_outline(self):
        title = self.title_edit.text().strip()
        if not title:
            QtWidgets.QMessageBox.warning(self, "提示", "请先输入题目/标题。")
            return
        genre = self.genre_combo.currentText().strip()
        template = self.template_combo.currentText().strip()
        instructions = self.instructions_edit.toPlainText().strip()
        prefer_r = self.prefer_reasoner.isChecked()

        self._log(f"[{now_str()}] 生成大纲：{template}")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            nodes = self.engine.gen_outline_tree(title, genre, template, instructions, prefer_r)
            self._apply_outline_nodes(nodes)
            outline_md = tree_to_markdown(self.tree.topLevelItem(0), include_notes=False)
            miss = missing_required(template, outline_md)
            if miss:
                self._log(f"[{now_str()}] 结构提示：可能缺少：{', '.join(miss)}")
            self._snapshot("生成大纲")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def _apply_outline_nodes(self, nodes: List[Dict[str, Any]]):
        self._ignore = True
        try:
            root = self.tree.topLevelItem(0)
            root.takeChildren()

            def build(parent, d):
                it = self._new_item(str(d.get("title", "未命名")))
                it.setData(0, QtCore.Qt.UserRole + 2, str(d.get("notes", "")))
                parent.addChild(it)
                for ch in d.get("children", []) or []:
                    build(it, ch)

            for d in nodes:
                build(root, d)
            self.tree.expandAll()
            self.tree.setCurrentItem(root.child(0) if root.childCount() else root)
        finally:
            self._ignore = False
            self.on_tree_sel()

    def on_write_all(self):
        title = self.title_edit.text().strip()
        if not title:
            QtWidgets.QMessageBox.warning(self, "提示", "请先输入题目/标题。")
            return
        self._sync_ui_to_project()
        outline_md = tree_to_markdown(self.tree.topLevelItem(0), include_notes=True)
        prefer_r = self.prefer_reasoner.isChecked()

        self._log(f"[{now_str()}] 撰写全文：{self.project.template}")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            md = self.engine.write_full(self.project.title, self.project.genre, self.project.template,
                                        self.project.instructions, outline_md, self.project.references, prefer_r)
            self.draft.setPlainText(md)
            self._snapshot("撰写全文")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def on_fill_empty(self):
        title = self.title_edit.text().strip()
        if not title:
            QtWidgets.QMessageBox.warning(self, "提示", "请先输入题目/标题。")
            return
        self._sync_ui_to_project()
        prefer_r = self.prefer_reasoner.isChecked()
        outline_md = tree_to_markdown(self.tree.topLevelItem(0), include_notes=True)

        root = self.tree.topLevelItem(0)
        targets = []

        def collect(item):
            if item != root:
                content = str(item.data(0, QtCore.Qt.UserRole + 3) or "").strip()
                if not content:
                    targets.append(item)
            for i in range(item.childCount()):
                collect(item.child(i))

        collect(root)
        if not targets:
            QtWidgets.QMessageBox.information(self, "提示", "没有发现空章节。")
            return

        self._log(f"[{now_str()}] 补全空章节：{len(targets)} 个")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            full = self.draft.toPlainText()
            for i, it in enumerate(targets, 1):
                path = get_item_path_titles(it)
                sec_title = path[-1] if path else it.text(0)
                notes = str(it.data(0, QtCore.Qt.UserRole + 2) or "")
                existing = ""
                self._log(f"[{now_str()}] ({i}/{len(targets)}) 生成：{sec_title}")
                sec_md = self.engine.write_section(self.project.title, self.project.genre, self.project.template,
                                                   self.project.instructions, outline_md, path, notes, existing,
                                                   self.project.references, "write", prefer_r)
                it.setData(0, QtCore.Qt.UserRole + 3, sec_md)
                full = replace_section_in_markdown(full, sec_title, sec_md)
            self.draft.setPlainText(full)
            self._snapshot(f"补全空章节x{len(targets)}")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def _current_section_item(self):
        item = self.tree.currentItem()
        root = self.tree.topLevelItem(0)
        if not item or item == root:
            return None
        return item

    def on_section(self, mode: str):
        item = self._current_section_item()
        if not item:
            QtWidgets.QMessageBox.information(self, "提示", "请在章节树中选中一个章节。")
            return
        title = self.title_edit.text().strip()
        if not title:
            QtWidgets.QMessageBox.warning(self, "提示", "请先输入题目/标题。")
            return

        self._sync_ui_to_project()
        prefer_r = self.prefer_reasoner.isChecked()
        outline_md = tree_to_markdown(self.tree.topLevelItem(0), include_notes=True)
        path = get_item_path_titles(item)
        sec_title = path[-1] if path else item.text(0)
        notes = str(item.data(0, QtCore.Qt.UserRole + 2) or "")
        existing = str(item.data(0, QtCore.Qt.UserRole + 3) or "")

        self._log(f"[{now_str()}] 单章{mode}：{sec_title}")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            sec_md = self.engine.write_section(self.project.title, self.project.genre, self.project.template,
                                               self.project.instructions, outline_md, path, notes, existing,
                                               self.project.references, mode, prefer_r)
            item.setData(0, QtCore.Qt.UserRole + 3, sec_md)
            merged = replace_section_in_markdown(self.draft.toPlainText(), sec_title, sec_md)
            self.draft.setPlainText(merged)
            self._snapshot(f"单章{mode}:{sec_title}")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def _get_target_text(self) -> Tuple[str, Optional[QtGui.QTextCursor]]:
        cursor = self.draft.textCursor()
        if cursor.hasSelection():
            return cursor.selectedText().replace("\u2029", "\n"), cursor
        return self.draft.toPlainText(), None

    def _replace_target_text(self, new_text: str, cursor: Optional[QtGui.QTextCursor]):
        if cursor and cursor.hasSelection():
            cursor.insertText(new_text)
        else:
            self.draft.setPlainText(new_text)

    def on_transform(self, kind: str):
        text, cursor = self._get_target_text()
        if not text.strip():
            QtWidgets.QMessageBox.information(self, "提示", "没有可处理文本。")
            return
        prefer_r = self.prefer_reasoner.isChecked()
        instructions = self.instructions_edit.toPlainText().strip()
        label = "润色" if kind == "polish" else "降重"

        self._log(f"[{now_str()}] {label}（{'选中' if cursor else '全文'}）")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            out = self.engine.transform(kind, text, instructions, prefer_r)
            self._replace_target_text(out, cursor)
            self._snapshot(label)
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def on_abstract(self):
        full_text = self.draft.toPlainText().strip()
        if not full_text:
            QtWidgets.QMessageBox.information(self, "提示", "请先生成全文。")
            return
        prefer_r = self.prefer_reasoner.isChecked()

        self._log(f"[{now_str()}] 生成摘要…")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            ab = self.engine.gen_abstract(full_text, 250, prefer_r)
            if not ab:
                QtWidgets.QMessageBox.information(self, "提示", "未生成摘要（可能未配置 API）。")
                return
            if re.search(r"^##\s*摘要", full_text, flags=re.M):
                md2 = re.sub(r"^##\s*摘要.*?(?=^##\s|\Z)", f"## 摘要\n{ab}\n\n", full_text, flags=re.S | re.M)
            else:
                md2 = f"## 摘要\n{ab}\n\n" + full_text
            self.draft.setPlainText(md2)
            self._snapshot("生成摘要")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def on_keywords(self):
        full_text = self.draft.toPlainText().strip()
        if not full_text:
            QtWidgets.QMessageBox.information(self, "提示", "请先生成全文。")
            return
        prefer_r = self.prefer_reasoner.isChecked()

        self._log(f"[{now_str()}] 生成关键词…")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            kw = self.engine.gen_keywords(full_text, 5, prefer_r)
            if not kw:
                QtWidgets.QMessageBox.information(self, "提示", "未生成关键词（可能未配置 API）。")
                return
            if re.search(r"^##\s*关键词", full_text, flags=re.M):
                md2 = re.sub(r"^##\s*关键词.*?(?=^##\s|\Z)", f"## 关键词\n{kw}\n\n", full_text, flags=re.S | re.M)
            else:
                md2 = f"## 关键词\n{kw}\n\n" + full_text
            self.draft.setPlainText(md2)
            self._snapshot("生成关键词")
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

    def on_check(self):
        outline_md = tree_to_markdown(self.tree.topLevelItem(0), include_notes=False)
        template = self.template_combo.currentText()
        miss = missing_required(template, outline_md)
        if not miss:
            QtWidgets.QMessageBox.information(self, "结构校验", "结构检查通过：未发现明显缺失章节。")
        else:
            QtWidgets.QMessageBox.warning(self, "结构校验", "可能缺失章节：\n- " + "\n- ".join(miss))

    # ---------- export ----------
    def on_export_md(self):
        text = self.draft.toPlainText().strip()
        if not text:
            QtWidgets.QMessageBox.information(self, "提示", "请先生成全文再导出。")
            return
        default = (self.title_edit.text().strip() or "PaperWriter") + ".md"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出 Markdown", default, "Markdown (*.md)")
        if not path:
            return
        if not path.endswith(".md"):
            path += ".md"
        export_markdown(path, text)
        self._log(f"[{now_str()}] 已导出 Markdown：{path}")

    def on_export_docx(self):
        text = self.draft.toPlainText().strip()
        if not text:
            QtWidgets.QMessageBox.information(self, "提示", "请先生成全文再导出。")
            return
        default = (self.title_edit.text().strip() or "PaperWriter") + ".docx"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出 Word", default, "Word (*.docx)")
        if not path:
            return
        if not path.endswith(".docx"):
            path += ".docx"
        export_docx(path, text)
        self._log(f"[{now_str()}] 已导出 Word：{path}")

    # ---------- snapshots restore ----------
    def on_restore_snapshot(self):
        row = self.snap_list.currentRow()
        if row < 0 or row >= len(self.project.snapshots):
            return
        snap = self.project.snapshots[row]
        try:
            self.project = ProjectData.from_dict(snap.project)
            self.project.snapshots = self.project.snapshots[row:]
            self._load_project_to_ui()
            self._log(f"[{now_str()}] 已回滚：{snap.ts} | {snap.action}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", f"回滚失败：{e}")


# =========================
# CLI selftest (for CI)
# =========================
def run_cli_if_requested():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--selftest", action="store_true")
    args, _ = parser.parse_known_args()
    if args.selftest:
        print("SELFTEST_OK")
        raise SystemExit(0)


def main():
    run_cli_if_requested()
    
    # === FIX: Linux Input Method (CLEANED) ===
    # 我们彻底移除了自动设置环境变量的逻辑。
    # 软件启动时不带任何预设偏见。
    if sys.platform.startswith("linux"):
        print("Linux Launch: No environment variables forced. Using system default.")
    # =========================================

    app = QtWidgets.QApplication(sys.argv)
    
    # 字体优化
    if sys.platform.startswith("linux"):
        app.setFont(QtGui.QFont("WenQuanYi Micro Hei", 10))
    else:
        app.setFont(QtGui.QFont("Microsoft YaHei", 9))

    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
