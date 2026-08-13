# -*- coding: utf-8 -*-
"""
规范性引用文件章节审核器
=========================
依据《a4guifan_rules_summary.txt》的 A4R001~A4R015 规则，对标准文档中的
“规范性引用文件”章节以及正文引用关系进行自动化审核。

设计上与《术语章节审核器》保持同构：
    - 复用其 AuditIssue 数据模型（便于统一回写批注 / 导出 xlsx）。
    - 提供 ReferenceParser（解析章节结构）与 ReferenceAuditor（执行规则）。

联动说明：
    A4R012（标准仅在“来源”中提及却列入规范性引用文件）依赖“术语和定义”
    章节提取出的来源标准集合，由主控文件在调用本模块时通过
    terms_source_standards 参数传入。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Set, Tuple

# ---------------------------------------------------------------------------
# 复用术语章节审核器的数据模型与章节边界判定
# ---------------------------------------------------------------------------
try:  # 包内导入（推荐用法：from review_engine import references）
    from .terminology import AuditIssue, _is_top_level_heading  # noqa: F401
except ImportError:  # 兜底：作为单文件脚本直接运行时
    from terminology import AuditIssue, _is_top_level_heading  # noqa: F401


# ---------------------------------------------------------------------------
# 规则常量
# ---------------------------------------------------------------------------

REFERENCE_CHAPTER_TITLES = ["规范性引用文件"]

# 标准引导语（有引用文件 / 无引用文件）
# 注：采用前缀匹配（不强制结尾 $），以兼容 GB/T 1.1 长式引导语
# （“下列文件中的内容通过文中的规范性引用而构成本文件必不可少的条款。其中，……”）。
VALID_REF_GUIDE_PATTERNS = [
    r"^下列文件对于本文件的应用是必不可少的。",
    r"^下列文件中的内容通过文中的规范性引用而构成本文件必不可少的条款。",
    r"^下列文件中的条款通过本文件的引用而成为本文件的条款。",
    r"^本文件没有规范性引用文件。",
]
NO_REF_GUIDE = r"^本文件没有规范性引用文件。$"

# 序号前缀 / 书名号 / 引号等非法符号
REF_FORBIDDEN_PREFIX_RE = re.compile(r"^\s*[\d]+[\.\、]\s*")          # 人工序号前缀
REF_FORBIDDEN_MARK_RE = re.compile(r"[《》“”\"]")                     # 书名号 / 引号
# 占位符编号（XXXX / XXXX-XXXX / XXXX:XXXX）
REF_PLACEHOLDER_RE = re.compile(r"\bX{2,4}(?:[-:]\d{2,4})?\b")
# 年份号使用“X-X”等非标准格式（四位年份范围，如 2020-2021）
REF_YEAR_RANGE_RE = re.compile(r"\b\d{4}\s*[-/]\s*\d{4}\b")

# 国际组织代号（国际标准应以这些代号开头）
INTL_PREFIXES = ("ISO", "IEC", "ITU", "ISO/IEC", "3GPP", "IEEE")

# 标准编号正则（尽量覆盖国标/行标/地标/团标/国际）
# 机构代号之后允许两种形式：
#   - 数字型：GB/T 12345—2020、YD/T 1990-2023
#   - 字母.数字型（国际电联等）：ITU-T M.3010、ITU-T H.621、ISO/IEC 21827:2022
STD_CODE_RE = re.compile(
    r"(?:GB|GB/T|GBZ|GBZ/T|DL/T|YD/T|DB\d{2}/T|DB\d{2}|Q/GDW|T/[A-Za-z]+|"
    r"IEEE|ISO|IEC|ISO/IEC|ITU-T|ITU-R|ITU|3GPP)\s*"
    r"(?:[A-Z]\.\d[\d.]*|\d[\w.]*(?:[.—-]\d{2,6})?(?::\d{4})?)"
)

# 正文规范性引用措辞（用于判定标准是否被“规范性”引用）
NORMATIVE_CUES = ["应符合", "应遵循", "应满足", "按照", "依据", "应采用", "应执行", "应参照"]


# ---------------------------------------------------------------------------
# 标准编号工具
# ---------------------------------------------------------------------------

def normalize_std_code(raw: str) -> str:
    """将标准编号归一化为“基础代号”（大写、去空格、去掉年代号），用于跨章节比对。"""
    s = raw.upper().replace(" ", "").replace(" ", "")
    s = s.replace("—", "-")  # 中文破折号 -> 连字符，便于比较
    # 去掉末尾的年代号（1999:2020 / -2020）
    s = re.sub(r"(?::|-)\d{4}$", "", s)
    return s


def extract_std_codes(text: str) -> List[str]:
    """从文本中提取标准基础代号列表（已归一化）。"""
    codes = []
    for m in STD_CODE_RE.finditer(text or ""):
        code = normalize_std_code(m.group(0))
        if code and code not in codes:
            codes.append(code)
    return codes


def classify_std(code: str) -> int:
    """返回标准层级序号，用于排序检查 A4R004。

    顺序：国家标准(0) < 行业标准(1) < 地方标准(2) < 团体标准(3)
          < 国际(ISO/IEC/ITU 等)(4) < 其他机构/文献(5)
    """
    if code.startswith("GB"):
        return 0
    if code.startswith("DB"):
        return 2
    if code.startswith("T/"):
        return 3
    if code.startswith(INTL_PREFIXES):
        return 4
    if code.startswith(("IEEE", "3GPP")):
        return 5
    return 1  # 行业标准（YD/T、DL/T 等）


def is_international(code: str) -> bool:
    return code.startswith(INTL_PREFIXES)


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------

@dataclass
class ReferenceEntry:
    """解析出的规范性引用文件条目。"""
    raw_text: str = ""            # 原始文本（整段）
    code: str = ""                # 归一化基础代号
    has_number: bool = True       # 是否带标准发布编号
    category: int = 1             # 层级（classify_std）
    paragraph_index: int = -1     # 章节段落索引
    year: str = ""                # 年代号（原始）


@dataclass
class ReferenceChapter:
    """规范性引用文件章节结构。"""
    title: str = ""
    guide_sentence: str = ""
    guide_paragraph_index: int = -1
    entries: List[ReferenceEntry] = field(default_factory=list)
    body_paragraphs: List[str] = field(default_factory=list)
    source_paragraphs: List[str] = field(default_factory=list)
    is_table_format: bool = False


# ---------------------------------------------------------------------------
# 解析器
# ---------------------------------------------------------------------------

class ReferenceParser:
    """将“规范性引用文件”章节段落解析为 ReferenceChapter。"""

    def __init__(self, paragraphs: List[str]):
        self.paragraphs = [p.strip() for p in paragraphs if p.strip()]
        self.chapter_title: str = ""

    def parse(self, chapter_title: str = "") -> ReferenceChapter:
        chapter = ReferenceChapter()
        chapter.title = chapter_title or self._detect_title()
        self.chapter_title = chapter.title
        chapter.source_paragraphs = self.paragraphs
        if not self.paragraphs:
            return chapter

        start_idx = 1 if self.paragraphs[0] == chapter.title else 0

        # 引导语：前 3 段内查找合规引导语，否则回退首段
        guide_idx, guide_text = self._find_guide(start_idx)
        chapter.guide_sentence = guide_text
        chapter.guide_paragraph_index = guide_idx

        # 表格形式检测（含 $$ 或大量制表符）
        chapter.is_table_format = any("$$" in p or p.count("\t") >= 3 for p in self.paragraphs)

        # 解析条目
        chapter.entries = self._parse_entries(guide_idx)

        chapter.body_paragraphs = [
            p for i, p in enumerate(self.paragraphs)
            if i > guide_idx and p != chapter.title
        ]
        return chapter

    def _detect_title(self) -> str:
        if not self.paragraphs:
            return ""
        first = self.paragraphs[0]
        for t in REFERENCE_CHAPTER_TITLES:
            if t in first:
                return t
        return first

    def _find_guide(self, start_idx: int) -> Tuple[int, str]:
        for i in range(start_idx, min(start_idx + 3, len(self.paragraphs))):
            if any(re.match(p, self.paragraphs[i]) for p in VALID_REF_GUIDE_PATTERNS):
                return i, self.paragraphs[i]
        if start_idx < len(self.paragraphs):
            return start_idx, self.paragraphs[start_idx]
        return -1, ""

    def _parse_entries(self, guide_idx: int) -> List[ReferenceEntry]:
        entries: List[ReferenceEntry] = []
        if guide_idx < 0:
            return entries
        for i in range(guide_idx + 1, len(self.paragraphs)):
            text = self.paragraphs[i]
            if text == self.chapter_title:
                continue
            if not self._looks_like_entry(text):
                continue
            codes = extract_std_codes(text)
            entry = ReferenceEntry(
                raw_text=text,
                code=codes[0] if codes else "",
                has_number=bool(codes),
                category=classify_std(codes[0]) if codes else 6,
                paragraph_index=i,
            )
            # 提取年代号
            ym = re.search(r"(?::|-)(\d{4})\b", text)
            if ym:
                entry.year = ym.group(1)
            entries.append(entry)
        return entries

    def _looks_like_entry(self, text: str) -> bool:
        """判断段落是否像一条规范性引用文件条目。"""
        if not text:
            return False
        # 含标准代号
        if STD_CODE_RE.search(text):
            return True
        # 或以标准代号常见前缀开头（即便未被正则完全命中）
        if re.match(r"^(GB|YD|DL|DB|ISO|IEC|ITU|IEEE|T/|Q/)", text):
            return True
        return False


# ---------------------------------------------------------------------------
# 审核器
# ---------------------------------------------------------------------------

class ReferenceAuditor:
    """依据 A4R001~A4R015 对规范性引用文件章节进行审核。"""

    def __init__(
        self,
        full_text: str = "",
        body_paragraphs: Optional[List[str]] = None,
        chapter: Optional[ReferenceChapter] = None,
        terms_source_standards: Optional[Set[str]] = None,
    ):
        self.full_text = full_text
        self.body_paragraphs = body_paragraphs or []
        self.chapter = chapter or ReferenceChapter()
        self.terms_source_standards = terms_source_standards or set()
        self.issues: List[AuditIssue] = []

    # 对外接口 -----------------------------------------------------------
    def audit(self) -> List[AuditIssue]:
        self.issues = []
        self._check_chapter_structure()      # A4R001
        self._check_table_format()           # A4R002
        self._check_guide_sentence()         # A4R003
        self._check_sorting()                # A4R004 / A4R015
        self._check_missing_number()         # A4R005
        self._check_intl_domestic_format()   # A4R006
        self._check_missing_org_code()       # A4R007
        self._check_year_format()            # A4R008
        self._check_forbidden_marks()        # A4R009
        self._check_body_citation()          # A4R010 / A4R011 / A4R013 / A4R014
        return self.issues

    def _add(self, rule, location, description, suggestion,
             severity="error", details=None, paragraph_index=-1):
        self.issues.append(AuditIssue(
            rule=rule, location=location, description=description,
            suggestion=suggestion, severity=severity,
            details=details or {}, paragraph_index=paragraph_index,
        ))

    # 规则实现 -----------------------------------------------------------

    def _check_chapter_structure(self):  # A4R001
        if not self.chapter.title or self.chapter.title not in REFERENCE_CHAPTER_TITLES:
            self._add(
                rule="规范性引用文件-章节",
                location="文档",
                description="规范性引用文件章节结构错误（缺失或名称不规范）",
                suggestion="检查章节是否缺失，格式是否错误；无引用文件时保留第2章并只写引导语：本文件没有规范性引用文件",
                paragraph_index=-2,
            )
            return
        # 无任何条目且无“无引用文件”引导语
        if not self.chapter.entries:
            if not re.match(NO_REF_GUIDE, self.chapter.guide_sentence or ""):
                self._add(
                    rule="规范性引用文件-章节",
                    location=f"章节：{self.chapter.title}",
                    description="规范性引用文件章节结构错误（无引用条目且未给出无引用文件引导语）",
                    suggestion="补充引用文件条目，或无引用文件时保留第2章并注明“本文件没有规范性引用文件”",
                    paragraph_index=-2,
                )

    def _check_table_format(self):  # A4R002
        if self.chapter.is_table_format:
            self._add(
                rule="规范性引用文件-格式",
                location=f"章节：{self.chapter.title}",
                description="规范性引用文件章节格式错误，不应使用表格形式呈现",
                suggestion="改为普通段落列表形式，不要使用表格排版",
                paragraph_index=-2,
            )

    def _check_guide_sentence(self):  # A4R003
        guide = self.chapter.guide_sentence
        has_entries = bool(self.chapter.entries)
        if has_entries:
            if not any(re.match(p, guide or "") for p in VALID_REF_GUIDE_PATTERNS):
                self._add(
                    rule="规范性引用文件-引导语",
                    location=f"引导语：{guide}",
                    description="存在引用文件时缺少标准引导语或引导语使用不当",
                    suggestion="使用国家标准规定引导语，如“下列文件对于本文件的应用是必不可少的。”",
                    paragraph_index=self.chapter.guide_paragraph_index,
                )
            elif re.match(NO_REF_GUIDE, guide or ""):
                self._add(
                    rule="规范性引用文件-引导语",
                    location=f"引导语：{guide}",
                    description="有引用文件却出现“本文件没有规范性引用文件”引导语",
                    suggestion="有引用文件时请使用“下列文件对于本文件的应用是必不可少的。”等标准引导语",
                    paragraph_index=self.chapter.guide_paragraph_index,
                )
        else:
            if guide and not re.match(NO_REF_GUIDE, guide):
                self._add(
                    rule="规范性引用文件-引导语",
                    location=f"引导语：{guide}",
                    description="无规范性引用文件时引导语使用不当",
                    suggestion="无引用文件时保留第2章并注明“本文件没有规范性引用文件”",
                    paragraph_index=self.chapter.guide_paragraph_index,
                )

    def _check_sorting(self):  # A4R004 / A4R015
        entries = self.chapter.entries
        last_cat = -1
        last_code = ""
        for e in entries:
            if not e.code:  # A4R015：编号无法解析
                if REF_PLACEHOLDER_RE.search(e.raw_text):
                    self._add(
                        rule="规范性引用文件-编号",
                        location=f"引用条目：{e.raw_text}",
                        description="引用文件编号无法解析或为占位符，请核对编号格式",
                        suggestion="替换为真实标准编号",
                        paragraph_index=e.paragraph_index,
                    )
                continue
            # 同类型内未按编号升序（best-effort）
            if e.category == last_cat and last_code:
                if e.code.upper() < last_code.upper():
                    self._add(
                        rule="规范性引用文件-排序",
                        location=f"引用条目：{e.raw_text}",
                        description="规范性引用文件排序错误（同类型内未按标准号从小到大排列）",
                        suggestion="按规定的层级和编号顺序重新排列",
                        paragraph_index=e.paragraph_index,
                    )
            # 层级不应回退
            if last_cat >= 0 and e.category < last_cat:
                self._add(
                    rule="规范性引用文件-排序",
                    location=f"引用条目：{e.raw_text}",
                    description="规范性引用文件排序错误（层级顺序不符合规定）",
                    suggestion="按 国家标准→行业标准→地方标准→团体标准→ISO/IEC/ITU→其他机构→其他文献 的顺序排列",
                    paragraph_index=e.paragraph_index,
                )
            last_cat = e.category
            last_code = e.code

    def _check_missing_number(self):  # A4R005
        for e in self.chapter.entries:
            # 占位符条目（如 YD/T XXXX-XXXX）已由 A4R015 处理，避免重复报告
            if REF_PLACEHOLDER_RE.search(e.raw_text):
                continue
            if not e.has_number:
                self._add(
                    rule="规范性引用文件-编号",
                    location=f"引用条目：{e.raw_text}",
                    description="引用的文件应使用标准发布编号及标准发布名称",
                    suggestion="为每个引用文件补充标准发布编号",
                    paragraph_index=e.paragraph_index,
                )

    def _check_intl_domestic_format(self):  # A4R006
        for e in self.chapter.entries:
            if not e.code:
                continue
            if is_international(e.code):
                # 国际标准应以“标准发布号 中文名称（英文名称）”结尾
                if "（" not in e.raw_text and "(" not in e.raw_text:
                    self._add(
                        rule="规范性引用文件-格式",
                        location=f"引用条目：{e.raw_text}",
                        description="国际引用文件应为“标准发布号 中文（英文）”结尾格式",
                        suggestion="补充中文名称及英文名称，如“标准发布号 中文名称（英文名称）”",
                        paragraph_index=e.paragraph_index,
                    )
            else:
                # 国内标准不应以“（英文）”结尾
                if re.search(r"（英文）\s*$", e.raw_text) or re.search(r"\([A-Za-z]+\)$", e.raw_text):
                    self._add(
                        rule="规范性引用文件-格式",
                        location=f"引用条目：{e.raw_text}",
                        description="国内引用文件不应以（英文）结尾",
                        suggestion="国内引用文件按对应格式书写，去掉末尾（英文）",
                        paragraph_index=e.paragraph_index,
                    )

    def _check_missing_org_code(self):  # A4R007
        for e in self.chapter.entries:
            # 编号以 H. 开头但缺少应有的国际组织代号
            if re.search(r"(?<![\w/])\d*\.\d+", e.code) and e.code.startswith("H."):
                if not any(e.code.startswith(p) for p in ("ITU", "ISO/IEC", "IEC")):
                    self._add(
                        rule="规范性引用文件-编号",
                        location=f"引用条目：{e.raw_text}",
                        description="国际引用文件缺少国际组织代号",
                        suggestion="补充如 ITU-T H.XXX、ISO/IEC H.XXX 等代号",
                        paragraph_index=e.paragraph_index,
                    )

    def _check_year_format(self):  # A4R008
        for e in self.chapter.entries:
            if REF_YEAR_RANGE_RE.search(e.raw_text):
                self._add(
                    rule="规范性引用文件-年份",
                    location=f"引用条目：{e.raw_text}",
                    description="引用文件年份号格式错误（使用了 X-X 等非标准格式）",
                    suggestion="使用四位数年份号，如 GB/T 12345—2020",
                    paragraph_index=e.paragraph_index,
                )

    def _check_forbidden_marks(self):  # A4R009
        for e in self.chapter.entries:
            if REF_FORBIDDEN_PREFIX_RE.match(e.raw_text) or REF_FORBIDDEN_MARK_RE.search(e.raw_text):
                self._add(
                    rule="规范性引用文件-符号",
                    location=f"引用条目：{e.raw_text}",
                    description="引用文件不应有序号前缀，也不应使用书名号、引号等符号",
                    suggestion="删除人工序号、书名号、引号，仅保留标准编号和名称",
                    paragraph_index=e.paragraph_index,
                )

    def _check_body_citation(self):  # A4R010 / A4R011 / A4R013 / A4R014
        # 提取正文引用情况（排除引用文件章节与术语章节自身）
        normative_cited: Set[str] = set()
        note_only: Set[str] = set()
        for para in self.body_paragraphs:
            codes = extract_std_codes(para)
            if not codes:
                continue
            is_note = para.strip().startswith("注")
            has_cue = any(cue in para for cue in NORMATIVE_CUES)
            for c in codes:
                if is_note:
                    note_only.add(c)
                if has_cue:
                    normative_cited.add(c)

        ref_codes = {e.code for e in self.chapter.entries if e.code}

        # A4R010：列在规范性引用文件中但正文未正确引用
        for e in self.chapter.entries:
            if e.code and e.code not in normative_cited:
                self._add(
                    rule="规范性引用文件-正文引用",
                    location=f"引用条目：{e.raw_text}",
                    description="该标准在正文中未被正确引用，请检查正文是否需要引用该标准，或从规范性引用文件中移除",
                    suggestion="在正文中引用该标准，或将其移除",
                    severity="warning",
                    paragraph_index=e.paragraph_index,
                )

        # A4R011：仅在“注”中提及
        for e in self.chapter.entries:
            if e.code and e.code in note_only and e.code not in normative_cited:
                self._add(
                    rule="规范性引用文件-正文引用",
                    location=f"引用条目：{e.raw_text}",
                    description="该标准仅在“注”中提及，不应放在规范性引用文件中，应放在参考文献章节中",
                    suggestion="将该标准从规范性引用文件移至参考文献",
                    paragraph_index=e.paragraph_index,
                )

        # A4R013：正文引用编号与条目不一致（同基础代号但年代号不同）
        ref_base = {re.sub(r"(?::|-)\d{4}$", "", c) for c in ref_codes}
        for c in normative_cited:
            base = re.sub(r"(?::|-)\d{4}$", "", c)
            if base in ref_base and c not in ref_codes:
                self._add(
                    rule="规范性引用文件-一致性",
                    location=f"正文引用标准：{c}",
                    description="正文中引用的标准编号与规范性引用文件中的条目不一致（如多写/少写年代号）",
                    suggestion="核对并保持一致",
                    paragraph_index=-1,
                )

        # A4R014：正文有规范性引用但未在第二章列出
        for c in normative_cited:
            if c not in ref_codes and re.sub(r"(?::|-)\d{4}$", "", c) not in ref_base:
                self._add(
                    rule="规范性引用文件-一致性",
                    location=f"正文引用标准：{c}",
                    description="正文有规范性引用文件，需要在第二章给出规范性引用",
                    suggestion="在第二章补充该标准条目",
                    paragraph_index=-1,
                )

# ---------------------------------------------------------------------------
# 便捷入口
# ---------------------------------------------------------------------------

def audit_references(
    paragraphs: List[str],
    chapter_title: str = "",
    full_text: str = "",
    body_paragraphs: Optional[List[str]] = None,
    terms_source_standards: Optional[Set[str]] = None,
) -> List[AuditIssue]:
    """对“规范性引用文件”章节段落进行审核。"""
    parser = ReferenceParser(paragraphs)
    chapter = parser.parse(chapter_title)
    auditor = ReferenceAuditor(
        full_text=full_text,
        body_paragraphs=body_paragraphs,
        chapter=chapter,
        terms_source_standards=terms_source_standards,
    )
    return auditor.audit()


# ---------------------------------------------------------------------------
# 术语“来源”标准提取（供主控文件做联动）
# ---------------------------------------------------------------------------

SOURCE_RE = re.compile(r"[\[【]来源[:：]\s*(.*?)[】\]]", re.S)


def extract_term_source_standards(term_definitions: List[str]) -> Set[str]:
    """从术语定义文本（含 [来源：…] / 【来源：…】）中提取标准基础代号集合。"""
    codes: Set[str] = set()
    for d in term_definitions:
        for m in SOURCE_RE.finditer(d or ""):
            codes.update(extract_std_codes(m.group(1)))
    return codes
