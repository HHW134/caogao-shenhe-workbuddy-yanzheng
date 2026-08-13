# -*- coding: utf-8 -*-
"""
标准文档审核主控
================
按章节顺序调用各专项审核模块，并把“术语和定义”与“规范性引用文件”
两个章节的结果聚合，输出带批注的 docx 与问题汇总 xlsx。

调用关系（与 form_validation_term_ref_rules_summary.txt 一致）：
    1. 先处理“术语和定义”章节，得到术语条目及其“来源”标准集合；
    2. 再处理“规范性引用文件”章节，将术语“来源”标准集合传入，
       用于 A4R012 / FVT006 联动检查（仅出现在术语来源中的标准
       不应列入规范性引用文件）；
    3. 在主函数中显式实现两章节联动（术语“来源”标准若同时出现在
       规范性引用文件列表中，则分别在术语章节与引用文件章节给出批注）。

用法：
    python 审核主控.py <input.docx> [-o JSON] [--docx-out ...] [--xlsx-out ...] [--no-comments]
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import List, Optional, Set, Tuple

# ---------------------------------------------------------------------------
# 模块加载
# ---------------------------------------------------------------------------

try:  # 包内导入（推荐用法：from review_engine import pipeline）
    from . import terminology as terms_mod
    from . import references as ref_mod
except ImportError:  # 兜底：作为单文件脚本直接运行时
    import terminology as terms_mod  # type: ignore
    import references as ref_mod  # type: ignore

AuditIssue = terms_mod.AuditIssue
_is_top_level_heading = terms_mod._is_top_level_heading
write_annotated_docx = terms_mod.write_annotated_docx
write_issues_xlsx = terms_mod.write_issues_xlsx
TerminologyParser = terms_mod.TerminologyParser
TerminologyAuditor = terms_mod.TerminologyAuditor
TERMINOLOGY_CHAPTER_TITLES = terms_mod.TERMINOLOGY_CHAPTER_TITLES

ReferenceParser = ref_mod.ReferenceParser
ReferenceAuditor = ref_mod.ReferenceAuditor
REFERENCE_CHAPTER_TITLES = ref_mod.REFERENCE_CHAPTER_TITLES
extract_term_source_standards = ref_mod.extract_term_source_standards
extract_std_codes = ref_mod.extract_std_codes


# ---------------------------------------------------------------------------
# 章节定位
# ---------------------------------------------------------------------------

def _load_items(docx_path: str):
    """返回 [(Paragraph对象, 文本), ...]（仅非空段落）。"""
    from docx import Document
    doc = Document(docx_path)
    items = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            items.append((p, t))
    return doc, items


def _locate_chapter(items, titles, allowed_in_chapter):
    """定位某章节，返回 (start_index, title, span[(p,t)...])。

    span 从章节标题之后取到下一个顶层章节标题之前。
    """
    for i, (p, t) in enumerate(items):
        if t in titles:
            span = []
            for j in range(i + 1, len(items)):
                pj, tj = items[j]
                if _is_top_level_heading(pj.style.name, tj, allowed_in_chapter):
                    break
                span.append((pj, tj))
            return i, t, span
    return -1, "", []


# ---------------------------------------------------------------------------
# 联动：术语“来源”标准 vs 规范性引用文件（主函数实现）
# ---------------------------------------------------------------------------

def linkage_term_source_vs_references(terms_chapter, refs_chapter,
                                       para_objects_terms) -> List[AuditIssue]:
    """FVT006 / A4R012：术语“来源”中的标准若同时列入规范性引用文件，给出批注。

    该联动逻辑放在主函数中：批注锚定到术语条目所在段落（术语章节侧）。
    """
    issues: List[AuditIssue] = []
    if not terms_chapter or not refs_chapter:
        return issues
    ref_codes: Set[str] = {e.code for e in refs_chapter.entries if e.code}
    if not ref_codes:
        return issues

    for entry in terms_chapter.entries:
        for code in extract_std_codes(entry.definition):
            if code in ref_codes:
                issues.append(AuditIssue(
                    rule="术语和定义-来源联动",
                    location=f"术语'{entry.name_cn or entry.raw_text}'（来源标准：{code}）",
                    description="该标准出现在术语和定义章节中，不应放在规范性引用文件中，应放在参考文献章节中",
                    suggestion="将该标准从第二章“规范性引用文件”中移除，放入“参考文献”章节",
                    paragraph_index=entry.paragraph_index,
                ))
                break
    return issues


# ---------------------------------------------------------------------------
# 结构性联动（FVT003 / FVT007 等）
# ---------------------------------------------------------------------------

def structural_linkage_checks(terms_span, refs_span,
                              para_objects_terms, para_objects_refs,
                              terms_idx, refs_idx) -> List[AuditIssue]:
    """FVT003：术语类章节包含二级目次；FVT007：章节顺序提示。"""
    issues: List[AuditIssue] = []

    # FVT003：术语类章节下方出现二级编号（如 3.1、3.2）
    for k, (_, t) in enumerate(terms_span):
        if re.match(r"^\d+\.\d+", t) and len(t) <= 30:
            issues.append(AuditIssue(
                rule="术语和定义-章节结构",
                location=f"术语章节二级目次：{t}",
                description="术语和定义类章节不应包含二级目次",
                suggestion="删除术语章节下的二级目次，术语条目直接以条目形式呈现",
                paragraph_index=k,
            ))

    # FVT007：标准结构中第2章（规范性引用文件）应在第3章（术语和定义）之前
    if terms_idx >= 0 and refs_idx >= 0 and refs_idx > terms_idx:
        issues.append(AuditIssue(
            rule="章节-结构顺序",
            location="章节顺序",
            description="结构提示——术语和定义章节位置需符合标准结构要求（通常在规范性引用文件之后）",
            suggestion="标准结构中第2章为规范性引用文件，第3章通常为术语和定义，二者顺序固定",
            severity="info",
            paragraph_index=-2,
        ))
    return issues


# ---------------------------------------------------------------------------
# 主控流程
# ---------------------------------------------------------------------------

def audit_document(docx_path: str):
    doc, items = _load_items(docx_path)
    full_text = "\n".join(t for _, t in items)

    all_issues: List[AuditIssue] = []
    meta = {"术语和定义": {}, "规范性引用文件": {}}

    # ---- 1. 术语和定义章节 ----
    terms_idx, terms_title, terms_span = _locate_chapter(
        items, TERMINOLOGY_CHAPTER_TITLES,
        allowed_in_chapter={*TERMINOLOGY_CHAPTER_TITLES, "缩略语"},
    )
    para_objects_terms = [p for p, _ in terms_span]
    title_para_terms = items[terms_idx][0] if terms_idx >= 0 else None

    terms_chapter = None
    if terms_idx < 0:
        all_issues.append(AuditIssue(
            rule="章节-术语和定义",
            location="文档",
            description="术语和定义章节结构错误（缺失）",
            suggestion="保留第3章并给出引导语：本文件没有术语和定义",
        ))
    else:
        ch_paragraphs = [t for _, t in terms_span]
        parser = TerminologyParser(ch_paragraphs)
        terms_chapter = parser.parse(terms_title)
        auditor = TerminologyAuditor(full_text=full_text, chapter=terms_chapter)
        issues_t = auditor.audit()
        all_issues.extend(issues_t)
        meta["术语和定义"] = {
            "章节标题": terms_title,
            "引导语": terms_chapter.guide_sentence,
            "术语条目数": len(terms_chapter.entries),
            "子章节": terms_chapter.sub_sections,
        }

    # ---- 2. 规范性引用文件章节 ----
    refs_idx, refs_title, refs_span = _locate_chapter(
        items, REFERENCE_CHAPTER_TITLES,
        allowed_in_chapter={*REFERENCE_CHAPTER_TITLES},
    )
    para_objects_refs = [p for p, _ in refs_span]
    title_para_refs = items[refs_idx][0] if refs_idx >= 0 else None

    refs_chapter = None
    if refs_idx < 0:
        all_issues.append(AuditIssue(
            rule="规范性引用文件-章节",
            location="文档",
            description="规范性引用文件章节结构错误（缺失）",
            suggestion="保留第2章并给出引导语：本文件没有规范性引用文件",
        ))
    else:
        # 正文段落 = 排除两个章节 span 后的其余段落
        occupied = set(range(terms_idx, terms_idx + 1 + len(terms_span))) if terms_idx >= 0 else set()
        occupied |= set(range(refs_idx, refs_idx + 1 + len(refs_span))) if refs_idx >= 0 else set()
        body_paragraphs = [t for idx, (_, t) in enumerate(items) if idx not in occupied]

        ch_paragraphs = [t for _, t in refs_span]
        parser = ReferenceParser(ch_paragraphs)
        refs_chapter = parser.parse(refs_title)

        # 术语“来源”标准集合（联动输入）
        term_source = extract_term_source_standards(
            [e.definition for e in (terms_chapter.entries if terms_chapter else [])]
        )
        auditor = ReferenceAuditor(
            full_text=full_text,
            body_paragraphs=body_paragraphs,
            chapter=refs_chapter,
            terms_source_standards=term_source,
        )
        issues_r = auditor.audit()
        all_issues.extend(issues_r)
        meta["规范性引用文件"] = {
            "章节标题": refs_title,
            "引导语": refs_chapter.guide_sentence,
            "引用条目数": len(refs_chapter.entries),
            "术语来源标准数": len(term_source),
        }

    # ---- 3. 两章节联动（主函数实现）----
    linkage_issues = linkage_term_source_vs_references(
        terms_chapter, refs_chapter, para_objects_terms
    )
    all_issues.extend(linkage_issues)

    # ---- 4. 结构性联动（FVT003 / FVT007）----
    struct_issues = structural_linkage_checks(
        terms_span, refs_span, para_objects_terms, para_objects_refs,
        terms_idx, refs_idx,
    )
    all_issues.extend(struct_issues)

    # 按章节拆分，便于回写批注时各自锚定
    terms_related = [i for i in all_issues
                     if i.rule.startswith("术语和定义") or i.rule.startswith("章节-术语和定义")
                     or i.rule == "术语和定义-来源联动"
                     or (i.rule.startswith("章节-结构顺序"))]
    refs_related = [i for i in all_issues
                    if i.rule.startswith("规范性引用文件")]
    # 其余（如跨章节 info）归入术语侧一并回写
    terms_side = [i for i in all_issues if i in terms_related or i not in refs_related]
    # 简化：terms_side 含全部非 refs 类问题；refs_side 仅 refs 类
    terms_side = [i for i in all_issues if i not in refs_related]

    return {
        "doc": doc,
        "items": items,
        "all_issues": all_issues,
        "terms_issues": terms_side,
        "refs_issues": refs_related,
        "para_objects_terms": para_objects_terms,
        "title_para_terms": title_para_terms,
        "para_objects_refs": para_objects_refs,
        "title_para_refs": title_para_refs,
        "meta": meta,
    }


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="标准文档审核主控（术语 + 规范性引用文件）")
    parser.add_argument("input", help="输入 .docx 文件")
    parser.add_argument("-o", "--output", help="输出 JSON 结果文件路径（可选）")
    parser.add_argument("--docx-out", help="带批注 docx 输出路径（默认 <输入名>_双章节审核批注.docx）")
    parser.add_argument("--xlsx-out", help="问题汇总 xlsx 输出路径（默认 <输入名>_双章节审核问题.xlsx）")
    parser.add_argument("--no-comments", action="store_true", help="仅输出 xlsx，不写批注")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在 {input_path}")
        return 1

    result = audit_document(str(input_path))
    issues = result["all_issues"]

    summary = {
        "术语和定义": result["meta"]["术语和定义"],
        "规范性引用文件": result["meta"]["规范性引用文件"],
        "问题总数": len(issues),
        "审核结论": "通过" if not issues else "存在问题",
        "问题列表": [i.to_dict() for i in issues],
    }

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(json.dumps(summary, ensure_ascii=False, indent=2))
        print(f"JSON 结果已保存：{args.output}")

    # xlsx
    xlsx_path = args.xlsx_out or str(input_path.with_name(input_path.stem + "_双章节审核问题.xlsx"))
    meta_xlsx = {
        "章节标题": "术语和定义 / 规范性引用文件",
        "引导语": f"术语:{summary['术语和定义'].get('引导语','')} | 引用:{summary['规范性引用文件'].get('引导语','')}",
        "术语条目数": summary["术语和定义"].get("术语条目数", ""),
        "子章节": "、".join(summary["术语和定义"].get("子章节", []) or []),
        "问题总数": len(issues),
        "审核结论": summary["审核结论"],
    }
    write_issues_xlsx(issues, xlsx_path, meta=meta_xlsx)
    print(f"xlsx 汇总已保存：{xlsx_path}")

    # docx 批注（两章节分别锚定）
    if not args.no_comments:
        docx_out = args.docx_out or str(input_path.with_name(input_path.stem + "_双章节审核批注.docx"))
        doc = result["doc"]
        if result["para_objects_terms"] or result["title_para_terms"]:
            write_annotated_docx(
                doc, result["terms_issues"], result["para_objects_terms"],
                result["title_para_terms"], docx_out, author="术语审核",
            )
        if result["para_objects_refs"] or result["title_para_refs"]:
            write_annotated_docx(
                doc, result["refs_issues"], result["para_objects_refs"],
                result["title_para_refs"], docx_out, author="引用审核",
            )
        print(f"带批注 docx 已保存：{docx_out}")

    print(f"审核完成：共 {len(issues)} 个问题。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
