#!/usr/bin/env python3
"""
AI 需求共创产品经理 Agent — 评审报告生成器
风格：产品经理分析纪要，不是模板报告
"""

import sys
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ── Color palette ──
COLOR_PRIMARY   = "1A3C6E"
COLOR_SECONDARY = "4A6FA5"
COLOR_ACCENT    = "2E75B6"
COLOR_PASS      = "2D7D46"
COLOR_WARN      = "8B6914"
COLOR_FAIL      = "8B3A3A"
COLOR_BODY      = "2C2C2C"
COLOR_MUTED     = "888888"
COLOR_BORDER    = "D8D8D8"
COLOR_HEADER_BG = "1A3C6E"
COLOR_ALT_ROW   = "F7F9FC"
COLOR_CALLOUT   = "F0F4FA"


def hex_rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_cell_bg(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>')
    cell._element.get_or_add_tcPr().append(shd)


def set_cell_text(cell, text, bold=False, color=None, size=Pt(9.5)):
    cell.text = ""
    para = cell.paragraphs[0]
    para.paragraph_format.space_before = Pt(3)
    para.paragraph_format.space_after = Pt(3)
    run = para.add_run(str(text))
    run.font.size = size
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if bold:
        run.bold = True
    if color:
        run.font.color.rgb = color


def add_section_title(doc, text):
    """一级章节标题"""
    h = doc.add_heading(text, level=1)
    for run in h.runs:
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        run.font.color.rgb = hex_rgb(COLOR_PRIMARY)
    return h


def add_sub_title(doc, text):
    """二级小标题，比一级轻"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(10)
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.color.rgb = hex_rgb(COLOR_SECONDARY)
    return para


def add_body(doc, text, indent=None, color=None, size=Pt(10.5), space_after=Pt(6)):
    """正文段落"""
    para = doc.add_paragraph()
    run = para.add_run(str(text))
    run.font.size = size
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.color.rgb = hex_rgb(color or COLOR_BODY)
    if indent:
        para.paragraph_format.left_indent = Cm(indent)
    para.paragraph_format.space_after = space_after
    return para


def add_bullet(doc, text, color=None, indent=0.5):
    """简洁 bullet，不用 Word 列表样式"""
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(indent)
    para.paragraph_format.space_after = Pt(3)
    run = para.add_run(f"· {text}")
    run.font.size = Pt(10.5)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.color.rgb = hex_rgb(color or COLOR_BODY)
    return para


def add_numbered(doc, num, text, color=None, indent=0.5):
    """编号条目"""
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(indent)
    para.paragraph_format.space_after = Pt(4)
    run = para.add_run(f"{num}. {text}")
    run.font.size = Pt(10.5)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.color.rgb = hex_rgb(color or COLOR_BODY)
    return para


def add_divider(doc, light=False):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after = Pt(4)
    color = "E8E8E8" if light else COLOR_BORDER
    pPr = para._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{"2" if light else "4"}" w:space="1" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def add_callout(doc, text, color_hex=None):
    """高亮引用块，用于核心判断"""
    bg = color_hex or COLOR_CALLOUT
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(0.5)
    para.paragraph_format.right_indent = Cm(0.5)
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    # 左边框
    pPr = para._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="12" w:space="4" w:color="{COLOR_ACCENT}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    run = para.add_run(str(text))
    run.font.size = Pt(10.5)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.color.rgb = hex_rgb(COLOR_BODY)
    return para


def build_simple_table(doc, rows, col_widths=None):
    """简洁两列表格，用于行动项"""
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:left w:val="none"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="2" w:space="0" w:color="E8E8E8"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)
    for r_idx, (label, content) in enumerate(rows):
        cell0 = table.rows[r_idx].cells[0]
        cell1 = table.rows[r_idx].cells[1]
        set_cell_text(cell0, label, bold=True, color=hex_rgb(COLOR_SECONDARY))
        set_cell_text(cell1, content, color=hex_rgb(COLOR_BODY))
        if r_idx % 2 == 1:
            set_cell_bg(cell0, COLOR_ALT_ROW)
            set_cell_bg(cell1, COLOR_ALT_ROW)
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return table


def generate_report_pm(data: dict, output_path: str):
    """产品经理分析纪要风格报告"""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.8)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.5)

    # ── 标题区 ──
    title = doc.add_heading('需求分析纪要', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in title.runs:
        run.font.color.rgb = hex_rgb(COLOR_PRIMARY)
        run.font.size = Pt(20)
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = meta.add_run(f"{datetime.now().strftime('%Y年%m月%d日')}  ·  内部文件")
    r1.font.size = Pt(9)
    r1.font.color.rgb = hex_rgb(COLOR_MUTED)
    meta.paragraph_format.space_after = Pt(2)
    add_divider(doc)

    # ── 第一部分：核心判断 ──
    add_section_title(doc, "一、核心判断")

    # 需求本质
    essence = data.get("requirement_essence", "")
    if essence:
        add_sub_title(doc, "当前需求本质")
        add_callout(doc, essence)

    # 当前建议
    verdict = data.get("verdict", "")
    verdict_reason = data.get("verdict_reason", "")
    preconditions = data.get("preconditions", [])

    if verdict:
        add_sub_title(doc, "当前阶段建议")
        verdict_color = COLOR_PASS if "推进" in verdict and "条件" not in verdict else (
            COLOR_WARN if "条件" in verdict else COLOR_FAIL
        )
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(verdict)
        r.bold = True
        r.font.size = Pt(12)
        r.font.name = "微软雅黑"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        r.font.color.rgb = hex_rgb(verdict_color)

        if verdict_reason:
            add_body(doc, verdict_reason)
        for cond in preconditions:
            add_bullet(doc, cond, color=COLOR_SECONDARY)

    # 最大风险
    top_risk = data.get("top_risk", "")
    if top_risk:
        add_sub_title(doc, "当前最大风险")
        add_callout(doc, top_risk)

    add_divider(doc)

    # ── 第二部分：核心问题分析 ──
    core = data.get("core_analysis", {})
    core_title = core.get("title", "二、核心问题分析") if core else "二、核心问题分析"
    add_section_title(doc, f"二、{core_title}" if not core_title.startswith("二") else core_title)

    # 核心矛盾
    cc = data.get("core_contradiction", {})
    if cc:
        ctype = cc.get("type", "")
        desc = cc.get("description", "")
        consequence = cc.get("consequence", "")
        if ctype:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(f"核心矛盾：{ctype}")
            r.bold = True
            r.font.size = Pt(10.5)
            r.font.name = "微软雅黑"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
            r.font.color.rgb = hex_rgb(COLOR_FAIL)
        if desc:
            add_body(doc, desc)
        if consequence:
            add_body(doc, f"如果不解决：{consequence}", color=COLOR_WARN, size=Pt(10))

    # 核心分析段落
    if core:
        paragraphs = core.get("paragraphs", [])
        for para_text in paragraphs:
            if para_text:
                add_body(doc, para_text)

    # 兼容旧版 focused_analysis
    focused = data.get("focused_analysis", [])
    if focused and not core:
        for item in focused:
            if not isinstance(item, dict):
                continue
            focus = item.get("focus", "")
            if focus:
                add_sub_title(doc, focus)
            for field, label, color in [
                ("current_state", "现状", COLOR_BODY),
                ("core_problem", "核心问题", COLOR_FAIL),
                ("risk", "风险", COLOR_WARN),
                ("recommendation", "建议", COLOR_SECONDARY),
            ]:
                val = item.get(field, "")
                if val:
                    add_body(doc, f"{label}：{val}", color=color, size=Pt(10))

    add_divider(doc)

    # ── 第三部分：AI 适配性 ──
    ai_fit = data.get("ai_fit", {})
    # 兼容旧版 ai_necessity
    ai_necessity = data.get("ai_necessity", "")
    ai_reason = data.get("ai_necessity_reason", "")

    add_section_title(doc, "三、AI 适配性")

    if ai_necessity:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        color = COLOR_PASS if ai_necessity == "必要" else (COLOR_WARN if "部分" in ai_necessity else COLOR_FAIL)
        r = p.add_run(f"总体判断：{ai_necessity}")
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.name = "微软雅黑"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        r.font.color.rgb = hex_rgb(color)
        if ai_reason:
            add_body(doc, ai_reason, size=Pt(10))

    suitable = ai_fit.get("suitable", []) if ai_fit else []
    rule_based = ai_fit.get("rule_based", []) if ai_fit else []

    if suitable:
        add_sub_title(doc, "适合 AI 的场景")
        for item in suitable:
            add_bullet(doc, item, color=COLOR_PASS)

    if rule_based:
        add_sub_title(doc, "更适合规则的场景")
        for item in rule_based:
            add_bullet(doc, item, color=COLOR_SECONDARY)

    add_divider(doc)

    # ── 第四部分：MVP 建议 ──
    mvp = data.get("mvp", {})
    # 兼容旧版 next_steps
    ns = data.get("next_steps", {})

    add_section_title(doc, "四、MVP 建议")

    defer = mvp.get("defer", []) if mvp else ns.get("defer_to_later", [])
    phase1 = mvp.get("phase1", []) if mvp else []

    # 兼容旧版 phase_plan
    if not phase1 and ns.get("phase_plan"):
        phase1 = [f"{p.get('phase', '')}：{p.get('focus', '')}" for p in ns.get("phase_plan", [])]

    mvp_scope = ns.get("mvp_scope", "") if ns else ""

    if mvp_scope:
        add_body(doc, mvp_scope)

    if phase1:
        add_sub_title(doc, "建议第一阶段先做")
        for i, item in enumerate(phase1, 1):
            add_numbered(doc, i, item, color=COLOR_BODY)

    if defer:
        add_sub_title(doc, "当前不建议做")
        for item in defer:
            add_bullet(doc, item, color=COLOR_MUTED)

    add_divider(doc)

    # ── 第五部分：下一步行动 ──
    next_actions = data.get("next_actions", [])
    # 兼容旧版 next_steps.must_do_before_start
    if not next_actions and ns:
        must_do = ns.get("must_do_before_start", [])
        next_actions = [{"action": item, "why": "", "output": ""} for item in must_do]

    add_section_title(doc, "五、下一步行动")

    if next_actions:
        for i, action in enumerate(next_actions, 1):
            if isinstance(action, dict):
                act = action.get("action", "")
                why = action.get("why", "")
                output = action.get("output", "")
                if act:
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(6)
                    p.paragraph_format.space_after = Pt(2)
                    r = p.add_run(f"{i}. {act}")
                    r.bold = True
                    r.font.size = Pt(10.5)
                    r.font.name = "微软雅黑"
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
                    r.font.color.rgb = hex_rgb(COLOR_BODY)
                    if why:
                        add_body(doc, why, indent=0.5, color=COLOR_MUTED, size=Pt(10), space_after=Pt(2))
                    if output:
                        add_body(doc, f"产出：{output}", indent=0.5, color=COLOR_SECONDARY, size=Pt(10), space_after=Pt(4))
            else:
                add_numbered(doc, i, str(action))

    # 兼容旧版 blocking_issues / main_risks
    blocking = data.get("blocking_issues", [])
    main_risks = data.get("main_risks", data.get("high_risks", []))
    if blocking or main_risks:
        add_divider(doc, light=True)
        add_sub_title(doc, "附：需关注的问题")
        all_issues = [(item, COLOR_FAIL) for item in blocking] + [(item, COLOR_WARN) for item in main_risks]
        for item, color in all_issues:
            if isinstance(item, dict):
                iss = item.get("issue", item.get("risk", ""))
                impact = item.get("impact", "")
                action_text = item.get("required_action", item.get("mitigation", ""))
                if iss:
                    add_bullet(doc, iss, color=color)
                if impact:
                    add_body(doc, f"影响：{impact}", indent=1.0, color=COLOR_MUTED, size=Pt(9.5), space_after=Pt(2))
                if action_text:
                    add_body(doc, f"建议：{action_text}", indent=1.0, color=COLOR_SECONDARY, size=Pt(9.5), space_after=Pt(4))

    # ── 页脚 ──
    doc.add_paragraph()
    add_divider(doc)
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run("— 本文件由 AI 产品经理 Agent 生成，供内部参考 —")
    fr.font.size = Pt(8)
    fr.font.color.rgb = hex_rgb(COLOR_MUTED)
    fr.italic = True

    doc.save(output_path)
    print(f"评审报告已生成: {output_path}")


def generate_report_legacy(data: dict, output_path: str):
    """旧版评分结构兼容报告（保留原有逻辑）"""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    title = doc.add_heading('业务需求文档评审意见', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in title.runs:
        run.font.color.rgb = hex_rgb(COLOR_PRIMARY)
        run.font.size = Pt(22)

    meta = doc.add_paragraph()
    r1 = meta.add_run(f"评审日期：{datetime.now().strftime('%Y年%m月%d日')}")
    r1.font.size = Pt(9)
    r1.font.color.rgb = hex_rgb(COLOR_MUTED)
    meta.paragraph_format.space_after = Pt(4)
    add_divider(doc)

    add_section_title(doc, "一、评审概要")
    total_score = data.get("total_score", 0)
    max_score = data.get("max_score", 100)
    passed = total_score >= 60
    result_para = doc.add_paragraph()
    r = result_para.add_run("评审结论：建议通过" if passed else "评审结论：建议驳回修订")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = hex_rgb(COLOR_PASS if passed else COLOR_FAIL)
    add_body(doc, f"综合评分：{total_score} / {max_score}（通过线 ≥ 60 分）")

    summary = data.get("executive_summary", "")
    if summary:
        add_body(doc, summary)

    add_divider(doc)
    add_section_title(doc, "二、评审意见")

    analyses = data.get("analyses", {})
    dim_order = data.get("dim_order", list(analyses.keys()))
    dimension_scores = data.get("dimensions", {})
    for dim_name in dim_order:
        if dim_name not in analyses:
            continue
        analysis = analyses[dim_name]
        dim_score = dimension_scores.get(dim_name, {})
        score = dim_score.get("score", 0)
        max_s = dim_score.get("max", 0)
        rating = dim_score.get("rating", "-")
        add_sub_title(doc, f"{dim_name}（{score}/{max_s}，{rating}）")
        if isinstance(analysis, dict):
            for obs in analysis.get("observations", []):
                add_bullet(doc, obs)
            for sug in analysis.get("suggestions", []):
                add_body(doc, f"→ {sug}", color=COLOR_SECONDARY, size=Pt(10))
        elif isinstance(analysis, str):
            for line in analysis.strip().split("\n"):
                if line.strip():
                    add_body(doc, line.strip())

    doc.add_paragraph()
    add_divider(doc)
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run("— 本评审意见由需求评审系统自动生成 —")
    fr.font.size = Pt(8)
    fr.font.color.rgb = hex_rgb(COLOR_MUTED)
    fr.italic = True

    doc.save(output_path)
    print(f"评审报告已生成: {output_path}")


def generate_report(data: dict, output_path: str):
    """自动判断格式"""
    new_keys = {"verdict", "blocking_issues", "focused_analysis", "core_contradiction",
                "requirement_essence", "ai_fit", "mvp", "next_actions"}
    if any(k in data for k in new_keys):
        generate_report_pm(data, output_path)
    else:
        generate_report_legacy(data, output_path)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python generate_report.py --file <JSON文件路径> [输出路径]")
        sys.exit(1)

    output_path = "需求分析纪要.docx"

    if sys.argv[1] == "--file":
        json_path = sys.argv[2]
        output_path = sys.argv[3] if len(sys.argv) > 3 else output_path
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = json.loads(sys.argv[1])
        output_path = sys.argv[2] if len(sys.argv) > 2 else output_path

    generate_report(data, output_path)
