#!/usr/bin/env python3
"""
生成专业格式的业务需求文档评审报告（Word .docx）
风格：正式评审意见 / 产品需求评审 / 咨询式文档
"""

import sys
import os
import json
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# ── Color palette (subdued, professional) ──
COLOR_PRIMARY = "1A3C6E"      # Deep navy for headings
COLOR_SECONDARY = "4A6FA5"    # Medium blue for subheadings
COLOR_ACCENT = "2E75B6"       # Accent blue
COLOR_PASS = "2D7D46"         # Muted green for pass
COLOR_FAIL = "8B3A3A"         # Muted red for fail
COLOR_BODY = "333333"         # Body text
COLOR_MUTED = "888888"        # Muted/footer text
COLOR_BORDER = "D0D0D0"       # Table borders
COLOR_HEADER_BG = "1A3C6E"    # Table header background
COLOR_ALT_ROW = "F5F7FA"      # Alternating row background


def hex_to_rgb(hex_color: str):
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def set_cell_bg(cell, hex_color: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>')
    cell._element.get_or_add_tcPr().append(shading)


def set_cell_text(cell, text: str, bold=False, color=None, size=Pt(10), alignment=None):
    """Set cell text with formatting."""
    cell.text = ""
    para = cell.paragraphs[0]
    if alignment is not None:
        para.alignment = alignment
    para.paragraph_format.space_before = Pt(2)
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run(text)
    run.font.size = size
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    if bold:
        run.bold = True
    if color:
        run.font.color.rgb = color


def add_styled_heading(doc, text: str, level: int):
    """Add a heading with professional styling."""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        run.font.color.rgb = hex_to_rgb(COLOR_PRIMARY)
    return h


def add_body_para(doc, text: str, bold=False, indent=None):
    """Add a body paragraph with professional styling."""
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.font.size = Pt(10.5)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.color.rgb = hex_to_rgb(COLOR_BODY)
    if bold:
        run.bold = True
    if indent:
        para.paragraph_format.left_indent = Cm(indent)
    para.paragraph_format.space_after = Pt(4)
    return para


def add_divider(doc):
    """Add a thin horizontal line."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    pPr = para._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="4" w:space="1" w:color="{COLOR_BORDER}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def build_style_table(doc, headers: list, rows: list, col_widths: list = None):
    """Build a professionally styled table."""
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set table borders
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{COLOR_BORDER}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, h, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), size=Pt(9))
        set_cell_bg(cell, COLOR_HEADER_BG)

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_data in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            if isinstance(cell_data, tuple):
                text, is_bold, color = cell_data
            else:
                text, is_bold, color = str(cell_data), False, hex_to_rgb(COLOR_BODY)
            set_cell_text(cell, text, bold=is_bold, color=color, size=Pt(9))
            if r_idx % 2 == 1:
                set_cell_bg(cell, COLOR_ALT_ROW)

    # Column widths
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)

    doc.add_paragraph()  # spacing
    return table


def generate_report(scores: dict, output_path: str):
    doc = Document()

    # Page setup
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # ==================== COVER / TITLE ====================
    title = doc.add_heading('业务需求文档评审意见', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in title.runs:
        run.font.color.rgb = hex_to_rgb(COLOR_PRIMARY)
        run.font.size = Pt(22)

    # Meta line
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = meta.add_run(f"评审日期：{datetime.now().strftime('%Y年%m月%d日')}")
    run.font.size = Pt(9)
    run.font.color.rgb = hex_to_rgb(COLOR_MUTED)
    meta.paragraph_format.space_after = Pt(4)

    run2 = meta.add_run("     |     ")
    run2.font.size = Pt(9)
    run2.font.color.rgb = hex_to_rgb(COLOR_MUTED)

    run3 = meta.add_run("密级：内部")
    run3.font.size = Pt(9)
    run3.font.color.rgb = hex_to_rgb(COLOR_MUTED)

    add_divider(doc)

    # ==================== 一、评审概要 ====================
    add_styled_heading(doc, "一、评审概要", 1)

    total_score = scores.get("total_score", 0)
    max_score = scores.get("max_score", 100)
    passed = total_score >= 60

    # Overall result box
    result_para = doc.add_paragraph()
    result_para.paragraph_format.space_before = Pt(8)
    result_para.paragraph_format.space_after = Pt(8)

    result_text = "评审结论：建议通过" if passed else "评审结论：建议驳回修订"
    run_result = result_para.add_run(result_text)
    run_result.bold = True
    run_result.font.size = Pt(13)
    run_result.font.color.rgb = hex_to_rgb(COLOR_PASS if passed else COLOR_FAIL)

    add_body_para(doc, f"综合评分：{total_score} / {max_score}（通过线 ≥ 60 分）")

    # Dimension scores summary table
    dimension_scores = scores.get("dimensions", {})
    if dimension_scores:
        add_styled_heading(doc, "各维度评估概览", 2)

        headers = ["评估维度", "满分", "得分", "达成率", "评价"]
        rows = []
        for dim_name, dim_data in dimension_scores.items():
            max_s = dim_data.get("max", 0)
            score = dim_data.get("score", 0)
            rate = f"{score / max_s * 100:.0f}%" if max_s > 0 else "-"
            rating = dim_data.get("rating", "-")

            # Color code the rating
            rating_lower = rating
            if "低" in rating_lower or "较差" in rating_lower:
                rate_color = hex_to_rgb(COLOR_FAIL)
            elif "高" in rating_lower or "良好" in rating_lower:
                rate_color = hex_to_rgb(COLOR_PASS)
            else:
                rate_color = hex_to_rgb(COLOR_ACCENT)

            rows.append([
                dim_name,
                str(max_s),
                str(score),
                rate,
                (rating, False, rate_color),
            ])

        build_style_table(doc, headers, rows, col_widths=[3.5, 1.5, 1.5, 1.5, 2.5])

    # Executive summary
    summary = scores.get("executive_summary", "")
    if summary:
        add_styled_heading(doc, "摘要", 2)
        add_body_para(doc, summary)

    add_divider(doc)

    # ==================== 二、逐项评审意见 ====================
    add_styled_heading(doc, "二、逐项评审意见", 1)

    analyses = scores.get("analyses", {})
    dim_order = scores.get("dim_order", list(analyses.keys()))

    for dim_name in dim_order:
        if dim_name not in analyses:
            continue

        analysis = analyses[dim_name]
        dim_score = dimension_scores.get(dim_name, {})
        max_s = dim_score.get("max", 0)
        score = dim_score.get("score", 0)
        rating = dim_score.get("rating", "-")

        add_styled_heading(doc, dim_name, 2)

        # Score line
        score_para = doc.add_paragraph()
        score_run = score_para.add_run(f"得分：{score}/{max_s}    评价：{rating}")
        score_run.bold = True
        score_run.font.size = Pt(10)
        score_run.font.color.rgb = hex_to_rgb(COLOR_SECONDARY)

        # Observations
        if isinstance(analysis, dict):
            observations = analysis.get("observations", [])
            suggestions = analysis.get("suggestions", [])
            key_points = analysis.get("key_points", [])

            if observations:
                for obs in observations:
                    add_body_para(doc, f"• {obs}", indent=0.5)

            if key_points:
                add_body_para(doc, "", bold=False)
                for kp in key_points:
                    add_body_para(doc, f"— {kp}", indent=1)

            if suggestions:
                add_body_para(doc, "", bold=False)
                add_body_para(doc, "优化方向：", bold=True)
                for sug in suggestions:
                    add_body_para(doc, f"→ {sug}", indent=0.5)

        elif isinstance(analysis, list):
            for item in analysis:
                if isinstance(item, dict):
                    for k, v in item.items():
                        add_body_para(doc, f"• {k}：{v}", indent=0.5)
                else:
                    add_body_para(doc, f"• {item}", indent=0.5)
        elif isinstance(analysis, str):
            # Split by paragraphs
            for line in analysis.strip().split("\n"):
                line = line.strip()
                if line:
                    add_body_para(doc, line)

    add_divider(doc)

    # ==================== 三、关键发现 ====================
    add_styled_heading(doc, "三、关键发现与改进建议", 1)

    issues = scores.get("issues", [])
    suggestions_list = scores.get("suggestions", [])

    if issues:
        add_styled_heading(doc, "需关注事项", 2)

        headers = ["编号", "位置", "问题描述", "影响", "关注级别"]
        rows = []
        for i, issue in enumerate(issues, 1):
            location = issue.get("location", "-")
            description = issue.get("description", "-")
            impact = issue.get("impact", "-")
            severity = issue.get("severity", "建议关注")

            severity_lower = severity
            if "重点" in severity_lower or "优先" in severity_lower:
                sev_color = hex_to_rgb(COLOR_FAIL)
            else:
                sev_color = hex_to_rgb(COLOR_ACCENT)

            rows.append([
                str(i),
                location,
                description,
                impact,
                (severity, True, sev_color),
            ])

        build_style_table(doc, headers, rows, col_widths=[1, 2.5, 5, 3, 2])

    if suggestions_list:
        add_styled_heading(doc, "改进建议", 2)
        for i, sug in enumerate(suggestions_list, 1):
            if isinstance(sug, dict):
                title = sug.get("title", f"建议 {i}")
                detail = sug.get("detail", "")

                sug_para = doc.add_paragraph()
                sug_run = sug_para.add_run(f"{i}. {title}")
                sug_run.bold = True
                sug_run.font.size = Pt(10.5)
                sug_run.font.name = "微软雅黑"
                sug_run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

                if detail:
                    add_body_para(doc, detail, indent=0.5)
            else:
                add_body_para(doc, f"{i}. {sug}")

    add_divider(doc)

    # ==================== 四、流程图专项评审 ====================
    flowchart_review = scores.get("flowchart_review")
    if flowchart_review:
        add_styled_heading(doc, "四、流程图专项评审", 1)

        if isinstance(flowchart_review, dict):
            headers = ["评审维度", "得分", "满分", "评审意见"]
            rows = []
            for dim_key, dim_label in [("granularity", "颗粒度"), ("completeness", "完整性"), ("clarity", "清晰度")]:
                dim_data = flowchart_review.get(dim_key, {})
                if isinstance(dim_data, dict):
                    rows.append([
                        dim_label,
                        str(dim_data.get("score", "-")),
                        str(dim_data.get("max", 4 if dim_key == "granularity" else 3)),
                        dim_data.get("comment", "-"),
                    ])

            if rows:
                build_style_table(doc, headers, rows, col_widths=[2.5, 1.5, 1.5, 8])

            # Overall flowchart comment
            overall = flowchart_review.get("overall", "")
            if overall:
                add_body_para(doc, overall)
        elif isinstance(flowchart_review, str):
            add_body_para(doc, flowchart_review)

    # ==================== Footer ====================
    doc.add_paragraph()
    add_divider(doc)
    footer_para = doc.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_para.add_run("— 本评审意见由需求评审系统自动生成 —")
    footer_run.font.size = Pt(8)
    footer_run.font.color.rgb = hex_to_rgb(COLOR_MUTED)
    footer_run.italic = True

    doc.save(output_path)
    print(f"评审报告已生成: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python generate_report.py '<JSON数据>' [输出路径]")
        print("  或: python generate_report.py --file <JSON文件路径> [输出路径]")
        sys.exit(1)

    output_path = "需求评审意见.docx"

    if sys.argv[1] == "--file":
        json_path = sys.argv[2]
        output_path = sys.argv[3] if len(sys.argv) > 3 else output_path
        with open(json_path, "r", encoding="utf-8") as f:
            scores = json.load(f)
    else:
        scores = json.loads(sys.argv[1])
        output_path = sys.argv[2] if len(sys.argv) > 2 else output_path

    generate_report(scores, output_path)
