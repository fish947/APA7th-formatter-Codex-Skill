#!/usr/bin/env python3
"""AI-reviewed APA workflow. This validates records; it is not an AI model."""
from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path
import re
import subprocess
import sys
from datetime import datetime, timezone

import apa7_format as engine


def write_new_json(path, value):
    path = Path(path).expanduser().absolute()
    if path.exists() or path.is_symlink():
        raise ValueError(f"输出已存在，不能覆盖：{path}")
    if path.suffix.lower() != ".json":
        raise ValueError("记录输出必须使用 .json 扩展名。")
    with path.open("x", encoding="utf-8") as output:
        json.dump(value, output, ensure_ascii=False, indent=2)
    return path


def load_json(path):
    value = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError("JSON 顶层必须是对象。")
    return value


def prepare(source, profile):
    draft = engine.prepare_config(source, profile)
    draft.update(review_status="pending", profile=profile,
                 decision_notes={"paragraphs": {}, "tables": {}}, unresolved=[])
    return draft


def check_note(notes, key, role):
    note = notes.get(key)
    if not isinstance(note, dict) or note.get("confidence") not in {"high", "uncertain"}:
        raise ValueError(f"对象 {key} 缺少 high／uncertain 置信分类。")
    if not isinstance(note.get("reason"), str) or not note["reason"].strip():
        raise ValueError(f"对象 {key} 缺少分类依据。")
    if note["confidence"] == "uncertain" and role != "preserve":
        raise ValueError(f"不确定对象 {key} 必须设为 preserve，不能强行排版。")


def validate_review(source, config, profile):
    if not isinstance(config, dict) or config.get("review_status") != "reviewed":
        raise ValueError("结构判断尚未完成；请由 AI 阅读文档并填写 reviewed 配置。")
    if config.get("profile") != profile:
        raise ValueError("配置的论文模式与本次执行不一致。")
    fresh = engine.prepare_config(source, profile)
    if config.get("source_sha256") != fresh["source_sha256"]:
        raise ValueError("结构配置与当前原稿版本不一致，请重新准备并判断。")
    paragraphs, tables = fresh["_review"]["paragraphs"], fresh["_review"]["tables"]
    roles, table_roles = config.get("roles"), config.get("table_roles")
    notes = config.get("decision_notes", {})
    if not isinstance(notes, dict):
        raise ValueError("decision_notes 必须是对象。")
    for label, values, expected in (("roles", roles, {str(p["paragraph"]) for p in paragraphs}),
                                    ("table_roles", table_roles, {str(t["table"]) for t in tables})):
        if not isinstance(values, dict) or set(values) != expected:
            raise ValueError(f"{label} 必须覆盖全部对应对象，不能遗漏或使用旧编号。")
    paragraph_notes, table_notes = notes.get("paragraphs"), notes.get("tables")
    if not isinstance(paragraph_notes, dict) or set(paragraph_notes) != set(roles):
        raise ValueError("每个段落必须有 decision_notes.paragraphs 记录。")
    if not isinstance(table_notes, dict) or set(table_notes) != set(table_roles):
        raise ValueError("每个表格必须有 decision_notes.tables 记录。")
    for key, role in roles.items():
        if role not in engine.ROLES:
            raise ValueError(f"无效段落角色：{key}={role}")
        check_note(paragraph_notes, key, role)
        if role in {"heading4", "heading5"}:
            prefix = config.get("run_in_headings", {}).get(key)
            text = paragraphs[int(key) - 1]["text"]
            if not isinstance(prefix, str) or not prefix.endswith(".") or not text.startswith(prefix) or not text[len(prefix):].strip():
                raise ValueError(f"同行标题 {key} 必须有精确前缀、结尾句号及同段正文。")
    header_rows = config.get("table_header_rows", {})
    if not isinstance(header_rows, dict) or not set(header_rows).issubset(set(table_roles)):
        raise ValueError("table_header_rows 包含无效编号。")
    for table in tables:
        key = str(table["table"])
        role = table_roles[key]
        if role not in {"data", "preserve"}:
            raise ValueError(f"无效表格角色：{key}={role}")
        check_note(table_notes, key, role)
        if table["complex"] and role != "preserve":
            raise ValueError(f"复杂表格 {key} 当前只能 preserve，需专门处理和复核。")
        if role == "data":
            count = header_rows.get(key)
            if type(count) is not int or not 1 <= count <= table["rows"]:
                raise ValueError(f"数据表 {key} 需要明确且有效的表头行数。")
    cover = config.get("title_page")
    if cover is not None:
        if not isinstance(cover, list) or len(cover) != 2 or any(type(i) is not int for i in cover):
            raise ValueError("title_page 应为两个整数段号。")
        start, end = cover
        if not 1 <= start <= end <= len(paragraphs):
            raise ValueError("title_page 超出段落范围。")
        for index in range(start, end + 1):
            if paragraphs[index - 1]["text"].strip():
                expected = "title" if index == start else "title_meta"
                if roles[str(index)] != expected or paragraph_notes[str(index)]["confidence"] != "high":
                    raise ValueError("title_page 与已判断角色冲突，不能用范围覆盖不确定段落。")
    unresolved = config.get("unresolved", [])
    if not isinstance(unresolved, list) or any(not isinstance(x, str) for x in unresolved):
        raise ValueError("unresolved 必须是待确认事项字符串列表。")
    validated = copy.deepcopy(config)
    validated.pop("_review", None)
    validated["ai_review"] = {"status": "classification_record_validated",
                              "source_sha256": fresh["source_sha256"], "decisions": notes,
                              "unresolved": unresolved,
                              "limitation": "Record consistency is validated; semantic correctness is not proven."}
    return validated


def apply_reviewed(source, config, *, profile, **options):
    validated = validate_review(source, config, profile)
    return engine.format_file(source, profile=profile, config=validated, **options)


def render_document(source, renderer, output_dir):
    """Call a trusted host documents renderer, then bind page files to this DOCX."""
    source, renderer = Path(source).resolve(), Path(renderer).resolve()
    output_dir = Path(output_dir).expanduser().absolute()
    if output_dir.exists() or output_dir.is_symlink():
        raise ValueError("渲染目录必须是新目录，不能复用旧页面。")
    if not source.is_file() or source.suffix.lower() != ".docx" or not renderer.is_file():
        raise ValueError("需要现有 .docx 和可信的 render_docx.py 路径。")
    before = engine.digest(source)
    result = subprocess.run([sys.executable, str(renderer), str(source), "--output_dir", str(output_dir), "--emit_pdf"],
                            capture_output=True, text=True, timeout=300)
    if result.returncode or engine.digest(source) != before:
        raise RuntimeError("渲染失败或文档在渲染时改变；不得记录为通过。\n" + (result.stderr or result.stdout)[-2000:])
    pages = []
    for path in output_dir.glob("page-*.png"):
        match = re.fullmatch(r"page-(\d+)\.png", path.name)
        if match:
            pages.append({"page": int(match[1]), "file": path.name, "sha256": engine.digest(path)})
    pages.sort(key=lambda p: p["page"])
    if not pages or [p["page"] for p in pages] != list(range(1, len(pages) + 1)):
        raise RuntimeError("渲染页缺失或不连续；不得记录为通过。")
    manifest = {"document": str(source), "document_sha256": before, "pages": pages,
                "status": "RENDERED_AWAITING_VISUAL_REVIEW"}
    write_new_json(output_dir / "render_manifest.json", manifest)
    return manifest


def record_qa(document, render_dir, review):
    """Check that a human/AI page review addresses every unchanged rendered page."""
    document, render_dir = Path(document).resolve(), Path(render_dir).resolve()
    manifest = load_json(render_dir / "render_manifest.json")
    current = engine.digest(document)
    if manifest.get("document_sha256") != current or review.get("document_sha256") != current:
        raise ValueError("视觉复核记录、渲染页和当前 Word 版本不一致。")
    pages = manifest.get("pages", [])
    if not pages or [p.get("page") for p in pages] != list(range(1, len(pages) + 1)):
        raise ValueError("渲染清单页号无效。")
    for page in pages:
        filename = page.get("file")
        if filename != f"page-{page['page']}.png":
            raise ValueError("渲染清单含不安全或不匹配的文件名。")
        path = render_dir / filename
        if path.is_symlink() or not path.is_file() or engine.digest(path) != page.get("sha256"):
            raise ValueError("渲染图片缺失或已改变，必须重新渲染并复核。")
    records = review.get("pages")
    if not isinstance(records, list) or any(not isinstance(p, dict) or type(p.get("page")) is not int for p in records):
        raise ValueError("pages 应为逐页复核记录。")
    if sorted(p["page"] for p in records) != list(range(1, len(pages) + 1)):
        raise ValueError("必须逐页复核全部页面，不能遗漏或重复页号。")
    for page in records:
        if page.get("status") not in {"pass", "issue"} or not isinstance(page.get("notes"), str) or not page["notes"].strip():
            raise ValueError("每页必须有 pass／issue 及实际检查说明。")
    return {"document": str(document), "document_sha256": current,
            "render_manifest_sha256": engine.digest(render_dir / "render_manifest.json"),
            "status": "REVISION_REQUIRED" if any(p["status"] == "issue" for p in records) else "VISUAL_REVIEW_RECORDED",
            "created_at": datetime.now(timezone.utc).isoformat(), "pages": records,
            "limitation": "This records the supplied page review, not full APA compliance or proof of semantic correctness."}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    prep = commands.add_parser("prepare", help="准备供 AI 判断的结构草稿")
    prep.add_argument("source", type=Path)
    prep.add_argument("--profile", choices=("student", "professional"), required=True)
    prep.add_argument("--output", type=Path, required=True)
    apply = commands.add_parser("apply", help="验证完整的 AI 分类记录后执行排版")
    apply.add_argument("source", type=Path)
    apply.add_argument("--config", type=Path, required=True)
    apply.add_argument("--profile", choices=("student", "professional"), required=True)
    apply.add_argument("--output", type=Path)
    apply.add_argument("--font", choices=tuple(engine.FONTS), default="Times New Roman")
    apply.add_argument("--running-head", default="")
    apply.add_argument("--export-visuals", action="store_true")
    render = commands.add_parser("render", help="调用当前环境的可信 documents 渲染器")
    render.add_argument("source", type=Path)
    render.add_argument("--renderer", type=Path, required=True)
    render.add_argument("--output-dir", type=Path, required=True)
    qa = commands.add_parser("record-qa", help="校验并记录 AI／人工逐页复核结果")
    qa.add_argument("document", type=Path)
    qa.add_argument("--render-dir", type=Path, required=True)
    qa.add_argument("--review", type=Path, required=True)
    qa.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "prepare":
            print(write_new_json(args.output, prepare(args.source, args.profile)))
        elif args.command == "apply":
            out, report = apply_reviewed(args.source, load_json(args.config), profile=args.profile, output=args.output,
                                         font=args.font, running_head=args.running_head, export_visuals=args.export_visuals)
            print(json.dumps({"document": str(out), "document_sha256": engine.digest(out),
                              "status": report["status"], "feedback": report["feedback"]}, ensure_ascii=False, indent=2))
        elif args.command == "render":
            print(json.dumps(render_document(args.source, args.renderer, args.output_dir), ensure_ascii=False, indent=2))
        else:
            print(write_new_json(args.output, record_qa(args.document, args.render_dir, load_json(args.review))))
        return 0
    except Exception as exc:
        print("未完成：" + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
