# Execution and review contract

Commands below assume the current directory is this skill folder. Otherwise substitute absolute paths. `python3` means the selected compatible runtime, preferably the host's bundled document Python. Do not change global Python packages or use `--break-system-packages`.

Before running these commands, the user must explicitly choose `student` or `professional`. Do not infer the profile from the manuscript. Use a newly created temporary working directory for the structure JSON, trial DOCX files, renders and QA records; only the final reviewed DOCX is placed beside the source.

## Preparation

```sh
python3 scripts/apa7_format.py --sources
python3 scripts/apa7_workflow.py prepare "/path/paper.docx" --profile student --output "/path/new.structure.json"
python3 scripts/apa7_format.py "/path/paper.docx" --inspect-visuals
```

`prepare` is read-only for the Word file. It creates a fresh JSON draft and rejects an existing destination. `_review.paragraphs` contains 1-based direct-body paragraph IDs (including empty paragraphs); `_review.tables` contains 1-based top-level Word table IDs, cells and complexity flags; `_review.body_order` shows their interleaving and unsupported containers. Nested cell paragraphs do not use top-level paragraph IDs.

The draft starts with empty classifications and `review_status: pending`. Do not mechanically copy every inferred role and mark it reviewed without reading the content. Read large drafts in chunks until every relevant object is accounted for.

## Configuration to fill

The following demonstrates the shape for a two-paragraph, one-table input, not universal IDs. Retain the actual 64-character `source_sha256` from preparation.

```json
{
  "source_sha256": "use_the_actual_hash_from_prepare",
  "profile": "student",
  "review_status": "reviewed",
  "roles": {"1": "heading2", "2": "body"},
  "table_roles": {"1": "data"},
  "table_header_rows": {"1": 1},
  "run_in_headings": {},
  "replace_headers": false,
  "decision_notes": {
    "paragraphs": {
      "1": {"confidence": "high", "reason": "Participants is a subsection within the Method section."},
      "2": {"confidence": "high", "reason": "Complete prose sentences describing the participant sample."}
    },
    "tables": {
      "1": {"confidence": "high", "reason": "Condition/Mean/SD header followed by numeric result rows; one header row."}
    }
  },
  "unresolved": []
}
```

All paragraphs and top-level tables must have one valid role and one reason. Confidence is `high` or `uncertain`; the latter requires role `preserve`. Reasons should identify the actual contextual evidence, not just say “APA.” Unsupported complex tables must be preserved. For a table with no confidently identified header, preserve it rather than invent a header count.

Supported paragraph roles are returned by `_review.allowed_roles`. `run_in_headings` maps a heading-4/5 paragraph ID to the exact prefix, e.g. `Response Accuracy.`; the original paragraph must already continue with prose. `title_page`, when supplied, must be two valid paragraph IDs and agree with the explicit paragraph roles. `replace_headers` defaults to false; only enable it when replacement of the existing content is within the user's request.

The input hash binds the whole source version. If the source is edited or re-saved, re-prepare and re-assess the IDs. Never replace the hash merely to force an old configuration through. `_review` is preparation data, not evidence of completed AI review. The gate checks consistency and coverage; it cannot prove that an AI reason is correct.

## Apply

```sh
python3 scripts/apa7_workflow.py apply "/path/paper.docx" --config "/path/new.structure.json" --profile student --output "/path/paper_APA7.docx"
```

For professional mode, use `--profile professional --running-head "APPROVED SHORT TITLE"`. The configuration profile must match. Add `--export-visuals` only when visual export is requested. Font options are shown in the engine's `--help`.

The apply command returns one output DOCX and a concise `feedback` object. It does not create audit files by default. Read its four sections—changes, APA sources, changed locations and remaining checks—plus any unresolved items in the reviewed configuration. The engine may preserve an unsupported component even when the classification is correct.

For `.doc`, first create a `.docx` copy using a trusted conversion path and check conversion fidelity, then prepare that stable DOCX. The bare engine supports `--soffice` for conversion, but the reviewed workflow uses DOCX-bound IDs. Encrypted files, macro documents and unaccepted revisions are not supported; ask the author to resolve these on a copy rather than accepting revisions automatically.

## Render and record actual visual review

Use the available documents skill's renderer, resolved from the current host instead of a hardcoded user/machine path. On Codex desktop, load workspace dependencies first; the renderer must use bundled LibreOffice/Poppler, not silently fall back to the user's desktop application. If that skill/runtime is absent, explain the missing rendering capability and do not mark visual QA complete.

```sh
python3 scripts/apa7_workflow.py render "/path/paper_APA7.docx" --renderer "/trusted/documents/render_docx.py" --output-dir "/path/new_render"
```

This writes page PNGs, a QA PDF and `render_manifest.json`. Use the image-viewing capability to inspect **each** PNG, including any blank-looking page. Do not infer visual success from text extraction, file creation or a source hash.

After inspection, create a new page-review JSON with the actual output hash from the apply result/render manifest. Include every page exactly once:

```json
{
  "document_sha256": "use_the_actual_output_hash",
  "pages": [
    {"page": 1, "status": "pass", "notes": "Describe the title-page and header checks actually performed."},
    {"page": 2, "status": "issue", "notes": "Describe the actual defect and its location."}
  ]
}
```

Use `pass` only for a visually acceptable page. Then record the review:

```sh
python3 scripts/apa7_workflow.py record-qa "/path/paper_APA7.docx" --render-dir "/path/new_render" --review "/path/page-review.json" --output "/path/paper_APA7.qa.json"
```

Missing/duplicate page records, stale document hashes or changed page images are rejected. `VISUAL_REVIEW_RECORDED` means the supplied page review was consistently recorded, not that APA content or scientific accuracy has been certified. `REVISION_REQUIRED` means at least one visible defect remains; do not deliver it as finished. Keep unresolved content items in the short “还要审核” section even if all pages visually pass.

After review, copy only the accepted DOCX from the temporary folder to a new destination beside the source. Do not overwrite the source or keep multiple trial copies there. In the reply, use only: “改了什么”, “APA 来源”, “改了原稿哪里”, and “还要审核”.
