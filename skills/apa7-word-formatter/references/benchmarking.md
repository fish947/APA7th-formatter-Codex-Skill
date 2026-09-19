# Repeatable benchmark and regression workflow

Use this only for formatter maintenance, large-corpus testing, or release decisions. Ordinary paper formatting follows `workflow.md`.

## What counts as good

A case passes only when all applicable gates pass:

- the source hash is unchanged;
- exactly one readable DOCX is created;
- saved-format and content-preservation checks pass;
- expected machine-check statuses and issue codes match;
- every required rendered page receives a fresh `pass` review.

An engine pass with missing visual review is `needs_visual_review`, not a complete pass. Averages never override a preservation failure or a page marked `issue`.

## Synthetic regression suite

Run from the project repository, or replace the script path with this installed Skill's `scripts/apa7_benchmark.py`:

```sh
python3 tools/apa7_benchmark.py generate benchmark/generated/smoke
python3 tools/apa7_benchmark.py run benchmark/generated/smoke/cases.json --output-dir benchmark/runs/smoke
```

The generated documents contain only fictional content. The suite covers a normal student paper, a professional paper with expected warnings, and a complex merged table that must be preserved.

## Add a private real-paper case

Do not copy a private manuscript into the repository. Create an ignored case folder that references it in place:

```sh
python3 tools/apa7_benchmark.py prepare-case "/private/paper.docx" --id case_001 --profile student --output-dir benchmark/private/case_001
```

For professional mode, also pass `--running-head`. The command creates a pending structure configuration and a private manifest; it does not copy or edit the paper. Complete the structure classification and compliance review using `workflow.md`, then set explicit expectations in the manifest.

For a release-quality run, use the trusted documents renderer:

```sh
python3 tools/apa7_benchmark.py run benchmark/private/case_001/cases.private.json --output-dir benchmark/runs/case_001 --renderer "/trusted/documents/render_docx.py"
```

Inspect every rendered page. Fill the generated `page-review.template.json` with `pass` or `issue` and concrete notes, then record it:

```sh
python3 tools/apa7_benchmark.py record-qa benchmark/runs/case_001 case_001 "/path/page-review.json"
```

## Public held-out sources

Openly licensed public DOCX files are useful for compatibility testing, but do not commit the documents themselves. Before downloading, verify the license on the source landing page. Record only the title, landing page, direct download URL, license, intended use, exact size, MD5 and SHA-256 in `benchmark/public-sources.json`, and keep the downloaded file under the ignored `benchmark/private/` directory.

After download, verify that every local file matches the reviewed registry entry:

```sh
python3 tools/apa7_benchmark.py verify-sources benchmark/public-sources.json --corpus-dir benchmark/private/open-corpus
```

A checksum mismatch is a different test document and must not silently replace the held-out case. Public availability does not remove the need for license review, attribution, privacy screening, or a fresh page-by-page QA record.

## Turn a failure into a permanent regression

1. Confirm the defect against the unchanged source and latest render.
2. Remove private wording and reduce it to the smallest fictional DOCX structure that still reproduces the defect.
3. Add the fixture construction and its expected machine status or namespaced issue code to the synthetic suite.
4. Add a focused unit or integration test that fails before the fix.
5. Make the narrow formatter correction.
6. Run all tests, the synthetic benchmark, and the held-out real-paper cases again.

Keep some real cases out of day-to-day development as a held-out set. Never mark AI classification accuracy, bibliographic facts, scientific meaning, or APA compliance as proven by the Python gates alone.
