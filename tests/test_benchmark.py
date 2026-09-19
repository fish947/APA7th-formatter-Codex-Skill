"""Regression tests for the repeatable corpus benchmark."""
import json
import hashlib
from pathlib import Path
import tempfile
import unittest

from docx import Document

from tools import apa7_benchmark as benchmark


class BenchmarkTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="apa7_benchmark_test_")
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def test_generated_suite_passes_all_hard_gates_and_expectations(self):
        manifest = benchmark.generate_synthetic_suite(self.root / "fixtures")
        results = benchmark.run_suite(manifest, self.root / "run")
        self.assertEqual(results["summary"]["status"], "passed")
        self.assertEqual(results["summary"]["cases"], 3)
        for case in results["cases"]:
            self.assertEqual(case["status"], "passed", case)
            self.assertTrue(all(case["hard_gates"].values()), case)
            self.assertEqual(case["expectation_status"], "passed")
        self.assertIn("Overall: **PASS**", (self.root / "run" / "benchmark-summary.md").read_text())

    def test_changed_expectation_becomes_a_visible_regression_failure(self):
        manifest = benchmark.generate_synthetic_suite(self.root / "fixtures")
        value = benchmark.read_json(manifest)
        value["cases"] = [value["cases"][0]]
        value["cases"][0]["expect"]["exact_issue_codes"] = ["statistics:p_zero"]
        bad_manifest = benchmark.write_new_json(self.root / "fixtures" / "bad-cases.json", value)
        results = benchmark.run_suite(bad_manifest, self.root / "failed-run")
        self.assertEqual(results["summary"]["status"], "failed")
        self.assertEqual(results["cases"][0]["expectation_status"], "failed")
        self.assertTrue(any("issue codes differ" in item for item in results["cases"][0]["failures"]))

    def test_private_case_scaffold_references_but_does_not_copy_source(self):
        source = self.root / "private-paper.docx"
        doc = Document()
        doc.add_paragraph("Private synthetic placeholder")
        doc.save(source)
        before = source.read_bytes()
        manifest, config = benchmark.prepare_private_case(
            source, "private_001", "student", self.root / "private-case"
        )
        value = benchmark.read_json(manifest)
        self.assertEqual(value["cases"][0]["source"], str(source.resolve()))
        self.assertEqual(benchmark.read_json(config)["review_status"], "pending")
        self.assertEqual(source.read_bytes(), before)
        self.assertFalse(any(path.suffix == ".docx" for path in manifest.parent.iterdir()))

    def test_failed_private_inspection_leaves_no_case_directory(self):
        source = self.root / "damaged.docx"
        source.write_bytes(b"not a Word package")
        destination = self.root / "failed-private-case"

        with self.assertRaises(Exception):
            benchmark.prepare_private_case(source, "damaged", "student", destination)

        self.assertFalse(destination.exists())

    def test_public_source_registry_verifies_exact_local_corpus_files(self):
        corpus = self.root / "corpus"
        corpus.mkdir()
        document = corpus / "public-paper.docx"
        document.write_bytes(b"public synthetic Word placeholder")
        data = document.read_bytes()
        registry = {
            "schema_version": 1,
            "sources": [{
                "id": "public_paper",
                "title": "Public Paper",
                "landing_page": "https://example.org/record",
                "download_url": "https://example.org/paper.docx",
                "license": "CC-BY-4.0",
                "license_url": "https://creativecommons.org/licenses/by/4.0/",
                "local_filename": document.name,
                "size_bytes": len(data),
                "md5": hashlib.md5(data).hexdigest(),
                "sha256": hashlib.sha256(data).hexdigest(),
                "intended_use": "Compatibility regression",
            }],
        }
        registry_path = benchmark.write_new_json(self.root / "public-sources.json", registry)

        result = benchmark.verify_public_sources(registry_path, corpus)

        self.assertEqual(result["status"], "passed")
        self.assertEqual(result["sources"][0]["status"], "passed")

    def test_public_source_registry_reports_checksum_drift(self):
        corpus = self.root / "corpus"
        corpus.mkdir()
        document = corpus / "public-paper.docx"
        document.write_bytes(b"changed download")
        registry = {
            "schema_version": 1,
            "sources": [{
                "id": "public_paper",
                "title": "Public Paper",
                "landing_page": "https://example.org/record",
                "download_url": "https://example.org/paper.docx",
                "license": "CC-BY-4.0",
                "license_url": "https://creativecommons.org/licenses/by/4.0/",
                "local_filename": document.name,
                "size_bytes": 1,
                "md5": "0" * 32,
                "sha256": "0" * 64,
                "intended_use": "Compatibility regression",
            }],
        }
        registry_path = benchmark.write_new_json(self.root / "public-sources.json", registry)

        result = benchmark.verify_public_sources(registry_path, corpus)

        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["sources"][0]["status"], "mismatch")

    def test_visual_case_cannot_pass_until_fresh_page_review_is_recorded(self):
        manifest = benchmark.generate_synthetic_suite(self.root / "fixtures")
        value = benchmark.read_json(manifest)
        value["cases"] = [value["cases"][0]]
        value["cases"][0]["require_visual_review"] = True
        visual_manifest = benchmark.write_new_json(self.root / "fixtures" / "visual-cases.json", value)
        renderer = self.root / "fake_renderer.py"
        renderer.write_text(
            """from pathlib import Path
import sys
out = Path(sys.argv[sys.argv.index('--output_dir') + 1])
out.mkdir(parents=True)
(out / 'page-1.png').write_bytes(b'synthetic rendered page')
(out / 'output.pdf').write_bytes(b'%PDF-1.4 synthetic')
""",
            encoding="utf-8",
        )
        results = benchmark.run_suite(visual_manifest, self.root / "visual-run", renderer)
        case = results["cases"][0]
        self.assertEqual(case["status"], "awaiting_visual_review")
        self.assertEqual(results["summary"]["status"], "needs_visual_review")
        template_path = self.root / "visual-run" / "cases" / case["id"] / "page-review.template.json"
        review = benchmark.read_json(template_path)
        for page in review["pages"]:
            page.update(status="pass", notes="Inspected the complete synthetic page at full size.")
        review_path = benchmark.write_new_json(self.root / "visual-review.json", review)
        qa, summary = benchmark.record_case_qa(self.root / "visual-run", case["id"], review_path)
        self.assertEqual(qa["status"], "VISUAL_REVIEW_RECORDED")
        self.assertEqual(summary["status"], "passed")


if __name__ == "__main__":
    unittest.main()
