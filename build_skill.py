"""Build an installable skill from the maintained local engines; no external code."""
from pathlib import Path
import argparse
import hashlib
import json
import re
from zipfile import ZipFile, ZIP_DEFLATED


def build(archive=None):
    root = Path(__file__).resolve().parent
    skill = root / "skills" / "apa7-word-formatter"
    sources = {
        "apa7_format.py": root / "apa7_format.py",
        "apa7_results.py": root / "apa7_results.py",
        "apa7_statistics.py": root / "apa7_statistics.py",
        "apa7_visuals.py": root / "apa7_visuals.py",
        "apa7_workflow.py": root / "apa7_workflow.py",
        "apa7_benchmark.py": root / "tools" / "apa7_benchmark.py",
        "requirements.txt": root / "requirements.txt",
    }
    hashes = {}
    for name, source in sources.items():
        data = source.read_bytes()
        (skill / "scripts" / name).write_bytes(data)
        hashes["scripts/" + name] = hashlib.sha256(data).hexdigest()
    engine_text = (root / "apa7_format.py").read_text(encoding="utf-8")
    version = re.search(r'^VERSION = "([^"]+)"$', engine_text, re.MULTILINE)
    if not version:
        raise RuntimeError("Could not read the formatter version")
    manifest = {"version": version.group(1), "engine_sha256": hashes}
    (skill / "build-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    if archive:
        archive = Path(archive)
        with ZipFile(archive, "x", ZIP_DEFLATED) as output:
            for path in sorted(skill.rglob("*")):
                if path.is_file() and "__pycache__" not in path.parts and path.name != ".DS_Store" and path.suffix != ".pyc":
                    output.write(path, Path(skill.name) / path.relative_to(skill))
        with ZipFile(archive) as verify:
            if verify.testzip():
                raise RuntimeError("Archive integrity check failed")
    return skill


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archive", type=Path)
    args = parser.parse_args()
    print(build(args.archive))
