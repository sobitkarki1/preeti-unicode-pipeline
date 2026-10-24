"""Batch conversion of legacy Nepali font documents to Unicode Devanagari."""

from __future__ import annotations

import argparse
import csv
import html
import importlib
import logging
import random
import re
import shutil
import sys
import tempfile
import urllib.request
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.parse import urlparse

SUPPORTED = {".docx", ".txt", ".pdf", ".html", ".htm"}
LEGACY_FONTS = {"preeti", "sagarmatha", "kalimati"}
DEVANAGARI_RE = re.compile(r"[\u0900-\u097f]")
LEGACY_NGRAMS = (
    "g]kfn", "dfhUd", "g]kf]", "s+u", "k|", "cf", "If", "q", "z|", "lj",
)
STRONG_NGRAMS = {"g]kfn", "g]kf]", "dfhUd", "s+u"}
# These are deliberately conservative: ordinary English must not be treated as Preeti.
LEGACY_RE = re.compile("|".join(re.escape(value) for value in (
    "g]kfn", "g]kf]", "dfhUd", "s+u", "z|", "k|", "b\\",
)))
PDF_LEGACY_RE = re.compile("|".join(re.escape(value) for value in (
    "g]kfn", "dfhUd", ";DaGwL", "sfof{", "vf]Ng]", "kmf/fd",
    "k|i6", "lgb{", ";+:yf", "cjl:yt", "k|j4{g", "pmhf{",
    "l;+rfO", "dGqfno", "nfu\"", "Aff]8{", "lg0f{o", "k|:tfjgf",
    "Joj:yf", "sfg\"gL", "agfO{", "ug]{", "ahf/nfO{",
)))


@dataclass
class Result:
    source: str
    output: str | None = None
    status: str = "failed"
    reason: str = ""
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ManifestSource:
    value: str
    filename: str | None = None


class Converter:
    """Small adapter around npttf2utf, keeping its package API in one place."""

    def __init__(self) -> None:
        try:
            package = importlib.import_module("npttf2utf")
            self.mapper = package.FontMapper(str(Path(package.__file__).with_name("map.json")))
            self.txt_handler = package.TxtHandler(str(Path(package.__file__).with_name("map.json")))
            self.docx_handler = package.DocxHandler(str(Path(package.__file__).with_name("map.json")))
        except (ImportError, OSError, AttributeError) as exc:
            raise RuntimeError(
                "npttf2utf is required. Install dependencies with: python -m pip install -r requirements.txt"
            ) from exc

    def text(self, value: str) -> str:
        return self.mapper.map_to_unicode(value, from_font="Preeti")

    def docx(self, source: Path, target: Path) -> None:
        self.docx_handler.map_fonts(
            original_file_path=str(source),
            output_file_path=str(target),
            from_font="auto",
            to_font="unicode",
        )

    def pdf(self, source: Path, target: Path) -> None:
        try:
            with importlib.import_module("pdfplumber").open(source) as document:
                pages = []
                for page in document.pages:
                    raw = page.extract_text() or ""
                    pages.append("\n".join(
                        self.text(line)
                        if PDF_LEGACY_RE.search(line) or re.search(r"[\[\]{};|`~]", line)
                        else line
                        for line in raw.splitlines()
                    ))
            target.write_text("\n\n".join(pages), encoding="utf-8")
        except (ImportError, OSError, ValueError):
            try:
                module = importlib.import_module("preeti_unicode")
                ok = module.file_converter(
                    input_file=str(source), input_format="pdf",
                    output_file=str(target), output_format="txt",
                )
                if ok is False:
                    raise RuntimeError("preeti_unicode reported PDF conversion failure")
            except ImportError as exc:
                raise RuntimeError(
                    "PDF conversion requires pdfplumber or preeti-unicode-converter; "
                    "install requirements.txt"
                ) from exc


def detect_text(value: str, threshold: float) -> tuple[bool, float]:
    compact = re.sub(r"\s+", "", value)
    if not compact:
        return False, 0.0
    hits = sum(compact.count(ngram) for ngram in LEGACY_NGRAMS)
    strong_hits = sum(compact.count(ngram) for ngram in STRONG_NGRAMS)
    score = min(1.0, max(strong_hits, hits / 2))
    return score >= threshold, score


def detect_docx(path: Path) -> tuple[bool, float]:
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError("DOCX detection requires python-docx") from exc
    document = Document(path)
    fonts: set[str] = set()
    text_parts: list[str] = []
    for paragraph in document.paragraphs:
        text_parts.append(paragraph.text)
        fonts.update(
            run.font.name.lower() for run in paragraph.runs
            if run.font.name
        )
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text_parts.append(paragraph.text)
                    fonts.update(
                        run.font.name.lower() for run in paragraph.runs
                        if run.font.name
                    )
    matched = any(any(font == name or name in font for name in LEGACY_FONTS) for font in fonts)
    _, score = detect_text("\n".join(text_parts), 0.0)
    return matched or score > 0, 1.0 if matched else score


def extract_pdf_text(path: Path) -> str:
    try:
        import pdfplumber
    except ImportError as exc:
        raise RuntimeError("PDF detection requires pdfplumber") from exc
    with pdfplumber.open(path) as document:
        return "\n".join(page.extract_text() or "" for page in document.pages)


class HtmlConverter(HTMLParser):
    def __init__(self, converter: Converter, threshold: float) -> None:
        super().__init__(convert_charrefs=False)
        self.converter, self.threshold = converter, threshold
        self.stack: list[bool] = []
        self.parts: list[str] = []

    @property
    def in_legacy(self) -> bool:
        return any(self.stack)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_dict = dict(attrs)
        style = (attrs_dict.get("style") or "").lower()
        marked = "preeti" in (attrs_dict.get("face") or "").lower() or "preeti" in style
        self.stack.append(marked)
        raw = self.get_starttag_text() or f"<{tag}>"
        self.parts.append(raw)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.parts.append(self.get_starttag_text() or f"<{tag}/>")

    def handle_endtag(self, tag: str) -> None:
        self.parts.append(f"</{tag}>")
        if self.stack:
            self.stack.pop()

    def handle_data(self, data: str) -> None:
        marked, _ = detect_text(data, self.threshold)
        self.parts.append(self.converter.text(data) if self.in_legacy or marked else data)

    def handle_entityref(self, name: str) -> None:
        self.parts.append(f"&{name};")

    def handle_charref(self, name: str) -> None:
        self.parts.append(f"&#{name};")


def validate(source: Path, output: Path, sample_count: int = 5) -> list[str]:
    text = output.read_text(encoding="utf-8", errors="replace") if output.suffix.lower() in {".txt", ".html", ".htm"} else ""
    if output.suffix.lower() == ".docx":
        from docx import Document
        text = "\n".join(p.text for p in Document(output).paragraphs)
    if output.suffix.lower() == ".pdf":
        text = output.read_text(encoding="utf-8", errors="replace")
    warnings: list[str] = []
    if not DEVANAGARI_RE.search(text):
        warnings.append("output contains no Devanagari characters")
    if LEGACY_RE.search(text):
        warnings.append("output still contains a likely Preeti sequence")
    lines = [line for line in text.splitlines() if line.strip()]
    if lines:
        random.Random(str(source)).sample(lines, min(sample_count, len(lines)))
    return warnings


def iter_sources(input_value: Path | None, manifest: Path | None) -> Iterable[Path | str]:
    if input_value:
        if input_value.is_dir():
            yield from (p for p in input_value.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED)
        elif input_value.is_file():
            yield input_value
        else:
            raise FileNotFoundError(input_value)
    if manifest:
        with manifest.open(newline="", encoding="utf-8-sig") as stream:
            if manifest.suffix.lower() == ".csv":
                for row in csv.DictReader(stream):
                    value = row.get("path") or row.get("url") or row.get("source") or row.get("source_url")
                    if value:
                        yield ManifestSource(value, row.get("filename") or None)
            else:
                try:
                    import yaml
                except ImportError as exc:
                    raise RuntimeError("YAML manifests require PyYAML") from exc
                data = yaml.safe_load(stream) or []
                data = data.get("sources", data) if isinstance(data, dict) else data
                for item in data:
                    if isinstance(item, dict):
                        value = item.get("path") or item.get("url") or item.get("source") or item.get("source_url")
                        yield ManifestSource(value, item.get("filename") or None) if value else None
                    else:
                        yield item


def localize(source: Path | str, temp_dir: Path) -> tuple[Path, str]:
    value = str(source)
    if urlparse(value).scheme in {"http", "https"}:
        target = temp_dir / Path(urlparse(value).path).name
        urllib.request.urlretrieve(value, target)
        return target, value
    return Path(value), value


def process(source: Path | str, output_root: Path, input_root: Path | None,
            threshold: float, converter: Converter, temp_dir: Path) -> Result:
    requested_name = source.filename if isinstance(source, ManifestSource) else None
    source_value = source.value if isinstance(source, ManifestSource) else source
    display = str(source_value)
    try:
        local, display = localize(source_value, temp_dir)
    except (OSError, ValueError) as exc:
        return Result(display, status="failed", reason=f"download failed: {exc}")
    suffix = local.suffix.lower()
    if suffix not in SUPPORTED:
        return Result(display, status="skipped", reason="unsupported extension")
    relative = requested_name or (local.name if input_root is None else str(local.relative_to(input_root)))
    target = output_root / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        if suffix == ".docx":
            detected, _ = detect_docx(local)
        elif suffix in {".txt", ".html", ".htm"}:
            sample = local.read_text(encoding="utf-8", errors="replace")
            detected, _ = detect_text(re.sub(r"<[^>]+>", " ", sample), threshold)
        elif suffix == ".pdf":
            sample = extract_pdf_text(local)
            pdf_hits = len(PDF_LEGACY_RE.findall(sample))
            detected = min(1.0, pdf_hits / 2) >= threshold
        else:
            detected = True
        if not detected:
            return Result(display, status="skipped", reason="already Unicode or not detected as Preeti")
        if suffix == ".docx":
            converter.docx(local, target)
        elif suffix == ".pdf":
            target = target.with_suffix(".pdf.txt")
            converter.pdf(local, target)
        elif suffix in {".html", ".htm"}:
            parser = HtmlConverter(converter, threshold)
            parser.feed(local.read_text(encoding="utf-8", errors="replace"))
            target.write_text("".join(parser.parts), encoding="utf-8")
        else:
            lines = local.read_text(encoding="utf-8", errors="replace").splitlines(keepends=True)
            target.write_text("".join(converter.text(line) if detect_text(line, threshold)[0] else line for line in lines), encoding="utf-8")
        warnings = validate(local, target)
        return Result(display, str(target), "ok", "", warnings)
    except (OSError, RuntimeError, ValueError, ImportError) as exc:
        return Result(display, str(target), "failed", str(exc))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--input", type=Path, help="file or directory to scan recursively")
    group.add_argument("--manifest", type=Path, help="CSV/YAML manifest with path, url, or source values")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--threshold", type=float, default=0.7)
    parser.add_argument("--log", type=Path, help="optional log file")
    args = parser.parse_args(argv)
    if not 0 <= args.threshold <= 1:
        parser.error("--threshold must be between 0 and 1")
    args.output.mkdir(parents=True, exist_ok=True)
    if args.log:
        args.log.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s",
                        filename=str(args.log) if args.log else None)
    converter = Converter()
    results: list[Result] = []
    with tempfile.TemporaryDirectory(prefix="preeti-pipeline-") as temp:
        for source in iter_sources(args.input, args.manifest):
            result = process(source, args.output, args.input if args.input and args.input.is_dir() else None,
                             args.threshold, converter, Path(temp))
            results.append(result)
            logging.info("source=%s output=%s status=%s warnings=%s", result.source, result.output, result.status, result.warnings)
    summary = {
        "total": len(results),
        "converted_ok": sum(r.status == "ok" and not r.warnings for r in results),
        "converted_with_warnings": sum(r.status == "ok" and bool(r.warnings) for r in results),
        "failed": sum(r.status == "failed" for r in results),
        "skipped": sum(r.status == "skipped" for r in results),
    }
    print(summary)
    for result in results:
        if result.status != "ok" or result.warnings:
            print(f"{result.status}: {result.source}: {result.reason or '; '.join(result.warnings)}")
    return 1 if summary["failed"] else 0


if __name__ == "__main__":
    sys.exit(main())
