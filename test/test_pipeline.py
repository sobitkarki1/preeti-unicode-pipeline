from pathlib import Path

import pytest

import preeti_pipeline as pipeline


def test_heuristic_detects_pretti_and_rejects_english():
    assert pipeline.detect_text("g]kfn dfhUd", 0.7)[0]
    assert not pipeline.detect_text("This is ordinary English.", 0.7)[0]


def test_txt_conversion_preserves_nonlegacy_lines(tmp_path, monkeypatch):
    class FakeConverter:
        def text(self, value):
            return value.replace("g]kfn", "नेपाल")

    source = tmp_path / "sample.txt"
    source.write_text("g]kfn\nEnglish line\n", encoding="utf-8")
    result = pipeline.process(source, tmp_path / "out", tmp_path, 0.7, FakeConverter(), tmp_path / "tmp")
    assert result.status == "ok"
    assert (tmp_path / "out" / "sample.txt").read_text(encoding="utf-8") == "नेपाल\nEnglish line\n"


def test_fixture_matches_expected():
    fixture = Path(__file__).parent / "fixtures" / "sample_preti.txt"
    expected = (Path(__file__).parent / "fixtures" / "sample_unicode.txt").read_text(encoding="utf-8")
    assert fixture.read_text(encoding="utf-8").strip() == "g]kfn"
    assert expected.strip() == "नेपाल"


@pytest.mark.skipif(not pipeline.DEVANAGARI_RE, reason="sanity")
def test_validation_requires_devanagari(tmp_path):
    source = tmp_path / "source.txt"
    output = tmp_path / "output.txt"
    source.write_text("g]kfn", encoding="utf-8")
    output.write_text("नेपाल", encoding="utf-8")
    assert pipeline.validate(source, output) == []
