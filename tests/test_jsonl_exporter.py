import json
import tempfile
from pathlib import Path
from kami_excel_extractor.jsonl_exporter import JsonlExporter


def test_to_jsonl_string():
    chunks = [
        {"content": "chunk 1", "metadata": {"id": 1}},
        {"content": "chunk 2", "metadata": {"id": 2}},
    ]
    jsonl_str = JsonlExporter.to_jsonl_string(chunks)
    lines = jsonl_str.strip().split("\n")

    assert len(lines) == 2
    data1 = json.loads(lines[0])
    assert data1["content"] == "chunk 1"
    assert data1["metadata"]["id"] == 1

    data2 = json.loads(lines[1])
    assert data2["content"] == "chunk 2"
    assert data2["metadata"]["id"] == 2


def test_export():
    chunks = [
        {"content": "chunk 1", "metadata": {"id": 1}},
    ]
    with tempfile.TemporaryDirectory() as tmp_dir:
        output_file = Path(tmp_dir) / "output.jsonl"
        JsonlExporter.export(chunks, output_file)

        assert output_file.exists()
        with open(output_file, "r", encoding="utf-8") as f:
            content = f.read()

        data = json.loads(content.strip())
        assert data["content"] == "chunk 1"
        assert data["metadata"]["id"] == 1
