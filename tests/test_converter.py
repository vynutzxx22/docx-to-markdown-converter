from pathlib import Path

from docx import Document

from app import convert_docx


def test_converts_headings_paragraphs_tables_and_links(tmp_path: Path):
    source = tmp_path / "sample.docx"
    output = tmp_path / "out"

    doc = Document()
    doc.add_heading("Main heading", level=1)
    doc.add_paragraph("A normal paragraph.")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Sample"
    table.cell(1, 1).text = "42"
    doc.save(source)

    result = convert_docx(source, output, extract_images=False)
    content = result.read_text(encoding="utf-8")

    assert "# Main heading" in content
    assert "A normal paragraph." in content
    assert "| Name | Value |" in content
    assert "| Sample | 42 |" in content


def test_extracts_embedded_image(tmp_path: Path):
    from PIL import Image

    image_path = tmp_path / "pixel.png"
    Image.new("RGB", (4, 4), "red").save(image_path)
    source = tmp_path / "with_image.docx"
    output = tmp_path / "out"

    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run().add_picture(str(image_path))
    doc.save(source)

    result = convert_docx(source, output, extract_images=True)
    content = result.read_text(encoding="utf-8")

    assert "![image_" in content
    media = output / "with_image_media"
    assert media.exists()
    assert list(media.iterdir())
