from __future__ import annotations

import os
import re
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

APP_TITLE = "DOCX to Markdown Converter"
APP_VERSION = "1.0.0"


def clean_text(value: str) -> str:
    return re.sub(r"[ \t]+", " ", value.replace("\xa0", " ")).strip()


def escape_table_cell(value: str) -> str:
    return clean_text(value).replace("\\", "\\\\").replace("|", r"\|").replace("\n", " ")


def iter_document_blocks(parent):
    """Yield paragraphs and tables in the order they occur in the DOCX body."""
    if hasattr(parent, "element") and hasattr(parent.element, "body"):
        element = parent.element.body
    elif hasattr(parent, "_tc"):
        element = parent._tc
    else:
        element = parent
    for child in element.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def paragraph_markdown_text(paragraph: Paragraph) -> str:
    """Preserve text and external hyperlinks in paragraph order."""
    rels = paragraph.part.rels
    parts: list[str] = []
    for child in paragraph._p.iterchildren():
        if child.tag == qn("w:r"):
            value = "".join(node.text or "" for node in child.iter() if node.tag == qn("w:t"))
            if value:
                parts.append(value)
        elif child.tag == qn("w:hyperlink"):
            value = "".join(node.text or "" for node in child.iter() if node.tag == qn("w:t"))
            rid = child.get(qn("r:id"))
            target = rels[rid].target_ref if rid and rid in rels else ""
            if value and target:
                parts.append(f"[{value}]({target})")
            elif value:
                parts.append(value)
    return clean_text("".join(parts))


def extract_paragraph_images(paragraph: Paragraph, media_dir: Path, media_prefix: str) -> list[str]:
    links: list[str] = []
    for blip in paragraph._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        try:
            image_part = paragraph.part.related_parts[rid]
            content_type = image_part.content_type.lower()
            ext = content_type.split("/")[-1]
            ext = {"jpeg": "jpg", "tiff": "tif", "x-emf": "emf", "x-wmf": "wmf"}.get(ext, ext)
            if not re.fullmatch(r"[a-z0-9]+", ext):
                ext = "bin"
            media_dir.mkdir(parents=True, exist_ok=True)
            filename = f"image_{len(list(media_dir.iterdir())) + 1:03d}.{ext}"
            (media_dir / filename).write_bytes(image_part.blob)
            links.append(f"![{filename}]({media_prefix}/{filename})")
        except Exception:
            continue
    return links


def list_prefix(paragraph: Paragraph) -> tuple[int, str] | None:
    """Return indentation level and Markdown marker for a Word list paragraph."""
    ppr = paragraph._p.pPr
    if ppr is None or ppr.numPr is None:
        return None
    ilvl = ppr.numPr.ilvl
    level = int(ilvl.val) if ilvl is not None else 0
    num_id_node = ppr.numPr.numId
    num_id = int(num_id_node.val) if num_id_node is not None else None
    marker = "-"
    try:
        numbering = paragraph.part.numbering_part.element
        for num in numbering.findall(qn("w:num")):
            if int(num.get(qn("w:numId"))) == num_id:
                abstract_id = num.find(qn("w:abstractNumId"))
                if abstract_id is not None:
                    for abstract in numbering.findall(qn("w:abstractNum")):
                        if abstract.get(qn("w:abstractNumId")) == abstract_id.get(qn("w:val")):
                            for level_node in abstract.findall(qn("w:lvl")):
                                if int(level_node.get(qn("w:ilvl"))) == level:
                                    fmt = level_node.find(qn("w:numFmt"))
                                    if fmt is not None and fmt.get(qn("w:val")) != "bullet":
                                        marker = "1."
                                    break
                break
    except Exception:
        pass
    return level, marker


def table_to_markdown(table: Table) -> str:
    rows: list[list[str]] = []
    for row in table.rows:
        cells = [escape_table_cell(" ".join(paragraph_markdown_text(p) for p in cell.paragraphs)) for cell in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    width = max(map(len, rows))
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = [
        "| " + " | ".join(rows[0]) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(lines)


def convert_docx(source: str | Path, output_dir: str | Path, extract_images: bool = True) -> Path:
    source = Path(source)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    document = Document(str(source))
    media_folder_name = f"{source.stem}_media"
    media_dir = output_dir / media_folder_name
    output: list[str] = []

    for block in iter_document_blocks(document):
        if isinstance(block, Paragraph):
            text = paragraph_markdown_text(block)
            images = extract_paragraph_images(block, media_dir, media_folder_name) if extract_images else []
            style = (block.style.name or "").lower()
            if style == "title":
                rendered = f"# {text}" if text else ""
            elif style.startswith("heading "):
                match = re.search(r"(\d+)$", style)
                level = min(int(match.group(1)), 6) if match else 1
                rendered = f"{'#' * level} {text}" if text else ""
            else:
                rendered = text

            marker = list_prefix(block)
            if marker and rendered:
                level, bullet = marker
                rendered = f"{'  ' * level}{bullet} {rendered}"
            if rendered:
                output.append(rendered)
            output.extend(images)
            if rendered or images:
                output.append("")
        elif isinstance(block, Table):
            table_md = table_to_markdown(block)
            if table_md:
                output.extend([table_md, ""])

    target = output_dir / f"{source.stem}.md"
    target.write_text("\n".join(output).strip() + "\n", encoding="utf-8")
    return target


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} {APP_VERSION}")
        self.root.geometry("850x610")
        self.root.minsize(700, 500)
        self.files: list[Path] = []
        self.dark = True
        self.busy = False
        self._build()
        self._apply_theme()

    def _build(self):
        self.main = ttk.Frame(self.root, padding=22)
        self.main.pack(fill="both", expand=True)

        header = ttk.Frame(self.main)
        header.pack(fill="x")
        self.heading = ttk.Label(header, text="DOCX  →  Markdown", style="Heading.TLabel")
        self.heading.pack(side="left")
        self.theme_button = ttk.Button(header, text="Light mode", command=self._toggle_theme)
        self.theme_button.pack(side="right")

        self.subtitle = ttk.Label(self.main, text="Convert Word documents while keeping useful document structure.")
        self.subtitle.pack(anchor="w", pady=(5, 18))

        panel = ttk.Frame(self.main, style="Panel.TFrame", padding=16)
        panel.pack(fill="x")
        ttk.Label(panel, text="Select Word documents", style="PanelHeading.TLabel").pack(anchor="w")
        ttk.Label(panel, text="Add individual .docx files or select a folder for batch conversion.").pack(anchor="w", pady=(4, 12))
        actions = ttk.Frame(panel, style="Panel.TFrame")
        actions.pack(anchor="w")
        ttk.Button(actions, text="Add DOCX files", command=self._add_files).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Add folder", command=self._add_folder).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Clear list", command=self._clear).pack(side="left")

        row = ttk.Frame(self.main)
        row.pack(fill="x", pady=(18, 6))
        ttk.Label(row, text="FILES QUEUED").pack(side="left")
        self.count_label = ttk.Label(row, text="0 files")
        self.count_label.pack(side="right")
        self.file_list = tk.Listbox(self.main, height=8, relief="flat", borderwidth=0, highlightthickness=1,
                                    selectmode="extended", font=("Segoe UI", 10))
        self.file_list.pack(fill="both", expand=True)

        ttk.Label(self.main, text="OUTPUT FOLDER").pack(anchor="w", pady=(14, 5))
        output_row = ttk.Frame(self.main)
        output_row.pack(fill="x")
        self.output_var = tk.StringVar()
        ttk.Entry(output_row, textvariable=self.output_var).pack(side="left", fill="x", expand=True)
        ttk.Button(output_row, text="Browse", command=self._choose_output).pack(side="left", padx=(8, 0))

        self.images_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.main, text="Extract images into a companion media folder", variable=self.images_var).pack(anchor="w", pady=(10, 12))
        self.progress = ttk.Progressbar(self.main, mode="determinate")
        self.progress.pack(fill="x")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.main, textvariable=self.status_var).pack(anchor="w", pady=(6, 12))

        footer = ttk.Frame(self.main)
        footer.pack(fill="x")
        ttk.Button(footer, text="Open output folder", command=self._open_output).pack(side="right", padx=(0, 8))
        self.convert_button = ttk.Button(footer, text="Convert documents", style="Accent.TButton", command=self._start)
        self.convert_button.pack(side="right")
        ttk.Label(self.main, text="Complex Word layouts and merged table cells may require manual review.", style="Muted.TLabel").pack(anchor="w", pady=(14, 0))

    def _apply_theme(self):
        style = ttk.Style()
        style.theme_use("clam")
        if self.dark:
            bg, panel, fg, muted, entry, select = "#111827", "#1F2937", "#F3F4F6", "#9CA3AF", "#0B1220", "#334155"
        else:
            bg, panel, fg, muted, entry, select = "#F4F6FA", "#FFFFFF", "#172033", "#64748B", "#FFFFFF", "#DCE7F5"
        self.root.configure(bg=bg)
        style.configure(".", background=bg, foreground=fg, font=("Segoe UI", 10))
        style.configure("TFrame", background=bg)
        style.configure("Panel.TFrame", background=panel)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("Heading.TLabel", font=("Segoe UI Semibold", 23), background=bg, foreground=fg)
        style.configure("PanelHeading.TLabel", font=("Segoe UI Semibold", 12), background=panel, foreground=fg)
        style.configure("Muted.TLabel", foreground=muted)
        style.configure("TButton", padding=(12, 8), background=panel, foreground=fg, borderwidth=0)
        style.map("TButton", background=[("active", select)])
        style.configure("Accent.TButton", padding=(16, 10), background="#2563EB", foreground="#FFFFFF")
        style.map("Accent.TButton", background=[("active", "#1D4ED8")])
        style.configure("TEntry", fieldbackground=entry, foreground=fg, insertcolor=fg)
        style.configure("TCheckbutton", background=bg, foreground=fg)
        style.configure("Horizontal.TProgressbar", troughcolor=panel, background="#3B82F6")
        self.file_list.configure(bg=entry, fg=fg, selectbackground=select, selectforeground=fg,
                                 highlightbackground=panel, highlightcolor="#3B82F6")
        self.theme_button.configure(text="Light mode" if self.dark else "Dark mode")

    def _toggle_theme(self):
        self.dark = not self.dark
        self._apply_theme()

    def _add_files(self):
        selected = filedialog.askopenfilenames(filetypes=[("Word documents", "*.docx")])
        self._add_paths(selected)

    def _add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self._add_paths([str(p) for p in Path(folder).glob("*.docx")])

    def _add_paths(self, paths):
        known = {str(p.resolve()).lower() for p in self.files}
        for value in paths:
            path = Path(value)
            if path.suffix.lower() == ".docx" and str(path.resolve()).lower() not in known:
                self.files.append(path)
                known.add(str(path.resolve()).lower())
        self._refresh()

    def _refresh(self):
        self.file_list.delete(0, "end")
        for path in self.files:
            self.file_list.insert("end", path.name)
        self.count_label.configure(text=f"{len(self.files)} files")

    def _clear(self):
        self.files.clear()
        self._refresh()
        self.status_var.set("File list cleared.")

    def _choose_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_var.set(folder)

    def _start(self):
        if self.busy:
            return
        if not self.files:
            messagebox.showwarning(APP_TITLE, "Add at least one DOCX file first.")
            return
        output = self.output_var.get().strip() or str(self.files[0].parent / "Markdown_Output")
        self.output_var.set(output)
        self.busy = True
        self.convert_button.configure(state="disabled")
        self.progress.configure(maximum=len(self.files), value=0)
        threading.Thread(target=self._worker, args=(list(self.files), Path(output), self.images_var.get()), daemon=True).start()

    def _worker(self, files, output, extract_images):
        succeeded, failed = [], []
        for index, path in enumerate(files, start=1):
            try:
                succeeded.append(convert_docx(path, output, extract_images).name)
            except Exception as error:
                failed.append(f"{path.name}: {error}")
            self.root.after(0, self.progress.configure, {"value": index})
            self.root.after(0, self.status_var.set, f"Processed {index} of {len(files)}...")
        self.root.after(0, self._finish, succeeded, failed)

    def _finish(self, succeeded, failed):
        self.busy = False
        self.convert_button.configure(state="normal")
        summary = f"Finished. {len(succeeded)} converted, {len(failed)} failed."
        self.status_var.set(summary)
        details = "\n".join(succeeded[:12])
        if failed:
            details += "\n\nErrors:\n" + "\n".join(failed[:8])
            messagebox.showwarning("Conversion finished", summary + "\n\n" + details)
        else:
            messagebox.showinfo("Conversion finished", summary + ("\n\n" + details if details else ""))

    def _open_output(self):
        folder = Path(self.output_var.get().strip()) if self.output_var.get().strip() else (
            self.files[0].parent / "Markdown_Output" if self.files else None
        )
        if folder and folder.exists():
            os.startfile(folder)
        else:
            messagebox.showinfo(APP_TITLE, "The output folder does not exist yet.")


if __name__ == "__main__":
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()
