"""Patent Text Extractor utilities.

This module converts patent PDF files to images, extracts text from each
page, parses the relevant patent metadata and injects the results into an
already opened PowerPoint presentation.  The implementation is intentionally
modular so individual steps can be reused in other workflows.

The code relies on ``pdf2image`` for PDF rasterisation and ``pytesseract``
for optical character recognition.  PowerPoint automation requires the
``win32com`` package and thus only works on Windows systems with Microsoft
Office installed.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from tqdm import tqdm

try:  # Windows only dependency
    import win32com.client as win32
except ImportError:  # pragma: no cover - not available on all platforms
    win32 = None  # type: ignore

import tkinter as tk
from tkinter import filedialog


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class PatentInfo:
    """Container for parsed patent metadata."""

    title: str
    number: str
    date: str
    inventors: str
    abstract: str


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def get_tesseract_path() -> str:
    """Return the platform specific path to the ``tesseract`` executable."""

    if getattr(sys, "frozen", False):
        # PyInstaller bundle location
        return os.path.join(sys._MEIPASS, "tesseract.exe")  # type: ignore[attr-defined]
    return r"C:/Program Files/Tesseract-OCR/tesseract.exe"


def show_progress(iterable: Iterable, description: str) -> Iterable:
    """Wrap an iterable with a :class:`tqdm.tqdm` progress bar."""

    return tqdm(iterable, desc=description, unit="step")


def select_pdf_file() -> Path | None:
    """Open a file chooser dialog and return the selected PDF path."""

    root = tk.Tk()
    root.withdraw()
    filename = filedialog.askopenfilename(
        title="Select a PDF Patent File", filetypes=[("PDF Files", "*.pdf")]
    )
    return Path(filename) if filename else None


def extract_text_from_image(image: Image.Image) -> str:
    """Extract text from a single image using ``pytesseract``."""

    return pytesseract.image_to_string(image)


def split_image(image: Image.Image) -> Tuple[Image.Image, Image.Image]:
    """Split an image into two equally sized columns."""

    width, height = image.size
    column_width = width // 2
    left_bbox = (0, 0, column_width, height)
    right_bbox = (column_width, 0, width, height)
    return image.crop(left_bbox), image.crop(right_bbox)


def save_image(image: Image.Image, path: Path) -> None:
    """Persist an image to ``path``."""

    image.save(path)


def convert_pdf_to_images(pdf_path: Path, output_dir: Path) -> Tuple[List[Path], List[str]]:
    """Convert the pages of *pdf_path* to images and extract their text."""

    if hasattr(sys, "_MEIPASS"):
        poppler_path = Path(sys._MEIPASS) / "poppler"  # type: ignore[attr-defined]
    else:
        poppler_path = Path("path/to/your/poppler")

    dpi_value = 600
    images = convert_from_path(pdf_path, poppler_path=str(poppler_path), dpi=dpi_value)
    image_paths: List[Path] = []
    texts: List[str] = []

    for i, image in show_progress(enumerate(images, start=1), "Converting PDF"):
        full_image_path = output_dir / f"page_{i}.png"
        save_image(image, full_image_path)

        left, right = split_image(image)
        save_image(left, output_dir / f"page_{i}_col1.png")
        save_image(right, output_dir / f"page_{i}_col2.png")

        combined_text = f"{extract_text_from_image(left)}\n\n{extract_text_from_image(right)}"
        (output_dir / f"page_{i}.txt").write_text(combined_text, encoding="utf-8")

        image_paths.append(full_image_path)
        texts.append(combined_text)

    return image_paths, texts


def parse_patent_info(text: str) -> PatentInfo:
    """Extract metadata from raw patent text."""

    title_match = re.search(r"\(54\)\s*([\s\S]+?)(?=\(\d{2}\)|\n{2,}|\Z)", text, re.IGNORECASE)
    title = title_match.group(1).strip().replace("\n", " ") if title_match else "Title N/A"
    title = title.lstrip(")")

    number_match = re.search(r"US\s\d{1,3},\d{3},\d{3}\s\w\d", text)
    number = number_match.group(0) if number_match else "PATENT # N/A"

    date_match = re.search(
        r"\(45\)\s*Date of Patent:\s*(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.\s\d{1,2},\s\d{4}\b)",
        text,
    )
    date = date_match.group(1) if date_match else "Date N/A"

    inventors_match = re.search(r"Inventors?:\s*([\s\S]+?)(?=\(\d{2}\)|\n\n|\Z)", text, re.IGNORECASE)
    inventors = (
        inventors_match.group(1).strip().replace("\n", " ")
        if inventors_match
        else "Inventors N/A"
    )

    abstract_match = re.search(r"Abstract:?\s*([\s\S]+?)(?=\n\n|\Z)", text, re.IGNORECASE)
    abstract = abstract_match.group(1).strip() if abstract_match else "Abstract N/A"

    return PatentInfo(title, number, date, inventors, abstract)


def modify_active_ppt(images: Sequence[Path], texts: Sequence[str]) -> None:
    """Insert *images* and *texts* into the active PowerPoint presentation."""

    if win32 is None:
        raise RuntimeError("win32com is required for PowerPoint automation")

    powerpoint = win32.Dispatch("PowerPoint.Application")
    presentation = powerpoint.ActivePresentation

    for slide_index in range(presentation.Slides.Count, 1, -1):
        presentation.Slides(slide_index).Delete()

    for image_path, text in show_progress(zip(images, texts), "Updating slides"):
        slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12)
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = 0x1C1C1C

        info = parse_patent_info(text)

        picture = slide.Shapes.AddPicture(
            str(image_path), LinkToFile=False, SaveWithDocument=True,
            Left=500, Top=50, Width=374.4, Height=459.6
        )
        picture.name = "patent_image"

        title_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=40, Width=400, Height=50)
        title_shape.TextFrame.TextRange.Text = info.title
        title_shape.TextFrame.TextRange.Font.Bold = True
        title_shape.TextFrame.TextRange.Font.Size = 16
        title_shape.TextFrame.TextRange.Font.Name = "Calibri"
        title_shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

        patent_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=120, Width=400, Height=30)
        patent_shape.TextFrame.TextRange.Text = f"PATENT #: {info.number}     {info.date}"
        patent_shape.TextFrame.TextRange.Font.Size = 16
        patent_shape.TextFrame.TextRange.Font.Name = "Calibri"
        patent_shape.TextFrame.TextRange.Font.Color.RGB = 0xD3D3D3
        text_range = patent_shape.TextFrame.TextRange
        start_pos = text_range.Text.find(info.number)
        if start_pos != -1:
            text_range.Characters(start_pos + 1, len(info.number)).Font.Bold = True

        inventors_shape = slide.Shapes.AddTextbox(Orientation=1, Left=50, Top=160, Width=400, Height=60)
        inventors_shape.TextFrame.TextRange.Text = f"INVENTORS: {info.inventors}"
        inventors_shape.TextFrame.TextRange.Font.Size = 14
        inventors_shape.TextFrame.TextRange.Font.Name = "Calibri"
        inventors_shape.TextFrame.TextRange.Font.Color.RGB = 0xD3D3D3
        inventors_start = len("INVENTORS: ") + 1
        inventors_shape.TextFrame.TextRange.Characters(
            inventors_start, len(info.inventors)
        ).Font.Bold = True

        line_top = inventors_shape.Top + inventors_shape.Height + 10
        line = slide.Shapes.AddLine(BeginX=50, BeginY=line_top, EndX=450, EndY=line_top)
        line.Line.ForeColor.RGB = 0x888888
        line.Line.Weight = 0.75

    presentation.Save()


def main() -> None:
    """Entry point for command line usage."""

    output_dir = Path("images")
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir()

    pdf_path = select_pdf_file()
    if not pdf_path:
        print("No file selected. Exiting.")
        return

    image_paths, texts = convert_pdf_to_images(pdf_path, output_dir)
    modify_active_ppt(image_paths, texts)


if __name__ == "__main__":  # pragma: no cover - script entry point
    main()
