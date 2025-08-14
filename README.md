# Patent Text Extractor

Utilities for turning a patent PDF into a set of images, extracting text from
those images and injecting the results into an already opened PowerPoint
presentation.  The main entry point is the `patent_text_extractor.py` script
which can be run directly:

```bash
python patent_text_extractor.py
```

The application will ask the user to select a PDF file.  It requires the
following third party tools to be available on the system:

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [Poppler](https://poppler.freedesktop.org/) for `pdf2image`
- Microsoft PowerPoint with the `win32com` Python package (Windows only)

The extracted images and text files are written to an `images/` directory.
