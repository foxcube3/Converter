# Universal File to Text Converter

A simple Flask web app that converts many file types into plain text.

Supported formats (best-effort):
- Text/Markdown/CSV/JSON/YAML
- PDF (pdfminer.six)
- DOCX (docx2txt)
- PPTX (python-pptx)
- XLSX (openpyxl)
- HTML (BeautifulSoup)
- EPUB (ebooklib + BeautifulSoup)
- MOBI/AZW3 (via Calibre's ebook-convert, if installed)
- RTF (striprtf)
- Images: PNG/JPG/TIFF/BMP/GIF via OCR (pytesseract)
- Audio/Video: MP3/WAV/M4A/MP4/WEBM/OGG via Whisper (optional)

Notes:
- OCR requires Tesseract to be installed on your system.
- Whisper may download a model on first run and benefits from FFmpeg for more formats.

## Setup

1) Create a virtual environment and install dependencies:
```
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

2) Optional system dependencies:
- Tesseract OCR (for image text extraction)
  - macOS: `brew install tesseract`
  - Ubuntu/Debian: `sudo apt-get install tesseract-ocr`
  - Windows: install from https://github.com/UB-Mannheim/tesseract/wiki
- FFmpeg (helps Whisper with more audio/video formats)
  - macOS: `brew install ffmpeg`
  - Ubuntu/Debian: `sudo apt-get install ffmpeg`
  - Windows: https://ffmpeg.org/download.html
- Calibre (for MOBI/AZW3/FB2 via ebook-convert)
  - macOS: `brew install --cask calibre` or download from https://calibre-ebook.com/download
  - Ubuntu/Debian: `sudo apt-get install calibre` (or use official binary installer)
  - Windows: download from https://calibre-ebook.com/download
  - After install, ensure `ebook-convert` is in your PATH
  - Server-side setting: set `PREFER_CALIBRE_TXT=true` (default) to prefer TXT conversion; set to `false` to prefer HTML parsing fallback

3) Run the app:
```
python app.py
```
Open http://localhost:5000 and upload a file. You'll get a downloadable `.txt` of the extracted content.

## Environment

- APP_SECRET_KEY: Optional Flask secret key
- WHISPER_MODEL: Whisper model name (e.g., tiny, small, medium). Default: `small`.

## Limitations

This is a best-effort extractor using pure-Python libraries where possible. Some formats or edge cases may need additional system tools or yield partial results.