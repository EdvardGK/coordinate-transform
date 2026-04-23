"""Convert a Markdown file to PDF using Pandoc + headless Microsoft Edge.

Pandoc renders the .md to a self-contained HTML with a print-friendly stylesheet,
then Edge in headless mode prints the HTML to PDF. No LaTeX required.

Usage:
    python md_to_pdf.py input.md [output.pdf]
"""
import os, sys, subprocess, tempfile, shutil

PANDOC = r'C:\Program Files\Pandoc\pandoc.exe'
CHROME_CANDIDATES = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
]


def find_chrome():
    for c in CHROME_CANDIDATES:
        if os.path.exists(c):
            return c
    raise RuntimeError('No Chrome or Edge found for headless PDF rendering')

CSS = """
@page { size: A4; margin: 18mm 15mm; }
html { font-size: 11pt; }
body {
    font-family: -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: #1a1a1a;
    line-height: 1.5;
    max-width: none;
}
h1 { font-size: 22pt; border-bottom: 2px solid #333; padding-bottom: 6px; margin-top: 0; }
h2 { font-size: 16pt; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 24pt; }
h3 { font-size: 13pt; margin-top: 18pt; }
h4 { font-size: 11pt; margin-top: 12pt; }
p, li { font-size: 10.5pt; }
code {
    font-family: "Cascadia Mono", "Consolas", "Menlo", monospace;
    font-size: 9.5pt;
    background: #f4f4f4;
    padding: 1px 4px;
    border-radius: 3px;
}
pre { background: #f4f4f4; padding: 8px 12px; border-radius: 4px; overflow-x: auto; }
pre code { background: none; padding: 0; }
table {
    border-collapse: collapse;
    margin: 8pt 0;
    font-size: 9pt;
    width: 100%;
}
th, td {
    border: 1px solid #bbb;
    padding: 4px 8px;
    text-align: left;
    vertical-align: top;
}
th { background: #eee; font-weight: 600; }
tr:nth-child(even) td { background: #fafafa; }
hr { border: none; border-top: 1px solid #ccc; margin: 18pt 0; }
blockquote { border-left: 3px solid #ccc; padding-left: 12px; color: #555; margin: 8pt 0; }
"""


def md_to_html(md_path, html_path):
    """Pandoc: markdown -> standalone HTML with embedded CSS."""
    with tempfile.NamedTemporaryFile('w', suffix='.css', delete=False, encoding='utf-8') as f:
        f.write(CSS)
        css_path = f.name
    try:
        result = subprocess.run(
            [PANDOC, md_path,
             '-f', 'gfm',
             '-t', 'html5',
             '-s',
             '--metadata', f'title={os.path.splitext(os.path.basename(md_path))[0]}',
             '--embed-resources',
             '--resource-path', os.path.dirname(os.path.abspath(md_path)),
             '--css', css_path,
             '-o', html_path],
            capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode != 0:
            print('Pandoc stderr:', result.stderr)
            raise RuntimeError(f'Pandoc failed: {result.returncode}')
    finally:
        try:
            os.unlink(css_path)
        except OSError:
            pass


def html_to_pdf(html_path, pdf_path):
    """Chromium-based headless browser: HTML -> PDF."""
    chrome = find_chrome()
    abs_html = os.path.abspath(html_path).replace('\\', '/')
    file_uri = f'file:///{abs_html}'
    abs_pdf = os.path.abspath(pdf_path)
    with tempfile.TemporaryDirectory() as user_data:
        result = subprocess.run(
            [chrome,
             '--headless=new',
             '--disable-gpu',
             '--no-pdf-header-footer',
             f'--user-data-dir={user_data}',
             f'--print-to-pdf={abs_pdf}',
             file_uri],
            capture_output=True, text=True, timeout=120
        )
        if not os.path.exists(abs_pdf):
            print('Browser stdout:', result.stdout)
            print('Browser stderr:', result.stderr)
            raise RuntimeError('Headless browser did not produce a PDF')


def md_to_pdf(md_path, pdf_path=None):
    if pdf_path is None:
        pdf_path = os.path.splitext(md_path)[0] + '.pdf'
    with tempfile.TemporaryDirectory() as tmp:
        html_path = os.path.join(tmp, 'doc.html')
        md_to_html(md_path, html_path)
        html_to_pdf(html_path, pdf_path)
    print(f'PDF: {pdf_path}')
    return pdf_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    md = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    md_to_pdf(md, out)
