"""Simple DOCX to PDF converter using Chrome."""
import sys
import html
import subprocess
import tempfile
from pathlib import Path
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

def convert(docx_path, pdf_path):
    doc = docx.Document(docx_path)

    html_parts = ['''<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
@page { size: A4; margin: 2.54cm; }
body { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.4; max-width: 21cm; margin: 0 auto; padding: 2cm; }
h1 { font-size: 14pt; margin-top: 18pt; margin-bottom: 12pt; }
h2 { font-size: 13pt; margin-top: 14pt; margin-bottom: 10pt; }
p { margin: 6pt 0; text-align: justify; }
table { border-collapse: collapse; width: 100%; margin: 12pt 0; font-size: 10pt; }
th, td { border: 1px solid #000; padding: 5px 6px; overflow-wrap: break-word; }
th { background: #D9E2F3; font-weight: bold; }
.center { text-align: center; }
.code { font-family: 'Courier New', monospace; font-size: 8pt; background: #f5f5f5;
        padding: 8px; margin: 6pt 0; white-space: pre-wrap; border: 1px solid #ddd; }
.page-break { page-break-after: always; }
.pb { page-break-before: always; }
ul, ol { margin: 6pt 0; padding-left: 20pt; }
li { margin: 3pt 0; }
</style></head><body>
''']

    def get_text(para):
        text = ''
        for run in para.runs:
            t = html.escape(run.text)
            if run.bold:
                t = f'<strong>{t}</strong>'
            if run.italic:
                t = f'<em>{t}</em>'
            text += t
        return text

    def is_code(para):
        for run in para.runs:
            if run.font.name and 'Courier' in run.font.name:
                return True
        return False

    in_code = False
    code_lines = []

    pending_break = False

    def emit(chunk):
        nonlocal pending_break
        if pending_break:
            chunk = f'<div class="pb">{chunk}</div>'
            pending_break = False
        html_parts.append(chunk)

    for element in doc.element.body:
        if element.tag.endswith('}p'):
            para = Paragraph(element, doc)
            text = get_text(para)

            for run in para.runs:
                if run._element.xml and 'w:br' in run._element.xml and 'page' in run._element.xml:
                    if in_code and code_lines:
                        html_parts.append(f'<pre class="code">{chr(10).join(code_lines)}</pre>')
                        code_lines = []
                        in_code = False
                    pending_break = True

            if not text.strip():
                continue

            if para.style and para.style.name.startswith('Heading'):
                if in_code and code_lines:
                    html_parts.append(f'<pre class="code">{chr(10).join(code_lines)}</pre>')
                    code_lines, in_code = [], False
                level = para.style.name[-1] if para.style.name[-1].isdigit() else '1'
                emit(f'<h{level}>{text}</h{level}>')
            elif is_code(para):
                if not in_code:
                    in_code = True
                code_lines.append(para.text)
            elif para.style and 'List' in para.style.name:
                if in_code and code_lines:
                    html_parts.append(f'<pre class="code">{chr(10).join(code_lines)}</pre>')
                    code_lines, in_code = [], False
                emit(f'<li>{text}</li>')
            else:
                if in_code and code_lines:
                    html_parts.append(f'<pre class="code">{chr(10).join(code_lines)}</pre>')
                    code_lines, in_code = [], False
                align = para.alignment
                cls = ' class="center"' if align and align.name == 'CENTER' else ''
                emit(f'<p{cls}>{text}</p>')

        elif element.tag.endswith('}tbl'):
            if in_code and code_lines:
                html_parts.append(f'<pre class="code">{chr(10).join(code_lines)}</pre>')
                code_lines, in_code = [], False
            table = Table(element, doc)
            rows_html = []
            for i, row in enumerate(table.rows):
                cells = []
                tag = 'th' if i == 0 else 'td'
                for cell in row.cells:
                    cell_text = ' '.join(p.text for p in cell.paragraphs)
                    cells.append(f'<{tag}>{html.escape(cell_text)}</{tag}>')
                rows_html.append('<tr>' + ''.join(cells) + '</tr>')
            emit('<table>' + ''.join(rows_html) + '</table>')

    if in_code and code_lines:
        html_parts.append(f'<pre class="code">{chr(10).join(code_lines)}</pre>')

    html_parts.append('</body></html>')
    html_content = '\n'.join(html_parts)

    with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as f:
        f.write(html_content)
        html_file = f.name

    chrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    cmd = [chrome, '--headless=new', '--disable-gpu', '--no-sandbox',
           '--no-pdf-header-footer', '--print-to-pdf-no-header',
           f'--print-to-pdf={pdf_path}', f'file://{html_file}']
    subprocess.run(cmd, check=True, capture_output=True)
    Path(html_file).unlink()
    print(f"PDF saved to {pdf_path}")

if __name__ == '__main__':
    convert(sys.argv[1], sys.argv[2])
