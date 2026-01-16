from lxml import etree
import zipfile
import html
from pathlib import Path

def hwpx_xml_to_html(hwpx_path: str, output_html="index.html"):
    blocks = []

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        for info in zf.infolist():
            if not info.filename.startswith("Contents/") or not info.filename.endswith(".xml"):
                continue

            raw = zf.read(info.filename)

            try:
                root = etree.fromstring(raw)
                pretty = etree.tostring(
                    root,
                    pretty_print=True,
                    encoding="unicode"
                )
            except Exception:
                continue

            blocks.append(f"""
            <section>
              <h2>{info.filename}</h2>
              <pre>{html.escape(pretty)}</pre>
            </section>
            """)

    html_doc = f"""
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body {{ background:#111; color:#eee; font-family:Consolas; padding:20px }}
        h2 {{ color:#ffd479 }}
        pre {{ background:#1e1e1e; padding:16px; overflow:auto }}
      </style>
    </head>
    <body>
      <h1>HWPX XML Viewer (Pretty)</h1>
      {''.join(blocks)}
    </body>
    </html>
    """

    Path(output_html).write_text(html_doc, encoding="utf-8")

if __name__ == "__main__":
    hwpx_xml_to_html(
        hwpx_path="template.hwpx",
        output_html="index.html"
    )
