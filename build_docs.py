import os
import base64
import zipfile
import re
from pathlib import Path

doc_path = Path(r"C:\smart_parking_app\docs\PROJECT_DOCUMENTATION.md")
screenshots_dir = Path(r"C:\smart_parking_app\docs\screenshots")
out_html = Path(r"C:\smart_parking_app\DeccanPulse_Project_Report.html")
out_md = Path(r"C:\smart_parking_app\DeccanPulse_Project_Report.md")
out_zip = Path(r"C:\smart_parking_app\DeccanPulse_Complete_Project_Docs.zip")

content = doc_path.read_text(encoding="utf-8")
out_md.write_text(content, encoding="utf-8")

# Embed screenshots as base64 for self-contained single-file HTML report
images_b64 = {}
if screenshots_dir.exists():
    for img in sorted(screenshots_dir.glob("*.png")):
        b64 = base64.b64encode(img.read_bytes()).decode("utf-8")
        images_b64[img.name] = f"data:image/png;base64,{b64}"

html_lines = [
    "<!DOCTYPE html>",
    "<html lang=\"en\">",
    "<head>",
    "  <meta charset=\"UTF-8\" />",
    "  <title>DeccanPulse — Hyderabad Smart Parking Project Report</title>",
    "  <style>",
    "    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #1e293b; max-width: 960px; margin: 40px auto; padding: 0 24px; background: #ffffff; }",
    "    h1, h2, h3 { color: #0f3347; font-weight: 700; margin-top: 1.8em; }",
    "    h1 { font-size: 2.2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 0.4em; }",
    "    h2 { font-size: 1.5rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.3em; margin-top: 1.4em; }",
    "    h3 { font-size: 1.2rem; }",
    "    p { margin: 12px 0; }",
    "    ul, ol { margin: 12px 0 12px 24px; }",
    "    li { margin: 6px 0; }",
    "    table { width: 100%; border-collapse: collapse; margin: 24px 0; font-size: 0.92rem; }",
    "    th, td { border: 1px solid #cbd5e1; padding: 10px 14px; text-align: left; vertical-align: top; }",
    "    th { background: #f1f5f9; color: #0f3347; font-weight: 700; }",
    "    tr:nth-child(even) { background: #f8fafc; }",
    "    code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: 'Cascadia Mono', Consolas, monospace; font-size: 0.9em; color: #c05020; }",
    "    pre { background: #0f172a; color: #f8fafc; padding: 16px; border-radius: 8px; overflow-x: auto; font-size: 0.9em; line-height: 1.45; }",
    "    pre code { background: none; color: inherit; padding: 0; }",
    "    .img-box { margin: 28px 0; text-align: center; background: #f8fafc; border: 1.5px solid #e2e8f0; border-radius: 12px; padding: 16px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); }",
    "    .img-box img { max-width: 100%; height: auto; border-radius: 8px; border: 1px solid #cbd5e1; }",
    "    .img-caption { margin-top: 10px; font-size: 0.9rem; color: #475569; font-weight: 600; }",
    "    hr { border: 0; border-top: 1px solid #e2e8f0; margin: 36px 0; }",
    "  </style>",
    "</head>",
    "<body>"
]

in_code = False
in_table = False

for line in content.splitlines():
    if line.startswith("```"):
        if in_code:
            html_lines.append("</code></pre>")
            in_code = False
        else:
            lang = line[3:].strip()
            html_lines.append(f"<pre><code class=\"{lang}\">")
            in_code = True
        continue
    if in_code:
        html_lines.append(line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        continue

    if line.startswith("|") and line.endswith("|"):
        if not in_table:
            html_lines.append("<table>")
            in_table = True
        cells = [c.strip() for c in line.split("|")[1:-1]]
        if all(set(c) <= set("-: ") for c in cells if c):
            continue
        is_header = not any("<td>" in l for l in html_lines[-3:])
        tag = "th" if is_header else "td"
        row_str = "<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>"
        html_lines.append(row_str)
        continue
    else:
        if in_table:
            html_lines.append("</table>")
            in_table = False

    if line.startswith("# "):
        html_lines.append(f"<h1>{line[2:]}</h1>")
    elif line.startswith("## "):
        html_lines.append(f"<h2>{line[3:]}</h2>")
    elif line.startswith("### "):
        html_lines.append(f"<h3>{line[4:]}</h3>")
    elif line.startswith("---"):
        html_lines.append("<hr />")
    elif line.startswith("!["):
        m = re.match(r"!\[(.*?)\]\((.*?)\)", line)
        if m:
            cap, p = m.groups()
            fn = Path(p).name
            src = images_b64.get(fn, p)
            html_lines.append(f"<div class=\"img-box\"><img src=\"{src}\" alt=\"{cap}\" /><div class=\"img-caption\">Figure: {cap}</div></div>")
    elif line.startswith("* ") or line.startswith("- "):
        html_lines.append(f"<li>{line[2:]}</li>")
    elif line.strip():
        html_lines.append(f"<p>{line}</p>")

if in_table:
    html_lines.append("</table>")
if in_code:
    html_lines.append("</code></pre>")

html_lines.append("</body></html>")
out_html.write_text("\n".join(html_lines), encoding="utf-8")

# Create ZIP archive
with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(out_md, arcname="DeccanPulse_Project_Report.md")
    z.write(out_html, arcname="DeccanPulse_Project_Report.html")
    if screenshots_dir.exists():
        for img in sorted(screenshots_dir.glob("*.png")):
            z.write(img, arcname=f"screenshots/{img.name}")

print("Built successfully:")
print("Markdown:", out_md, f"({out_md.stat().st_size} bytes)")
print("HTML:", out_html, f"({out_html.stat().st_size} bytes)")
print("ZIP:", out_zip, f"({out_zip.stat().st_size} bytes)")
