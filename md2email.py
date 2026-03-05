#!/usr/bin/env python3
"""md2email - Convert a markdown file to HTML and copy to clipboard for Outlook."""

import sys
import win32clipboard
import markdown

CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")

# Minimal styling that survives Outlook's paste renderer
STYLE = """<style>
  body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; }
  h1   { font-size: 16pt; margin-top: 16px; margin-bottom: 6px; }
  h2   { font-size: 13pt; margin-top: 14px; margin-bottom: 6px; }
  h3   { font-size: 11pt; margin-top: 12px; margin-bottom: 4px; }
  p    { margin: 8px 0; }
  ul, ol { margin: 6px 0 6px 1.5em; padding: 0; }
  li   { margin: 4px 0; }
  blockquote {
    border-left: 3px solid #cccccc;
    margin: 12px 0 12px 1em;
    padding: 4px 1em;
    color: #444444;
  }
  code { font-family: Consolas, monospace; background: #f4f4f4; padding: 1px 4px; }
  pre  { font-family: Consolas, monospace; background: #f4f4f4; padding: 8px; }
</style>"""


def build_cf_html(fragment: str) -> bytes:
    """
    Wrap an HTML fragment in the Windows CF_HTML clipboard format.
    See: https://docs.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
    """
    # Placeholder header — exact length must match the real header
    header_template = (
        "Version:0.9\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_frag:010d}\r\n"
        "EndFragment:{end_frag:010d}\r\n"
    )
    dummy = header_template.format(start_html=0, end_html=0, start_frag=0, end_frag=0)
    header_len = len(dummy.encode("utf-8"))

    pre_fragment = f"<html><head>{STYLE}</head><body>\r\n<!--StartFragment-->\r\n"
    post_fragment = "\r\n<!--EndFragment-->\r\n</body></html>"

    start_html  = header_len
    start_frag  = header_len + len(pre_fragment.encode("utf-8"))
    end_frag    = start_frag + len(fragment.encode("utf-8"))
    end_html    = end_frag   + len(post_fragment.encode("utf-8"))

    header = header_template.format(
        start_html=start_html,
        end_html=end_html,
        start_frag=start_frag,
        end_frag=end_frag,
    )

    return (header + pre_fragment + fragment + post_fragment).encode("utf-8")


def main():
    if len(sys.argv) != 2:
        print("Usage: md2email <file.md>")
        sys.exit(1)

    md_path = sys.argv[1]
    try:
        with open(md_path, encoding="utf-8") as f:
            md_text = f.read()
    except FileNotFoundError:
        print(f"Error: file not found: {md_path}")
        sys.exit(1)

    md = markdown.Markdown(extensions=["extra"])
    body_html = md.convert(md_text)
    body_html = body_html.replace("<hr />", "<br><br>").replace("<hr>", "<br><br>")

    cf_data = build_cf_html(body_html)

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(CF_HTML, cf_data)
    finally:
        win32clipboard.CloseClipboard()

    print(f"Copied to clipboard — paste into Outlook: {md_path}")


if __name__ == "__main__":
    main()
