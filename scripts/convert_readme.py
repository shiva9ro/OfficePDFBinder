#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""日本語・英語READMEをアプリ内マニュアル用HTMLへ変換する。"""

import markdown
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

# HTML テンプレート
html_template = """<!DOCTYPE html>
<html lang="{lang}">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            font-family: 'Meiryo UI', 'Segoe UI', 'Yu Gothic', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        h1 {{
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #34495e;
            border-bottom: 2px solid #ecf0f1;
            padding-bottom: 5px;
            margin-top: 30px;
        }}
        h3 {{
            color: #555;
            margin-top: 20px;
        }}
        code {{
            background-color: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 0.9em;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 15px 0;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
        }}
        th {{
            background-color: #3498db;
            color: white;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        ul, ol {{
            margin: 10px 0;
            padding-left: 30px;
        }}
        li {{
            margin: 5px 0;
        }}
        hr {{
            border: none;
            border-top: 1px solid #ddd;
            margin: 20px 0;
        }}
        img {{
            max-width: 100%;
            height: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
            margin: 10px 0;
        }}
    </style>
</head>
<body>
{body}
</body>
</html>"""


def convert_readme(source, destination, lang, title):
    source_path = PROJECT_ROOT / source
    destination_path = PROJECT_ROOT / destination
    with source_path.open("r", encoding="utf-8") as source_file:
        markdown_content = source_file.read()
    html_body = markdown.markdown(
        markdown_content, extensions=["tables", "fenced_code"]
    )
    # GitHubではMarkdown同士、同梱マニュアルではHTML同士を移動する。
    html_body = html_body.replace('href="README.ja.md"', 'href="README.ja.html"')
    html_body = html_body.replace('href="README.md"', 'href="README.html"')
    html_content = html_template.format(lang=lang, title=title, body=html_body)
    with destination_path.open("w", encoding="utf-8") as destination_file:
        destination_file.write(html_content)
    print(f"{destination} を生成しました。")


convert_readme(
    "README.md",
    "README.html",
    "en",
    "Office PDF Binder - User Manual",
)
convert_readme(
    "README.ja.md",
    "README.ja.html",
    "ja",
    "Office PDF Binder - ユーザーマニュアル",
)
