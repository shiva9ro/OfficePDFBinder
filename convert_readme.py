#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
README.md を README.html に変換するスクリプト
"""

import markdown

# README.md を読み込む
with open("README.md", "r", encoding="utf-8") as f:
    md_content = f.read()

# Markdown を HTML に変換
html_body = markdown.markdown(md_content, extensions=["tables", "fenced_code"])

# HTML テンプレート
html_template = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Office PDF Binder - ユーザーマニュアル</title>
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
    </style>
</head>
<body>
{body}
</body>
</html>"""

# HTML を生成
html_content = html_template.format(body=html_body)

# README.html に書き込む
with open("README.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("README.html を生成しました。")
