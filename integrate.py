#!/usr/bin/env python3
"""Integrate generated content into index.html."""

import re
import subprocess

subprocess.run(['python3', 'build_content.py'], check=True)

with open('content_ro.html', 'r', encoding='utf-8') as f:
    content = f.read()

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

new_title = 'Scrisoarea Elenei Matei despre cistita cronică, prediabet și Cystiolla'
html = re.sub(r'<title>.*?</title>', f'<title>{new_title}</title>', html, count=1)

pattern = r'(<span class="mainText" style="">)(.*?)(</span>\s*<span id="oneform">)'
replacement = r'\1\n' + content + r'\n\3'
html_new = re.sub(pattern, replacement, html, count=1, flags=re.DOTALL)

if html_new == html:
    raise SystemExit('Failed to replace mainText content')

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html_new)

print('index.html updated successfully')
