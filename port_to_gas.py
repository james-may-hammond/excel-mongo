"""
Module: port_to_gas.py
Description: Script to port the application code to Google Apps Script (GAS) format.
Dependencies: os, re
"""
import os
import re

base_dir = '/Users/rmgx/Actice Projects/Excel-Mongo'
html_path = os.path.join(base_dir, 'fe/taskpane.html')
js_path = os.path.join(base_dir, 'fe/taskpane.js')
css_path = os.path.join(base_dir, 'fe/taskpane.css')
out_path = os.path.join(base_dir, 'fe_sheets/Sidebar.html')

with open(html_path, 'r') as f:
    html_content = f.read()

with open(js_path, 'r') as f:
    js_content = f.read()

with open(css_path, 'r') as f:
    css_content = f.read()

# 1. Remove Office.js
html_content = re.sub(r'<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js".*?</script>', '', html_content)
html_content = re.sub(r'<script src="config.js"></script>', '<script>\nconst API_BASE = "http://localhost:8000";\n</script>', html_content)
html_content = re.sub(r'<link rel="stylesheet" href="taskpane.css">', f'<style>{css_content}</style>', html_content)
html_content = re.sub(r'<script src="taskpane.js\?v=5"></script>', '', html_content)

# 2. Modify JS
js_content = re.sub(r'const API_BASE = .*?;', '', js_content) # Handled above

# Replace Office.onReady
js_content = re.sub(r'Office\.onReady\(async \(info\) => \{', 'window.addEventListener("load", async () => {', js_content)
js_content = re.sub(r'if \(info\.host === Office\.HostType\.Excel\) \{', 'if (true) {', js_content)

# We will need to rewrite the Excel.run parts manually, but let's do a basic wrap
final_html = html_content.replace('</body>', f'<script>\n{js_content}\n</script>\n</body>')

with open(out_path, 'w') as f:
    f.write(final_html)

print("Generated Sidebar.html")
