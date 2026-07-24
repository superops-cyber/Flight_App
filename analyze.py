import re
from collections import defaultdict

def analyze_file(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    print(f"\n--- {filename} ({len(lines)} lines) ---")
    
    empty = sum(1 for l in lines if not l.strip())
    print(f"Empty lines: {empty}")
    
    if filename.endswith('.js') or filename.endswith('.gs'):
        comments = sum(1 for l in lines if l.strip().startswith('//'))
        print(f"Comment lines (//): {comments}")
        
        # Find functions
        funcs = []
        for l in lines:
            m = re.match(r'^(?:async )?function\s+(\w+)', l.strip())
            if m:
                funcs.append(m.group(1))
        print(f"Total functions: {len(funcs)}")
        print(f"Sample functions: {funcs[:10]}")
        
    if filename.endswith('.html'):
        scripts = len(re.findall(r'<script.*?>', "".join(lines), re.IGNORECASE))
        styles = len(re.findall(r'<style.*?>', "".join(lines), re.IGNORECASE))
        print(f"<script> tags: {scripts}")
        print(f"<style> tags: {styles}")

analyze_file('Code.js')
analyze_file('PilotApp.html')
analyze_file('Tab3_WB.html')
