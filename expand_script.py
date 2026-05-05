import re
import sys
import os

def expand_includes(filename, mapping, rendered_lines, seen_files):
    if filename in seen_files:
        return
    seen_files.add(filename)
    
    # Try .html extension if file not found
    path = filename
    if not os.path.exists(path) and not path.endswith('.html'):
        path += '.html'
        
    if not os.path.exists(path):
        rendered_lines.append(f"<!-- ERROR: Could not find {filename} -->")
        mapping.append((filename, 0))
        return

    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for i, line in enumerate(lines):
        line_num = i + 1
        # Match <?!= include('FileName'); ?>
        match = re.search(r'<\?!= include\(\s*[\'\"](.+?)[\'\"]\s*\);\s*\?>', line)
        if match:
            inc_file = match.group(1)
            expand_includes(inc_file, mapping, rendered_lines, seen_files)
        else:
            rendered_lines.append(line.rstrip('\n'))
            mapping.append((path, line_num))

rendered_lines = []
mapping = []
expand_includes('PilotApp.html', mapping, rendered_lines, set())

# Task 2: Report mapping for 3841 (1-indexed) and print lines 3833-3849
target_idx = 3840 # 0-indexed
start_idx = 3832
end_idx = 3849

print("--- Rendered Lines 3833-3849 with Source Mapping ---")
for i in range(start_idx, end_idx):
    if i < len(rendered_lines):
        src_file, src_line = mapping[i]
        print(f"R{i+1:4} | {src_file}:{src_line:<4} | {rendered_lines[i]}")

# Task 3: Print exact source file lines for mapped source line +/- 8
if target_idx < len(mapping):
    target_file, target_line = mapping[target_idx]
    print(f"\n--- Source View: {target_file} around line {target_line} ---")
    with open(target_file, 'r', encoding='utf-8') as f:
        src_lines = f.readlines()
        s = max(0, target_line - 9)
        e = min(len(src_lines), target_line + 8)
        for i in range(s, e):
            ln = i + 1
            marker = ">>" if ln == target_line else "  "
            print(f"{marker} {ln:4}: {src_lines[i].rstrip()}")
