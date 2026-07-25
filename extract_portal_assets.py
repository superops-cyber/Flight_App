import os

def extract_assets(filename, prefix):
    with open(filename, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    html_lines = []
    css_lines = []
    js_lines = []
    
    in_style = False
    in_script = False
    
    css_filename = f"{prefix}_CSS"
    js_filename = f"{prefix}_JS"
    
    for line in lines:
        if '</style>' in line:
            in_style = False
            css_lines.append(line)
            html_lines.append(f'    <?!= include("{css_filename}"); ?>\n')
            continue
            
        if '</script>' in line:
            in_script = False
            js_lines.append(line)
            html_lines.append(f'    <?!= include("{js_filename}"); ?>\n')
            continue
            
        if '<style>' in line or '<style ' in line:
            in_style = True
            css_lines.append(line)
            continue
            
        if '<script>' in line or '<script ' in line:
            in_script = True
            js_lines.append(line)
            continue
            
        if in_style:
            css_lines.append(line)
        elif in_script:
            js_lines.append(line)
        else:
            html_lines.append(line)
            
    with open(filename, 'w', encoding='utf-8') as f:
        f.writelines(html_lines)
        
    with open(f"{css_filename}.html", 'w', encoding='utf-8') as f:
        f.writelines(css_lines)
        
    with open(f"{js_filename}.html", 'w', encoding='utf-8') as f:
        f.writelines(js_lines)
        
    print(f"Extracted {filename} into {filename} ({len(html_lines)} lines), {css_filename}.html ({len(css_lines)} lines), {js_filename}.html ({len(js_lines)} lines).")

if __name__ == "__main__":
    extract_assets('View_Tools.html', 'View_Tools')
    extract_assets('View_FlightFollowing.html', 'View_FlightFollowing')