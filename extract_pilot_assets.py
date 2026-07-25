import os

def extract():
    with open('PilotApp.html', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    html_lines = []
    css_lines = []
    js_lines = []
    
    in_style = False
    in_script = False
    
    for line in lines:
        # Check closures first
        if '</style>' in line:
            in_style = False
            css_lines.append(line)
            # Add the include tag to HTML
            html_lines.append('    <?!= include("PilotApp_CSS"); ?>\n')
            continue
            
        if '</script>' in line:
            in_script = False
            js_lines.append(line)
            # Add the include tag to HTML
            html_lines.append('    <?!= include("PilotApp_JS"); ?>\n')
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
            
    with open('PilotApp.html', 'w', encoding='utf-8') as f:
        f.writelines(html_lines)
        
    with open('PilotApp_CSS.html', 'w', encoding='utf-8') as f:
        f.writelines(css_lines)
        
    with open('PilotApp_JS.html', 'w', encoding='utf-8') as f:
        f.writelines(js_lines)
        
    print(f"Extracted to PilotApp.html ({len(html_lines)} lines), PilotApp_CSS.html ({len(css_lines)} lines), PilotApp_JS.html ({len(js_lines)} lines).")

if __name__ == "__main__":
    extract()