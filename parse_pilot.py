import re

def parse_pilot_app():
    with open('PilotApp.html', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    print(f"Total lines: {len(lines)}")
    
    in_style = False
    in_script = False
    style_lines = 0
    script_lines = 0
    html_lines = 0
    
    for line in lines:
        if '<style>' in line or '<style ' in line:
            in_style = True
        if '</style>' in line:
            in_style = False
            style_lines += 1
            continue
            
        if '<script>' in line or '<script ' in line:
            in_script = True
        if '</script>' in line:
            in_script = False
            script_lines += 1
            continue
            
        if in_style:
            style_lines += 1
        elif in_script:
            script_lines += 1
        else:
            html_lines += 1
            
    print(f"Lines of HTML: {html_lines}")
    print(f"Lines of CSS (<style>): {style_lines}")
    print(f"Lines of JS (<script>): {script_lines}")
    
if __name__ == "__main__":
    parse_pilot_app()