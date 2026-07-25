import os

def extract_assets(filename, prefix):
    if not os.path.exists(filename):
        print(f"{filename} not found.")
        return
        
    with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
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
            
    # Only write if we actually extracted something significant (avoiding tiny inline scripts)
    if len(css_lines) > 20 or len(js_lines) > 20:
        with open(filename, 'w', encoding='utf-8') as f:
            f.writelines(html_lines)
            
        if len(css_lines) > 0:
            with open(f"{css_filename}.html", 'w', encoding='utf-8') as f:
                f.writelines(css_lines)
                
        if len(js_lines) > 0:
            with open(f"{js_filename}.html", 'w', encoding='utf-8') as f:
                f.writelines(js_lines)
                
        print(f"Extracted {filename} into {filename} ({len(html_lines)} lines), {css_filename}.html ({len(css_lines)} lines), {js_filename}.html ({len(js_lines)} lines).")
    else:
        print(f"Skipped {filename} (Not enough embedded JS/CSS to warrant extraction)")

if __name__ == "__main__":
    tabs_to_extract = [
        'Tab2_Briefing.html',
        'Tab3_WB.html',
        'Tab4_Performance.html',
        'Tab5_Release.html',
        'Tab6_Enroute.html',
        'Tab7_Arrival.html',
        'Tab8_Debrief.html',
        'RunwayBriefingCard.html'
    ]
    
    for tab in tabs_to_extract:
        # Prefix is just the filename without .html
        extract_assets(tab, tab.replace('.html', ''))