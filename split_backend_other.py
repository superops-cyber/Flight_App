import re
import os
from collections import OrderedDict

def split_backend_other():
    with open('Backend_Other.js', 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
        
    print(f"Original lines in Backend_Other.js: {len(lines)}")
    
    in_func = False
    func_name = None
    func_lines = []
    global_code = []
    
    functions = OrderedDict()
    brace_depth = 0
    
    for line in lines:
        m = re.match(r'^\s*(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(', line)
        if not in_func and m:
            in_func = True
            func_name = m.group(1)
            func_lines = [line]
            brace_depth = line.count('{') - line.count('}')
            if brace_depth == 0 and '{' in line and '}' in line:
                if func_name not in functions:
                    functions[func_name] = "".join(func_lines)
                in_func = False
        elif in_func:
            func_lines.append(line)
            brace_depth += line.count('{') - line.count('}')
            if brace_depth <= 0:
                functions[func_name] = "".join(func_lines)
                in_func = False
        else:
            global_code.append(line)
            
    print(f"Total functions in Backend_Other: {len(functions)}")
    
    # Categorize
    categories = {
        'Backend_Maintenance.js': ['maintenance', 'mx', 'squawk', 'cva', 'inspect', 'defect', 'repair'],
        'Backend_Dispatch.js': ['dispatch', 'mission', 'scheduler', 'pax', 'passenger', 'manifest', 'schedule'],
        'Backend_Pilot.js': ['pilot', 'duty', 'flightdeck', 'wb', 'briefing', 'weight', 'balance'],
        'Backend_Flight.js': ['flight', 'route', 'waypoint', 'fuel', 'enroute', 'arrival', 'debrief', 'leg'],
        'Backend_ANAC.js': ['anac', 'rab', 'diario', 'civil', 'logbook', 'civ'],
        'Backend_Portal.js': ['portal', 'view', 'render', 'html', 'page', 'dashboard']
    }
    
    files_content = {
        'Backend_Maintenance.js': "",
        'Backend_Dispatch.js': "",
        'Backend_Pilot.js': "",
        'Backend_Flight.js': "",
        'Backend_ANAC.js': "",
        'Backend_Portal.js': "",
        'Backend_Other.js': "".join(global_code) # Keep global code in Other
    }
    
    for fname, body in functions.items():
        fname_lower = fname.lower()
        placed = False
        
        for cat_file, keywords in categories.items():
            if any(kw in fname_lower for kw in keywords):
                files_content[cat_file] += "\n" + body
                placed = True
                break
                
        if not placed:
            files_content['Backend_Other.js'] += "\n" + body
            
    # Write files
    for filename, content in files_content.items():
        if content.strip():
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content.strip() + "\n")
            print(f"Created/Updated {filename} with {len(content.splitlines())} lines.")

    print("Splitting Backend_Other.js complete.")

if __name__ == "__main__":
    split_backend_other()