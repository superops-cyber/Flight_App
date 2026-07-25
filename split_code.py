import re
import os
from collections import OrderedDict

def clean_and_split():
    with open('Code.js', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    print(f"Original lines: {len(lines)}")
    
    # Pass 1: Strip excessive blank lines (more than 1 in a row)
    cleaned_lines = []
    blank_count = 0
    for line in lines:
        if not line.strip():
            blank_count += 1
            if blank_count <= 1:
                cleaned_lines.append(line)
        else:
            blank_count = 0
            cleaned_lines.append(line)
            
    print(f"Lines after removing extra blanks: {len(cleaned_lines)}")
    
    # Pass 2: Identify functions and their content
    # We will build a dictionary of function_name -> function_body
    # to deduplicate exactly identical functions.
    
    in_func = False
    func_name = None
    func_lines = []
    
    # We also have global variables or top-level code that we need to keep
    global_code = []
    
    functions = OrderedDict() # to keep order of first appearance
    
    brace_depth = 0
    
    for line in cleaned_lines:
        # Detect function start
        # This regex matches `function foo(...) {` or `async function foo(...) {`
        m = re.match(r'^(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(', line.strip())
        
        if not in_func and m:
            in_func = True
            func_name = m.group(1)
            func_lines = [line]
            brace_depth = line.count('{') - line.count('}')
            # If function is entirely on one line
            if brace_depth == 0 and '{' in line and '}' in line:
                if func_name not in functions:
                    functions[func_name] = "".join(func_lines)
                in_func = False
        elif in_func:
            func_lines.append(line)
            brace_depth += line.count('{') - line.count('}')
            
            # Function ends
            if brace_depth <= 0:
                body = "".join(func_lines)
                # If function is duplicated, keep the FIRST or LAST? 
                # Usually in GAS, the last definition overrides the first.
                # So we should overwrite if it exists to keep the latest implementation.
                functions[func_name] = body
                in_func = False
        else:
            global_code.append(line)
            
    print(f"Total unique functions extracted: {len(functions)}")
    
    # Split into categories based on name patterns
    categories = {
        'Backend_API.js': ['doGet', 'doPost', 'include', 'getPortalProbeHtml'],
        'Backend_Dropdowns.js': ['getDropdownData', 'getPilotStartupData', 'getPilotAirportData'],
        'Backend_Sync.js': ['sync', 'export', 'import'],
        'Backend_Utils.js': ['safe', '_passenger', '_airport', 'appLog_', 'getRequiredSheet_']
    }
    
    files_content = {
        'Backend_Main.js': "".join(global_code),
        'Backend_API.js': "",
        'Backend_Dropdowns.js': "",
        'Backend_Sync.js': "",
        'Backend_Utils.js': "",
        'Backend_Other.js': ""
    }
    
    for fname, body in functions.items():
        placed = False
        for cat_file, patterns in categories.items():
            if any(p.lower() in fname.lower() for p in patterns):
                files_content[cat_file] += "\n" + body
                placed = True
                break
        
        if not placed:
            if fname.startswith('_') or 'helper' in fname.lower():
                files_content['Backend_Utils.js'] += "\n" + body
            else:
                files_content['Backend_Other.js'] += "\n" + body

    # Write files
    for filename, content in files_content.items():
        if content.strip():
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content.strip() + "\n")
            print(f"Created {filename} with {len(content.splitlines())} lines.")

    # Move original Code.js to backup
    os.rename('Code.js', 'Code.js.backup')
    print("Renamed Code.js to Code.js.backup")

if __name__ == "__main__":
    clean_and_split()
