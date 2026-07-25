import re
import os

def extract_functions(lines):
    funcs = {}
    in_func = False
    func_name = None
    func_lines = []
    brace_depth = 0
    
    other_lines = []
    
    for line in lines:
        m = re.match(r'^\s*(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(', line)
        if not in_func and m:
            in_func = True
            func_name = m.group(1)
            func_lines = [line]
            brace_depth = line.count('{') - line.count('}')
            if brace_depth == 0 and '{' in line and '}' in line:
                funcs[func_name] = "".join(func_lines)
                in_func = False
        elif in_func:
            func_lines.append(line)
            brace_depth += line.count('{') - line.count('}')
            if brace_depth <= 0:
                funcs[func_name] = "".join(func_lines)
                in_func = False
        else:
            other_lines.append(line)
            
    return funcs, other_lines

def dedupe():
    files = ['PilotApp_JS.html', 'View_Tools_JS.html', 'View_FlightFollowing_JS.html']
    
    file_funcs = {}
    
    for f in files:
        with open(f, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        funcs, others = extract_functions(lines)
        file_funcs[f] = funcs
        print(f"Extracted {len(funcs)} functions from {f}")

    # Find shared function names
    func_counts = {}
    func_bodies = {}
    
    for f, funcs in file_funcs.items():
        for name, body in funcs.items():
            if name not in func_counts:
                func_counts[name] = 0
                func_bodies[name] = body # take the first implementation we see
            func_counts[name] += 1
            
    shared_funcs = {name: func_bodies[name] for name, count in func_counts.items() if count > 1}
    print(f"Found {len(shared_funcs)} functions shared across multiple files.")
    print("Shared:", list(shared_funcs.keys())[:20])
    
    # Create Shared_JS.html
    with open('Shared_JS.html', 'w', encoding='utf-8') as f:
        f.write('<script>\n')
        for name, body in shared_funcs.items():
            f.write(body + "\n")
        f.write('</script>\n')
        
    # Re-write original files removing the shared funcs by NAME
    for filename in files:
        with open(filename, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            
        new_lines = []
        in_shared_func = False
        brace_depth = 0
        
        for line in lines:
            m = re.match(r'^\s*(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(', line)
            if not in_shared_func and m:
                func_name = m.group(1)
                if func_name in shared_funcs:
                    in_shared_func = True
                    brace_depth = line.count('{') - line.count('}')
                    new_lines.append(f"// Function {func_name} moved to Shared_JS\n")
                    if brace_depth == 0 and '{' in line and '}' in line:
                        in_shared_func = False
                    continue
            
            if in_shared_func:
                brace_depth += line.count('{') - line.count('}')
                if brace_depth <= 0:
                    in_shared_func = False
                continue
                
            new_lines.append(line)
                
        with open(filename, 'w', encoding='utf-8') as file:
            file.writelines(new_lines)
            
    # Inject <?!= include('Shared_JS'); ?> into the parent HTML files
    parents = ['PilotApp.html', 'View_Tools.html', 'View_FlightFollowing.html']
    for p in parents:
        with open(p, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if 'Shared_JS' not in content:
            js_include = f'<?!= include("{p.replace(".html", "")}_JS"); ?>'
            if js_include in content:
                content = content.replace(js_include, f'<?!= include("Shared_JS"); ?>\n    {js_include}')
                with open(p, 'w', encoding='utf-8') as f:
                    f.write(content)
                    
    print("Deduplication complete.")

if __name__ == "__main__":
    dedupe()