import os
import re
from collections import defaultdict

def find_dead_code():
    # 1. Collect all JS/GS files (ignore node_modules, backup files, etc.)
    code_files = []
    html_files = []
    for root, dirs, files in os.walk('.'):
        if 'node_modules' in root or '.git' in root or '.venv' in root:
            continue
        for file in files:
            if file.endswith('.js') or file.endswith('.gs'):
                if not file.endswith('.backup') and not file.startswith('_'):
                    code_files.append(os.path.join(root, file))
            elif file.endswith('.html'):
                html_files.append(os.path.join(root, file))

    print(f"Scanning {len(code_files)} JS/GS files and {len(html_files)} HTML files...")

    # 2. Extract all defined functions
    # function name -> list of files where it's defined
    defined_functions = defaultdict(list)
    func_regex = re.compile(r'^(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(')
    
    for file in code_files:
        with open(file, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                # Top level functions
                m = func_regex.search(line.strip())
                if m:
                    func_name = m.group(1)
                    # Ignore special GAS entry points
                    if func_name not in ['doGet', 'doPost']:
                        defined_functions[func_name].append(file)

    print(f"Found {len(defined_functions)} unique functions.")

    # 3. Search for usages
    # We will search for \bfunc_name\b in all files
    # If the count of usages across all files > number of definitions, it's used.
    # We have to be careful about the definition itself counting as a usage.
    
    usage_counts = {func: 0 for func in defined_functions}
    
    all_files = code_files + html_files
    for file in all_files:
        with open(file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            for func in defined_functions:
                # find all occurrences of the function name as a whole word
                matches = len(re.findall(r'\b' + re.escape(func) + r'\b', content))
                usage_counts[func] += matches

    # 4. Filter dead code
    # If usage_counts == number of times it was defined, it means it's never called!
    # (Because the definition itself counts as 1 usage per file where it's defined)
    dead_functions = []
    for func, count in usage_counts.items():
        if count <= len(defined_functions[func]):
            dead_functions.append(func)

    dead_functions.sort()
    
    print(f"\nFound {len(dead_functions)} potentially unused functions!")
    
    with open('dead_code_report.txt', 'w', encoding='utf-8') as out:
        out.write(f"Potentially unused functions ({len(dead_functions)}):\n")
        out.write("=========================================\n")
        for f in dead_functions:
            out.write(f"- {f} (Defined in: {', '.join(defined_functions[f])})\n")
            
    print("Report saved to dead_code_report.txt")

if __name__ == "__main__":
    find_dead_code()