import os
import re

def remove_junk_functions():
    junk_funcs = {
        'buildDispatchInfrastructure', 'buildMultiPaxDispatch', 'buildNewAviationDB',
        'createDBPilotAuthorizations', 'setupChecksDatabase', 'setupFuelSystem',
        'setupMaintenanceSchedulingFramework', 'setupMaintenanceSchedulingSheets',
        'upgradeMaintenanceSystem', 'syncAnacData_legacy', 'syncAnacPrivateData_legacy',
        'testFFDebug', 'testFFDebugFromOutboundMap', 'testFFDebugKPS',
        'testFlightFollowSheetMessages', 'testFlightFollowSyncHealth', 'testMissionData',
        'testSyncFlightFollowInboxToSheet', 'logSheetsAndHeaders', 'diagnosticANAC'
    }

    files_modified = 0

    for root, dirs, files in os.walk('.'):
        if 'node_modules' in root or '.git' in root or '.venv' in root:
            continue
        for file in files:
            if not (file.endswith('.js') or file.endswith('.gs')): continue
            if file.endswith('.backup') or file.startswith('_'): continue
            
            filepath = os.path.join(root, file)
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
                
            new_lines = []
            in_junk = False
            brace_depth = 0
            file_changed = False
            
            for line in lines:
                m = re.match(r'^(?:async\s+)?function\s+([a-zA-Z0-9_$]+)\s*\(', line.strip())
                if not in_junk and m:
                    func_name = m.group(1)
                    if func_name in junk_funcs:
                        in_junk = True
                        file_changed = True
                        brace_depth = line.count('{') - line.count('}')
                        if brace_depth == 0 and '{' in line and '}' in line:
                            in_junk = False
                        continue
                        
                if in_junk:
                    brace_depth += line.count('{') - line.count('}')
                    if brace_depth <= 0:
                        in_junk = False
                    continue
                
                new_lines.append(line)
                
            if file_changed:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.writelines(new_lines)
                files_modified += 1
                print(f"Removed junk functions from {filepath}")

    print(f"Done. Modified {files_modified} files.")

if __name__ == "__main__":
    remove_junk_functions()