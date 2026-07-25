import os
import glob

def cleanup_temp_files():
    # Patterns for temp files
    patterns = [
        '__*.js',
        '_*.js',
        'line_map.txt',
        'temp_script_check.js',
        'extracted_script.js',
        'extracted_scripts.js',
        '*.backup'
    ]
    
    deleted_count = 0
    
    for pattern in patterns:
        for file in glob.glob(pattern):
            try:
                os.remove(file)
                print(f"Deleted: {file}")
                deleted_count += 1
            except Exception as e:
                print(f"Error deleting {file}: {e}")
                
    print(f"\nSuccessfully deleted {deleted_count} temporary/debug files.")

if __name__ == "__main__":
    cleanup_temp_files()