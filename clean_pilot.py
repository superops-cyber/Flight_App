import re

def clean_pilot_app():
    with open('PilotApp.html', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    print(f"Original lines in PilotApp.html: {len(lines)}")
    
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
    
    with open('PilotApp.html', 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)
        
    print("Saved cleaned PilotApp.html")

if __name__ == "__main__":
    clean_pilot_app()