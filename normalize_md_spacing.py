from pathlib import Path
import re

FILES = [Path('135/Asas SOP.md'), Path('135/Asas MGM.md')]

heading_re = re.compile(r'^#{2,6}\s+')
bullet_re = re.compile(r'^\s*-\s+')


def normalize(path: Path) -> None:
    lines = path.read_text(encoding='utf-8').splitlines()
    out: list[str] = []

    for line in lines:
        stripped = line.strip()
        is_heading = bool(heading_re.match(line))
        is_bullet = bool(bullet_re.match(line))

        if is_heading:
            if out and out[-1].strip() != '':
                out.append('')
            out.append(line)
            continue

        if is_bullet:
            if out:
                prev = out[-1]
                prev_nonblank = prev.strip() != ''
                prev_is_heading = bool(heading_re.match(prev))
                prev_is_bullet = bool(bullet_re.match(prev))
                if prev_nonblank and not prev_is_heading and not prev_is_bullet:
                    out.append('')
            out.append(line)
            continue

        out.append(line)

    # collapse excessive blank lines
    cleaned: list[str] = []
    prev_blank = False
    for line in out:
        blank = line.strip() == ''
        if blank and prev_blank:
            continue
        cleaned.append(line)
        prev_blank = blank

    path.write_text('\n'.join(cleaned).strip() + '\n', encoding='utf-8')


if __name__ == '__main__':
    for file_path in FILES:
        normalize(file_path)
        print(f'normalized spacing: {file_path}')
