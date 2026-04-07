from pathlib import Path
import re

FILES = [Path("135/Asas MGM.md"), Path("135/Asas SOP.md")]

bullet_re = re.compile(r"^(\s*)-\s+(.*)$")
heading_re = re.compile(r"^#{2,6}\s+")


def next_nonblank(lines: list[str], start: int) -> int | None:
    index = start
    while index < len(lines):
        if lines[index].strip() != "":
            return index
        index += 1
    return None


def adjust_file(path: Path) -> int:
    lines = path.read_text(encoding="utf-8").splitlines()
    changes = 0
    index = 0

    while index < len(lines):
        line = lines[index]
        match = bullet_re.match(line)
        if not match:
            index += 1
            continue

        indent = len(match.group(1))
        text = match.group(2).rstrip()
        if not text.endswith(":"):
            index += 1
            continue

        next_idx = next_nonblank(lines, index + 1)
        if next_idx is None:
            break

        next_match = bullet_re.match(lines[next_idx])
        if not next_match:
            index += 1
            continue

        if len(next_match.group(1)) != indent:
            index += 1
            continue

        run_index = next_idx
        while run_index < len(lines):
            curr = lines[run_index]
            if curr.strip() == "":
                run_index += 1
                continue
            if heading_re.match(curr):
                break
            curr_match = bullet_re.match(curr)
            if not curr_match:
                break
            curr_indent = len(curr_match.group(1))
            if curr_indent != indent:
                break

            lines[run_index] = " " * (indent + 2) + curr.lstrip()
            changes += 1
            run_index += 1

        index = run_index

    if changes:
        path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    return changes


if __name__ == "__main__":
    total = 0
    for file_path in FILES:
        changed = adjust_file(file_path)
        total += changed
        print(f"{file_path}: adjusted {changed} bullet lines")
    print(f"total adjusted: {total}")
