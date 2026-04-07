from pathlib import Path
import re


def format_to_md(src: Path, dst: Path) -> None:
    lines = src.read_text(encoding="utf-8", errors="ignore").splitlines()
    out: list[str] = []

    for raw in lines:
        line = raw.replace("\u00a0", " ").rstrip()
        stripped = line.strip()

        if re.match(r"^-{20,}$", stripped):
            out.append("")
            continue

        if stripped == "":
            out.append("")
            continue

        m = re.match(r"^Seção\s+(\d+)\s*\|\s*(.+)$", stripped, re.IGNORECASE)
        if m:
            out.append(f"## Seção {m.group(1)} — {m.group(2).strip()}")
            continue

        m = re.match(r"^(\d+\.\d+(?:\.\d+)?)\s+(.+)$", stripped)
        if m:
            title = m.group(2).strip()
            if title.endswith(":"):
                title = title[:-1].rstrip()
            out.append(f"### {m.group(1)} {title}")
            continue

        if re.match(r"^(?:•|◦)\s*$", stripped):
            continue

        m = re.match(r"^(?:•)\s*(.+)$", stripped)
        if m:
            text = m.group(1).strip()
            if text:
                out.append(f"- {text}")
            continue

        m = re.match(r"^(?:◦)\s*(.+)$", stripped)
        if m:
            text = m.group(1).strip()
            if text:
                out.append(f"  - {text}")
            continue

        m = re.match(r"^(\d+)\s+(.+)$", stripped)
        if m:
            text = m.group(2).strip()
            if text:
                out.append(f"- {text}")
            continue

        if stripped.lower().startswith("fim da seção"):
            out.append(f"**{stripped}**")
            continue

        out.append(stripped)

    cleaned: list[str] = []
    prev_blank = False
    for line in out:
        blank = line.strip() == ""
        if blank and prev_blank:
            continue
        cleaned.append(line)
        prev_blank = blank

    dst.write_text("\n".join(cleaned).strip() + "\n", encoding="utf-8")


if __name__ == "__main__":
    format_to_md(Path("135/Asas SOP.txt"), Path("135/Asas SOP.md"))
    format_to_md(Path("135/Asas MGM.txt"), Path("135/Asas MGM.md"))
    print("created 135/Asas SOP.md")
    print("created 135/Asas MGM.md")
