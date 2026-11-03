import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = Path.home() / "OneDrive" / "Desktop" / "Flight_App_Codebase_Map.md"

EXTENSIONS = {".js", ".gs", ".html", ".py", ".json", ".md"}
EXCLUDE = {
    ".clasp.json",
    ".clasp-test-deploy-id",
    ".env",
    "package-lock.json",
}
SKIP_PARTS = {
    ".git", "node_modules", ".next", ".venv",
    "offline_build", "logbook-scanner-v2-backup",
    ".mx-push-worktree",
}

def tracked_files():
    result = subprocess.run(
        ["git", "ls-files", "-z"],
        cwd=ROOT,
        capture_output=True,
        check=True,
    )
    return [
        ROOT / name
        for name in result.stdout.decode("utf-8").split("\0")
        if name
    ]

def functions_in(text, suffix):
    if suffix == ".py":
        pattern = r"^\s*(?:async\s+)?def\s+([A-Za-z_]\w*)\s*\("
        return re.findall(pattern, text, re.MULTILINE)

    pattern = (
        r"^\s*(?:export\s+)?(?:async\s+)?function\s+"
        r"([A-Za-z_$][\w$]*)\s*\("
    )
    return re.findall(pattern, text, re.MULTILINE)

def main():
    lines = [
        "# Flight_App Codebase Map",
        "",
        "Generated from Git-tracked files.",
        "This report lists file structure and function names,",
        "not a complete semantic or security audit.",
        "",
    ]

    files = []
    for path in tracked_files():
        relative = path.relative_to(ROOT)
        if path.suffix.lower() not in EXTENSIONS:
            continue
        if path.name in EXCLUDE:
            continue
        if any(part in SKIP_PARTS for part in relative.parts):
            continue
        if not path.is_file():
            continue
        files.append((relative, path))

    files.sort(key=lambda item: str(item[0]).lower())
    lines += [
        f"## Inventory ({len(files)} files)",
        "",
    ]

    for relative, path in files:
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue

        count = len(text.splitlines())
        lines.append(f"- `{relative}` — {count:,} lines")

    lines += ["", "## Function index", ""]

    for relative, path in files:
        if path.suffix.lower() not in {".js", ".gs", ".py"}:
            continue

        text = path.read_text(encoding="utf-8", errors="replace")
        names = sorted(set(functions_in(text, path.suffix.lower())))

        if names:
            lines.append(f"### `{relative}`")
            lines.append(", ".join(f"`{name}`" for name in names))
            lines.append("")

    lines += ["## HTML integration clues", ""]

    for relative, path in files:
        if path.suffix.lower() != ".html":
            continue

        text = path.read_text(encoding="utf-8", errors="replace")
        ids = sorted(set(re.findall(
            r'\bid\s*=\s*["\']([^"\']+)["\']', text
        )))
        calls = sorted(set(re.findall(
            r"google\.script\.run\.(\w+)\s*\(", text
        )))

        if ids or calls:
            lines.append(f"### `{relative}`")
            if calls:
                lines.append(
                    "- Apps Script calls: "
                    + ", ".join(f"`{name}`" for name in calls)
                )
            if ids:
                lines.append(
                    f"- HTML element IDs: {len(ids)} "
                    "(names omitted to keep the report compact)"
                )
            lines.append("")

    OUTPUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report created: {OUTPUT}")
    print(f"Files indexed: {len(files)}")

if __name__ == "__main__":
    main()
