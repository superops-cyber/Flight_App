from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import subprocess
import tempfile
import shutil

ref_patched = Path('135/manuals/reference.patched.docx')
ref_default = Path('135/manuals/reference.docx')
ref_doc = ref_patched if ref_patched.exists() else ref_default

ns_uri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': ns_uri}
ET.register_namespace('w', ns_uri)


def patch_bullets(out_doc: Path) -> None:
    with tempfile.TemporaryDirectory() as td_str:
        td = Path(td_str)
        with ZipFile(out_doc, 'r') as zin:
            zin.extractall(td)

        numbering = td / 'word' / 'numbering.xml'
        tree = ET.parse(numbering)
        root = tree.getroot()

        for lvl in root.findall('.//w:lvl', ns):
            numfmt = lvl.find('w:numFmt', ns)
            ppr = lvl.find('w:pPr', ns)
            ind = ppr.find('w:ind', ns) if ppr is not None else None
            if numfmt is None or ind is None:
                continue
            if numfmt.get('{%s}val' % ns_uri) != 'bullet':
                continue
            left = ind.get('{%s}left' % ns_uri)
            if left is not None:
                ind.set('{%s}left' % ns_uri, str(int(left) + 540))

        tree.write(numbering, encoding='utf-8', xml_declaration=True)

        rebuilt = out_doc.with_suffix('.tmp.docx')
        with ZipFile(rebuilt, 'w', ZIP_DEFLATED) as zout:
            for path in td.rglob('*'):
                if path.is_file():
                    zout.write(path, path.relative_to(td).as_posix())
        shutil.move(rebuilt, out_doc)


def build(md_path: Path, out_doc: Path, include_toc: bool) -> None:
    cmd = [
        'pandoc',
        str(md_path),
        '--from', 'markdown+fenced_divs+raw_attribute',
        '--reference-doc=' + str(ref_doc),
    ]
    if include_toc:
        cmd.append('--toc')
    cmd += ['-o', str(out_doc)]
    subprocess.run(cmd, check=True)
    patch_bullets(out_doc)


jobs = [
    (Path('135/Asas SOP.md'), Path('135/Asas_SOP.docx'), False),
    (Path('135/Asas SOP.md'), Path('135/Asas_SOP_with_toc.docx'), True),
    (Path('135/Asas MGM.md'), Path('135/Asas_MGM.docx'), False),
    (Path('135/Asas MGM.md'), Path('135/Asas_MGM_with_toc.docx'), True),
]

for md, out, toc in jobs:
    build(md, out, toc)
    print(out)
    print(out.stat().st_size)
