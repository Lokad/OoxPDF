"""Public continuous-endnote controls derived from pinned footnote combinations."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
import hashlib,json

root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
sources=json.loads((cases/'math-footnotes.json').read_text(encoding='utf-8'))
from DocxFixtureNotes import replace_footnotes_with_endnotes

manifest=[]
for case in sources:
    source=root/case['Input'];assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==case['Sha256']
    with ZipFile(source) as archive:entries={name:archive.read(name) for name in archive.namelist()}
    entries=replace_footnotes_with_endnotes(entries)
    target=cases/(source.stem.replace('math-footnote-','math-endnote-')+'.docx')
    with ZipFile(target,'w',ZIP_DEFLATED) as archive:
        for name,data in sorted(entries.items()):
            info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
    manifest.append(case|dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper()))
(cases/'math-endnotes.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned four continuous endnote/body/cell controls; default lower-Roman numbering retained.')
