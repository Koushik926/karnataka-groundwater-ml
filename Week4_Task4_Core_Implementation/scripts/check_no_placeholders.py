"""Broad placeholder sweep: anything that tells a human to supply content."""
import glob, re, sys, zipfile
from pypdf import PdfReader
from pptx import Presentation

PAT = re.compile(
    r"\[[^\]\n]{0,120}(add|insert|your|fill|paste|enter|screenshot|tbd|todo|xx)"
    r"[^\]\n]{0,120}\]"                      # bracketed instruction
    r"|before printing|to be filled|fill in|fill after|paste here|insert here"
    r"|enter actual|enter your|add your|\bTODO\b|\bTBD\b|\[TO FILL\]"
    r"|lorem ipsum|_{4,}|xxxx", re.I)

def text_of(f):
    if f.endswith(".pdf"):
        return "\n".join(p.extract_text() or "" for p in PdfReader(f).pages)
    if f.endswith(".pptx"):
        out = []
        for sl in Presentation(f).slides:
            for sh in sl.shapes:
                if sh.has_text_frame: out.append(sh.text_frame.text)
                if sh.has_table:
                    for r in sh.table.rows:
                        out.extend(c.text for c in r.cells)
        return "\n".join(out)
    if f.endswith(".docx"):
        z = zipfile.ZipFile(f)
        xml = "".join(z.read(n).decode("utf8", "ignore")
                      for n in z.namelist() if n.endswith(".xml"))
        return re.sub(r"<[^>]+>", " ", xml)
    return open(f, errors="ignore").read()

bad = 0
targets = (sorted(glob.glob("Phase1_Reports/**/*.pdf", recursive=True))
           + sorted(glob.glob("Phase1_Reports/**/*.docx", recursive=True))
           + sorted(glob.glob("Week4_Task4_Core_Implementation/presentation/decks/Phase1_*.pptx")))
for f in targets:
    try:
        hits = sorted(set(m.group(0).strip()[:70] for m in PAT.finditer(text_of(f))))
    except Exception as e:
        print("ERR", f, e); continue
    if hits:
        bad += len(hits)
        print(f"{f.split('/')[-1]:45s} {hits}")
print(f"\nscanned {len(targets)} files; placeholder hits: {bad}")
sys.exit(1 if bad else 0)
