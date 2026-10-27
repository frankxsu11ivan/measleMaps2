# measleMaps2
lotsThings and failed Potholes
bash -lc python - <<'PY' import csv, re from pathlib import Path src=Path('/mnt/data/261026 OSC FAILED PH CSV JBL .csv') out=Path('/mnt/data/261026_OSC_Potholes_Cleaned.csv') removed=Path('/mnt/data/261026_OSC_Potholes_Removed_Outliers.csv') rows=[]; rem=[] with src.open(newline='', encoding='utf-8-sig') as f: for i,r in enumerate(csv.reader(f),1): if not r or all(not x.strip() for x in r): continue r=[x.strip() for x in r] if len(r)<5: continue ident,n,e,z,desc=r[:5] d=' '.join(desc.upper().replace('_',' ').split()) # user-defined outliers: clear, no hydro, no vac if 'CLEAR' in d or 'NO HYDRO' in d or 'NO VAC' in d: reason='CLEAR' if 'CLEAR' in d else ('NO HYDRO' if 'NO HYDRO' in d else 'NO VAC') rem.append([ident,n,e,z,desc,reason]) continue # normalize retained failure descriptions if 'CONCRETE' in d and ('FAIL' in d or 'FAILED' in d): norm='FAILED - CONCRETE TOP' elif 'FAIL' in d or 'FAILED' in d: norm='FAILED' else: norm=d rows.append([ident,n,e,z,norm]) with out.open('w',newline='',encoding='utf-8') as f: w=csv.writer(f); w.writerow(['Unique Identifier','Northing','Easting','Elevation','Status']); w.writerows(rows) with removed.open('w',newline='',encoding='utf-8') as f: w=csv.writer(f); w.writerow(['Unique Identifier','Northing','Easting','Elevation','Original Description','Removal Reason']); w.writerows(rem) from collections import Counter ids=Counter(r[0] for r in rows) print('retained',len(rows),'removed',len(rem),'duplicate_ids',{k:v for k,v in ids.items() if v>1}) print('removal_reasons',Counter(r[-1] for r in rem)) print('statuses',Counter(r[-1] for r in rows)) PY
bash -lc cat > /mnt/data/add_footnote.py <<'PY'
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

src='/mnt/data/Utility_Observation_Report_2026-10-26.docx'
out='/mnt/data/Utility_Observation_Report_2026-10-26_Updated.docx'
doc=Document(src)

# Locate the closing line and insert a clearly styled report footnote immediately before it.
target=None
for p in doc.paragraphs:
    if 'End of Entry | October 26, 2026' in p.text:
        target=p
        break
if target is None:
    target=doc.add_paragraph()

new_p = OxmlElement('w:p')
target._p.addprevious(new_p)
foot = doc.add_paragraph()
foot._p.getparent().remove(foot._p)
new_p.getparent().replace(new_p, foot._p)
foot.alignment=WD_ALIGN_PARAGRAPH.LEFT
foot.paragraph_format.space_before=Pt(8)
foot.paragraph_format.space_after=Pt(4)
foot.paragraph_format.left_indent=Pt(8)
foot.paragraph_format.right_indent=Pt(8)

r=foot.add_run('Footnote: ')
r.bold=True; r.italic=True; r.font.name='Aptos'; r.font.size=Pt(8.5); r.font.color.rgb=RGBColor(90,90,90)
r=foot.add_run('Data from failed utility investigations is being cleaned. The initial export identified 198 noted failures and 26 clear pothole locations.')
r.italic=True; r.font.name='Aptos'; r.font.size=Pt(8.5); r.font.color.rgb=RGBColor(90,90,90)

# Keep the note and closing entry together where possible.
pPr=foot._p.get_or_add_pPr(); keep=OxmlElement('w:keepNext'); pPr.append(keep)

doc.save(out)
print(out)
PY
python /mnt/data/add_footnote.py
rm -rf /mnt/data/rendered_oct26_updated
python /home/oai/skills/docx/render_docx.py /mnt/data/Utility_Observation_Report_2026-10-26_Updated.docx --output_dir /mnt/data/rendered_oct26_updated
ls -la /mnt/data/rendered_oct26_updated
