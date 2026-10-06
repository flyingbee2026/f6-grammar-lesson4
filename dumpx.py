import docx, sys
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
d=docx.Document('L3_Teachers.docx')
def sh(el):
    tcPr=el.find(qn('w:tcPr'))
    if tcPr is None: return None
    s=tcPr.find(qn('w:shd'))
    if s is None: return None
    return s.get(qn('w:fill'))
def pf(p):
    out=[]
    pPr=p._p.find(qn('w:pPr'))
    if pPr is not None:
        sp=pPr.find(qn('w:spacing'))
        if sp is not None: out.append('spacing='+str(dict((k.split('}')[1],v) for k,v in sp.attrib.items())))
        j=pPr.find(qn('w:jc'))
        if j is not None: out.append('jc='+j.get(qn('w:val')))
        bdr=pPr.find(qn('w:pBdr'))
        if bdr is not None: out.append('pBdr='+str([ (c.tag.split('}')[1], c.get(qn('w:val')), c.get(qn('w:sz')), c.get(qn('w:color'))) for c in bdr]))
    return ' '.join(out)
def rf(r):
    f=r.font
    col=None
    if f.color is not None and f.color.type is not None: col=str(f.color.rgb)
    return f"[{r.text!r} b={r.bold} i={r.italic} u={r.underline} sz={f.size.pt if f.size else None} col={col}]"
i=0
for child in d.element.body.iterchildren():
    if child.tag.endswith('}p'):
        p=Paragraph(child,d)
        if not p.text.strip():
            print(f'--BLANK-- {pf(p)}')
            continue
        print(f'P{i} style={p.style.name} {pf(p)}')
        print('   ', ' '.join(rf(r) for r in p.runs))
    elif child.tag.endswith('}tbl'):
        t=Table(child,d)
        print(f'TABLE {len(t.rows)}x{len(t.columns)} style={t.style.name if t.style else None}')
        tblPr=child.find(qn('w:tblPr'))
        print('   tblPr:', [(c.tag.split("}")[1], c.attrib) for c in tblPr] if tblPr is not None else None)
        for ri,row in enumerate(t.rows):
            print('   R%d h=%s'%(ri, row.height))
            for ci,c in enumerate(row.cells):
                print(f'      C{ci} shd={sh(c._tc)} w={c.width}')
                for p in c.paragraphs:
                    print(f'         p{rf2 if False else ""} {pf(p)}', ' '.join(rf(r) for r in p.runs))
    i+=1
