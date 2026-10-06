import sys, docx
from docx.table import Table
from docx.text.paragraph import Paragraph
d=docx.Document(sys.argv[1])
for child in d.element.body.iterchildren():
    if child.tag.endswith('}p'):
        p=Paragraph(child,d)
        print('P:',p.text)
    elif child.tag.endswith('}tbl'):
        t=Table(child,d)
        print(f'TABLE {len(t.rows)}x{len(t.columns)}')
        for ri,row in enumerate(t.rows):
            print('  R%d: %s'%(ri,[c.text.replace(chr(10),' \\n ') for c in row.cells]))
