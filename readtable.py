from docx.api import Document

doc= Document('docs/Jemi.docx')

for table in doc.tables:
    data=[]
    keys=None
    for i,row in enumerate(table.rows):
        text=(cell.text for cell in row.cells)
        keys='kvtystring'
        row_data = dict(zip(text,keys))
        data.append(row_data)
    for rw in data:
        print rw
    print('--------------------------------------')
        
