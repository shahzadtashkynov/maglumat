import docx
def gettext(filename):
    doc=docx.Document(filename)
    fulltext=[]
    for p in doc.paragraphs:
        fulltext.append(p.text)
    return '\n'.join(fulltext)