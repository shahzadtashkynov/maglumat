import docxpy
import os
files=os.listdir('docs')
for f in files:
    print(f)
    target="docs/"+str(f)
    # print(type(target))
    text=docxpy.process(target,"img")
