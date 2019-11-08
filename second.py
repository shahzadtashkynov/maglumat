import mysql.connector
mydb = mysql.connector.connect(
host="localhost",
user="root",
passwd="321478569",
database="kadrlar")
cursor= mydb.cursor()

# files=os.listdir('docs')
# files.sort()
# for f in files:
#     doc=docx.Document('docs/'+f)
#     table = doc.tables[1]
#     wplaces=[]
#     keys='abcdefghij'
#     for i,row in enumerate(table.rows):
#         text=(cell.text for cell in row.cells)
#         rdata = dict(zip(keys,text))
#         wplaces.append(rdata)
sql="INSERT INTO wpcs(name) VALUES (%s)"
val=("Shahzad")
cursor.execute(sql,val)
mydb.commit()
print(cursor.rowcount," record inserted")