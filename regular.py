import re 
f = open('data.txt')
data = f.read()
f.close()
print(data)
print(re.search(r'\d{2}.\d{2}.\d{4}',data))