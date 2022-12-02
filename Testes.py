import pandas

xl_file = pandas.ExcelFile('C:\\Users\\Matheus\\Downloads\\TESTE PROGRAMA.xlsx', header=None)
dfs = []
for sheet_name in xl_file.sheet_names:
	dfs.append(xl_file.parse(sheet_name))

df = dfs[0]

#for xlsx in folder

df = pandas.read_excel('C:\\Users\\Matheus\\Downloads\\TESTE PROGRAMA.xlsx', header=None)
text = df[1].tolist()[5]