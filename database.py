import openpyxl
import codecs

# "파일네임 바꿔라"
filename = "지역별 회사별 귀금속 비율(28열) (2).xlsx"
book = openpyxl.load_workbook(filename)
# 워크시트 번호 잘확인
sheet = book.worksheets[0]

def replaceToQuery(text):
	if text == None:
		return '0'
	else:
		return text
# file이름 바꿔라
sqlFile = codecs.open('region_company_metalratio_pra.sql', 'w',  'utf-8')

for row in sheet.rows:
	# 여기서도 이름 바꿔야됨
	queryString = "insert into region_company_metalratio values ("
	for i in row:
		queryString +=  "\'"+str(replaceToQuery(i.value))+"\'" + ","
	queryString = queryString.rstrip(',')
	queryString += ");"
	sqlFile.write(queryString + '\n')
sqlFile.close()
