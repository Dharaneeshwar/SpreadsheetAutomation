from openpyxl import load_workbook
import json

with open('Json/counsellors details.json') as cd:
  	counsellor_data = json.load(cd)
current_counsellor = ""
for year in counsellor_data:
	for student in counsellor_data[year]:
		try:
			current_counsellor = student["Counselor's Name"]
		except:
			pass 
		