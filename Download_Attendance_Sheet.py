from gsheets import Sheets

sheets = Sheets.from_files('client_secrets.json')
print(sheets)  #doctest: +ELLIPSIS
url = 'https://docs.google.com/spreadsheets/d/1Zi3Q4Td6sXIQIRkmgtM_No_AYpc8HpMGb653ARxLpa8'
s = sheets.get(url)
print(s)
s.sheets[0].to_csv('Dragonboat_Spring_2020_Attendance.csv', encoding='utf-8', dialect='excel')