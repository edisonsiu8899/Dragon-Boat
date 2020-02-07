from gsheets import Sheets
from pyexcel.cookbook import merge_all_to_a_book
# import pyexcel.ext.xlsx # no longer required if you use pyexcel >= 0.2.2
import glob


name = "DB_Spr_2020_Attendance"

sheets = Sheets.from_files('client_secrets.json')
print(sheets)  #doctest: +ELLIPSIS
url = 'https://docs.google.com/spreadsheets/d/1Zi3Q4Td6sXIQIRkmgtM_No_AYpc8HpMGb653ARxLpa8'
s = sheets.get(url)
print(s)
s.sheets[0].to_csv(name + '.csv', encoding='utf-8', dialect='excel')

merge_all_to_a_book(glob.glob(name + ".csv"), name + ".xlsx")
