# from persiantools.jdatetime import JalaliDate
#
# # Get today's date in Shamsi
# today_shamsi = JalaliDate.today()
# print("Today's Shamsi Date:", today_shamsi.strftime("%Y/%m/%d"))
#
# # Calculate 5 years before today
# five_years_ago = today_shamsi.replace(year=today_shamsi.year - 5)
# print("5 Years Ago (Shamsi):", five_years_ago.strftime("%Y/%m/%d"))
import pandas as pd

df = pd.read_excel('test.xls', engine='xlrd')

print(df)
