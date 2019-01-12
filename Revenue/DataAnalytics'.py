import inline as inline
import matplotlib
import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt
#% matplotlib inline
excel_file = 'salesData.xlsx'

originalFile = pd.read_excel(excel_file)

movies_sheet1 = pd.read_excel(excel_file, sheet_name=0, index_col=0)
movies_sheet1.head()

#sorted_by_gross = originalFile.sort_values(['Quantity'], ascending=False)

#sorted_by_gross['Quantity'].head(10).plot(kind="barh")
#plt.show()

year_data1 = (originalFile['Product.line']=='Personal Accessories')

year_data = originalFile[year_data1]
print(year_data['Revenue'].sum())

sorted_by_country = year_data.sort_values(['Retailer.country'], ascending=False)

#sorted_by_country[sorted_by_country["Retailer.country"] == "United States"]["Revenue"].plot(kind="hist")
#plt.show()

sorted_by_us1 = (sorted_by_country['Retailer.country']=='Austria')
sorted_austria = sorted_by_country[sorted_by_us1]
print(sorted_austria['Revenue'].sum())
a = sorted_austria['Revenue'].sum();

sorted_by_us2 = (sorted_by_country['Retailer.country']=='Australia')
sorted_australia = sorted_by_country[sorted_by_us2]
print(sorted_australia['Revenue'].sum())
b = sorted_australia['Revenue'].sum();


sorted_by_us3 = (sorted_by_country['Retailer.country']=='Belgium')
sorted_belgium = sorted_by_country[sorted_by_us3]
print(sorted_belgium['Revenue'].sum())
c = sorted_belgium['Revenue'].sum();

sorted_by_us4 = (sorted_by_country['Retailer.country']=='Brazil')
sorted_brazil = sorted_by_country[sorted_by_us4]
print(sorted_brazil['Revenue'].sum())
d = sorted_brazil['Revenue'].sum();

sorted_by_us5 = (sorted_by_country['Retailer.country']=='Canada')
sorted_canada = sorted_by_country[sorted_by_us5]
print(sorted_canada['Revenue'].sum())
e = sorted_canada['Revenue'].sum();

sorted_by_us6 = (sorted_by_country['Retailer.country']=='China')
sorted_china = sorted_by_country[sorted_by_us6]
print(sorted_china['Revenue'].sum())
f = sorted_china['Revenue'].sum();

sorted_by_us7 = (sorted_by_country['Retailer.country']=='Denmark')
sorted_denmark = sorted_by_country[sorted_by_us7]
print(sorted_denmark['Revenue'].sum())
g = sorted_denmark['Revenue'].sum();

sorted_by_us8 = (sorted_by_country['Retailer.country']=='Finland')
sorted_finland = sorted_by_country[sorted_by_us8]
print(sorted_finland['Revenue'].sum())
h = sorted_finland['Revenue'].sum();

sorted_by_us9 = (sorted_by_country['Retailer.country']=='France')
sorted_france = sorted_by_country[sorted_by_us9]
print(sorted_france['Revenue'].sum())
i = sorted_france['Revenue'].sum();

sorted_by_us10 = (sorted_by_country['Retailer.country']=='Germany')
sorted_germany = sorted_by_country[sorted_by_us10]
print(sorted_germany['Revenue'].sum())
j = sorted_germany['Revenue'].sum();

sorted_by_us11 = (sorted_by_country['Retailer.country']=='Italy')
sorted_italy = sorted_by_country[sorted_by_us11]
print(sorted_italy['Revenue'].sum())
k = sorted_italy['Revenue'].sum()

sorted_by_us12 = (sorted_by_country['Retailer.country']=='Japan')
sorted_japan = sorted_by_country[sorted_by_us12]
print(sorted_japan['Revenue'].sum())
l = sorted_japan['Revenue'].sum();

sorted_by_us13 = (sorted_by_country['Retailer.country']=='Korea')
sorted_korea = sorted_by_country[sorted_by_us13]
print(sorted_korea['Revenue'].sum())
m = sorted_korea['Revenue'].sum();

sorted_by_us14 = (sorted_by_country['Retailer.country']=='Mexico')
sorted_mexico = sorted_by_country[sorted_by_us14]
print(sorted_mexico['Revenue'].sum())
n = sorted_mexico['Revenue'].sum();


sorted_by_us15 = (sorted_by_country['Retailer.country']=='Netherlands')
sorted_netherlands = sorted_by_country[sorted_by_us15]
print(sorted_netherlands['Revenue'].sum())
o = sorted_netherlands['Revenue'].sum();

sorted_by_us16 = (sorted_by_country['Retailer.country']=='Singapore')
sorted_singapore = sorted_by_country[sorted_by_us16]
print(sorted_singapore['Revenue'].sum())
p = sorted_singapore['Revenue'].sum();


sorted_by_us17 = (sorted_by_country['Retailer.country']=='Spain')
sorted_spain = sorted_by_country[sorted_by_us17]
print(sorted_spain['Revenue'].sum())
q = sorted_spain['Revenue'].sum();


sorted_by_us18 = (sorted_by_country['Retailer.country']=='Sweden')
sorted_sweden = sorted_by_country[sorted_by_us18]
print(sorted_sweden['Revenue'].sum())
r = sorted_sweden['Revenue'].sum();


sorted_by_us19 = (sorted_by_country['Retailer.country']=='Switzerland')
sorted_switzerland = sorted_by_country[sorted_by_us19]
print(sorted_switzerland['Revenue'].sum())
s = sorted_switzerland['Revenue'].sum();

sorted_by_us20 = (sorted_by_country['Retailer.country']=='United Kingdom')
sorted_unitedkindom = sorted_by_country[sorted_by_us20]
print(sorted_unitedkindom['Revenue'].sum())
t = sorted_unitedkindom['Revenue'].sum();

sorted_by_us21 = (sorted_by_country['Retailer.country']=='United States')
sorted_unitedstates = sorted_by_country[sorted_by_us21]
print(sorted_unitedstates['Revenue'].sum())
u = sorted_unitedstates['Revenue'].sum();



print(sum)

#year_data.to_excel('output.xlsx', index=False)
writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

sorted_by_country.to_excel(writer, sheet_name='report')

workbook = writer.book

worksheet = writer.sheets['report']

header_fmt = workbook.add_format({'bold': True})

worksheet.set_row(0, None, header_fmt)
writer.save()


#to print the total revnue of each country by product line

workbook_1 = xlsxwriter.Workbook('Revenue.xlsx')
worksheet_1 = workbook_1.add_worksheet()
bold = workbook_1.add_format({'bold': 1})
worksheet_1.write('A1', 'Country', bold)
worksheet_1.write('B1', 'Revenue', bold)
expenses = (
     ['Austria', a],
     ['Australia',  b],
     ['Belgium', c],
     ['Brazil',  d],
     ['Canada', e],
     ['China',  f],
     ['Denmark', g],
     ['Finland',  h],
     ['France', i],
     ['Germany',  j],
     ['Italy', k],
     ['Japan',  l],
     ['Korea',  m],
     ['Mexico', n],
     ['Netherlands',  o],
     ['Singapore',  p],
     ['Spain', q],
     ['Sweden',  r],
     ['Switzerland',  s],
     ['UK', t],
     ['US',  u],
)

row = 1
col = 0

for x, y in (expenses):
     # Convert the date string into a datetime object.
     worksheet_1.write_string  (row, col,x)
     worksheet_1.write_number(row, col + 1, y )
     row += 1;



workbook_1.close()


#graph

excel_file_1 = 'Revenue.xlsx'

originalFile_1 = pd.read_excel(excel_file_1)


plt.bar(originalFile_1['Country'], originalFile_1['Revenue'], align='center')
plt.xticks(originalFile_1['Country'])
plt.show()
