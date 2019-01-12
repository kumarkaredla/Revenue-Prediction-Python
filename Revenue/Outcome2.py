import inline as inline
import matplotlib
import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt


#Product1
excel_file_1 = 'Revenue.xlsx'
originalFile_1 = pd.read_excel(excel_file_1)
sorted_revenue = originalFile_1.sort_values(['Revenue'], ascending=True)
revenue_1 = sorted_revenue['Revenue']
print(sorted_revenue)
#1, 6, 17

#product2
excel_file_2 = 'Revenue1.xlsx'
originalFile_2 = pd.read_excel(excel_file_2)
sorted_revenu1 = originalFile_2.sort_values(['Revenue'], ascending=True)
revenue_2 = sorted_revenu1['Revenue']
print(sorted_revenu1)
#18, 17, 6

#product3
excel_file_3 = 'Revenue2.xlsx'
originalFile_3 = pd.read_excel(excel_file_3)
sorted_revenu2 = originalFile_3.sort_values(['Revenue'], ascending=True)
revenue_3 = sorted_revenu2['Revenue']
print(sorted_revenu2)
#6, 18, 2

#product4
excel_file_4 = 'Revenue3.xlsx'
originalFile_4 = pd.read_excel(excel_file_4)
sorted_revenu3 = originalFile_4.sort_values(['Revenue'], ascending=True)
revenue_4 = sorted_revenu3['Revenue']
print(sorted_revenu3)
#6, 17, 18

#product5
excel_file_5 = 'Revenue4.xlsx'
originalFile_5 = pd.read_excel(excel_file_5)
sorted_revenu4 = originalFile_5.sort_values(['Revenue'], ascending=True)
revenue_5 = sorted_revenu4['Revenue']
print(sorted_revenu4)
#6, 17, 18

revenue_6 = revenue_1[6] + revenue_2[6] + revenue_3[6] + revenue_4[6] + revenue_5[6]
revenue_7 = revenue_1[17] + revenue_2[17] + revenue_4[17] + revenue_5[17]
revenue_8 = revenue_3[18] + revenue_4[18] + revenue_5[18]
revenue_9 = revenue_1[1]
revenue_10 = revenue_3[2]


revenue = [revenue_10,revenue_9,revenue_8,revenue_7,revenue_6]
print(revenue)
LABELS = ["Belgium", "Australia", "Sweden", "Switzerland", "Denmark"]

plt.bar(LABELS, revenue, align='center', width=0.3)
plt.show()
