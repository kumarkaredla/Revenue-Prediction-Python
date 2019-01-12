import inline as inline
import matplotlib
import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt


#Product1
excel_file_1 = 'Revenue.xlsx'
originalFile_1 = pd.read_excel(excel_file_1)
sorted_revenue = originalFile_1.sort_values(['Revenue'], ascending=False)
revenue_1 = sorted_revenue['Revenue']
print(sorted_revenue)
#20, 11, 8

#product2
excel_file_2 = 'Revenue1.xlsx'
originalFile_2 = pd.read_excel(excel_file_2)
sorted_revenu1 = originalFile_2.sort_values(['Revenue'], ascending=False)
revenue_2 = sorted_revenu1['Revenue']
print(sorted_revenu1)
#20, 11, 5

#product3
excel_file_3 = 'Revenue2.xlsx'
originalFile_3 = pd.read_excel(excel_file_3)
sorted_revenu2 = originalFile_3.sort_values(['Revenue'], ascending=False)
revenue_3 = sorted_revenu2['Revenue']
print(sorted_revenu2)
#20,11,5

#product4
excel_file_4 = 'Revenue3.xlsx'
originalFile_4 = pd.read_excel(excel_file_4)
sorted_revenu3 = originalFile_4.sort_values(['Revenue'], ascending=False)
revenue_4 = sorted_revenu3['Revenue']
print(sorted_revenu3)
#20, 11, 5

#product5
excel_file_5 = 'Revenue4.xlsx'
originalFile_5 = pd.read_excel(excel_file_5)
sorted_revenu4 = originalFile_5.sort_values(['Revenue'], ascending=False)
revenue_5 = sorted_revenu4['Revenue']
print(sorted_revenu4)
#20, 11, 5

revenue_6 = revenue_1[20] + revenue_2[20] + revenue_3[20] + revenue_4[20] + revenue_5[20]
revenue_7 = revenue_1[11] + revenue_2[11] + revenue_3[11] + revenue_4[11] + revenue_5[11]
revenue_8 = revenue_2[5] + revenue_3[5] + revenue_4[5] + revenue_5[5]
revenue_9 = revenue_1[8]

revenue = [revenue_6,revenue_7,revenue_8,revenue_9]
print(revenue)
LABELS = ["US", "Japan", "China", "France"]

plt.bar(LABELS, revenue, align='center')
plt.show()
