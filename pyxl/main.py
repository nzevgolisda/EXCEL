from openpyxl import Workbook
workbook = Workbook()
sheet = workbook.active
Names = ['Nikos', 'Tolis', 'Adreas', 'Maria']
AM = [4156, 4223, 5184, 7935]
Sex = ['m', 'm', 'm', 'f']
Grades = [1, 8, 5, 2]
Final = [6, 7, 5, 7]

def fillGaps(i, Names, AM, Sex, Grades, Final):
    col_1 = 'A' + str(i+1)
    sheet[col_1] = Names[i]
    col_2 = 'B' + str(i+1)
    sheet[col_2] = AM[i]
    col_3 = 'C' + str(i+1)
    sheet[col_3] = Grades[i]
    col_4 = 'D' + str(i+1)
    sheet[col_4] = Final[i]
    col_5 = 'E' + str(i+1)
    sheet[col_5] = Sex[i]
    col_6 = 'F' + str(i+1)
    sheet[col_6] = Grades[i]*0.4 + Final[i]*0.6
for i in range(len(Names)):
    fillGaps(i, Names, AM, Sex, Grades, Final)
workbook.save(filename="data.xlsx")
    
