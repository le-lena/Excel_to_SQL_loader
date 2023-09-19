#Подключаем модули
import xlrd, xlwt,openpyxl
import pymssql
import datetime
import os
from openpyxl import load_workbook
from sys import exit

#Считываем файл с настройками

y=open('config.txt','r')

for line in y:
    line_sep=line.split('=')

    if line_sep[0]=='processed_directory' or line_sep[0]=='processed_directory ':
        if line_sep[1][-1]=='\n':
         processed_directory=line_sep[1][:-1]
        else:
         processed_directory = line_sep[1]


    elif line_sep[0] == 'directory' or line_sep[0] == 'directory ':
        if line_sep[1][-1] == '\n':
         directory = line_sep[1][:-1]
        else:
         directory = line_sep[1]


    elif line_sep[0] == 'sql_directory' or line_sep[0] == 'sql_directory ':
        if line_sep[1][-1] == '\n':
         sql_directory = line_sep[1][:-1]
        else:
         sql_directory = line_sep[1]


    elif line_sep[0] == 'server' or line_sep[0] == 'server ':
        if line_sep[1][-1] == '\n':
            server = line_sep[1][:-1]
        else:
            server = line_sep[1]


    elif line_sep[0] == 'database' or line_sep[0] == 'database ':
        if line_sep[1][-1] == '\n':
            database = line_sep[1][:-1]
        else:
            database = line_sep[1]



    elif line_sep[0] == 'username' or line_sep[0] == 'username ':
        if line_sep[1][-1] == '\n':
            username = line_sep[1][:-1]
        else:
            username = line_sep[1]


    elif line_sep[0] == 'password' or line_sep[0] == 'password ':
        if line_sep[1][-1] == '\n':
            password = line_sep[1][:-1]
        else:
            password = line_sep[1]

    else:
        continue


#Получаем список файлов/директорий в переменную files
files = os.listdir(directory)

#Ищем в папке файлы формата xls, xlsx
excel = [x for x in files if x.endswith('.xlsx') or x.endswith('.xls')]

#Проверяем что в папке всего один файл формата xls, xlsx
if len(excel)==0:
    print("В папке",directory, "нет excel файла")
    input("Press Enter to continue...")
    exit()
if len(excel)>1:
    print("В папке",directory, len(excel), "excel файла(ов)")
    print(excel)
    input("Press Enter to continue...")
    exit()

#рабочие файлы
file_name = directory + excel[0]
sql_file_name = directory+'sql.txt'
print("Исходный файл находится здесь:",file_name)

#открываем исходный файл
rb = xlrd.open_workbook(file_name)
wb = openpyxl.load_workbook(file_name)

#Проверяем что в файле всего 1 лист
sheet_all=rb.sheet_names()
if len(sheet_all)==0:
    print('В книге 0 листов, файл пустой')
    input("Press Enter to continue...")
    exit()
elif len(sheet_all)>1:
    print('В книге более 1 листа, файл некорректный')
    input("Press Enter to continue...")
    exit()

#Выбираем активный лист
sheet = rb.sheet_by_index(0)

#Проверяем что расположение столбцов в файле верное
sheet_ranges=wb[sheet.name]
column_a = sheet_ranges['A']


for f in range(len(column_a)):
    if column_a[f].value=='InternalCode':
        #print(f)
        break


if (sheet.row_values(f,0)[0]=='InternalCode') and (sheet.row_values(f,8)[0]=='Fee') and (sheet.row_values(f,19)[0]=='TotalFee')  and (sheet.row_values(f,21)[0]=='Contr') and (sheet.row_values(f,22)[0]=='Folder'):
  print('Порядок столбцов верный')
else:
  print('порядок столбцов НЕверный')
  input("Press Enter to continue...")
  exit()



#Находим текущую дату
date=[]
date=str(datetime.date.today())
date_time=str(datetime.datetime.today())
for i in date_time:
    if i in (' ', ':','.'):
        date_time=date_time.replace(i,'-')
#print(date_time)
for i in date:
    if i=='-':
        date=date.replace(i,'')
print("Сегодняшняя дата - ",date)


print("Сервер:",server,'\n', "База:",database)
print("На базе данных будет запущено следующее:")

#заполняем первую часть (1 и 2 столбцы)
for k in range(len(column_a)):
  if (sheet.row_values(k, 0)[0], sheet.row_values(k, 19)[0], sheet.row_values(k, 8)[0],sheet.row_values(k, 21)[0]) not in ('', ' ')  and len(str(sheet.row_values(k,0)[0]))!=0 and len(str(sheet.row_values(k,19)[0]))!=0 and len(str(sheet.row_values(k,8)[0]))!=0 and len(str(sheet.row_values(k,21)[0]))!=0 and 'k/' in sheet.row_values(k, 0)[0]:
   a=[]
   a='exec dbo.sp_loader '+"'"+str(sheet.row_values(k,0)[0])+"',"+"'"+str(sheet.row_values(k,19)[0]*(-1))+"','"+date+"','"+str(sheet.row_values(k,8)[0])+"',"+'NULL'+",'"+str(sheet.row_values(k,21)[0]).split('.')[0]+"'"
   print(a)
   s=[]
   s='Y'+str(k+1)
   sheet_ranges[s] = a

  if (sheet.row_values(k, 0)[0], sheet.row_values(k, 19)[0], sheet.row_values(k, 8)[0], sheet.row_values(k, 22)[0]) not in ('', ' ')  and len(str(sheet.row_values(k,0)[0]))!=0 and len(str(sheet.row_values(k,19)[0]))!=0 and len(str(sheet.row_values(k,8)[0]))!=0 and len(str(sheet.row_values(k,22)[0]))!=0 and 'k/' in sheet.row_values(k, 0)[0]:
   b=[]
   b='exec dbo.sp_loader '+"'"+str(sheet.row_values(k,0)[0])+"',"+"'"+str(sheet.row_values(k,19)[0])+"','"+date+"','"+str(sheet.row_values(k,8)[0])+"',"+'NULL'+",'"+str(sheet.row_values(k,22)[0]).split('.')[0]+"'"
   print(b)
   r=[]
   r='Z'+str(k+1)
   sheet_ranges[r] = b

  else:
   continue


#сохраняем файл
wb.save(file_name)
wb.close()

rb = xlrd.open_workbook(file_name)
sheet1 = rb.sheet_by_index(0)

#открываем текстовый файл для записи скрипта на базе
f = open(sql_file_name, 'w')
f.write(
'BEGIN TRY''\n'
)
f.write('USE '+database+'\n\n')
f.write(
'declare''\n'
	'@id_before float,''\n'
	'@id_after float''\n'
'select @id_before = MAX(id) from TDB (nolock)''\n'
'\n'
'\n'
)
#записываем скрипт в файл
for num in sheet1.col_values(24):
    if num != '' and num != None and 'k/' in num:
      f.write(num+ '\n')
#f.write('\n')
for num in sheet1.col_values(25):
    if num != '' and num != None and 'k/' in num:
      f.write(num+ '\n')
f.write('\n')
#открываем файл с селектом- проверкой
d=open('cheking_select.txt','r')

for line in d:
    f.write(line)

f.close()



#перемещаем обработанные файлы
os.rename(file_name,processed_directory+sheet.row_values(0,0)[0]+' '+date_time+'.xlsx')
os.rename(file_name,processed_directory+sheet.row_values(0,0)[0]+' '+date_time+'.xlsx')


f=open(sql_file_name, 'r')

agree = str(input("Загрузить проводки?Y/N"))
if agree not in (' Y','Y','YES','Yes','y',' y'):
    print('Вы не подтвердили загрузку проводок. Проводки не загружены')
    input("Press Enter to continue...")
    exit()

#запускаем скрипт на базе
cnxn = pymssql.connect(server=server, user=username, password=password, database=database)


cursor = cnxn.cursor()

command=f.read()
cursor.execute(command)





while 1:
    num_fields = len(cursor.description)
    field_names = [i[0] for i in cursor.description]
    print(field_names)
    print('=====')
    for row in cursor:
        print(row)
    print('\n')
    if not cursor.nextset():
        break


cnxn.commit()
cnxn.close()
f.close()



os.rename(sql_file_name,sql_directory+sheet.row_values(0,0)[0]+' '+date_time+'.txt')

#print("Количество строк в файле равно ",amount1+amount2)
print("Проводки загружены успешно.")
print("Обработанный excel файл находится здесь",processed_directory+sheet.row_values(0,0)[0]+' '+date_time+'.xlsx')
print("Обработанный sql файл находится здесь",sql_directory+sheet.row_values(0,0)[0]+' '+date_time+'.txt')
y.close()


input("Press Enter to continue...")