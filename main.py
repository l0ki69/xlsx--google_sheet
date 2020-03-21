from pprint import pprint
import apiclient
import httplib2
import xlrd
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import gspread

#-------------------------------------------------------------------------------------------------------------

import re


def get_substrings(string):
    """Функция разбивки на слова"""
    return re.split('\W+', string)


def get_distance(s1, s2):
    """Расстояние Дамерау-Левенштейна"""
    d, len_s1, len_s2 = {}, len(s1), len(s2)
    for i in range(-1, len_s1 + 1):
        d[(i, -1)] = i + 1
    for j in range(-1, len_s2 + 1):
        d[(-1, j)] = j + 1
    for i in range(len_s1):
        for j in range(len_s2):
            if s1[i] == s2[j]:
                cost = 0
            else:
                cost = 1
            d[(i, j)] = min(
                d[(i - 1, j)] + 1,
                d[(i, j - 1)] + 1,
                d[(i - 1, j - 1)] + cost)
            if i and j and s1[i] == s2[j - 1] and s1[i - 1] == s2[j]:
                d[(i, j)] = min(d[(i, j)], d[i - 2, j - 2] + cost)
    return(d[len_s1 - 1, len_s2 - 1])


def check_substring(search_request, original_text, max_distance):
    """Проверка нечёткого вхождения одного набора слов в другой"""
    substring_list_1 = get_substrings(search_request)
    substring_list_2 = get_substrings(original_text)

    not_found_count = len(substring_list_1)

    for substring_1 in substring_list_1:
        for substring_2 in substring_list_2:
            if get_distance(substring_1, substring_2) <= max_distance:
                not_found_count -= 1

    if not not_found_count:
        return True


#-------------------------------------------------------------------------------------------------------------




def Univ_file():
    list_univ = []
    file = open('University.txt')
    for line in file:
        list_univ.append(line.replace('\n', '').split('|'))
        # print(list_univ)
    file.close()
    # print(list_univ)
    return list_univ


def KEY_D(d, key):
    if d == {}:
        # print("False")
        return False
    for i in d.keys():
        if Format_Univ(key) == Format_Univ(i):
            # print("True")
            return True
        else:
            # print("False")
            return False


def Format_Univ(Univ):
    return (str(Univ).lower()).replace(',', ' ')


def Sort_Dict_Name(list_inf):
    buf_lst = []  # sorted name
    Buf_Lst = []  # Buf_list
    for lst in list_inf:
        buf_lst.append(lst[0])
    buf_lst.sort()
    for i in buf_lst:
        for j in list_inf:
            if i == j[0]:
                break
        Buf_Lst.append(list_inf.pop(list_inf.index(j)))
    # print(Buf_Lst)
    return (Buf_Lst)


# Начало обработки таблицы excel

excel_data_file = xlrd.open_workbook('./Vygruzka_17_10.xlsx')
sheet = excel_data_file.sheet_by_index(0)

list_info_stud = []
num_rows = sheet.nrows
num_col = sheet.ncols

dict_univ = {}
univ = ""

list_univ = Univ_file()
k = 0


def Comp(lst_univ, str_univ):
    for i in lst_univ:
        result = check_substring(str_univ.lower().replace("'","").replace('"',''), i.lower().replace("'","").replace('"',''), max_distance=5)
        if result: return True
    return False


for i in range(1, num_rows):

    list_info_stud.append((str(sheet.cell(i, 2)).replace('text:', '').replace("'", "")))
    list_info_stud.append((str(sheet.cell(i, 6)).replace('text:', '').replace("'", "")))
    list_info_stud.append((str(sheet.cell(i, 7)).replace('text:', '').replace("'", "")))
    univ = str(sheet.cell(i, 5)).replace('text:', '').replace("'", "").rstrip()
    buf_list = list_info_stud[:]
    key = univ
    for j in range(0, len(list_univ)):
        #print(list_univ)
        if Comp(list_univ[j],univ.lower()):
            key = list_univ[j][0].upper()
            break


    if KEY_D(dict_univ, univ):
        dict_univ[Format_Univ(key).upper()].append(buf_list)
    else:
        dict_univ[Format_Univ(key).upper()] = [buf_list]
    list_info_stud.clear()

for key in dict_univ.keys():
    #print(key)
    dict_univ[key] = Sort_Dict_Name(dict_univ[key])

# Конец обработки excel

# Начало работы с google sheets

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)

client = gspread.authorize(creds)

Name_Sheets = "SheetsTest"

sheet = client.open(Name_Sheets).sheet1

data = sheet.get_all_records()

print(dict_univ)


k = 0

for i in dict_univ.keys():
    key = i
    k += 1
    #print(key)
    if k == 5: break
    print(dict_univ[key])

for i in range(0, len(dict_univ[key])):
    # row = [i - 1 , "Титаренко Алексей Андреевич" , "https://vk.com/l0ki69" , "https://sun9-46.userapi.com/c855232/v855232754/1eee16/DgBkCyRvxjs.jpg"]
    row = [i + 1]
    row.extend(dict_univ[key][i])
    # for j in range(0,len(row)):
    # print(type(row[j]))
    # print(row)
    sheet.insert_row(row, i+2)

# pprint(data)

# Конец работы с google sheets


"""

# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'creds.json'
# ID Google Sheets документа (можно взять из его URL)
spreadsheet_id = '1iX8QhJQOqJZfFpVPUeIPE4BD1YbBwdZ6HwcrCCVo-6w'

# Авторизуемся и получаем service — экземпляр доступа к API
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)


# Пример чтения файла
values = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range='A1:E10',
    majorDimension='COLUMNS'
).execute()
pprint(values)


# Пример записи в файл
values = service.spreadsheets().values().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": "B3:C4",
             "majorDimension": "ROWS",
             "values": [["This is B3", "This is C3"], ["This is B4", "This is C4"]]},
            {"range": "D5:E6",
             "majorDimension": "COLUMNS",
             "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
        ]
    }
).execute()
"""
