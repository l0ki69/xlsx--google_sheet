from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import gspread


class Google_Sheets:
    def __init__(self, Name_Google_Sheet):
        self.Name_Sheet = Name_Google_Sheet
        self.Sheet = self.authorization()

    def authorization(self):
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open(self.Name_Sheet).sheet1

        return sheet

    def __get_col_index(self, str_name):
        if str_name in self.Sheet.row_values(1):
            return self.Sheet.row_values(1).index(str_name) + 1
        else:
            print("column not found")
            return -1

    def get_data(self, str_col):
        index = self.__get_col_index(str_col)
        if index != -1:
            lst = self.Sheet.col_values(index)
            lst.remove(str_col)
            return lst
        else:
            return []


    def add_data(self, str_col, lst):
        index = self.__get_col_index(str_col)
        if index == -1:
            return 0
        cell_range = chr(64 + index) + "2" + ":" + chr(64 + index) + str(len(lst))
        cell_list = self.Sheet.range(cell_range)
        k = 0
        for cell in cell_list:
            cell.value = lst[k]
            k += 1

        self.Sheet.update_cells(cell_list)

Name = "SheetTest"

test = Google_Sheets(Name)

buf = test.get_data("strava_nickname")

print(buf)

test.add_data("18.03", [1, 2, 3, 4, 5, 6, 7])

