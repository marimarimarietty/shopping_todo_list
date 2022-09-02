import gspread
from oauth2client.service_account import ServiceAccountCredentials

class ControllGoogleSpreadsheet:
  def __init__(self, title):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('/my-shopping-project-361202-650ebef29f5a.json', scope)
    global gc
    gc = gspread.authorize(credentials)
    gc = gc.open("MyShoppingList")

    wks = gc.open('MyShoppingList').sheet1

    wks.update_acell('A2', 'Hello World!')
    print(wks.acell('A2'))

    try :
        #make new sheet and store object
        worksheet = gc.add_worksheet(title=title, rows="100", cols="2")
    except :
        #if the sheet already exists, store object
        worksheet = gc.worksheet(title)

    self.worksheet = worksheet #store worksheet
    self.Todo = 1
    self.Done = 2
    self.worksheet.update_cell(1, self.Todo, "Todo")
    self.worksheet.update_cell(1, self.Done, "Done")

  #return last row of Todo
  def detect_last_row(self):
    row_count = 1
    while self.worksheet.cell(row_count, self.Todo).value != "" or self.worksheet.cell(row_count, self.Done).value != "":
        row_count += 1
    return row_count

  def write_to_Todo(self, text):
    self.worksheet.update_cell(self.detect_last_row(), self.Todo, text[2:])

  def from_Todo_to_Done(self, text):
    #whether element exists in Todo or not
    is_there_cell = True
    try:
      #store target cell number
      num = self.worksheet.find(text[2:]).row

      self.worksheet.update_cell(num, self.Todo, "")
      self.worksheet.update_cell(num, self.Done, text[2:])

    except CellNotFound:#if can't find the cell
      is_there_cell = False

    return is_there_cell

  #return all Todo
  def get_Todo(self):
    #get all values of Todo
    TodoList = self.worksheet.col_values(self.Todo)

    #remove empty cell
    TodoList = [todo for todo in TodoList if todo != ""]

    #join todos br and・
    #remove TodoList[1]
    TodoList = "\n・".join(TodoList[1:])

    return TodoList

  #reset worksheet
  def clear_worksheet(self):
    self.worksheet.clear()
    self.worksheet.update_cell(1, self.Todo, "Todo")
    self.worksheet.update_cell(1, self.Done, "Done")

  #delete worksheet
  def delete_worksheet(self):
    gc.del_worksheet(self.worksheet)
