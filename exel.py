import openpyxl as xl
from openpyxl.styles import Alignment

import xlrd
class ExelWB:
    def __init__(self, name=None, res_file_name=None, temp_table_name: dict = None):
        if name != None:
            self.wb = xlrd.open_workbook(name)
            
        print(f"xlrd: Workbook opened: '{name}'")

        self.author = "Игорь"
        self.filename = name
        self.res_file_name = res_file_name
        self.res_premia_wb = xl.load_workbook(self.res_file_name)
        self.svod_table = temp_table_name
        self.emploes = {}     
    
    def do_prosrochki(self):

        sheet = self.wb.sheet_by_index(0)

        all_pros = {}
        month_pros = {}

        employees = [sheet.cell_value(row, 11) for row in range(2, sheet.nrows-1)] # col11
        dates = [int(sheet.cell_value(row, 8).split(".")[1]) for row in range(2, sheet.nrows-1)] # col8
        status = [sheet.cell_value(row, 4).split(", ") for row in range(2, sheet.nrows-1)]

        for i in range(len(employees)):
            cell = employees[i]
            print("task: ", end="")
            names = cell.split("\n")
            for name in names:
                if len(name) == 0:
                    continue
                try:
                    arr = name.split("/")
                    print("\tname:", arr[0].strip(), "\n\t\tstatus:", arr[1].strip())
                    employee = arr[0].strip()
                    if ("но не принят" in status[i]) and ("Отчет отклонен на доработку" in arr):
                        if employee not in self.svod_table.keys():
                            self.svod_table[employee] = {"y": 1}
                        else:
                            if "y" not in self.svod_table[employee].keys():
                                self.svod_table[employee]["y"] = 1

                            else:
                                self.svod_table[employee]["y"] += 1

                except:
                    print("\tname:", name, "\n\t\tstatus: NONE")
                    employee = name.strip()

                if employee not in all_pros.keys():
                    all_pros[employee] = 1
                else:
                    all_pros[employee] += 1

                if employee not in self.svod_table.keys():
                    self.svod_table[employee] = {"r": 1}
                else:
                    if "r" not in self.svod_table[employee].keys():
                        self.svod_table[employee]["r"] = 1
                    else:
                        self.svod_table[employee]["r"] += 1
                
                if employee not in month_pros.keys():
                    month_pros[employee] = {dates[i]: 1}
                else:
                    if dates[i] not in month_pros[employee].keys():
                        month_pros[employee][dates[i]] = 1
                    else:
                        month_pros[employee][dates[i]] += 1
                

            print()

        # Sorting by problems

        sorted_all_pros = dict(sorted(all_pros.items(), key=lambda item: item[1], reverse=True))
        sorted_month_pros = {}
        for employee, counts in sorted_all_pros.items():
            sorted_month_pros[employee] = month_pros[employee]
        # sorting by months
        for employee, dates in sorted_month_pros.items():
            sorted_month_pros[employee] = dict(sorted(dates.items(), key=lambda item: item[0], reverse=True))

        print("\n\n\nAll Problems:")

        for employee, count in sorted_all_pros.items():
            print(employee, " - ", count)


        print("\n\n\nMonth Problems:")

        for employee, dates in sorted_month_pros.items():
            print(employee)
            for date, count in dates.items():
                print("\t", date, " - ", count)


        
        res_sheet = self.res_premia_wb.active
        res_sheet.title = "Премия"
        res_sheet.cell(row=1, column=1).value = "Просрочки"
        res_sheet.cell(row=3, column=1).value = "Сотрудник"
        res_sheet.cell(row=3, column=2).value = "Всего просрочек"
        res_sheet.cell(row=3, column=3).value = "Месяц"
        res_sheet.cell(row=3, column=4).value = "Количество прострочек за месяц"

        for employee, dates in sorted_month_pros.items():
            res_sheet.cell(row=res_sheet.max_row + 1, column=1).value = employee
            res_sheet.cell(row=res_sheet.max_row, column=2).value = sorted_all_pros[employee]
            for date, count in dates.items():
                res_sheet.cell(row=res_sheet.max_row+1, column=3).value = date
                res_sheet.cell(row=res_sheet.max_row, column=4).value = count

        for col in res_sheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(wrap_text=True)  # Учитываем перенос строк
            res_sheet.column_dimensions[col_letter].width = max_length + 2

        for row in range(1, res_sheet.max_row + 1):
            res_sheet.row_dimensions[row].height = 20

                
        self.res_premia_wb.save(self.res_file_name)

        return self.svod_table

    
    def do_vipolneno(self):
        sheet = self.wb.sheet_by_index(0)
        
        all_vipo = {}

        employees = [sheet.cell_value(row, 10) for row in range(2, sheet.nrows-1)] # col10

        for i in range(len(employees)):
            cell = employees[i]
            print("task: ", end="")
            names = cell.split("\n")
            for name in names:
                if len(name) == 0:
                    continue
                try:
                    arr = name.split("/")
                    print("\tname:", arr[0].strip(), "\n\t\tstatus:", arr[1].strip())
                    employee = arr[0].strip()
                except:
                    print("\tname:", name, "\n\t\tstatus: NONE")
                    employee = name.strip()

                if employee not in all_vipo.keys():
                    all_vipo[employee] = 1
                else:
                    all_vipo[employee] += 1
                
                if employee not in self.svod_table.keys():
                    self.svod_table[employee] = {"g": 1}
                else:
                    if "g" not in self.svod_table[employee].keys():
                        self.svod_table[employee]["g"] = 1
                    else:
                        self.svod_table[employee]["g"] += 1

            print()

        sorted_all_vipo = dict(sorted(all_vipo.items(), key=lambda item: item[1], reverse=True))

        print("\n\n\nSorted by problems:")

        for employee, count in sorted_all_vipo.items():
            print(employee, " - ", count)

        # append to excel


        res_sheet = self.res_premia_wb.active
        res_sheet.cell(row=res_sheet.max_row+2, column=1).value = "Выполенные"
        res_sheet.cell(row=res_sheet.max_row+2, column=1).value = "Сотрудник"
        res_sheet.cell(row=res_sheet.max_row, column=2).value = "Всего выполенных"

        for employee, count in sorted_all_vipo.items():
            res_sheet.cell(row=res_sheet.max_row + 1, column=1).value = employee
            res_sheet.cell(row=res_sheet.max_row, column=2).value = count

        for col in res_sheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(wrap_text=True)  # Учитываем перенос строк
            res_sheet.column_dimensions[col_letter].width = max_length + 2

        for row in range(1, res_sheet.max_row + 1):
            res_sheet.row_dimensions[row].height = 20


        self.res_premia_wb.save(self.res_file_name)

        return self.svod_table

    def do_svod(self):

        print("\n\n\nСводная таблица:", self.svod_table)
        

        res_sheet = self.res_premia_wb.active
        
        res_sheet.cell(row=res_sheet.max_row+2, column=1).value = "Сводная таблица"
    
        res_sheet.cell(row=res_sheet.max_row+2, column=2).value = "Поручения"
        res_sheet.cell(row=res_sheet.max_row+1, column=1).value = "Сотрудник"
        res_sheet.cell(row=res_sheet.max_row, column=2).value = "Просрочено, отчет отклонен"
        res_sheet.cell(row=res_sheet.max_row, column=3).value = "Проверка"
        res_sheet.cell(row=res_sheet.max_row, column=4).value = "Выполнено"
        res_sheet.cell(row=res_sheet.max_row, column=5).value = "Итог по сотруднику"

        for employee, done in self.svod_table.items():
            res_sheet.cell(row=res_sheet.max_row + 1, column=1).value = employee
            res_sheet.cell(row=res_sheet.max_row, column=2).value = done.get("r", 0)
            res_sheet.cell(row=res_sheet.max_row, column=3).value = done.get("y", 0)
            res_sheet.cell(row=res_sheet.max_row, column=4).value = done.get("g", 0)
            res_sheet.cell(row=res_sheet.max_row, column=5).value = done.get("r", 0) + done.get("y", 0) + done.get("g", 0)

        res_sheet.cell(row=res_sheet.max_row+1, column=1).value = "Итого"
        r = []
        y = []
        g = []
        for val in self.svod_table.values():
            try:
                r.append(val.get("r", 0))
                y.append(val.get("y", 0))
                g.append(val.get("g", 0))
            except:
                pass

        res_sheet.cell(row=res_sheet.max_row, column=2).value = sum(r)
        res_sheet.cell(row=res_sheet.max_row, column=3).value = sum(y)
        res_sheet.cell(row=res_sheet.max_row, column=4).value = sum(g)
        res_sheet.cell(row=res_sheet.max_row, column=5).value = sum(r) + sum(y) + sum(g)

        for col in res_sheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(wrap_text=True)  # Учитываем перенос строк
            res_sheet.column_dimensions[col_letter].width = max_length + 2

        for row in range(1, res_sheet.max_row + 1):
            res_sheet.row_dimensions[row].height = 20

        self.res_premia_wb.save(self.res_file_name)
        
    def create_res_premia(res_file_name: str):
        xl.Workbook().save(res_file_name)

    

    def end(self):
        self.res_premia_wb.close()