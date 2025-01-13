import xlrd

wb = xlrd.open_workbook("2.xls")
sheet = wb.sheet_by_index(0)


all_pros = {}

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

        if employee not in all_pros.keys():
            all_pros[employee] = 1
        else:
            all_pros[employee] += 1  

    print()


# Sorting by problems

sorted_all_pros = dict(sorted(all_pros.items(), key=lambda item: item[1], reverse=True))

print("\n\n\nSorted by problems:")

for employee, count in sorted_all_pros.items():
    print(employee, " - ", count)
