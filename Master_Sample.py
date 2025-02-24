import calendar
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# Constants
DAYS_OF_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ATTENDANCE_OPTIONS = ["P", "L", "E", "A","N"]
REMARKS = {90: "Excellent", 75: "Good", 50: "Average", 0: "Poor"}

# Function to create attendance sheet for a month
def create_month_sheet(workbook, month, year):
    sheet = workbook.create_sheet(title=month)
    # sheet.freeze_panes = "A15"  # Freeze rows up to row 14

    # Set the header row
    sheet["A15"] = "Name"
    current_date = datetime(year, list(calendar.month_name).index(month), 1)
    col = 2
    total_days = 0

    while current_date.month == list(calendar.month_name).index(month):
        if current_date.strftime("%A") in DAYS_OF_WEEK:
            total_days += 1
            day_cell = sheet.cell(row=15, column=col)
            date_cell = sheet.cell(row=16, column=col)
            day_cell.value = current_date.strftime("%a")
            date_cell.value = current_date.strftime("%d-%b")
            day_cell.alignment = date_cell.alignment = Alignment(horizontal="center")
            col += 1
        current_date += timedelta(days=1)

    # Add average attendance and remarks columns
    for K in ATTENDANCE_OPTIONS:
        k = col + ATTENDANCE_OPTIONS.index(K) + 1
        sheet.cell(row=15, column=k).value = K

    attendance_col = col + 7
    avg_col = col + 8
    remarks_col = col + 9
    sheet.cell(row=15, column=attendance_col).value = "Attendance"
    sheet.cell(row=15, column=avg_col).value = "Average Attendance (%)"
    sheet.cell(row=15, column=remarks_col).value = "Remarks"

    return sheet, avg_col, remarks_col, total_days

# Function to calculate attendance summary and add dropdown options
def add_summary_and_chart(workbook, workers, total_days_dict):
    for sheet in workbook.worksheets:
        total_days= total_days_dict[sheet.title]
        avg_col = total_days + 10
        remarks_col = avg_col + 1

        dict_day = {
            23: {"P": 'Z', "L": 'AA', "E": 'AB', "A": 'AC', "N": 'AD'},
            22: {"P": 'Y', "L": 'Z', "E": 'AA', "A": 'AB', "N": 'AC'},
            21: {"P": 'X', "L": 'Y', "E": 'Z', "A": 'AA', "N": 'AB'},
            20: {"P": 'W', "L": 'X', "E": 'Y', "A": 'Z', "N": 'AA'}
        }

        dd_opts = list(list(dict_day.values())[0].keys())  # Dropdown options
        # Create a data validation for the dropdown
        dv = DataValidation(type="list", formula1=f'"{",".join(dd_opts)}"', showDropDown=True)

        # Add the data validation to each sheet
        sheet.add_data_validation(dv)

        for row in range(17, 17 + len(workers)):
            for column in range(total_days + 3, total_days + 8):
                position = column - total_days - 3
                if total_days == 23:
                    opt_list = list(dict_day[23].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(B{row}:X{row},${opt_list[position]}$15)"
                    sheet[f"AF{row}"].value = f"=SUM(Z{row}:AA{row})"
                    sheet[f"AG{row}"].value = f"=ROUND(AF{row}/{0.01 * total_days}, 1)"
                    #for cell in sheet[f"B{row}:X{row}"]:
                    dv.add(f"B{row}:X{row}")


                elif total_days == 22:
                    opt_list = list(dict_day[22].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(B{row}:W{row},${opt_list[position]}$15)"
                    sheet[f"AE{row}"].value = f"=SUM(Y{row}:Z{row})"
                    sheet[f"AF{row}"].value = f"=ROUND(AE{row}/{0.01 * total_days}, 1)"
                    dv.add(f"B{row}:W{row}")

                elif total_days == 21:
                    opt_list = list(dict_day[21].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(B{row}:V{row},${opt_list[position]}$15)"
                    sheet[f"AD{row}"].value = f"=SUM(X{row}:Y{row})"
                    sheet[f"AE{row}"].value = f"=ROUND(AD{row}/{0.01 * total_days}, 1)"
                    dv.add(f"B{row}:V{row}")

                elif total_days == 20:
                    opt_list = list(dict_day[20].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(B{row}:U{row},${opt_list[position]}$15)"
                    sheet[f"AC{row}"].value = f"=SUM(W{row}:X{row})"
                    sheet[f"AD{row}"].value = f"=ROUND(AC{row}/{0.01 * total_days}, 1)"
                    dv.add(f"B{row}:U{row}")

                else:
                    pass

            avg_cell = sheet.cell(row=row, column= avg_col)
            remark_cell = sheet.cell(row=row, column= remarks_col)
            remark_cell.value = (f'=IF({avg_cell.coordinate}>90.00,"Excellent",IF({avg_cell.coordinate}>75.00,"Good",'
                                 f'IF({avg_cell.coordinate}>50.00,"Average","Poor")))')

        # Add a chart
        chart = BarChart()
        chart.title = f"Average Daily Attendance ({sheet.title})"
        data = Reference(sheet, min_col=avg_col, min_row=17, max_row=16 + len(workers))
        chart.add_data(data, titles_from_data=True)
        sheet.add_chart(chart, f"B2")

        # Set column width for column A
        sheet.column_dimensions['A'].width = 20


# Main function
def main():
    workbook = Workbook()
    year = datetime.now().year
    workbook.remove(workbook.active)  # Remove default sheet

    # Example workers list
    workers = ['Ashani Akatugba', 'Bassey Nton N.', 'Anointing Obiora', 'Denis Esikpong',
               'Maurice Tchouncha', 'Nnamso Glory', 'Joseph Onyinye', 'Wilmont Agbor', 'Emmamuel Aniefiok',
               'Ofonime Ufot', 'Blessing Edet', 'Victor Emordi', 'Chris Okoh', 'Christopher Monday']

    total_days_dict = {}

    # Create sheets for each month
    for month in list(calendar.month_name)[1:]:
        sheet, avg_col, remarks_col, total_days = create_month_sheet(workbook, month, year)
        total_days_dict[month] = total_days

        # Add worker names
        for i, worker in enumerate(workers, start=17):
            sheet.cell(row=i, column=1).value = worker

    # Add attendance summary and charts
    add_summary_and_chart(workbook, workers, total_days_dict)

    # Save the workbook
    workbook.save("2026_Trainees_DAILY_MEETING_Attendance.xlsx")

if __name__ == "__main__":
    main()
