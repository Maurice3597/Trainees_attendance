import calendar
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# Constants
DAYS_OF_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ATTENDANCE_OPTIONS = ["P", "L", "E", "U", "N"]
REMARKS = {90: "Excellent", 75: "Good", 50: "Average", 0: "Poor"}


def create_month_sheet(workbook, month, year):
    sheet = workbook.create_sheet(title=month)
    sheet["A15"] = "Name"

    current_date = datetime(year, list(calendar.month_name).index(month), 1)
    col = 2
    total_days = 0

    while current_date.month == list(calendar.month_name).index(month):
        if current_date.strftime("%A") in DAYS_OF_WEEK:
            total_days += 1
            sheet.cell(row=15, column=col, value=current_date.strftime("%a")).alignment = Alignment(horizontal="center")
            sheet.cell(row=16, column=col, value=current_date.strftime("%d-%b")).alignment = Alignment(
                horizontal="center")
            col += 1
        current_date += timedelta(days=1)

    # Add dropdown for attendance options
    dv = DataValidation(type="list", formula1=f'"{','.join(ATTENDANCE_OPTIONS)}"', showDropDown=True)
    sheet.add_data_validation(dv)
    dv.add(f"B17:{chr(64 + col)}100")  # Apply dropdown to attendance columns

    # Add summary columns
    attendance_col, avg_col, remarks_col = col, col + 1, col + 2
    sheet.cell(row=15, column=attendance_col, value="Attendance")
    sheet.cell(row=15, column=avg_col, value="Average Attendance (%)")
    sheet.cell(row=15, column=remarks_col, value="Remarks")

    return sheet, avg_col, remarks_col, total_days


def add_summary_and_chart(workbook, workers, total_days_dict):
    for sheet in workbook.worksheets:
        total_days = total_days_dict[sheet.title]
        avg_col = total_days + 10
        remarks_col = avg_col + 1

        for row in range(17, 17 + len(workers)):
            sheet[f"{chr(65 + total_days + 7)}{row}"].value = f"=SUM(B{row}:{chr(64 + total_days)}{row})"
            sheet[
                f"{chr(65 + total_days + 8)}{row}"].value = f"=ROUND({chr(65 + total_days + 7)}{row}/{total_days}*100, 1)"
            sheet[f"{chr(65 + total_days + 9)}{row}"].value = (
                f'=IF({chr(65 + total_days + 8)}{row}>90,"Excellent", IF({chr(65 + total_days + 8)}{row}>75,"Good",'
                f'IF({chr(65 + total_days + 8)}{row}>50,"Average","Poor")))'
            )

        # Add chart
        chart = BarChart()
        chart.title = f"Average Daily Attendance ({sheet.title})"
        data = Reference(sheet, min_col=avg_col, min_row=17, max_row=16 + len(workers))
        chart.add_data(data, titles_from_data=False)
        sheet.add_chart(chart, "B2")

        sheet.column_dimensions['A'].width = 20


def main():
    workbook = Workbook()
    year = datetime.now().year
    workbook.remove(workbook.active)

    workers = ['Ashani Akatugba', 'Bassey Nton N,', 'Anointing Obiora', 'Denis Esikpong',
               'Maurice Tchouncha', 'Nnamso Glory', 'Joseph Onyinye', 'Wilmont Agbor', 'Emmamuel Aniefiok',
               'Ofonime Ufot', 'Blessing Edet', 'Victor Emordi', 'Chris Okoh', 'Christopher Monday']

    total_days_dict = {}

    for month in list(calendar.month_name)[1:]:
        sheet, avg_col, remarks_col, total_days = create_month_sheet(workbook, month, year)
        total_days_dict[month] = total_days

        for i, worker in enumerate(workers, start=17):
            sheet.cell(row=i, column=1, value=worker)

    add_summary_and_chart(workbook, workers, total_days_dict)
    workbook.save("Go_DAILY_MEETING_Attendance_Tracker.xlsx")


if __name__ == "__main__":
    main()
