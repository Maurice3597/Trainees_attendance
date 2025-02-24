import calendar
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.chart import BarChart,LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Color, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule

# Constants
DAYS_OF_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ATTENDANCE_OPTIONS = ["P", "L", "E", "A","N"]
REMARKS = {90: "Excellent", 75: "Good", 50: "Average", 0: "Poor"}

# Function to create attendance sheet for a month
def create_month_sheet(workbook, month, year):
    sheet = workbook.create_sheet(title=month)
    # sheet.freeze_panes = "A15"  # Freeze rows up to row 14

    # Set the header row
    sheet["D15"] = "Name"
    current_date = datetime(year, list(calendar.month_name).index(month), 1)
    col = 5
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
def add_summary(workbook, workers, total_days_dict):

    for sheet in workbook.worksheets:
        total_days= total_days_dict[sheet.title]
        avg_col = total_days + 13
        remarks_col = avg_col + 1

        dict_day = {
            23: {"P": 'AC', "L": 'AD', "E": 'AE', "A": 'AF', "N": 'AG'},
            22: {"P": 'AB', "L": 'AC', "E": 'AD', "A": 'AE', "N": 'AF'},
            21: {"P": 'AA', "L": 'AB', "E": 'AC', "A": 'AD', "N": 'AE'},
            20: {"P": 'Z', "L": 'AA', "E": 'AB', "A": 'AC', "N": 'AD'}
        }

        dd_opts = list(list(dict_day.values())[0].keys())  # Dropdown options
        # Create a data validation for the dropdown
        dv = DataValidation(type="list", formula1=f'"{",".join(dd_opts)}"', showDropDown=True)

        # Add the data validation to each sheet
        sheet.add_data_validation(dv)

        for row in range(17, 17 + len(workers)):
            for column in range(total_days + 6, total_days + 11):
                position = column - total_days - 6
                if total_days == 23:
                    opt_list = list(dict_day[23].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(E{row}:AA{row},${opt_list[position]}$15)"
                    sheet[f"AI{row}"].value = f"=SUM(AC{row}:AD{row})"
                    sheet[f"AJ{row}"].value = f"=ROUND(AI{row}/{0.01 * total_days}, 1)"
                    dv.add(f"E{row}:AA{row}")



                elif total_days == 22:
                    opt_list = list(dict_day[22].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(E{row}:Z{row},${opt_list[position]}$15)"
                    sheet[f"AH{row}"].value = f"=SUM(AB{row}:AC{row})"
                    sheet[f"AI{row}"].value = f"=ROUND(AH{row}/{0.01 * total_days}, 1)"
                    dv.add(f"E{row}:Z{row}")

                elif total_days == 21:
                    opt_list = list(dict_day[21].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(E{row}:Y{row},${opt_list[position]}$15)"
                    sheet[f"AG{row}"].value = f"=SUM(AA{row}:AB{row})"
                    sheet[f"AH{row}"].value = f"=ROUND(AG{row}/{0.01 * total_days}, 1)"
                    dv.add(f"E{row}:Y{row}")

                elif total_days == 20:
                    opt_list = list(dict_day[20].values())
                    sheet.cell(row=row, column=column ).value = f"=COUNTIF(E{row}:X{row},${opt_list[position]}$15)"
                    sheet[f"AF{row}"].value = f"=SUM(Z{row}:AA{row})"
                    sheet[f"AG{row}"].value = f"=ROUND(AF{row}/{0.01 * total_days}, 1)"
                    dv.add(f"E{row}:X{row}")

                else:
                    pass

            avg_cell = sheet.cell(row=row, column= avg_col)
            remark_cell = sheet.cell(row=row, column= remarks_col)
            remark_cell.value = (f'=IF({avg_cell.coordinate}>90.00,"Excellent",IF({avg_cell.coordinate}>75.00,"Good",'
                                f'IF({avg_cell.coordinate}>50.00,"Average","Poor")))')

        # Add a row for total attendance for each day
        sheet.cell(row=40, column=4).value = "Total Attendance"
        for col in range(5, total_days + 5):
            col_letter = get_column_letter(col)  # Convert column index to letter
            sheet.cell(row=40, column=col).value = (f'=COUNTIFS({col_letter}{17}:{col_letter}{len(workers)+17}, "P") +'
                                                    f'COUNTIFS({col_letter}{17}:{col_letter}{len(workers)+17}, "L")')


# Function to plot charts
def plot_charts(workbook, workers, total_days_dict):
    for sheet in workbook.worksheets:
        total_days = total_days_dict[sheet.title]
        avg_col = total_days + 13

        # Add a chart_1 for individual attendance
        chart_1 = BarChart()
        chart_1.title = f"Individual Average Daily Attendance ({sheet.title})"
        chart_1.x_axis.title = "Attendees"
        chart_1.y_axis.title = "Days_present"
        data_1 = Reference(sheet, min_col=avg_col, min_row=16, max_row=16 + len(workers))
        chart_1.add_data(data_1, titles_from_data=True)
        sheet.add_chart(chart_1, f"E2")
        sheet.column_dimensions['D'].width = 20 

        # Add line chart_1 for collective attendance
        chart_2 = BarChart()
        chart_2.title = "Total Daily Attendance Over"
        chart_2.x_axis.title = "Dates"
        chart_2.y_axis.title = "Total Attendance"
        # Define the data for the chart_1
        data_2 = Reference(sheet,
                         min_col=2,  # Starting from column B
                         min_row=40,
                         max_col=total_days + 1,  # Up to column X (which is 24)
                         max_row=40)  # Only the 40th row

        # Define the x-axis labels (dates) from row 1 (assuming headers are in row 1)
        dates = Reference(sheet,
                          min_col=2,
                          min_row=16,
                          max_col=total_days + 1,  # Same number as max_col data
                          max_row=16)  # Assume headers are in row
        chart_2.add_data(data_2, from_rows=False)
        chart_2.set_categories(dates)
        sheet.add_chart(chart_2, f"Q2")

def format_sheet(workbook, workers, total_days_dict):
    for sheet in workbook.worksheets:
        total_days = total_days_dict[sheet.title]
        # Create a light blue fill pattern
        light_blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
        for column in range(5, total_days + 5):
            col_letter = get_column_letter(column)
            cell_range = f"{col_letter}17:{col_letter}{17+len(workers)}"
            #for cell_range in cell_ranges:# Change this range as needed
            # Add conditional formatting rule
            sheet.conditional_formatting.add(
                cell_range,
                CellIsRule(operator='equal', formula=['"P"'], fill=light_blue_fill)
            )



# Main function
def main():
    workbook = Workbook()
    year = datetime.now().year
    workbook.remove(workbook.active)  # Remove default sheet

    # Example workers list
    workers = ['Ashani Akatugba', 'Bassey Nton N.', 'Anointing Obiora', 'Denis Esikpong',
               'Maurice Tchouncha', 'Nnamso Glory', 'Joseph Onyinye', 'Wilmont Agbor', 'Emmamuel Aniefiok',
               'Ofonime Ufot', 'Blessing Edet', 'Victor Emordi', 'Chris Okoh', 'Christopher Monday', "", "", "", "", ""]

    total_days_dict = {}

    # Create sheets for each month
    for month in list(calendar.month_name)[1:]:
        sheet, avg_col, remarks_col, total_days = create_month_sheet(workbook, month, year)
        total_days_dict[month] = total_days

        # Add worker names
        for i, worker in enumerate(workers, start=17):
            sheet.cell(row=i, column=4).value = worker

    # Add attendance summary
    add_summary(workbook, workers, total_days_dict)

    # Add Charts
    plot_charts(workbook, workers, total_days_dict)

    # Add formatting
    format_sheet(workbook, workers, total_days_dict)

    # Save the workbook
    workbook.save("ADJ_ATTENDANCE_PROTOTYPE.xlsx")

if __name__ == "__main__":
    main()
