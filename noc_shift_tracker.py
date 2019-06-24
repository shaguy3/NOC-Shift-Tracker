import openpyxl
from datetime import date, datetime, timedelta


def add_to_db(sheet):
    db_wb = openpyxl.load_workbook('db.xlsx')
    db_sheet = db_wb[db_wb.sheetnames[0]]
    for row in range(4, sheet.max_row + 1):
        db_sheet.append((sheet['A' + str(row)].value[-5:] + '/' + str(date.today().year)[2:],
                         sheet['E' + str(row)].value,
                         sheet['F' + str(row)].value))

    db_wb.save('db.xlsx')


def purge_db():
    db_wb = openpyxl.load_workbook('db.xlsx')
    db_sheet = db_wb[db_wb.sheetnames[0]]

    db_sheet.delete_cols(1, 3)

    db_sheet['A1'] = 'Date of shift'
    db_sheet['B1'] = 'Shift start'
    db_sheet['C1'] = 'Shift end'

    db_wb.save('db.xlsx')


def get_months_shifts(month):
    db_wb = openpyxl.load_workbook('db.xlsx')
    db_sheet = db_wb[db_wb.sheetnames[0]]
    this_months_shifts = []
    current_row = 2

    for row in range(current_row, db_sheet.max_row + 1):
        if db_sheet['A' + str(current_row)].value[3:5] == month:
            shift_date = db_sheet['A' + str(current_row)].value
            shift_start_time = db_sheet['B' + str(current_row)].value
            shift_end_time = db_sheet['C' + str(current_row)].value

            shift_start_datetime = datetime.strptime(shift_date + shift_start_time, '%d/%m/%y%H:%M')

            if int(shift_end_time[0:2]) < int(shift_start_time[0:2]):
                shift_end_datetime = datetime.strptime(shift_date, '%d/%m/%y') + \
                                     timedelta(days=1,
                                               hours=int(shift_end_time[0:2]),
                                               minutes=int(shift_end_time[3:]))
            else:
                shift_end_datetime = datetime.strptime(shift_date, '%d/%m/%y') + \
                                     timedelta(hours=int(shift_end_time[0:2]),
                                               minutes=int(shift_end_time[3:]))

            this_shift = [shift_start_datetime, shift_end_datetime]
            this_months_shifts.append(this_shift)
            current_row += 1

    return this_months_shifts


def split_to_minutes(shift):
    splitted_shift = []
    current_minute = shift[0]
    while current_minute <= shift[1]:
        splitted_shift.append(current_minute)
        current_minute += timedelta(minutes=1)

    return splitted_shift


def minute_cat(current_minute):
    if current_minute.weekday() == 4:
        if current_minute < datetime(current_minute.year, current_minute.month, current_minute.day, 6):
            return 2  # Night
        elif current_minute < datetime(current_minute.year, current_minute.month, current_minute.day, 16):
            return 3  # Friday morning
        else:
            return 5  # Saturday
    elif current_minute.weekday() == 5:
        if current_minute < datetime(current_minute.year, current_minute.month, current_minute.day, 17):
            return 5  # Saturday
        elif current_minute < datetime(current_minute.year, current_minute.month, current_minute.day, 22):
            return 4  # Saturday night
        else:
            return 2  # Night
    else:
        if current_minute < datetime(current_minute.year, current_minute.month, current_minute.day, 6):
            return 2  # Night
        elif current_minute < datetime(current_minute.year, current_minute.month, current_minute.day, 22):
            return 1  # Day
        else:
            return 2  # Night


def categorize_shift(shift):
    splited_shift = split_to_minutes(shift)
    minutes_worked = 0  # For the overtime calculations
    shift_counters = [0, 0, 0, 0, 0]  # [day, night, friday_morning, saturday_night, saturday]

    # TODO: Go over each minute, get it's category in the week, and increment the specific category's counter.

    # TODO: If the shift enters into overtime from a regular shift, increment the overtime counter.

    # TODO: If the shift enters continues after two overtime hours, increment the extended overtime counter.

    # TODO: If the shift enters into overtime from an irregular shift, increment the same irregular shift.

    return shift_counters


def shift_calc(shift):
    shift_counters = categorize_shift(shift)


my_wb = openpyxl.load_workbook('may_2019.xlsx')
my_sheet = my_wb['Sheet1']

may_shifts = get_months_shifts('05')
for may_shift in may_shifts:
    for minute in split_to_minutes(may_shift):
        print(minute, minute_cat(minute))




















