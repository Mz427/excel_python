from datetime import date, timedelta
import calendar



the_work_day = date.fromisoformat('2021-04-04')
current_week = the_work_day.weekday()
if current_week == 0:
    the_work_day += timedelta(days = -1)
the_month_calendar = calendar.Calendar(0).monthdatescalendar(the_work_day.year, the_work_day.month)
calendar.setfirstweekday(0)

def get_week_date():
    if the_work_day.weekday() == 6:
        for week_i in the_month_calendar:
            for day_i in week_i:
                if the_work_day == day_i:
                    return week_i

def sheet_name_date():
    sheet_name_day = []
    this_work_week = get_week_date()

    for day_i in this_work_week:
        sheet_name_day.append(day_i.strftime("%d"))

    return sheet_name_day

def get_zb_source_suffix():
    sheet_name_day = sheet_name_date()
    this_work_week = get_week_date()
    zb_source_suffix = []
    zb_source_suffix.append(this_work_week[6].strftime("%Y%m%d"))

    if sheet_name_day[0] > sheet_name_day[6]:
        for i in this_work_week:
            if i.strftime("%d") == max(sheet_name_day):
                zb_source_suffix.append(i.strftime("%Y%m%d"))

    return zb_source_suffix

print(the_work_day)
print(the_month_calendar)
print(current_week)
print(get_week_date())
print(sheet_name_date())
print(get_zb_source_suffix())
