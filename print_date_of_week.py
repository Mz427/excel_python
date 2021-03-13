from datetime import date, timedelta
import calendar

class DateOfWeek:
    def __init__(self):
        calendar.setfirstweekday(0)
        self.diferent_month = False
        self.bian_jie_zhi = 7
        self.the_work_day = date.today()
        self.current_week = self.the_work_day.weekday()
        if self.current_week != 6:
            self.the_work_day += timedelta(days = -(self.current_week + 1))
        self.the_month_calendar = calendar.Calendar(0).monthdatescalendar(self.the_work_day.year, self.the_work_day.month)

    def get_week_date(self):
        for week_i in self.the_month_calendar:
            for day_i in week_i:
                if self.the_work_day == day_i:
                    return week_i

    def sheet_name_date(self):
        sheet_name_day = []
        this_work_week = self.get_week_date()

        for day_i in this_work_week:
            sheet_name_day.append(day_i.strftime("%d"))

        return sheet_name_day

    def get_zb_source_suffix(self):
        sheet_name_day = self.sheet_name_date()
        this_work_week = self.get_week_date()
        zb_source_suffix = [this_work_week[6].strftime("%Y%m%d")]

        if sheet_name_day[0] > sheet_name_day[6]:
            self.diferent_month = True
            for i, j in this_work_week, range(1, 7):
                if i.strftime("%d") == max(sheet_name_day):
                    self.bian_jie_zhi = j
                    zb_source_suffix.append(i.strftime("%Y%m%d"))

        return zb_source_suffix
