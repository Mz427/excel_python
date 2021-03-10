from datetime import date
import calendar

class DateOfWeek:
    def __init__(self):
        self.the_work_day = date.today()
        self.the_month_calendar = calendar.Calendar(0).monthdatescalendar(self.the_work_day.year, self.the_work_day.month)
        self.current_week = self.the_work_day.isoweekday()
        calendar.setfirstweekday(0)

    def get_week_date(self):
        if self.current_week == 6:
            for i, week_i in zip(range(0, 5), self.the_month_calendar):
                for day_i in week_i:
                    if self.the_work_day == day_i:
                        return week_i
        elif self.current_week == 0:
            for i, week_i in zip(range(0, 5), self.the_month_calendar):
                for day_i in week_i:
                    if self.the_work_day == day_i:
                        return week_i[i-1]

    def sheet_name_date(self):
        sheet_name_day = []
        this_work_week = self.get_week_date()

        for day_i in this_work_week:
            sheet_name_day.append(day_i.day)

        return sheet_name_day
