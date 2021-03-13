from print_date_of_week import DateOfWeek

class ZhouBaoSource:
    def __init__(self):
        self.work_week = DateOfWeek()
        self.date_suffix = work_week.get_zb_source_suffix()
        self.current_month_dir = "深圳运维工作记录" + self.date_suffix[1] + "/zhang/"
        self.current_week_qs = self.current_month_dir + "清算异常流水监控工作记录表" + self.date_suffix[1]
        self.current_week_fk = self.current_month_dir + "发卡网点异常流水监控" + self.date_suffix[1]
        self.current_week_ph = self.current_month_dir + "加油站报表流水平衡监控" + self.date_suffix[1]
        self.current_week_black = self.current_month_dir + "V4.9.3油站黑名单文件接收及油机文件生成情况" + self.date_suffix[1]
        self.current_week_ydjy = self.current_month_dir + "异地交易传输情况监控" + self.date_suffix[1]
        self.current_week_ydph = self.current_month_dir + "异地清分报表平衡监控" + self.date_suffix[1]
        if self.work_week.diferent_month:
            self.last_month_dir = "深圳运维工作记录" + self.date_suffix[2]
            self.last_week_qs = self.last_month_dir + "清算异常流水监控工作记录表" + self.date_suffix[2]
            self.last_week_fk = self.last_month_dir + "发卡网点异常流水监控" + self.date_suffix[2]
            self.last_week_ph = self.last_month_dir + "加油站报表流水平衡监控" + self.date_suffix[2]
            self.last_week_black = self.last_month_dir + "V4.9.3油站黑名单文件接收及油机文件生成情况" + self.date_suffix[2]
            self.last_week_ydjy = self.last_month_dir + "异地交易传输情况监控" + self.date_suffix[2]
            self.last_week_ydph = self.last_month_dir + "异地清分报表平衡监控" + self.date_suffix[2]
