from openpyxl import Workbook
from openpyxl import load_workbook
from zhou_bao_files import ZhouBaoSource
#from print_date_of_week import DateOfWeek

zb_source_files = ZhouBaoSource()
current_month_dir = zb_source_files.current_month_dir
zb_source_files_sheets = zb_source_files.work_week.sheet_name_date()
zb_bian_jie_zhi = zb_source_files.work_week.bian_jie_zhi
hui_zong_qs = []

def qing_suan(wb_qs_name, ws_qs_names, hui_zong_qs):
    l_tag = True
    jin_e_column = 10
    chu_li_column = 36
    wb_qs = load_workbook(wb_qs_name)

    for i in ws_qs_names:
        ws_qs = wb_qs[i]
        hui_zong_qs_sheet = [0] * 6
        j = 4
        while j <= ws_qs.max_row:
            if ws_qs.cell(row = j, column = jin_e_column).value != None and l_tag:
                if ws_qs.cell(row = j, column = chu_li_column).value == "未处理":
                    hui_zong_qs_sheet[3] += 1
                    hui_zong_qs_sheet[5] += float(ws_qs.cell(row = j, column = jin_e_column).value)
                hui_zong_qs_sheet[0] += 1
                hui_zong_qs_sheet[2] += float(ws_qs.cell(row = j, column = jin_e_column).value)
            elif ws_qs.cell(row = j, column = jin_e_column).value == None and l_tag:
                j += 3
                jin_e_column += 1
                chu_li_column += 3
                l_tag = False
            elif not l_tag:
                if ws_qs.cell(row = j, column = jin_e_column).value != None:
                    if ws_qs.cell(row = j, column = chu_li_column).value == "未处理":
                        hui_zong_qs_sheet[4] += 1
                        hui_zong_qs_sheet[5] += float(ws_qs.cell(row = j, column = jin_e_column).value)
                    hui_zong_qs_sheet[1] += 1
                    hui_zong_qs_sheet[2] += float(ws_qs.cell(row = j, column = jin_e_column).value)
            j += 1

        hui_zong_qs.append(hui_zong_qs_sheet)

if zb_source_files.work_week.diferent_month:
    qing_suan(zb_source_files.last_week_qs, zb_source_files_sheets[:zb_bian_jie_zhi], hui_zong_qs)
    qing_suan(zb_source_files.current_week_qs, zb_source_files_sheets[zb_bian_jie_zhi:], hui_zong_qs)
else:
    qing_suan(zb_source_files.current_week_qs, zb_source_files_sheets[:zb_bian_jie_zhi], hui_zong_qs)

print(hui_zong_qs)
