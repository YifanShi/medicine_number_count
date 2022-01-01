import os
import time
from tkinter import messagebox
import xlwt


def _write_titile(excel, line, row, t1, t2, style):
    if row % 2:
        excel.write(line, row, t1, style)
    else:
        excel.write(line, row, t2, style)


def write_to_excel(room_dict, medicine_name, medicine_info, count_time):
    # 创建一个表格
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个 sheet
    worksheet = workbook.add_sheet('统计表')
    # 设置字体对齐样式
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED
    # , HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM,
    # VERT_JUSTIFIED, VERT_DISTRIBUTED
    # 定义边框样式
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40
    # 定义字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '黑体'
    # 生成正文样式
    style = xlwt.XFStyle()  # Create Style
    style.alignment = alignment  # Add Alignment to Style
    style.borders = borders  # Add Borders to Style
    # 生成抬头样式
    style_titile = xlwt.XFStyle()
    style_titile.alignment = alignment
    style_titile.borders = borders
    style_titile.font = font
    line = 0
    # 遍历科室字典
    medicine_total_numbers = 0
    for room_name in room_dict:
        for project_name in room_dict[room_name]:
            medicine_total_numbers += room_dict[room_name][project_name]["数量"]
    # 生成表格抬头
    sheet_title = medicine_info + '(' + count_time + ')' + '-用量统计：' + \
                  str(medicine_total_numbers)
    worksheet.write_merge(line, line, 0, 9, sheet_title, style_titile)
    line += 2

    for room_name in room_dict:
        for project in room_dict[room_name]:
            # 参数对应 行, 列, 值
            project_medicine_numbers = str(room_dict[room_name][project]["数量"])
            project_title = room_name + '-' + project + '-' + medicine_info \
                            + '(' + count_time + ')' \
                            + '-用量统计：' + project_medicine_numbers
            # 写入科室抬头
            worksheet.write_merge(line, line, 0, 9, project_title, style_titile)
            line += 1
            # 写入病人和数量抬头
            for row in range(0, 10):
                _write_titile(worksheet, line, row, "患者姓名", "数量", style_titile)
            line += 1
            column = 0
            # 遍历病人写入科室病人信息
            for patient in room_dict[room_name][project]['病人']:
                numbers = room_dict[room_name][project]["病人"][patient]
                worksheet.write(line, column, patient, style)
                column += 1
                worksheet.write(line, column, numbers, style)
                column += 1
                if column == 10:
                    column = 0
                    line += 1
            line += 1
            column = 0
            # 写入科室开单医生和数量
            for row in range(0, 10):
                _write_titile(worksheet, line, row, "开单医生", "数量", style_titile)
            line += 1
            # 遍历医生写入科室医生信息
            for doctor in room_dict[room_name][project]["医生"]:
                numbers = room_dict[room_name][project]["医生"][doctor]
                worksheet.write(line, column, doctor, style)
                column += 1
                worksheet.write(line, column, numbers, style)
                column += 1
                if column == 10:
                    column = 0
                    line += 1
            line += 3
    # 获取当前时间
    now = time.strftime("%Y%m%d%H%M%S", time.localtime())
    # 判断文件夹是否存在
    output_dir = '表格导出'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # 定义表单名称
    if count_time == "":
        excel_output_name = './表格导出/' + medicine_name + '-' + now + '.xls'
    else:
        excel_output_name = './表格导出/' + medicine_name + '-' + count_time + '-' + now + '.xls'
    # 保存表格
    workbook.save(excel_output_name)
    message = '数据统计完成, 导出文件名称为：' + excel_output_name
    messagebox.showinfo(title='提示', message=message)


def count_room(room_dict, room, project, doctor_name, doctor_medicine_numbers, patient_name):
    if room_dict.__contains__(room):
        pass
    else:
        room_dict[room] = {}
    # 判断该药房科室是否存在
    if room_dict[room].__contains__(project):
        pass
    else:
        room_dict[room][project] = {}
        room_dict[room][project]["医生"] = {}
        room_dict[room][project]["病人"] = {}
        room_dict[room][project]["数量"] = 0
    # 判断医生是否存在
    if room_dict[room][project]["医生"].__contains__(doctor_name):
        room_dict[room][project]["医生"][doctor_name] += doctor_medicine_numbers
    else:
        room_dict[room][project]["医生"][doctor_name] = doctor_medicine_numbers
    # 判断病人是否存在
    if room_dict[room][project]["病人"].__contains__(patient_name):
        room_dict[room][project]["病人"][patient_name] += doctor_medicine_numbers
    else:
        room_dict[room][project]["病人"][patient_name] = doctor_medicine_numbers
    # 统计科室药物总量
    room_dict[room][project]["数量"] += doctor_medicine_numbers
    return room_dict


def _check_data_info(patient_column, medicine_column, medicine_numbers_column, project_column,
                     room_column, doctor_column):
    if not all([patient_column, medicine_column, medicine_numbers_column, project_column,
                room_column, doctor_column]):
        return False
    return True


def get_data_column(line1_info):
    line1_info_len = len(line1_info)
    patient_column = None
    medicine_column = None
    medicine_numbers_column = None
    project_column = None
    room_column = None
    doctor_column = None
    for number in range(0, line1_info_len):
        # 获取病人列
        if line1_info[number] == '姓名':
            patient_column = number
            # 获取药剂列
        elif line1_info[number] == '项目名称':
            medicine_column = number
            # 获取药剂数量列
        elif line1_info[number] == '数量':
            medicine_numbers_column = number
            # 获取科室列
        elif line1_info[number] == '开单科室':
            project_column = number
            # 获取开单医生列
        elif line1_info[number] == '开单医生':
            doctor_column = number
        elif line1_info[number] == '执行科室':
            room_column = number
    if _check_data_info(patient_column, medicine_column, medicine_numbers_column,
                        project_column, room_column, doctor_column):
        return patient_column, medicine_column, medicine_numbers_column, \
               project_column, room_column, doctor_column
    else:
        messagebox.showinfo(title='提示', message="错误：请确认选择的 excel 表格中需要统计的列名为:"
                                                "姓名,项目名称,数量,开单科室，开单医生，执行科室")
        raise Exception("错误：请确认选择的 excel 表格中需要统计的列名为:"
                        " 姓名, 项目名称, 数量, 开单科室, 开单医生, 执行科室")


def get_medicine_info(line2_info, medicine_column):
    medicine_info = line2_info[medicine_column]
    medicine_name = medicine_info.split()[0].replace('/','_')
    return medicine_info, medicine_name
