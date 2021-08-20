import csv
from utils import utils


def count_data_from_excel(sheet1,
                          rows_numbers,
                          patient_column,
                          medicine_numbers_column,
                          project_column,
                          room_column,
                          doctor_column):
    room_dict = {}
    for row in range(rows_numbers):
        if row == 0:
            continue
        # 获取医生开药数量
        doctor_medicine_numbers = float(sheet1[row][medicine_numbers_column])
        # 获取本条数据病人姓名
        patient_name = sheet1[row][patient_column]
        # 获取本条数据医生姓名
        doctor_name = sheet1[row][doctor_column]
        # 获取药房名称
        room = sheet1[row][room_column]
        # 获取科室名称
        project = sheet1[row][project_column]
        room_dict = utils.count_room(room_dict, room, project, doctor_name, doctor_medicine_numbers, patient_name)
    return room_dict


class Data:
    def __init__(self, file_path):
        # 获取 excel 数据
        try:
            with open(file_path, 'r') as f:
                csv_data = csv.reader(f, delimiter='\t')
                data = []
                for line in csv_data:
                    data.append(line)
        except Exception:
            raise Exception("错误：输入的 excel 文件不存在，或文件格式不为 csv")
        # 获取 sheet1 数据
        sheet1 = data
        # 获取 sheet1 行数
        rows_numbers = len(data)
        # 获取列数量
        line1_info = data[0]
        line2_info = data[1]
        patient_column, medicine_column, medicine_numbers_column, \
        project_column, room_column, doctor_column = utils.get_data_column(line1_info)
        self.medicine_info, self.medicine_name = utils.get_medicine_info(line2_info, medicine_column)
        self.room_dict = count_data_from_excel(sheet1,
                                               rows_numbers,
                                               patient_column,
                                               medicine_numbers_column,
                                               project_column,
                                               room_column,
                                               doctor_column)
