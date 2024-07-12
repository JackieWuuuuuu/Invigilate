import xlrd
import pickle
from datetime import datetime

from invigilate.entity.rotate_invigilate_teacher import RotateInvigilateTeacher


# 轮流监考业务
class RotateInvigilateService:

    def __init__(self):
        self.ssh_index = 0
        self.ssh_temp_list = []
        self.ssh_teacher_list = []

        self.gc_index = 0
        self.gc_temp_list = []
        self.gc_teacher_list = []

        self.update_at = None

    def save_to_pickle(self):
        with open("service_pickle/rotate_invigilate.pickle", "wb") as file:
            pickle.dump(self, file)

    def load_in_pickle(self):
        with open("service_pickle/rotate_invigilate.pickle", "rb") as file:
            obj = pickle.load(file)

            self.ssh_index = obj.ssh_index
            self.ssh_temp_list = obj.ssh_temp_list
            self.ssh_teacher_list = obj.ssh_teacher_list

            self.gc_index = obj.gc_index
            self.gc_temp_list = obj.gc_temp_list
            self.gc_teacher_list = obj.gc_teacher_list

            self.update_at = obj.update_at

    def read_teacher_in_xls(self):
        # 文件路径
        path = "./xls/2021-2022-1/轮流监考名单.xls"
        # 工作表序号，从0开始
        sheet_index = 0
        # 数据开始行数，从0开始
        data_row = 2

        # 工号列数，从0开始
        number_col = 2

        # 姓名列数，从0开始
        name_col = 3

        # 校区列数，从0开始
        campus_col = 5

        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(sheet_index)

        ssh_teacher_list = []
        gc_teacher_list = []
        for i in range(data_row, sheet.nrows):
            row = sheet.row(i)

            teacher = RotateInvigilateTeacher()
            teacher.number = str(int(row[number_col].value))
            teacher.name = row[name_col].value
            teacher.campus = row[campus_col].value

            if teacher.campus == "松山湖":
                ssh_teacher_list.append(teacher)
            else:
                gc_teacher_list.append(teacher)

        return ssh_teacher_list, gc_teacher_list

    def update_teacher_list(self):
        self.ssh_teacher_list, self.gc_teacher_list = self.read_teacher_in_xls()
        self.update_at = datetime.now()

    def get_ssh_teacher(self):
        teacher = self.ssh_teacher_list[self.ssh_index]
        self.ssh_index = (self.ssh_index + 1) % len(self.ssh_teacher_list)
        return teacher

    def get_gc_teacher(self):
        teacher = self.gc_teacher_list[self.gc_index]
        self.gc_index = (self.gc_index + 1) % len(self.gc_teacher_list)
        return teacher
