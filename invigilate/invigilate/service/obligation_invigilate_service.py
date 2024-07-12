import xlrd
import pickle
from datetime import datetime

from invigilate.entity.obligation_invigilate_teacher import ObligationInvigilateTeacher


# 义务监考业务
class ObligationInvigilateService:

    def __init__(self):
        self.update_at = None
        self.teacher_dict = {}
        self.teacher_list = []

    def save_to_pickle(self):
        with open("service_pickle/obligation_invigilate.pickle", "wb") as file:
            pickle.dump(self, file)

    def load_in_pickle(self):
        with open("service_pickle/obligation_invigilate.pickle", "rb") as file:
            obj = pickle.load(file)
            self.update_at = obj.update_at
            self.teacher_dict = obj.teacher_dict
            self.teacher_list = list(self.teacher_dict.values())

    def read_teacher_in_xls(self):
        # 文件路径
        path = "./xls/2022-2023-2/课程表.xls"
        # 工作表序号，从0开始
        sheet_index = 0
        # 数据开始行数，从0开始
        data_row = 2

        # 课程列数，从0开始
        course_col = 2

        # 考核方式列数，从0开始
        method_col = 6

        # 姓名列数，从0开始
        name_col = 13

        # 校区列数，从0开始
        campus_col = 20

        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(sheet_index)

        teacher_dict = {}
        for i in range(data_row, sheet.nrows):
            row = sheet.row(i)

            if row[method_col].value == "考试":
                course = row[course_col].value.split("]")[1]

                teachers = row[name_col].value.split("、")
                for teacher in teachers:
                    # name = teacher.split("[")[0]
                    name = teacher.split("]")[1]
                    if name not in teacher_dict:
                        oi_teacher = ObligationInvigilateTeacher()
                        oi_teacher.name = name
                        if "莞城" in row[campus_col].value:
                            oi_teacher.campus = "莞城"

                        teacher_dict[name] = oi_teacher
                    else:
                        oi_teacher = teacher_dict[name]

                    oi_teacher.count += 1
                    oi_teacher.course.add(course)

        self.teacher_dict = teacher_dict
        self.teacher_list = list(teacher_dict.values())
        self.update_at = datetime.now()
