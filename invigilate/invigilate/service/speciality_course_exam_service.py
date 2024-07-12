import xlrd
from xlutils.copy import copy
from datetime import datetime

from invigilate.entity.speciality_course_exam import SpecialityCourseExam


# 专业课期末考试业务
class SpecialityCourseExamService:

    def __init__(self):
        self.exam_list = []

    def read_exam_in_xls(self):
        # 文件路径
        path = "./xls/2022-2023-2/期末考试安排表-专业课.xls"
        # 工作表序号，从0开始
        sheet_index = 0
        # 数据开始行数，从0开始
        data_row = 2

        # 科目列数，从0开始
        course_col = 0

        # 教师列数，从0开始
        teacher_col = 1

        # 日期列数，从0开始
        date_col = 6

        # 地点列数，从0开始
        place_col = 7

        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(sheet_index)

        exam_dict = {}
        for i in range(data_row, sheet.nrows):
            row = sheet.row(i)

            date = self.get_date(row[date_col].value)

            # key = 日期 + 时间 + 考场
            key = date + str(row[place_col].value)
            if key not in exam_dict:
                start = datetime.strptime(date.split(" ")[0] + " " + date.split(" ")[1].split("-")[0], "%Y-%m-%d %H:%M")
                end = datetime.strptime(date.split(" ")[0] + " " + date.split(" ")[1].split("-")[1], "%Y-%m-%d %H:%M")

                exam = SpecialityCourseExam()
                exam.course = row[course_col].value
                exam.times = (start, end)

                exam_dict[key] = exam
            else:
                exam = exam_dict[key]

            exam.row.append(i)
            exam.teacher.update(row[teacher_col].value.split("、"))

        self.exam_list = list(exam_dict.values())

    def get_date(self, value):
        value = value.replace("年", "-")
        value = value.replace("月", "-")
        value = value.replace("日", " ")
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        for weekday in weekdays:
            value = value.replace(weekday, "")
        value = value.replace("：", ":")
        return value

    def write_in_xls(self):
        # 文件路径
        path = "./xls/2022-2023-2/期末考试安排表-专业课.xls"
        # 工作表序号，从0开始
        sheet_index = 0

        # 保存路径
        save_path = "./xls/2022-2023-2/期末考试安排表-专业课-已安排.xls"

        # 标题行数，从0开始
        title_row = 1

        # 监考老师列数，从0开始
        invigilator_col = 9

        wb = copy(xlrd.open_workbook(path))
        sheet = wb.get_sheet(sheet_index)

        sheet.write(title_row, invigilator_col, "监考老师")

        for exam in self.exam_list:
            for row in exam.row:
                sheet.write(row, invigilator_col, "、".join(exam.invigilator))

        wb.save(save_path)
