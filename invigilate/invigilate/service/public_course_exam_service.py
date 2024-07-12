import xlrd
from xlutils.copy import copy
from datetime import datetime

from invigilate.entity.public_course_exam import PublicCourseExam


# 公共课期末考试业务
class PublicCourseExamService:

    def __init__(self):
        self.exam_list = []

    def read_exam_in_xls(self):
        # 文件路径
        path = "./xls/2022-2023-2/期末考试安排表-公共课.xls"
        # 工作表序号，从0开始
        sheet_index = 0
        # 数据开始行数，从0开始
        data_start_row = 3
        # 数据结束行数，从0开始
        data_end_row = 262

        # 日期列数，从0开始
        date_col = 3

        # 课程列数，从0开始
        course_col = 4

        # 教师列数，从0开始
        teacher_col = 6

        # 时间列数，从0开始
        time_col = 11

        # 地点列数，从0开始
        place_col = 13

        # 校区列数，从0开始
        campus_col = 14

        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(sheet_index)

        exam_dict = {}
        for i in range(data_start_row, data_end_row + 1):
            row = sheet.row(i)

            date = xlrd.xldate_as_datetime(row[date_col].value, wb.datemode).strftime("%Y-%m-%d")
            time = row[time_col].value.replace("：", ":")

            # key = 日期 + 时间 + 考场
            key = date + time + row[place_col].value
            if key not in exam_dict:
                start = datetime.strptime(date + " " + time.split("-")[0], "%Y-%m-%d %H:%M")
                end = datetime.strptime(date + " " + time.split("-")[1], "%Y-%m-%d %H:%M")

                exam = PublicCourseExam()
                exam.course = row[course_col].value
                exam.campus = row[campus_col].value
                exam.times = (start, end)

                exam_dict[key] = exam
            else:
                exam = exam_dict[key]

            exam.row.append(i)
            if row[teacher_col].value != "":
                exam.teacher.add(row[teacher_col].value)

        self.exam_list = list(exam_dict.values())

    def write_in_xls(self):
        # 文件路径
        path = "./xls/2022-2023-2/期末考试安排表-公共课.xls"
        # 工作表序号，从0开始
        sheet_index = 0

        # 保存路径
        save_path = "./xls/2022-2023-2/期末考试安排表-公共课-已安排.xls"

        # 标题行数，从0开始
        title_row = 2

        # 监考老师列数，从0开始
        invigilator_col = 19

        wb = copy(xlrd.open_workbook(path))
        sheet = wb.get_sheet(sheet_index)

        sheet.write(title_row, invigilator_col, "监考老师")

        for exam in self.exam_list:
            for row in exam.row:
                sheet.write(row, invigilator_col, "、".join(exam.invigilator))

        wb.save(save_path)
