import pickle

from invigilate.service.obligation_invigilate_service import ObligationInvigilateService
from invigilate.service.public_course_exam_service import PublicCourseExamService
from invigilate.service.rotate_invigilate_service import RotateInvigilateService
from invigilate.service.speciality_course_exam_service import SpecialityCourseExamService


# 分配业务
class AllocationService:

    def __init__(self):
        self.teacher_invigilate_time_dict = {}
        self.load_in_pickle()

        self.oi_service = ObligationInvigilateService()
        self.oi_service.load_in_pickle()

        self.ri_service = RotateInvigilateService()
        self.ri_service.load_in_pickle()

        self.pc_exam_service = PublicCourseExamService()
        self.sc_exam_service = SpecialityCourseExamService()

    def save_to_pickle(self):
        with open("service_pickle/allocation_service.pickle", "wb") as file:
            pickle.dump(self, file)

    def load_in_pickle(self):
        with open("service_pickle/allocation_service.pickle", "rb") as file:
            obj = pickle.load(file)
            self.teacher_invigilate_time_dict = obj.teacher_invigilate_time_dict

    def run(self):
        while True:
            print(f"义务监考名单更新时间：{self.oi_service.update_at}")
            print(f"轮流监考名单更新时间：{self.ri_service.update_at}")
            print()

            flag = int(input("1.公共课期末考试\n2.专业课期末考试\n3.研究生考试\n4.四六级考试\n5.更新义务监考名单\n6.保存并退出\n请选择考试类型："))
            if flag == 1:
                self.public_course_exam()
            elif flag == 2:
                self.speciality_course_exam()
            elif flag == 3:
                self.postgraduate_exam()
            elif flag == 4:
                self.cet_exam()
            elif flag == 5:
                self.update_obligation_invigilate()
            elif flag == 6:
                self.save_and_exit()
                break
            else:
                print("没有该类型，请重新选择")

    def public_course_exam(self):
        self.pc_exam_service.read_exam_in_xls()

        # 优先安排同校区同课程的其他任课教师监考
        for exam in self.pc_exam_service.exam_list:
            for teacher in self.oi_service.teacher_list:
                if (teacher.count > 0
                        and teacher.campus == exam.campus
                        and teacher.name not in exam.teacher
                        and teacher.name not in exam.invigilator
                        and exam.course in teacher.course
                        and not self.is_overlap(teacher.name, exam.times)):
                    teacher.count -= 1
                    self.invigilate_exam(teacher, exam)
                    break

        # 安排同校区义务监考教师监考
        for i in range(0, 2):
            for exam in self.pc_exam_service.exam_list:
                if len(exam.invigilator) == i:
                    for teacher in self.oi_service.teacher_list:
                        if (teacher.count > 0
                                and teacher.campus == exam.campus
                                and teacher.name not in exam.teacher
                                and teacher.name not in exam.invigilator
                                and not self.is_overlap(teacher.name, exam.times)):
                            teacher.count -= 1
                            self.invigilate_exam(teacher, exam)
                            break

        # 莞城轮流监考教师排满
        # for i in range(0, 2):
        #     for exam in self.pc_exam_service.exam_list:
        #         if exam.campus == "莞城" and len(exam.invigilator) == i:
        #             is_allocation = False
        #             # 先从轮流监考的 temp 中取教师
        #             for j in range(0, len(self.ri_service.gc_temp_list)):
        #                 teacher = self.ri_service.gc_temp_list[j]
        #                 if (teacher.name not in exam.teacher
        #                         and teacher.name not in exam.invigilator
        #                         and not self.is_overlap(teacher.name, exam.times)):
        #                     teacher.count += 1
        #                     self.invigilate_exam(teacher, exam)
        #                     self.ri_service.gc_temp_list.pop(j)
        #                     is_allocation = True
        #                     break
        #             if is_allocation:
        #                 continue
        #             # temp 中没有取到教师，从队列中取
        #             for j in range(0, len(self.ri_service.gc_teacher_list)):
        #                 teacher = self.ri_service.get_gc_teacher()
        #                 if (teacher.name not in exam.teacher
        #                         and teacher.name not in exam.invigilator
        #                         and not self.is_overlap(teacher.name, exam.times)):
        #                     teacher.count += 1
        #                     self.invigilate_exam(teacher, exam)
        #                     break
        #                 else:
        #                     self.ri_service.gc_temp_list.append(teacher)

        # 莞城轮流监考教师排满
        for i in range(0, 2):
            for exam in self.pc_exam_service.exam_list:
                if exam.campus == "莞城" and len(exam.invigilator) == i:
                    for j in range(0, len(self.ri_service.gc_teacher_list)):
                        teacher = self.ri_service.gc_teacher_list[j]
                        if (teacher.name not in exam.teacher
                                and teacher.name not in exam.invigilator
                                and not self.is_overlap(teacher.name, exam.times)):
                            teacher.count += 1
                            self.invigilate_exam(teacher, exam)
                            self.ri_service.gc_teacher_list.pop(j)
                            self.ri_service.gc_teacher_list.append(teacher)
                            break

        for i in range(0, 2):
            for exam in self.pc_exam_service.exam_list:
                if len(exam.invigilator) == i:
                    for j in range(0, len(self.ri_service.ssh_teacher_list)):
                        teacher = self.ri_service.ssh_teacher_list[j]
                        if (teacher.name not in exam.teacher
                                and teacher.name not in exam.invigilator
                                and not self.is_overlap(teacher.name, exam.times)):
                            teacher.count += 1
                            if teacher.campus != exam.campus:
                                teacher.offsite_count += 1
                            self.invigilate_exam(teacher, exam)
                            self.ri_service.ssh_teacher_list.pop(j)
                            self.ri_service.ssh_teacher_list.append(teacher)
                            break

        # print("=======================================================================================================")
        # for exam in self.pc_exam_service.exam_list:
        #     print(vars(exam))
        print("=======================================================================================================")
        for teacher in self.ri_service.gc_teacher_list:
            print(vars(teacher))
        print("=======================================================================================================")
        for teacher in self.ri_service.ssh_teacher_list:
            print(vars(teacher))

        self.pc_exam_service.write_in_xls()

    def speciality_course_exam(self):
        self.sc_exam_service.read_exam_in_xls()

        for exam in self.sc_exam_service.exam_list:
            for teacher_name in exam.teacher:
                teacher = self.oi_service.teacher_dict.get(teacher_name)
                if (teacher
                        and teacher.count > 0
                        and teacher.name not in exam.invigilator
                        and not self.is_overlap(teacher.name, exam.times)):
                    teacher.count -= 1
                    self.invigilate_exam(teacher, exam)
                    break

        for i in range(0, 2):
            for exam in self.sc_exam_service.exam_list:
                if len(exam.invigilator) == i:
                    for teacher in self.oi_service.teacher_list:
                        if (teacher.count > 0
                                and teacher.campus == exam.campus
                                and teacher.name not in exam.invigilator
                                and not self.is_overlap(teacher.name, exam.times)):
                            teacher.count -= 1
                            self.invigilate_exam(teacher, exam)
                            break

        for i in range(0, 2):
            for exam in self.sc_exam_service.exam_list:
                if len(exam.invigilator) == i:
                    for j in range(0, len(self.ri_service.ssh_teacher_list)):
                        teacher = self.ri_service.ssh_teacher_list[j]
                        if (teacher.name not in exam.invigilator
                                and not self.is_overlap(teacher.name, exam.times)):
                            teacher.count += 1
                            self.invigilate_exam(teacher, exam)
                            self.ri_service.ssh_teacher_list.pop(j)
                            self.ri_service.ssh_teacher_list.append(teacher)
                            break

        # print("=======================================================================================================")
        # for exam in self.sc_exam_service.exam_list:
        #     print(vars(exam))

        self.sc_exam_service.write_in_xls()

    def postgraduate_exam(self):
        print("研究生考试")

    def cet_exam(self):
        print("四六级考试")

    def update_obligation_invigilate(self):
        self.oi_service.read_teacher_in_xls()
        self.teacher_invigilate_time_dict = {}

    def save_and_exit(self):
        self.oi_service.save_to_pickle()
        self.ri_service.save_to_pickle()
        self.save_to_pickle()

    def is_overlap(self, teacher_name, time1):
        if teacher_name not in self.teacher_invigilate_time_dict:
            return False

        for time2 in self.teacher_invigilate_time_dict[teacher_name]:
            if max(time1[0], time2[0]) <= min(time1[1], time2[1]):
                return True

        return False

    def invigilate_exam(self, teacher, exam):
        exam.invigilator.add(teacher.name)

        invigilate_time = self.teacher_invigilate_time_dict.get(teacher.name)
        if not invigilate_time:
            invigilate_time = []
            self.teacher_invigilate_time_dict[teacher.name] = invigilate_time
        invigilate_time.append(exam.times)
