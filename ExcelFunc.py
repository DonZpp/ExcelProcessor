import os
import openpyxl


strScheduleDir = "D:\document\赣东职院\中职\课表"
strScheduleName = "2023下半学期单双周课表.xlsx"
strSheetName = "教师名单"


def ReadTeacherName():
    os.chdir(strScheduleDir)
    workbook = openpyxl.load_workbook(strScheduleName)
    sTeachDic = {}
    sTeachSheet = workbook[strSheetName]
    i = 0
    for sTeachCell in sTeachSheet[sTeachSheet.dimensions]:
        sTeachDic[i] = sTeachCell[0].value
        i = i + 1

    # print(sTeachDic)


















