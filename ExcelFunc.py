import os
import openpyxl


strScheduleDir = "D:\document\赣东职院\中职\课表"
strScheduleName = "2023下半学期单双周课表.xlsx"
strSheetName = "教师名单"
strScheduleSheetName = "4.13春季课表"


class CurrInfo:
    strName = ""
    strClass = ""
    nCount = 0

    def IsSame(self, cr):
        return self.strName == cr.strName and self.strClass == cr.strClass

    def GetKey(self):
        return self.strName + self.strClass


class TeacherData:
    strTeacher = ""
    dCurrInfo = {}


def OpenScheduleWB():
    os.chdir(strScheduleDir)
    return openpyxl.load_workbook(strScheduleName)


def ReadTeacherName():
    workbook = OpenScheduleWB()
    arrTeacherName = {}
    sTeachSheet = workbook[strSheetName]
    i = 0
    for sCols in sTeachSheet[sTeachSheet.dimensions]:
        for cell in sCols:
            arrTeacherName[i] = cell.value
            i = i + 1

    return arrTeacherName


def ScheduleStatistic():
    workbook = OpenScheduleWB()

    arrTeacherName = ReadTeacherName()

    # 遍历课程表
    sheet = workbook[strScheduleSheetName]
    for sCols in sheet:
        for cell in sCols:
            cell.value.find(arrTeacherName)


ReadTeacherName()
#ScheduleStatistic()












