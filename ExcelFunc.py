import os
import openpyxl


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


arrTchs = ReadTeacherName()
strScheduleDir = "D:\document\赣东职院\中职\课表"
strScheduleName = "2023下半学期单双周课表.xlsx"
strSheetName = "教师名单"
strScheduleSheetName = "4.13春季课表"
nMinRow = 4
nMaxRow = 24
nMinCol = 0
nMaxCol = 41


def ColIndex2Num(strIdx):
    # column 0 -> A, 1 -> B
    return openpyxl.utils.column_index_from_string(strIdx)


class CurrInfo:
    def __init__(self, strName, strClass, strTchName) -> None:
        self.strName = strName
        self.strClass = strClass
        self.strTchName = strTchName


    strName = ""
    strClass = ""
    strTchName = ""
    nCount = 0


    def IsSame(self, cr):
        return self.strName == cr.strName and self.strClass == cr.strClass


    def GetKey(self):
        return self.strName + self.strClass


class TchData:
    strTchName = ""
    dicCurrs = {}


class ClsData:
    strClsName = ""
    dicCurrs = {}


def Space_NextlineFilter(strOrg):
    if (type(strOrg) == type("")):
        strTmp = strOrg.replace(" ", "")
        strTmp = strTmp.replace("\n", "")
        strTmp = strTmp.replace("\r", "")
        return strTmp
    else:
        return strOrg


def CellVal(cell):
    return Space_NextlineFilter(cell.value)


def OpenScheduleWB():
    os.chdir(strScheduleDir)
    return openpyxl.load_workbook(strScheduleName)


# ret CurrName, TchName
def GetCurrNameAndTchName(strOrgName):
    if type(strOrgName) !=  type(""):
        return strOrgName
    i = 0
    nTchNum = len(arrTchs)
    while i < len(arrTchs):
        strTchName = arrTchs[i]
        nPos = strOrgName.find(strTchName)
        if nPos != -1:
            return strOrgName[0: nPos], strTchName
        i = i + 1


def IncCurr(TchMapCurr, ClsMapCurr, strCurrName, strTchName, strClsName):
    # create if not exist
    if not TchMapCurr.has_key(strTchName):
        sNewTchData = TchData()
        sNewTchData.strTchName = strTchName
        sNewCurr = CurrInfo(strCurrName, strClsName, strTchName)
        sNewCurr.nCount = 1
        sNewTchData.dicCurrs[sNewCurr.GetKey()] = sNewCurr
        TchMapCurr[strTchName] = sNewTchData

    if not ClsMapCurr.has_key(strTchName):
        sNewClsData = ClsData()
        sNewClsData.strClsName = strClsName
        sNewCurr = CurrInfo(strCurrName, strClsName, strTchName)
        sNewCurr.nCount = 1
        sNewClsData.dicCurrs[sNewCurr.GetKey()] = sNewCurr
        ClsMapCurr[strClsName] = sNewClsData

    return


def ScheduleStatistic():
    workbook = OpenScheduleWB()

    TchMapCurr = {}     # 教师及其课程
    ClsMapCurr = {}     # 班级及其课程

    # 遍历课程表
    sheet = workbook[strScheduleSheetName]
    i = nMinRow
    j = nMinCol
    while i <= nMaxRow:
        while j <= nMaxCol:
            # inc class count
            strClsName = CellVal(sheet[i][0])
            strCurrName, strTchName = GetCurrNameAndTchName(CellVal(sheet[i][j])))
            if strCurrName != None:
                IncCurr(TchMapCurr, ClsMapCurr, strCurrName, strTchName, strClsName)
            j = j + 1
        i = i + 1
        j = nMinCol
    # while i <= sheet.max_row:
    #     sRow = sheet[i]
    #     i = i + 1
    #     for cell in sRow:
    #         print(cell.value)








#ReadTeacherName()
ScheduleStatistic()