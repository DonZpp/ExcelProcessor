import os
import openpyxl


# 课程表
strScheduleDir = 'D:\document\GanDong Acadamy\中职\课表'
strScheduleName = '2023下半学期单双周课表.xlsx'
strScheduleSheetName = '4.13春季课表'

# 教师表
strTchWBDir = 'D:\document\GanDong Acadamy\中职'
strTchWBName = '教师名单.xlsx'
strTchSheetName = 'Sheet1'

# 输入表格
strInitDir = 'D:\document-unimportant\ExcelWorkBench'
strInitTable = '初始数据.xlsx'
strInitDataSheetName = 'Sheet'
strStatisticSheet = 'Statistic'
nInitDataRowBeg = 2
INIT_CLS_INDEX = 1
INIT_TCH_INDEX = 2
INIT_CURR_INDEX = 3
INIT_COUNT_INDEX = 4
INIT_CONSECUTIVE_INDEX = 5
INIT_BAN_INDEX = 6

#单元格格式数据
sInitDataFont = openpyxl.styles.Font(size=20)
nClsNameCellWidth = 40

nOldTableMinRow = 4
nOldTableMaxRow = 24
nOldTableMinCol = 1
nOldTableMaxCol = 40
strEmpty = 'Empty'


# 上课天数
WORK_DAY = 5
CURR_NUM_PER_DAY = 8


def OpenTchWB():
    os.chdir(strTchWBDir)
    return openpyxl.load_workbook(strTchWBName)


def ReadTeacherName():
    workbook = OpenTchWB()
    arrTeacherName = {}
    sTeachSheet = workbook[strTchSheetName]
    i = 0
    for sCols in sTeachSheet[sTeachSheet.dimensions]:
        for cell in sCols:
            arrTeacherName[i] = cell.value
            i = i + 1

    return arrTeacherName


# predefine
arrTchs = ReadTeacherName()


def ColIndex2Num(strIdx):
    # column 1 -> A, 2 -> B
    return openpyxl.utils.column_index_from_string(strIdx)


def GetCurrKey(strCurrName, strClsName):
    return strCurrName + strClsName


class CurrInfo:
    def __init__(self, strCurrName, strClsName, strTchName) -> None:
        self.strCurrName = strCurrName
        self.strClsName = strClsName
        self.strTchName = strTchName
        self.nCount = 1

    def IsSame(self, cr):
        return self.strCurrName == cr.strCurrName and self.strClsName == cr.strClsName

    def GetKey(self):
        return GetCurrKey(self.strCurrName, self.strClsName)

    def Copy(self, cpy):
        self.strCurrName = cpy.strCurrName
        self.strTchName = cpy.strTchName
        self.strClsName = cpy.strClsName
        self.nCount = cpy.nCount


class InitCurrInfo(CurrInfo):
    def __init__(self, strCurrName, strClsName, strTchName, nCount, strConsecutive, strBanList):
        CurrInfo.__init__(self, strCurrName, strClsName, strTchName)
        self.nCount = nCount

        if strConsecutive == '是':
            self.bIsConsecutive = True 
        else: 
            self.bIsConsecutive = False

        self.sBanList = set()
        self.AddBanStr(strBanList)

 
    def AddBanList(self, sBanList):
        for nBanCurr in sBanList:
            self.sBanList.add(nBanCurr)
        return


    def AddBanStr(self, strBanList):
        if (strBanList == None):
            return
        sBanStrList = strBanList.split(',')
        sBanList = set()
        for strBan in sBanStrList:
            sBanList.add(int(strBan))
        self.AddBanList(sBanList)


class TchData:
    def __init__(self, strTchName : str, sCI : CurrInfo):
        self.strTchName = strTchName
        self.dicCurrs = {}
        self.dicCurrs[sCI.GetKey()] = sCI


class ClsData:
    def __init__(self, strClsName : str, sCI : CurrInfo):
        self.strClsName = strClsName
        self.dicCurrs = {}
        self.dicCurrs[sCI.GetKey()] = sCI


class Schedule:
    def __init__(self):
        self.Sch = {}
    
    # get next empty position in the schedule 
    def GetEmptyPos(self, strClsName):
        if strClsName in self.Sch.keys():
            nIndex = 1
            while nIndex <= WORK_DAY * CURR_NUM_PER_DAY:
                if (self.Sch[strClsName][nIndex] == None):
                    return nIndex
                nIndex = nIndex + 1
        else:
            self.Sch[strClsName] = {}
            nIndex = 1
            while nIndex <= WORK_DAY * CURR_NUM_PER_DAY:
                self.Sch[strClsName][nIndex] = None
            return 1


class IterStack:
    def __init__(self):
        self.stack = list()


def Space_NextlineFilter(strOrg):
    if (type(strOrg) == type('')):
        strTmp = strOrg.replace(' ', '')
        strTmp = strTmp.replace('\n', '')
        strTmp = strTmp.replace('\r', '')
        return strTmp
    else:
        return strOrg


def CellVal(cell):
    return Space_NextlineFilter(cell.value)


def PrintCurrData(CurrMap):
    for sCurrInfos in CurrMap.values():
        for sCI in sCurrInfos.dicCurrs.values():
            if (sCI.strTchName == strEmpty and (sCI.strCurrName != '自习' and sCI.strCurrName != '班会')):
                print('error!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
            print(sCI.strClsName, sCI.strTchName, sCI.strCurrName, sCI.nCount)

def PrintInitCurrData(CurrMap):
    for sCurrInfos in CurrMap.values():
        for sCI in sCurrInfos.dicCurrs.values():
            print(sCI.strClsName, sCI.strTchName, sCI.strCurrName, sCI.nCount, sCI.bIsConsecutive)










