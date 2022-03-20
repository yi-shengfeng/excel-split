# -*- encoding: utf-8 -*-
# @Description: 处理excel
# @Date: 2022-03-19 23:27:07
# @Author: YiShengfeng < yishengfeng@qq.com >
# @LastEditors: YiShengfeng
# @LastEditTime: 2022-03-20 19:46:19

import sys
import os
from tkinter import FIRST
from tkinter.tix import DisplayStyle
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from Ui_convert import Ui_Excel
import pandas as pd


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


def convert_to_number(letter, columnA=0):
    """
    字母列号转数字
    columnA: 你希望A列是第几列(0 or 1)? 默认0
    return: int
    """
    ab = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    letter0 = letter.upper()
    w = 0
    for _ in letter0:
        w *= 26
        w += ab.find(_)
    return w - 1 + columnA


def convert_to_letter(number, columnA=0):
    """
    数字转字母列号
    columnA: 你希望A列是第几列(0 or 1)? 默认0
    return: str in upper case
    """
    ab = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    n = number - columnA
    x = n % 26
    if n >= 26:
        n = int(n / 26)
        return convert_to_letter(n, 1) + ab[x + 1]
    else:
        return ab[x + 1]


class ExcelProcess():
    _inputFile = ''
    _inputFileName = ''
    _outputFilePath = ''
    _headerNum = None
    _processNum = []
    _processStandard = []
    _isSameStandard = False

    def __init__(self, inputFile, outputFilePath, processNum, headerNum=None, processStandard=[1], sameStandard=False) -> None:
        self._inputFile = inputFile
        self._inputFileName = os.path.split(self._inputFile)[1].split('.')[0]
        self._outputFilePath = outputFilePath
        self._headerNum = headerNum
        self._processNum = processNum
        self._processStandard = processStandard
        self._isSameStandard = sameStandard
        pass

    def process(self):
        if len(self._processNum) > 1:
            # 处理多sheet
            if self._isSameStandard:
                tempData = None
                for i in range(len(self._processNum)):
                    if self._processNum[i] == ' ':
                        continue
                    else:
                        curSheetExcel = pd.read_excel(
                            self._inputFile, header=self._headerNum, sheet_name=i)
                        data1 = curSheetExcel.iloc[:, convert_to_number(
                            self._processNum[i])]
                        tempData = pd.concat([tempData, data1])
                uniqueList = tempData.unique()
                sheetNames = pd.read_excel(
                    self._inputFile, sheet_name=None).keys()
                sheetNames = list(sheetNames)
                if len (self._processNum) != len(sheetNames):
                    return False
                for i in uniqueList:
                    writer = pd.ExcelWriter(
                        self._outputFilePath+'/'+self._inputFileName+str(i)+'.xlsx')
                    for index in range(len(self._processNum)):
                        if self._processNum[index] == ' ':
                            tempsheet = pd.read_excel(
                                self._inputFile, sheet_name=sheetNames[index])
                            tempsheet.to_excel(
                                writer, sheet_name=sheetNames[index], index=False)
                        else:
                            tempsheet = pd.read_excel(
                                self._inputFile, sheet_name=sheetNames[index])
                            tempWriteData = tempsheet[tempsheet.iloc[:, convert_to_number(
                                self._processNum[index])] == i]
                            tempWriteData.to_excel(
                                writer, sheet_name=sheetNames[index], index=False)
                    writer.close()
                return True
            elif len(self._processNum) == len(self._processStandard):
                standard = self._processStandard.copy()
                if ' ' in standard:
                    standard.remove(' ')
                standard = list(set(standard))
                tempData = {}
                for i in range(len(standard)):
                    tempData.update({standard[i]: None})
                uniqueList = {}
                for i in range(len(self._processNum)):
                    if self._processNum[i] == ' ' or self._processStandard[i] == ' ':
                        continue
                    else:
                        curSheetExcel = pd.read_excel(
                            self._inputFile, header=self._headerNum, sheet_name=i)
                        data1 = curSheetExcel.iloc[:, convert_to_number(
                            self._processNum[i])]
                        tempData.update({self._processStandard[i]: pd.concat(
                            [tempData[self._processStandard[i]], data1])})
                        # tempData[self._processStandard[i]] = pd.concat([tempData[self._processStandard[i]], data1])
                for i in range(len(standard)):
                    uniqueList[standard[i]] = tempData[standard[i]].unique()
                sheetNames = pd.read_excel(
                    self._inputFile, sheet_name=None).keys()
                sheetNames = list(sheetNames)
                if len (self._processNum) != len(sheetNames):
                    return False
                for i in standard:
                    for j in uniqueList[i]:
                        writer = pd.ExcelWriter(
                            self._outputFilePath+'/'+self._inputFileName+'-elm'+str(j)+'-std'+i+'.xlsx')
                        for index in range(len(self._processStandard)):
                            if self._processStandard[index] != i or self._processStandard[index] == ' ':
                                tempsheet = pd.read_excel(
                                    self._inputFile, sheet_name=sheetNames[index])
                                tempsheet.to_excel(
                                    writer, sheet_name=sheetNames[index], index=False)
                            else:
                                tempsheet = pd.read_excel(
                                    self._inputFile, sheet_name=sheetNames[index])
                                tempWriteData = tempsheet[tempsheet.iloc[:, convert_to_number(
                                    self._processNum[index])] == j]
                                tempWriteData.to_excel(
                                    writer, sheet_name=sheetNames[index], index=False)
                        writer.close()
                return True
            else:
                return False
        elif len(self._processNum) == 1:
            # 处理单sheet
            for i in range(len(self._processNum)):
                if self._processNum[i] == ' ':
                    continue
                else:
                    excelFile = pd.read_excel(
                        self._inputFile, header=self._headerNum)
                    uniqueList = excelFile.iloc[:, convert_to_number(
                        self._processNum[i])].unique()
                    for i in uniqueList:
                        sourceFile = pd.read_excel(self._inputFile)
                        writer = pd.ExcelWriter(
                            self._outputFilePath + '/'+self._inputFileName + str(i) + '.xlsx')
                        tempWriteData = sourceFile[sourceFile.iloc[:, convert_to_number(
                            self._processNum[i])] == i]
                        tempWriteData.to_excel(writer, index=False)
                        writer.close()
                    return True
        else:
            return False


class MyMainForm(QMainWindow, Ui_Excel):
    _maxSheets = 20
    _inputFile = ''
    _outputFilePath = ''
    _haveHeader = False
    _isMutiSheets = False
    _processNum = []
    _inputNumberStr = ''
    _isSameStandard = False
    _processStandard = []
    _processStandardStr = ''
    _headerNum = None
    _canConvert = False

    def __init__(self, parent=None):
        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)
        self.inputLineEdit.setPlaceholderText("请选择输入文件")
        self.inputLineEdit.setReadOnly(True)
        self.outputLineEdit.setPlaceholderText("请选择输出文件夹")
        self.outputLineEdit.setReadOnly(True)
        self.inputSelect.clicked.connect(self.select_input)
        self.outputSelect.clicked.connect(self.select_output)
        self.checkFile.clicked.connect(self.check)
        self.convert.clicked.connect(self.process)
        self.processNumberInput.setPlaceholderText("逗号分隔。")
        self.processNumberInput.textChanged.connect(self.process_number)
        self.sheet.clicked.connect(self.muti_sheets_check)
        self.header.clicked.connect(self.header_check)
        self.sameStandard.clicked.connect(self.same_standard)
        self.processStandard.setPlaceholderText("设置每个sheet处理的标准。")
        self.processStandard.textChanged.connect(self.process_standard)
        self.headerNum.setReadOnly(True)
        self.headerNum.setPlaceholderText("不可设置！")
        self.headerNum.textChanged.connect(self.header_num)

    def header_num(self):
        if is_number(self.headerNum.text()):
            if int(self.headerNum.text()) - 1 >= 0:
                self._headerNum = int(self.headerNum.text()) - 1
            return 0
        else:
            self._headerNum = None
            return 0

    def same_standard(self):
        self._isSameStandard = self.sameStandard.isChecked()
        if self._isSameStandard:
            self.processStandard.setReadOnly(True)
            self.processStandard.setPlaceholderText("已设置相同标准，不可设置！")
        else:
            self.processStandard.setReadOnly(False)
            self.processStandard.setPlaceholderText("设置每个sheet处理的标准。")
            # self.processStandard.setClearButtonEnabled(True)
        return 0

    def process_number(self):
        self._inputNumberStr = self.processNumberInput.text()
        return 0

    def process_standard(self):
        self._processStandardStr = self.processStandard.text()
        return 0

    def muti_sheets_check(self):
        self._isMutiSheets = self.sheet.isChecked()
        return 0

    def header_check(self):
        self._haveHeader = self.header.isChecked()
        if self._haveHeader:
            self.headerNum.setReadOnly(False)
            self.headerNum.setPlaceholderText("设置表头行数")
            # self.headerNum.setClearButtonEnabled(True)
        else:
            self.headerNum.setReadOnly(True)
            self.headerNum.setPlaceholderText("不可设置！")
        return 0

    def display(self, msg: str):
        self.info.setText(msg)
        return 0

    def select_input(self):
        directory = QFileDialog.getOpenFileName(
            self, "getOpenFileName", "./", "所有文件 (*);;Excel (*.xlsx;*.xls)")
        self.inputLineEdit.setText(directory[0])
        self._inputFile = directory[0]
        return 0

    def select_output(self):
        directory = QFileDialog.getExistingDirectory(
            self, "getExistingDirectory", "./")
        # 当窗口非继承QtWidgets.QDialog时，self可替换成 None
        self.outputLineEdit.setText(directory)
        self._outputFilePath = directory
        return 0

    def check(self):
        displayMsg = ''
        processNum = self._inputNumberStr.split(',')
        processStandard = self._processStandardStr.split(',')
        if self._inputFile == '':
            displayMsg += '还没选择输入文件\n'
            self.display(displayMsg)
            return 0
        elif self._outputFilePath == '':
            displayMsg += '还没选择输出文件夹\n'
            self.display(displayMsg)
            return 0
        elif self._inputNumberStr == '':
            displayMsg += '还没输入sheet的处理列号\n'
            self.display(displayMsg)
            return 0
        elif len(processNum) > self._maxSheets:
            displayMsg += '超出最大处理sheets数' + str(self._maxSheets) + '个\n'
            self.display(displayMsg)
            return 0
        else:
            self._processNum.clear()
            self._processNum.extend(processNum)
            self._processStandard.clear()
            self._processStandard.extend(processStandard)
            inputFile = os.path.split(self._inputFile)
            fileExtendName = inputFile[1].split('.')
            if fileExtendName[1] != 'xls' and fileExtendName[1] != 'xlsx':
                displayMsg += '无法处理.' + fileExtendName[1] + '类型的文件\n'
                self.display(displayMsg)
                self._canConvert = False
                return 0
            elif self._isMutiSheets and (len(self._processNum) <= 1):
                displayMsg += '选择的多sheet与需要处理sheet数不符合\n'
                self.display(displayMsg)
                self._canConvert = False
                return 0
            elif len(self._processNum) < 1:
                displayMsg += '处理列号不能为空'
                self.display(displayMsg)
                self._canConvert = False
                return 0
            elif (not self._isMutiSheets) and (len(self._processNum) > 1):
                displayMsg += '单sheet不能设置多个处理列'
                self.display(displayMsg)
                self._canConvert = False
                return 0
            elif (not self._isSameStandard) and (len(self._processNum) != len(self._processStandard)):
                displayMsg += '处理列号和处理标准数目不对应\n'
                self.display(displayMsg)
                self._canConvert = False
                return 0
            else:
                displayMsg += '输入信息:\n'
                displayMsg += '处理文件名: '+inputFile[1] + '\n'
                displayMsg += '输出文件路径: '+self._outputFilePath + '\n'
                if self._haveHeader:
                    displayMsg += '有表头\n'
                else:
                    displayMsg += '无表头\n'

                if self._isMutiSheets:
                    displayMsg += '有多个sheets\n'
                else:
                    displayMsg += '单个sheet\n'
                displayMsg += '需要处理的列号:'
                for i in range(len(self._processNum)):
                    if self._processNum[i] == ' ':
                        continue
                    else:
                        displayMsg += 'sheet' + \
                            str(i + 1) + ':' + str(self._processNum[i]) + ' '
                displayMsg += '\n'

                self.display(displayMsg)
                self._canConvert = True
                return 0

    def process(self):
        displayMsg = ''
        if self._canConvert:
            ep = ExcelProcess(self._inputFile, self._outputFilePath, self._processNum, headerNum=self._headerNum,
                            processStandard=self._processStandard, sameStandard=self._isSameStandard)
            if ep.process():
                displayMsg += '转换成功\n'
                self.display(displayMsg)
                return 0
            else:
                displayMsg += '转换失败\n'
                self.display(displayMsg)
                return 0
        else:
            displayMsg += '先进行检查，无误再进行转换\n'
            self.display(displayMsg)
            return 0


if __name__ == "__main__":
    # 固定的，PyQt5程序都需要QApplication对象。sys.argv是命令行参数列表，确保程序可以双击运行
    app = QApplication(sys.argv)
    # 初始化
    myWin = MyMainForm()
    # 将窗口控件显示在屏幕上
    myWin.show()
    # 程序运行，sys.exit方法确保程序完整退出。
    sys.exit(app.exec_())
