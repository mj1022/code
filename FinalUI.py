import os
import pandas as pd
import sys
import time
from openpyxl import load_workbook
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QComboBox, QFileDialog, QProgressBar

class Parsing(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.title1 = QLabel('BIOS Select', self)
        self.title1.move(20, 15)
        self.title1.resize(100,12)

        self.title3 = QLabel('Status : IDLE', self)
        self.title3.move(170, 15)
        cb = QComboBox(self)
        cb.addItem('AMT Max')
        cb.addItem('Whitley SMT')
        cb.addItem('Purley SMT')
        cb.addItem('Mix AMT SMT')
        cb.move(20, 32)
        cb.activated[str].connect(self.biosSelect)

        self.title2 = QLabel('File Select', self)
        self.title2.move(20, 63)
        fileBtn = QPushButton(self)
        fileBtn.setText('file...')
        fileBtn.move(20,80)
        fileBtn.clicked.connect(self.FileBtn)

        folderBtn = QPushButton(self)
        folderBtn.setText('folder...')
        folderBtn.move(20,110)
        folderBtn.clicked.connect(self.FolderBtn)

        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(20, 150, 285, 15)

        runBtn = QPushButton(self)
        runBtn.setText('Run')
        runBtn.move(225, 175)
        runBtn.clicked.connect(self.RunBtn)

        self.setWindowTitle('Log Parsing')
        self.resize(325, 225)
        self.move(400, 300)
        self.show()

    def NewPNMapping(self):
        Gb = 1
        tech = ''
        memSize = 1
        dimmtype = ''
        freq = ''
        xX = 1
        rank = 1

        if "T1" in self.config:
            memSize = 256
        elif "G9" in self.config:
            memSize = 64

        if self.config[10] == 'B':
            Gb = 16

        if self.config[5] == '4':
            xX = 4
        elif self.config[5] == '8':
            xX = 8
        elif self.config[5] == '6':
            xX = 16

        if self.config[6] == 'M':
            if Gb == 8:
                tech = 'PL'
            elif Gb == 16:
                tech = 'AL1st'
        elif self.config[6] == 'A':
            if Gb == 8:
                tech = 'DE'
            elif Gb == 16:
                tech = 'ALprime'
        elif self.config[6] == 'C':
            if Gb == 8:
                tech = 'AL'
            elif Gb == 16:
                tech = 'RG'
        elif self.config[6] == 'D':
            tech = 'DA'
        elif self.config[6] == 'J':
            tech = 'DA'

        if self.config[7] == 'X':
            freq = '3200'
        elif self.config[7] == 'W':
            freq = '2933'

        if self.config[9] == 'R' :
            dimmtype = 'RD'
        elif self.config[9] == 'L' :
            dimmtype = 'LRD'
        elif self.config[9] == 'U' :
            dimmtype = 'UD'
        elif self.config[9] == 'S' :
            dimmtype = 'SD'

        rank = int(memSize * 8 / (Gb * (64/xX)))
        self.DimmType = tech+str(Gb)+"G_"+str(memSize)+"GB_"+dimmtype+"_"+str(rank)+"Rx"+str(xX)+"_"+freq

    def PNMapping(self):
        Gb = 1
        tech = ''
        memSize = 1
        dimmtype = ''
        freq = ''
        xX = 1
        rank = 1

        if self.config[3] == '4':
            Gb = 4
        elif self.config[3] == '8':
            Gb = 8
        elif self.config[3] == 'A':
            Gb = 16
        elif self.config[3] == 'B':
            Gb = 32
        elif self.config[3] == 'C':
            Gb = 64
        if self.config[9] == 'P' or self.config[9] == 'M' or self.config[9] == '2' or self.config[9] == 'B':
            Gb = int(Gb/2)
        elif self.config[9] == '4':                             
            Gb = int(Gb/4)

        if self.config[8] == 'M':
            if Gb == 8:
                tech = 'PL'
            elif Gb == 16:
                tech = 'AL1st'
        elif self.config[8] == 'A':
            if Gb == 8:
                tech = 'DE'
            elif Gb == 16:
                tech = 'AL2nd'
        elif self.config[8] == 'C':
            if Gb == 8:
                tech = 'AL'
            elif Gb == 16:
                tech = 'RG'
        elif self.config[8] == 'D':
            tech = 'DA'
        elif self.config[8] == 'E':
            tech = 'RG'
        elif self.config[8] == 'J':
            tech = 'DEprime'
    
        if self.config[6] == 'R' :
            dimmtype = 'RD'
        elif self.config[6] == 'L' :
            dimmtype = 'LRD'
        elif self.config[6] == 'U' :
            dimmtype = 'UD'
        elif self.config[6] == 'S' :
            dimmtype = 'SD'


        if self.config[4] == 'A':
            memSize = 16
        elif self.config[4] == 'B':
            memSize = 32
        else :
            memSize = int(self.config[4])
        memSize *= 8

        if self.config[14] == 'U':
            freq = '2400'
        elif self.config[14] == 'V':
            freq = '2666'
        elif self.config[14] == 'W':
            freq = '2933'
        elif self.config[14] == 'X':
            freq = '3200'

        if self.config[11] == '4':
            xX = 4
        elif self.config[11] == '8':
            xX = 8
        elif self.config[11] == '6':
            xX = 16

        rank = int(memSize * 8 / (Gb * (64/xX)))
        self.DimmType = tech+str(Gb)+"G_"+str(memSize)+"GB_"+dimmtype+"_"+str(rank)+"Rx"+str(xX)+"_"+freq

    def biosSelect(self, bios):
        self.title1.setText(bios)
        self.bios = bios   

    def FileBtn(self):
        self.path = QFileDialog.getOpenFileName(self, 'Open file', ",/")
        self.path = ''.join(self.path)[:-13]
        self.fileOrFolder = 'file'

    def FolderBtn(self):
        self.path = QFileDialog.getExistingDirectory(self, 'Open file', ",/")
        self.fileOrFolder = 'folder'

    def fileSelect(self):
        if self.fileOrFolder == 'file':
            self.ROOTLOGPATH = ''.join(self.path[:self.path.rfind("\\")+1])
            self.LOGNAME = ''.join(self.path[self.path.rfind("\\")+1:])
            self.testList = []
                            
        if self.fileOrFolder == 'folder':
            self.ROOTLOGPATH = self.path + '\\'
            self.testList = [f for f in os.listdir(self.ROOTLOGPATH) 
                             if os.path.isfile(os.path.join(self.ROOTLOGPATH, f))
                             and (os.path.join(self.ROOTLOGPATH, f).lower().rfind('.log') != -1
                                  or os.path.join(self.ROOTLOGPATH, f).lower().rfind('.txt') != -1)]

    def AMTmakeExcel(self):
        pprchk = 0
        startchk = 0
        endchk = 0
        dpcchk = 0
        n0Count = 0
        n1Count = 0
        n0Type = ''
        n1Type = ''
        chkDPC = ''
        Version = ''
        dimmSpeed = ''
        SmartTestKey = ''
        n0Result = 'N0:'
        n1Result = 'N1:'
        resultInfo = []
        failInfo = []

        excelName = self.LOGNAME.split('DPC')[0][:-2]

        if self.LOGNAME.find('.xlsx') == -1:
            with open(self.ROOTLOGPATH + self.LOGNAME,'rt', errors='ignore') as f:
                rawLines = f.readlines()
                f.close()
            for line in rawLines:
                if 'Open' not in line:
                    if "SmartTest Key" in line:
                        SmartTestKey = line.split('SmartTest Key')[1][:12]
                    if "SmartTest-" in line and "RetryCount" in line:
                        Version = line.split('SmartTest-(')[1]
                        Version = Version[:6]
                        startchk += 1
                    if "Stage 4 " in line:
                        if "N0:" in line:
                            n0Result = line.split("), Stage")[0]
                            n0Result = 'N0:' + n0Result[-4:]
                            endchk += 1
                        if "N1:" in line:
                            n1Result = line.split("), Stage")[0]
                            n1Result = 'N1:' + n1Result[-4:]
                            endchk += 1
                        if pprchk == 1:
                            Result = 'PPR ' + n0Result + ', ' + n1Result
                            pprchk = 0
                        else: Result = n0Result + ', ' + n1Result
                    if n0Result != "N0: E:0" or n1Result != "N1: E:0":
                        if "*[0" in line:
                            Node = line.split(', ')[0][-3:]
                            Channel = line.split(', ')[1]
                            Dimm = line.split(', ')[2]
                            if int(Dimm[-1]) < 4:
                                Dimm = 'D:0'
                            if int(Dimm[-1]) >= 4:
                                Dimm = 'D:1'
                            Rank = line.split(', ')[2]
                            Bank = line.split(', ')[4]
                            Row = line.split(', ')[5]
                            failAdd = ['', '', '', '', '', Node, Channel, Dimm, Rank, Bank, Row, '', '']
                            failInfo.append(failAdd)
                    if "DDR4-" in line:
                        dimmSpeed = line.split('DDR4-')[1]
                        dimmSpeed = dimmSpeed[:4]
                    if ": [HMA" in line:
                        if "N0." in line and n0Count == 0:
                            self.config = line[line.find(': [')+3:line.find(': [')+19]
                            if self.config[3] == 'T':
                                self.NewPNMapping()
                            else : 
                                self.PNMapping()
                            n0Type = 'N0:' + self.DimmType
                            n0Count += 1
                        if "N1." in line and n1Count == 0:
                            self.config = line[line.find(': [')+3:line.find(': [')+19]
                            if self.config[3] == 'T':
                                self.NewPNMapping()
                            else : 
                                self.PNMapping()
                            n1Type = 'N1:' + self.DimmType
                            n1Count += 1
                        if "D1:" in line:
                            chkDPC = '2DPC'
                            dpcchk = 1
                        elif "D0:" in line and dpcchk == 0:
                            chkDPC = '1DPC'
                    if startchk == 2 and endchk == 2:
                        if "SetAdvMemTestCondition Starts" in line:
                            new = ['', '', '', '', Result, '', '', '', '', '', '', '', '']
                            resultInfo.append(new)
                            for i in failInfo:
                                resultInfo.append(i)
                            endchk = 0
                            startchk = 0
                            failInfo = []
                    if startchk == 1 and endchk == 1:
                        if "MemTestScram TestType 10 Starts" in line:
                            new = ['', '', '', '', Result, '', '', '', '', '', '', '', '']
                            resultInfo.append(new)
                            for i in failInfo:
                                resultInfo.append(i)
                            endchk = 0
                            startchk = 0
                            failInfo = []
                    if "Execute PPR flow to repair row failures" in line:
                        pprchk = 1

        verinfo = [Version, SmartTestKey]
        testInfo = [excelName, n0Type + ', ' + n1Type, chkDPC, dimmSpeed, '', '', '', '', '', '', '', '', '']

        self.verInfo = pd.DataFrame(verinfo)
        self.verInfo = self.verInfo.transpose()
        log = pd.DataFrame(testInfo)
        log = log.transpose()
        resultInfo = pd.DataFrame(resultInfo)
        log = pd.concat([log, resultInfo])
        self.resultlog = pd.concat([self.resultlog, log])

    def WSmakeExcel(self):
        n0Count = 0
        n1Count = 0
        startchk = 0
        endchk = 0
        dpcchk = 0
        pprchk = 0
        chkDPC = ''
        n0Result = 'N0:'
        n1Result = 'N1:'
        n0TestTime = ''
        n1TestTime = ''
        n0Type = ''
        n1Type = ''
        resultInfo = []
        dimmSpeed = ''
        failInfo = []
        
        excelName = self.LOGNAME

        if self.LOGNAME.find('Test') == -1 and self.LOGNAME.find('.xlsx') == -1:
            with open(self.ROOTLOGPATH + self.LOGNAME,'rt', errors='ignore') as f:
                rawLines = f.readlines()
                f.close()
                TestTime = ''
            for line in rawLines:
                if 'Open' not in line:
                    if "SmartTest-" in line:
                        if "RC" in line and 'PPR' not in line:
                            Version = line.split('X(')[1]
                            Version = Version[:6]
                            RC = line.split('], ')[1]
                            RC = RC[:-1]
                            startchk += 1
                    if "SmartTest - Total Test Times =" in line:
                        if "N0: SmartTest"in line:
                            n0TestTime = 'N0:' + line.split(", Result [")[0][-9:]
                        if "N0: SmartTest"in line:
                            n1TestTime = 'N1:' + line.split(", Result [")[0][-9:]
                        TestTime = n0TestTime + ', ' + n1TestTime
                    if "finished" in line:
                        if "N:0(" in line:
                            n0Result = line.split("able,")[1][:4]
                            n0Result = 'N0:' + n0Result
                            endchk += 1
                        if "N:1(" in line:
                            n1Result = line.split("able,")[1][:4]
                            n1Result = 'N1:' + n1Result
                            endchk += 1
                        if pprchk == 1:
                            Result = 'PPR ' + n0Result + ', ' + n1Result
                            pprchk = 0
                        else: Result = n0Result + ', ' + n1Result
                    if n0Result != "N0:E: 0" or n1Result != "N1:E: 0":
                        if "subRank:" in line:
                            line = line.split('] - ')[1]
                            Node = line.split(', ')[0][-3:]
                            Channel = line.split(', ')[1]
                            Dimm = line.split(', ')[2]
                            Rank = line.split(', ')[3]
                            Bank = line.split(', ')[5]
                            Row = line.split(', ')[6]
                            Col = line.split(', ')[7]
                            Dq = line.split(', ')[8]
                            failAdd = ['', '', '', '', '', '', Node, Channel, Dimm, Rank, Bank, Row, Col, Dq]
                            failInfo.append(failAdd)
                    if "DDR4-" in line:
                        dimmSpeed = line.split('DDR4-')[1]
                        dimmSpeed = dimmSpeed[:4]
                    if "AD00:" in line:
                        if "N:0," in line and n0Count == 0:
                            self.config = line[line.find('AD00:')+5:line.find('AD00:')+21]
                            if "-" not in self.config:
                                self.NewPNMapping()
                            else : 
                                self.PNMapping()
                            n0Type = 'N0:' + self.DimmType
                            n0Count += 1
                        if "N:1," in line and n1Count == 0:
                            self.config = line[line.find('AD00:')+5:line.find('AD00:')+21]
                            if self.config[3] == 'T':
                                self.NewPNMapping()
                            else : 
                                self.PNMapping()
                            n1Type = 'N1:' + self.DimmType
                            n1Count += 1
                        if "D:1," in line:
                            chkDPC = '2DPC'
                            dpcchk = 1
                        elif "D:0," in line and dpcchk == 0:
                            chkDPC = '1DPC'
                    if startchk == 2 and endchk == 2:
                        if "MemTestScram TestType 10 Starts" in line:
                            new = ['', '', '', '', '', Result, '', '', '', '', '', '', '']
                            resultInfo.append(new)
                            for i in failInfo:
                                resultInfo.append(i)
                            endchk = 0
                            startchk = 0
                            failInfo = []
                    if startchk == 1 and endchk == 1:
                        if "MemTestScram TestType 10 Starts" in line:
                            new = ['', '', '', '', '', Result, '', '', '', '', '', '', '']
                            resultInfo.append(new)
                            for i in failInfo:
                                resultInfo.append(i)
                            endchk = 0
                            startchk = 0
                            failInfo = []
                    if "skExecuteSmartPPR" in line:
                        pprchk = 1

        verinfo = [Version, RC]
        testInfo = [excelName, n0Type + ', ' + n1Type, chkDPC, TestTime, dimmSpeed, '', '', '', '', '', '', '', '', '']
        self.verInfo =  pd.DataFrame(verinfo)
        self.verInfo = self.verInfo.transpose()
        log = pd.DataFrame(testInfo)
        log = log.transpose()
        resultInfo = pd.DataFrame(resultInfo)
        log = pd.concat([log, resultInfo])
        self.resultlog = pd.concat([self.resultlog, log])

    def PSmakeExcel(self):
        n0Count = 0
        n1Count = 0
        n0 = 0
        n1 = 0
        chk = 0
        startchk = 0
        endchk = 0
        pprchk = 0
        chkDPC = '1DPC'
        n0Result = 'N0:'
        n1Result = 'N1:'
        n0TestTime = ''
        n1TestTime = ''
        resultInfo = []
        failInfo = []

        excelName = self.LOGNAME.split('DPC')[0][:-2]

        if self.LOGNAME.find('Test') == -1 and self.LOGNAME.find('.xlsx') == -1:
            with open(self.ROOTLOGPATH + self.LOGNAME,'rt', errors='ignore') as f:
                rawLines = f.readlines()
                f.close()
                TestTime = ''
            for line in rawLines:
                if 'Open' not in line and '2.7' not in line:
                    if "SmartTest-" in line:
                        if "RC" in line and 'PPR' not in line:
                            Version = line.split('X(')[1]
                            Version = Version[:6]
                            RC = line.split('), ')[1]
                            startchk += 1
                    if "SKHYNIX SmartTest - " in line:
                        if n0 == 1:
                            n0TestTime = 'N0:' + line.split("SmartTest - ")[1]
                            n0 = 0
                        if n1 == 1:
                            n1TestTime = 'N1:' + line.split("SmartTest - ")[1]
                            n1 = 0
                        TestTime = n0TestTime + ', ' + n1TestTime
                        failAdd = []
                    if "SmartTest-" in line and "finished" in line:
                        if "N:0" in line:
                            n0Result = line.split(') finished')[0]
                            n0Result = 'N0:' + n0Result[-4:]
                            n0 = 1
                            endchk += 1
                        if "N:1" in line:
                            n1Result = line.split(') finished')[0]
                            n1Result = 'N1:' + n1Result[-4:]
                            n1 = 1
                            endchk += 1
                        if pprchk == 1:
                            Result = 'PPR ' + n0Result + ', ' + n1Result
                            pprchk = 0
                        else : Result = n0Result + ', ' + n1Result
                    if n0Result != "N0: E:0" or n1Result != "N1: E:0":
                        if "- N:" in line and "'C], " not in line:
                            Node = line.split(', ')[0][-3:]
                            Channel = line.split(', ')[1]
                            Dimm = line.split(', ')[2]
                            Rank = line.split(', ')[3]
                            Bank = line.split(', ')[5]
                            Row = line.split(', ')[6]
                            Col = line.split(', ')[7]
                            Dq = line.split(', ')[8]
                            failAdd = ['', '', '', '', '', '', Node, Channel, Dimm, Rank, Bank, Row, Col, Dq]
                            failInfo.append(failAdd)
                    if "DDR4-" in line:
                        dimmSpeed = line.split('DDR4-')[1]
                        dimmSpeed = dimmSpeed[:4]
                    if "AD00:" in line:
                        if "N:0," in line and n0Count == 0:
                            self.config = line[line.find('AD00:')+5:line.find('AD00:')+21]
                            if self.config[3] == 'T':
                                self.NewPNMapping()
                            else : 
                                self.PNMapping()
                            n0Type = 'N0:' + self.DimmType
                            n0Count += 1
                        if "D:1," in line and chk == 0:
                            chkDPC = '2DPC'
                            chk = 1
                        if "N:1," in line and n1Count == 0:
                            self.config = line[line.find('AD00:')+5:line.find('AD00:')+21]
                            if self.config[3] == 'T':
                                self.NewPNMapping()
                            else : 
                                self.PNMapping()
                            n1Type = 'N1:' + self.DimmType
                            n1Count += 1
                    if startchk == 2 and endchk == 2:
                        if "SetAdvMemTestCondition Starts" in line:
                            new = ['', '', '', '', '', Result, '', '', '', '', '', '', '']
                            resultInfo.append(new)
                            for i in failInfo:
                                resultInfo.append(i)
                            endchk = 0
                            startchk = 0
                            failInfo = []
                    if startchk == 1 and endchk == 1:
                        if "SetAdvMemTestCondition Starts" in line:
                            new = ['', '', '', '', '', Result, '', '', '', '', '', '', '']
                            resultInfo.append(new)
                            for i in failInfo:
                                resultInfo.append(i)
                            endchk = 0
                            startchk = 0
                            failInfo = []
                    if "skExecuteSmartPPR" in line:
                        pprchk = 1

        verinfo = [Version, RC]
        testInfo = [excelName, n0Type + ', ' + n1Type, chkDPC, TestTime, dimmSpeed, '', '', '', '', '', '', '', '', '']

        self.verInfo =  pd.DataFrame(verinfo)
        self.verInfo = self.verInfo.transpose()
        log = pd.DataFrame(testInfo)
        resultInfo = pd.DataFrame(resultInfo)
        log = log.transpose()
        log = pd.concat([log, resultInfo])
        self.resultlog = pd.concat([self.resultlog, log])

    def RunBtn(self):
        step = 0
        testList = ['']
        self.title3.setText('Status : Parsing')
        self.title3.resize(150,12)
        self.title3.repaint()
        self.resultlog = pd.DataFrame()

        self.fileSelect()
        os.makedirs(self.ROOTLOGPATH + "Parsing", exist_ok = True)
        self.pbar.setMaximum(len(self.testList))

        if len(self.testList) == 0 :
            try:
                if self.bios == 'AMT Max':
                    self.AMTmakeExcel()
                if self.bios == 'Whitley SMT':
                    self.WSmakeExcel()
                if self.bios == 'Purley SMT':
                    self.PSmakeExcel()
                if self.bios == 'Mix AMT SMT':
                    self.WSmakeExcel()
                    self.AMTmakeExcel()
            except(PermissionError):
                print('Error')

        elif len(self.testList) != 0  :
            for logname in self.testList:
                step = step +1
                self.pbar.setValue(step)
                self.LOGNAME = ''.join(logname[logname.rfind("/")+1:])
                try:
                    if self.bios == 'AMT Max':
                        self.AMTmakeExcel()
                    if self.bios == 'Whitley SMT':
                        self.WSmakeExcel()
                    if self.bios == 'Purley SMT':
                        self.PSmakeExcel()
                    if self.bios == 'Mix AMT SMT':
                        self.WSmakeExcel()
                        self.AMTmakeExcel()
                except(PermissionError):
                    print('Error')

        savePath  = self.ROOTLOGPATH + 'Parsing/' + 'CheckList.xlsx'

        if self.bios == 'Whitley SMT' or self.bios == 'Purley SMT':
            self.verInfo.columns = ['Version', 'RC']
            self.resultlog.columns = ['LogName', 'DimmType', 'DPC', 'TestTime', 'DimmSpeed', 'Result', 'Node', 'Channel', 'Dimm', 'Rank', 'Bank', 'Row', 'Col', 'Dq']
        if self.bios == 'AMT Max':
            self.verInfo.columns = ['Version', 'SmartTestKey']
            self.resultlog.columns = ['LogName', 'DimmType', 'DPC', 'DimmSpeed', 'Result', 'Node', 'Channel', 'Dimm', 'Rank', 'Bank', 'Row', 'Col', 'Dq']
        resultLog = self.verInfo.append(self.resultlog)
        resultLog.to_excel(savePath, index = False)

        wb = load_workbook(self.ROOTLOGPATH + 'Parsing/' + 'CheckList.xlsx')
        ws = wb.active
        #if self.bios == 'Whitley SMT' or self.bios == 'Purley SMT':
            #ws.move_range("C3:G99", rows=-1, translate=True)
            #ws.move_range("H4:H99", rows=-2, translate=True)
            #ws.move_range("I5:P99", rows=-3, translate=True)
        if self.bios == 'AMT Max':
            ws.move_range("C3:F99", rows=-1, translate=True)
            ws.move_range("G4:G99", rows=-2, translate=True)
            ws.move_range("H5:P99", rows=-3, translate=True)
        wb.save(self.ROOTLOGPATH + 'Parsing/' + 'CheckList.xlsx')

        self.title3.setText('Status : Parsing End')
        self.title3.repaint()

        time.sleep(1)

if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = Parsing()
   sys.exit(app.exec_())