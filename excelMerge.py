#################################### init #################################### 
import pandas as pd
import pathlib
import os
import sys
import xlsxwriter
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from IPython.display import display
import win32com.client
from datetime import datetime
from PyQt5 import uic
import time

###################################### UI ####################################
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

formSrc = resource_path("./myWindow.ui")
form_class = uic.loadUiType(formSrc)[0]

################################### function ##################################
# EXCEL RECOVER
def excelRecover(srcDir):
    try:
        # setting
        o = win32com.client.Dispatch("Excel.Application")
        o.Interactive = False
        o.Visible = False
        # find
        file_list = os.listdir(srcDir)
        # recover
        for file in file_list:
            name, ext = os.path.splitext(file)
            origin = f"{srcDir}/{file}"
            new = f"{srcDir}/{name}_recover.xlsx"
            wb = o.Workbooks.Open(origin)
            wb.ActiveSheet.SaveAs(new,51)
            os.remove(f"{srcDir}/{file}")
    except:
        print('error')
    # quit
    o.Quit()
    
# CATEGORY CLEAN OPTION
def categoryClean(x):
    xArr = str(x).split(' > ')
    return f"{xArr[0]} > 기타" if len(xArr)<3 else x

# EXCEL STYLE
def excelStyle(x,color,font_size,font_color,border):
    color = f"background-color:{color};font-size:{font_size};color:{font_color};border:{border}"
    return color
        
# EXCEL MERGE
def excelMerge(srcDir, resultDir, bar, needTransform):
    try:
        # recover
        fromTo(bar, 10, 0.02)
        if(needTransform): excelRecover(srcDir)
        fromTo(bar, 25, 0.03)
        # setting
        now = datetime.now()
        fromTo(bar, 27, 0.03)
        resultFileName = f"{now.strftime('%Y-%m-%d(%H_%M_%S)')}.xlsx"
        allData = pd.DataFrame()
        fromTo(bar, 30, 0.03)
        fromTo(bar, 35, 0.05)
        # find
        files = pathlib.Path(srcDir).glob('*.xlsx')
        fromTo(bar, 50, 0.05)
        # read
        for file in list(files): 
            file = os.path.normpath(file)
            allData = pd.concat([allData, pd.read_excel(file, usecols=[0,2,7,9,10,17])], ignore_index=True) # concat all dataFrames        
        fromTo(bar, 53, 0.05)
        if allData.shape[0]>1048575: return False 
        fromTo(bar, 55, 0.05)
        fromTo(bar, 58, 0.02)
        # transform
        allData = allData.dropna() #null check
        exceptList = [
            '낚시', 
            '스포츠/레저', 
            '애견/PET', 
            '캠핑', 
            '사은품', 
            '생활/건강', 
            '뷰티', 
            '가구/인테리어', 
            '디지털/가전', 
            '식품', 
            '자동차용품',
            '시트커버/매트',
            '0 > 0 > 0',
            '181',
            '물류 > DP상품',
            '물류 > 쿠팡로켓',
            '물류 > 직송거래처',
            '물류 > 슬리브',
            '물류 > RLH',
            '물류 > 쿠팡',
            '물류 > 아마존',
            '물류 > 이마트',
            '물류 > 트레이더스',
            '물류 > 로켓',
            '물류 > 신세계팩토리',
            '>  >  '
        ]
        exceptOption = '|'.join(exceptList)
        allData['상품수량'] = pd.to_numeric(allData['상품수량'], errors ='coerce').fillna(0).astype('int')
        allData['판매가'] = pd.to_numeric(allData['판매가'], errors ='coerce').fillna(0).astype('int')
        allData = allData[~allData['카테고리'].str.contains(exceptOption)]
        allData = allData.loc[allData['CS']=='정상'] #filtering CS
        allData['카테고리'] = allData['카테고리'].apply(lambda x:str(x).replace(' ','')).apply(lambda x:x.replace('>',' > ')).apply(categoryClean) #category clean
        allData['발주일'] = allData['발주일'].apply(lambda x:str(x)[0:7]) #date format
        allData['상품수량'] = allData['상품수량'].apply(lambda x:int(x)) #numeric filter
        sumData = allData.groupby(['발주일','판매처','카테고리']).sum()
        fromTo(bar, 79, 0.05)
        #save
        #sumData.to_excel(f"{resultDir}/{resultFileName}")
        fromTo(bar, 95, 0.03)
        file_path = f"{resultDir}/{resultFileName}"
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            sumData.to_excel(writer)
            ws = writer.sheets['Sheet1']
            #fix width of columns
            ws.set_column(0, 1, 20)
            ws.set_column(1, 2, 20)
            ws.set_column(2, 3, 40)
            ws.set_column(3, 4, 15)
            ws.set_column(4, 5, 15)
        fromTo(bar, 101, 0.02)
        return "success"
    except:
        return "error"

# PROGRESS BAR ANIMATION
def fromTo(obj, end, interval):
    start = obj.value()
    for i in range(start, end): 
        time.sleep(interval)
        obj.setValue(i)

# GUI
def controllBox():
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

#################################### class ####################################
class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        print(self)
        self.init_UI()

    def init_UI(self):
        self.center()
        self.setWindowIcon(QIcon(r"C:\Users\user\spyderZero\fixIcon.png"))
        self.findButton.clicked.connect(lambda: self.srchButton_clicked(self.srcText))
        self.setButton.clicked.connect(lambda: self.srchButton_clicked(self.resultText))
        self.execButton.clicked.connect(self.exec)
        self.cancelButton.clicked.connect(lambda: self.close())

    def srchButton_clicked(self, obj) :
    	folder = QFileDialog.getExistingDirectory(self, "Select Directory")
    	if(folder != ''): obj.setText(os.path.normpath(folder))
    	else: QMessageBox.about(self, "Error", "Not selected!")

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
    def exec(self):
        self.execButton.setDisabled(True)
        try:
            self.progressBar.setValue(0)
            srcDir = self.srcText.toPlainText()
            resultDir = self.resultText.toPlainText()
            if(not srcDir or not resultDir): return QMessageBox.about(self, "Error", "Not selected!")
            changeResult = excelMerge(srcDir, resultDir, self.progressBar, self.checkBox.isChecked())
            if(changeResult == 'success'): QMessageBox.about(self, "Success", "Merge complete")
            else: 
                QMessageBox.about(self, "Error", "Merge failed!")
                self.progressBar.setValue(0)
        except:
            self.progressBar.setValue(0)
            QMessageBox.about(self, "Error", "Merge failed!")
        self.execButton.setEnabled(True)
        
###################################### exec ####################################
controllBox()
