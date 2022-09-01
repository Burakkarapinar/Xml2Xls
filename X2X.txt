from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QInputDialog, QApplication, QPushButton, QVBoxLayout
import xml.etree.ElementTree as ET
import xlsxwriter
import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtWidgets import QListWidget, QListWidgetItem
from tkinter import filedialog
from PyQt5.QtWidgets import *
from PyQt5 import QtGui
import openpyxl as ox
import os
import tkinter as tk

check_flag = "Default"


class Ui_MainWindow(QtWidgets.QWidget):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(500, 500)
        MainWindow.setMinimumSize(QtCore.QSize(500, 500))
        MainWindow.setMaximumSize(QtCore.QSize(500, 500))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(180, 10, 150, 40))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(180, 60, 150, 40))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton3.setGeometry(QtCore.QRect(260, 110, 70, 31))
        self.pushButton3.setObjectName("pushButton_3")
        self.pushButton3.clicked.connect(self.clear)
        self.pushButton4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton4.setGeometry(QtCore.QRect(180, 110, 70, 31))
        self.pushButton4.setObjectName("pushButton_4")
        self.pushButton4.clicked.connect(self.select_all)
        self.left_listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.left_listWidget.setGeometry(QtCore.QRect(40, 180, 200, 280))
        self.left_listWidget.setObjectName("listWidget")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(210, 150, 130, 20))
        self.checkBox.setObjectName("checkBox")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.right_listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.right_listWidget.setGeometry(QtCore.QRect(260, 180, 200, 280))
        self.right_listWidget.setObjectName("listWidget")
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.sel_items = []
        self.seq_lw=[]  #sequence of left listwidget
        self.seq_rw=[]  #sequence of right listwidget
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Xml2Xls"))
        self.pushButton.setText(_translate("MainWindow", "Select"))
        self.pushButton_2.setText(_translate("MainWindow", "Save"))
        self.checkBox.setText(_translate("MainWindow", "Open File "))
        self.pushButton3.setText(_translate("MainWindow", "Clear"))
        self.pushButton4.setText(_translate("MainWindow", "Select All"))
        self.pushButton.clicked.connect(self.select)
        self.pushButton_2.clicked.connect(self.save)
    
        #function for selecting all
    def select_all(self):
        for i in range(self.left_listWidget.count()):
            temp_item=self.left_listWidget.item(i).clone()
            self.right_listWidget.addItem(temp_item)
        self.left_listWidget.clear()

    def clear(self):
        for i in range(self.right_listWidget.count()):
            
            temp_item=self.right_listWidget.item(i).clone()
            self.left_listWidget.addItem(temp_item)
            
        self.right_listWidget.clear()
        
    def select(self):
        try:
            self.fileopenerror = False
            aroot = tk.Tk()
            pat = filedialog.askopenfilename(initialdir="/",
                                             title="Select a File",
                                             filetypes=(("xml files",
                                                         "*.xml"),
                                                        ("Text files",
                                                         "*.txt")))
            aroot.destroy()

            tree = ET.parse(pat)
        except Exception as e:
            QMessageBox.about(self, "Error", "No file selected.")
            print("Burak")
            return 1
            pass

        root = tree.getroot()
        multiple_selection = []  # list for multiple_selection
        data = []

        outer_row = 0 #check : sequence of requirements that have 2 index
        # reach outer children tag
        for i in range(len(root)):
            if (i != outer_row):
                outer_row += 1
            for j in range(len(root[i])):
                if len(root[i][j].items()) >= 0 and root[i][j].tag != 'custom-field-value' and root[i][
                    j].tag != 'document-tree-node':
                    data.append([outer_row, root[i][j].tag, root[i][j].text])
                    if root[i][j].tag not in multiple_selection:
                        multiple_selection.append(root[i][j].tag)

        inner_row = 0

        # reach inner children tag

        for i in range(len(root)):
            if (i != inner_row):
                inner_row += 1
            for j in range(len(root[i])):
                try:
                    temp = root[i][j].items()
                    data.append([inner_row, temp[0][1], temp[1][1]])
                    if temp[0][1] not in multiple_selection:
                        multiple_selection.append(temp[0][1])
                except:
                    continue

        cnt = 0
        for i in multiple_selection:
            if i != 'false':
                if i != 'true':
                    cnt += 1

        # -----------------------------------GET THE SELECTION DATA----------------------------------------------
        # for outer index data
    
        iinner_row = 0 #check : sequence of requirements that have 4 index
        ttt = 0
        l = []
        bool = True
        for i in range(len(root)):
            if (i != iinner_row):
                iinner_row += 1
                temp_list = 0
                lc = 0
            for j in range(len(root[i])):
                try:
                    for k in range(len(root[i][j])):

                        if root[i][j].tag != 'document-tree-node':
                            if (root[i][j][k]).tag != 'multi-line-text':
                                try:

                                    for l in range(len(root[i][j][k])):
                                        if bool:
                                            ttt = root[i][j][k].tag
                                            bool = False

                                        if ttt != root[i][j][k].tag:
                                            ttt = root[i][j][k].tag
                                            temp_list = []

                                        temp_list.append(str(root[i][j][k][l].text))
                                        data.append([iinner_row, root[i][j][k].tag, str(temp_list)])
                                        if root[i][j][k].tag not in multiple_selection:
                                            multiple_selection.append(root[i][j][k].tag)
                                except:
                                    continue



                except:
                    continue
        self.root = root

        def add_to_filter():
            val = self.left_listWidget.currentItem()
            self.right_listWidget.addItem(val.text())
            self.left_listWidget.takeItem(self.left_listWidget.currentRow())


        def remove_Sel(): 
            val = self.right_listWidget.currentItem()
            self.left_listWidget.addItem(val.text())
            self.right_listWidget.takeItem(self.right_listWidget.currentRow())


        #importing headers to left_List_Widget
        for i in multiple_selection:
            QListWidgetItem(i, self.left_listWidget)
        self.left_listWidget.setSelectionMode(QAbstractItemView.MultiSelection)
        self.left_listWidget.itemPressed.connect(add_to_filter)
        self.left_listWidget.setSelectionMode(QAbstractItemView.MultiSelection)

        self.right_listWidget.itemPressed.connect(remove_Sel)   #######################################

        self.data = data
        #check for selecting xml file
        global check_flag
        check_flag = "Selected"

    #function for saving
    def save(self):
        self.sel_items = []
        for index in range(self.right_listWidget.count()):
            self.sel_items.append(self.right_listWidget.item(index).text())
        
        #check : before select xml file, cannot be saved
        global check_flag
        if check_flag == "Default":
            QMessageBox.about(self, "Error", "Please Select Your .xml Folder Firstly")
            return 1

        #ask directory for file that will be saved
        aroot=tk.Tk()
        outPat = filedialog.askdirectory()
        folder_name, ok = QInputDialog.getText(self, '', 'Enter your file name:')
        aroot.destroy()
        name = outPat + "/" + folder_name+ ".xlsx"
        workbook = xlsxwriter.Workbook(name)
        worksheet = workbook.add_worksheet()

        #check : folder name unwritten
        if folder_name=='':
            QMessageBox.about(self, "", "Process is UnSuccesfull")
            return 1

        #write data on excel
        coloumn = -1
        for s in range(len(self.sel_items)):
            if s != coloumn:
                coloumn += 1
                worksheet.write(0, s, self.sel_items[s])
            for i in self.data:
                if i[1] == self.sel_items[s]:
                    worksheet.write(i[0] + 1, s, i[2])
        
        workbook.close()
        xl_to_edit=ox.load_workbook(name)

        k=0
        max_col=xl_to_edit.worksheets[0].max_column-1
        while(k!=max_col):
            i=1
            while(i>=1):
                if(i!=max_col):
                    temp=0
                    j=1
                    while (j!=xl_to_edit.worksheets[0].max_row):
                        if xl_to_edit.worksheets[0].cell(j,i).value is None or xl_to_edit.worksheets[0].cell(j,i).value=='\n\t\t':
                            temp=temp
                        else:
                            temp=temp+1
                        j+=1
                    if temp==0 or temp==1:
                        xl_to_edit.worksheets[0].delete_cols(i,1)
                        i=1
                        max_col=max_col-1
                    else:
                        i+=1
                else:
                    i=0                
            k=max_col
        xl_to_edit.save(name)
        xl_to_edit.close()
        QMessageBox.about(self, " ", "Process is Successful\nWarning:Deleted empty columns")

        if self.checkBox.isChecked() == True:
            os.system(name)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setWindowIcon(QtGui.QIcon(str(os.getcwd())+'\Stackpole.ico'))
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    app.exec_()

