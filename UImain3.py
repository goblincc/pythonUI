import sys
import third
import PyQt5
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QMessageBox, QLineEdit, QComboBox, QPushButton
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import pandas as pd
import datetime
import time
import os
import pyperclip
import xlrd
import openpyxl
import images2

class Window(third.Ui_MainWindow, QtWidgets.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setWindowIcon(QIcon(':/hh.ico'))
        self.i = 1
        self.k = 0
        self.setupUi(self)
        self.path = ''
        self.data = pd.DataFrame([])
        self.build_set = set()
        self.project_set = set()
        # 一个按钮的点击事件，响应函数为 def msg(self):
        self.pushButton.clicked.connect(self.msg)
        self.pushButton_2.clicked.connect(self.calTableValue)
        self.lineEdit.setAcceptDrops(True)
        self.pushButton_3.clicked.connect(self.resultExport)
        self.pushButton_4.clicked.connect(self.add)
        self.pushButton_5.clicked.connect(self.fixDele)
        self.pushButton_6.clicked.connect(self.save)
        self.pushButton_7.clicked.connect(self.submit)
        self.pushButton_8.clicked.connect(self.export)

    def resultExport(self):
        path = QtWidgets.QFileDialog.getExistingDirectory()
        txt = self.comboBox_3.currentText()
        savePath = path + "/" + txt + "_计算结果.xlsx"
        try:
            if os.path.exists(savePath):
                os.remove(savePath)
        except:
            self.file_except()
        else:
            # print(self.pd_dict)
            pd.DataFrame(self.pd_dict, columns=['分组条件','当期回访量','当期有效回访量','当期有效回访率','当期回访满意工单量',
                                                '当期回访满意度','当年累计有效回访量','当年累计回访量','当年累计有效回访率','当年累计回访满意工单',
                                                '当年累计回访满意度']).to_excel(savePath)



    def export(self):
        path = QtWidgets.QFileDialog.getExistingDirectory()
        savePath = path + "/替换后文件导出.xlsx"
        try:
            if os.path.exists(savePath):
                os.remove(savePath)
        except:
            self.file_except()
        else:
            self.data.to_excel(savePath)


    def file_except(self):  # 消息：警告
        reply = QMessageBox.warning(self, "提示", "文件正在使用")


    def submit(self):

        comboBox = self.comboBox_4.currentText()
        line = self.lineEdit_4.text()
        if (comboBox == '楼栋' and line not in self.build_set) or (comboBox == '项目分期' and line not in self.project_set):
            lineEdit = QtWidgets.QLineEdit(self.tab_3)
            lineEdit.setGeometry(QtCore.QRect(10, self.k * 30 + 80, 200, 22))
            lineEdit.setObjectName(str(self.k) + "j_lineEdit")
            lineEdit.setText(comboBox + "#" + line)
            lineEdit.setFixedWidth(350)

            pushButton = QtWidgets.QPushButton(self.tab_3)
            pushButton.setGeometry(QtCore.QRect(380, self.k * 30 + 80, 75, 23))
            pushButton.setObjectName(str(self.k) + "j_pushButton")
            pushButton.setText("删除")

            lineEdit.show()
            pushButton.show()

            if comboBox == '楼栋':
                self.build_set.add(line)
            elif comboBox == '项目分期':
                self.project_set.add(line)
            self.k += 1
            n1 = lineEdit.objectName()
            n2 = pushButton.objectName()
            pushButton.clicked.connect(lambda: self.delete_tab3(n1, n2))
        else:
            QMessageBox.warning(self, "提示", "数据重复提交, 请检查!")

    def save(self):
        if self.data.empty:
            QMessageBox.warning(self, "提示", "没有导入数据，请检查!")
        else:
            # 替换数据条数记数
            m_cnt = 0
            # 设置集中整改条数记数
            n_cnt = 0
            # 规则替换
            for j in range(self.i):
                if j != 0:
                    line1 = self.tab_2.findChild(QLineEdit, str(j) + "_lineEdit").text()
                    line2 = self.tab_2.findChild(QLineEdit, str(j) + "_2lineEdit").text()
                    select = self.tab_2.findChild(QComboBox, str(j) + "comboBox").currentText()
                else:
                    line1 = self.lineEdit_2.text()
                    line2 = self.lineEdit_3.text()
                    select = self.comboBox.currentText()
                if line1 != '':
                    if select == "楼栋":
                        m_cnt += self.data[self.data['楼栋'] == line1]['楼栋'].count()
                        self.data['楼栋'] = self.data['楼栋'].apply(lambda x: line2 if x == line1 else x)
                    elif select == "项目分期":
                        line2s = line2.split("#")
                        m_cnt += self.data[(self.data['项目'] + self.data['分期']) == line1]['分期'].count()
                        self.data["项目"] = self.data.apply(lambda x: line2s[0] if (x['项目'] + x['分期']) == line1 else x['项目'], axis=1)
                        self.data["分期"] = self.data.apply(lambda x: line2s[1] if (x['项目'] + x['分期']) == line1 else x['分期'], axis=1)
                else:
                    pass
            # 集中整改替换
            for l in range(self.k):
                if self.tab_3.findChild(QLineEdit, str(l) + "j_lineEdit") is not None:
                    line = self.tab_3.findChild(QLineEdit, str(l) + "j_lineEdit").text()
                    # print("line:", line)
                    lines = line.split('#')
                    if lines[0] == '楼栋':
                        self.data['维保阶段'] = self.data['维保阶段'].apply(lambda x: '集中整改期' if x == lines[1] else x)
                        n_cnt += self.data[self.data['楼栋'] == lines[1]]['楼栋'].count()
                    elif lines[0] == '项目分期':
                        self.data['维保阶段'] = self.data.apply(lambda x: '集中整改期' if (x['项目'] + '&' + x['分期']) == lines[1] else x['维保阶段'], axis=1)
                        n_cnt += self.data[(self.data['项目'] + '&' + self.data['分期']) == lines[1]]['分期'].count()
            QMessageBox.information(self, "信息", "替换数据:" + str(m_cnt) + ";设置集中整改:" + str(n_cnt))



    def fixDele(self):
        self.lineEdit_2.deleteLater()
        self.lineEdit_3.deleteLater()
        self.comboBox.deleteLater()
        self.pushButton_5.deleteLater()

    def delete_tab3(self, n1, n2):
        qts1 = self.tab_3.findChild(QLineEdit, n1)
        qts2 = self.tab_3.findChild(QPushButton, n2)
        txt = qts1.text().split("#")[1]
        # print("self.project_set:", self.project_set)
        # print("qts1:", qts1.text())
        if txt in self.build_set:
            self.build_set.remove(txt)
        if txt in self.project_set:
            self.project_set.remove(txt)
        qts1.deleteLater()
        qts2.deleteLater()


    def delete(self, n1, n2, n3, n4):
        qts1 = self.tab_2.findChild(QLineEdit, n1)
        qts2 = self.tab_2.findChild(QLineEdit, n2)
        qts3 = self.tab_2.findChild(QComboBox, n3)
        qts4 = self.tab_2.findChild(QPushButton, n4)
        # print(isinstance(qts1, QLineEdit))
        qts1.deleteLater()
        qts2.deleteLater()
        qts3.deleteLater()
        qts4.deleteLater()


    def msg(self):
        filePath, filetype = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", "./", "*.*")
        self.path = filePath
        self.lineEdit.setText(filePath)
        self.data = pd.read_excel(self.path, keep_default_na=False)

    def getLineEdit(self):
        return self.lineEdit.text()

    def string2Month(self, strs):
        date = datetime.datetime.strptime(strs, '%Y-%m-%d %H:%M:%S')
        month = str(int(date.strftime("%m")))
        # print(month)
        return month

    def string2Year(self, strs):
        date = datetime.datetime.strptime(strs, '%Y-%m-%d %H:%M:%S')
        year = date.strftime("%Y")
        return year

    def date2Timestamp(self, strs):
        timearray = time.strptime(strs, "%Y-%m-%d %H:%M:%S")
        return int(time.mktime(timearray))

    def add(self):
        lineEdit = QtWidgets.QLineEdit(self.tab_2)
        lineEdit.setGeometry(QtCore.QRect(20, self.i * 30 + 60, 250, 22))
        lineEdit.setObjectName(str(self.i) + "_lineEdit")
        lineEdit.setFixedWidth(250)

        lineEdit2 = QtWidgets.QLineEdit(self.tab_2)
        lineEdit2.setGeometry(QtCore.QRect(280, self.i * 30 + 60, 250, 22))
        lineEdit2.setObjectName(str(self.i) + "_2lineEdit")
        lineEdit2.setFixedWidth(250)

        comboBox = QtWidgets.QComboBox(self.tab_2)
        comboBox.setGeometry(QtCore.QRect(540, self.i * 30 + 60, 100, 23))
        comboBox.setObjectName(str(self.i) + "comboBox")
        comboBox.addItem("楼栋")
        comboBox.addItem("项目分期")

        pushButton = QtWidgets.QPushButton(self.tab_2)
        pushButton.setGeometry(QtCore.QRect(650, self.i * 30 + 60, 75, 23))
        pushButton.setObjectName(str(self.i) +"pushButton_5")
        pushButton.setText("删除")

        lineEdit.show()
        lineEdit2.show()
        comboBox.show()
        pushButton.show()
        self.i += 1
        n1 = lineEdit.objectName()
        n2 = lineEdit2.objectName()
        n3 = comboBox.objectName()
        n4 = pushButton.objectName()
        pushButton.clicked.connect(lambda: self.delete(n1, n2, n3, n4))

    def cal(self, datas):
        data = datas
        dateStarts = str(self.dateStart.dateTime().toPyDateTime())
        dateEnds = str(self.dateEnd.dateTime().toPyDateTime())
        dateStart = dateStarts[0:10]
        dateEnd = dateEnds[0:10]
        if dateStart > dateEnd:
            QMessageBox.warning(self, "提示", "选取时间有误！！！")
        else:
            select = self.comboBox_2.currentText()
            cal_group_name = self.comboBox_3.currentText()
            # cal_group = np.unique(data[cal_group_name].values).tolist()
            cal_group = set(data[cal_group_name].values)
            # print("cal_group:", cal_group)
            # print(dateEnd)

            # ①当期回访量：根据“时间段”判断“列T - 回访时间”包含在“开始时间”及“结束时间”之间的excel行数（即工单数）；
            data_cur_visit = data[(data['回访时间'] <= dateEnd) & (data['回访时间'] >= dateStart)]
            # print("data_cur_visit:", data_cur_visit['回访时间'].count())
            # ②当期有效回访量：根据“时间段”判断“列T - 回访时间”包含在“开始时间”及“结束时间”之间，且“列S - 回访状态”为“有效回访”的excel行数（即工单数）
            data_cur_valid_visit = data[
                (data['回访时间'] <= dateEnd) & (data['回访时间'] >= dateStart) & (data['回访状态'] == '有效回访')]

            # ④当期回访满意工单量：根据“时间段”判断“列T - 回访时间”包含在“开始时间”及“结束时间”之间，且“列S - 回访状态”为“有效回访”，且“列X - 您对本次维修总体的满意度感受如何？”为“非常满意”或“满意”的excel行数（即工单数）；
            data_cur_visit_satisfy = data[(data['回访时间'] <= dateEnd) & (data['回访时间'] >= dateStart) &
                                          (data['回访状态'] == '有效回访') &
                                          ((data['您对本次维修总体的满意度感受如何？'] == '非常满意') |
                                           (data['您对本次维修总体的满意度感受如何？'] == '满意'))]

            # ⑤当年累计回访量
            dateStartYear = self.string2Year(dateStarts)
            dateEndYear = self.string2Year(dateEnds)
            if dateEndYear == dateStartYear:
                dateEndYear = str(int(dateEndYear) + 1)
            elif dateEndYear > dateStartYear:
                pass
            elif dateEndYear < dateStartYear:
                pass

            data_cur_visit_year = data[(data['回访时间'] <= dateEndYear) & (data['回访时间'] >= dateStartYear)]

            # ⑥当年累计有效回访量
            data_cur_valid_visit_year = data[
                (data['回访时间'] <= dateEndYear) & (data['回访时间'] >= dateStartYear) & (data['回访状态'] == '有效回访')]

            # ⑧当年累计回访满意工单
            data_cur_valid_visit_satisfy_year = data[(data['回访时间'] <= dateEndYear) & (data['回访时间'] >= dateStartYear) &
                                                     (data['回访状态'] == '有效回访') &
                                                     ((data['您对本次维修总体的满意度感受如何？'] == '非常满意') |
                                                      (data['您对本次维修总体的满意度感受如何？'] == '满意'))]

            group_name = []
            cur_visit_list = []
            cur_valie_visit_list = []
            cur_valie_visit_rate_list = []
            cur_visit_satisfy_list = []
            cur_visit_satisfy_rate_list = []
            cur_visit_year_list = []
            cur_valid_visit_year_list = []
            cur_valid_visit_year_rate_list = []
            cur_valid_visit_satisfy_year_list = []
            cur_valid_visit_stasify_year_rate_list = []

            items = []
            for group in cal_group:
                cur_visit = 0
                cur_valie_visit = 0
                cur_valie_visit_rate = 0.0
                cur_visit_satisfy = 0
                cur_visit_satisfy_rate = 0.0
                cur_visit_year = 0
                cur_valid_visit_year = 0
                cur_valid_visit_year_rate = 0.0
                cur_valid_visit_satisfy_year = 0
                cur_valid_visit_stasify_year_rate = 0.0

                # print("group:", group)
                if cal_group_name == "分期":
                    pre_group = set(data[data[cal_group_name] == group]['项目'].values)
                    for pre in pre_group:
                        item = []
                        item.append(pre + group)
                        # 1.当前回访量
                        data_cur_tmp = data_cur_visit[
                            (data_cur_visit[cal_group_name] == group) & (data_cur_visit['项目'] == pre)]
                        cur_visit = data_cur_tmp["回访时间"].count()

                        # 2.当前有效回访量
                        data_cur_valid_visit_tmp = data_cur_valid_visit[
                            (data_cur_valid_visit[cal_group_name] == group) & (data_cur_valid_visit['项目'] == pre)]
                        cur_valie_visit = data_cur_valid_visit_tmp["回访时间"].count()

                        # print("cur_valie_visit:", cur_valie_visit)
                        # 3.当期有效回访率=②当期有效回访量/①当期回访量
                        if cur_visit != 0:
                            cur_valie_visit_rate = round(cur_valie_visit / cur_visit, 4)

                        # 4.当期回访满意工单量
                        data_cur_visit_satisfy_tmp = data_cur_visit_satisfy[
                            (data_cur_visit_satisfy[cal_group_name] == group) & (data_cur_visit_satisfy['项目'] == pre)]
                        cur_visit_satisfy = data_cur_visit_satisfy_tmp['回访时间'].count()

                        # 当期回访满意度
                        if cur_valie_visit != 0:
                            cur_visit_satisfy_rate = round(cur_visit_satisfy / cur_valie_visit, 4)

                        # 5.当年累计回访量
                        data_cur_visit_year_tmp = data_cur_visit_year[
                            (data_cur_visit_year[cal_group_name] == group) & (data_cur_visit_year['项目'] == pre)]
                        cur_visit_year = data_cur_visit_year_tmp['回访时间'].count()

                        # print("cur_visit_year:", cur_visit_year)
                        # 6.当年累计有效回访量
                        data_cur_valid_visit_year_tmp = data_cur_valid_visit_year[
                            (data_cur_valid_visit_year[cal_group_name] == group) & (
                                        data_cur_valid_visit_year['项目'] == pre)]
                        cur_valid_visit_year = data_cur_valid_visit_year_tmp['回访时间'].count()

                        # 7.当年累计有效回访率
                        if cur_valid_visit_year != 0:
                            cur_valid_visit_year_rate = round(cur_valid_visit_year / cur_visit_year, 4)

                        # 8.当年累计回访满意工单
                        data_cur_valid_visit_satisfy_year_tmp = data_cur_valid_visit_satisfy_year[
                            (data_cur_valid_visit_satisfy_year[cal_group_name] == group) & (
                                        data_cur_valid_visit_satisfy_year['项目'] == pre)]
                        cur_valid_visit_satisfy_year = data_cur_valid_visit_satisfy_year_tmp['回访时间'].count()

                        # 9.当年累计回访满意度
                        if cur_valid_visit_year != 0:
                            cur_valid_visit_stasify_year_rate = round(cur_valid_visit_satisfy_year / cur_valid_visit_year, 4)
                        else:
                            pass

                        group_name.append(pre + group)
                        cur_visit_list.append(cur_visit)
                        cur_valie_visit_list.append(cur_valie_visit)
                        cur_valie_visit_rate_list.append(cur_valie_visit_rate)
                        cur_visit_satisfy_list.append(cur_visit_satisfy)
                        cur_visit_satisfy_rate_list.append(cur_visit_satisfy_rate)
                        cur_visit_year_list.append(cur_visit_year)
                        cur_valid_visit_year_list.append(cur_valid_visit_year)
                        cur_valid_visit_year_rate_list.append(cur_valid_visit_year_rate)
                        cur_valid_visit_satisfy_year_list.append(cur_valid_visit_satisfy_year)
                        cur_valid_visit_stasify_year_rate_list.append(cur_valid_visit_stasify_year_rate)

                        item.append(cur_visit)
                        item.append(cur_valie_visit)
                        item.append(cur_valie_visit_rate)
                        item.append(cur_visit_satisfy)
                        item.append(cur_visit_satisfy_rate)
                        item.append(cur_visit_year)
                        item.append(cur_valid_visit_year)
                        item.append(cur_valid_visit_year_rate)
                        item.append(cur_valid_visit_satisfy_year)
                        item.append(cur_valid_visit_stasify_year_rate)

                        items.append(item)

                        for i in range(len(items)):
                            item = items[i]
                            row = self.tableWidget.rowCount()
                            self.tableWidget.insertRow(row)
                            for j in range(len(item)):
                                item = QTableWidgetItem(str(items[i][j]))
                                self.tableWidget.setItem(row, j, item)
                        items = []
                        self.pd_dict = {'分组条件': group_name,
                                        '当期回访量': cur_visit_list,
                                        '当期有效回访量': cur_valie_visit_list,
                                        '当期有效回访率': cur_valie_visit_rate_list,
                                        '当期回访满意工单量': cur_visit_satisfy_list,
                                        '当期回访满意度': cur_visit_satisfy_rate_list,
                                        '当年累计有效回访量': cur_valid_visit_year_list,
                                        '当年累计回访量': cur_visit_year_list,
                                        '当年累计有效回访率': cur_valid_visit_year_rate_list,
                                        '当年累计回访满意工单': cur_valid_visit_satisfy_year_list,
                                        '当年累计回访满意度': cur_valid_visit_stasify_year_rate_list
                                        }
                else:
                    item = []
                    item.append(group)
                    # 1.当前回访量
                    data_cur_tmp = data_cur_visit[data_cur_visit[cal_group_name] == group]
                    cur_visit = data_cur_tmp["回访时间"].count()

                    # 2.当前有效回访量
                    data_cur_valid_visit_tmp = data_cur_valid_visit[data_cur_valid_visit[cal_group_name] == group]
                    cur_valie_visit = data_cur_valid_visit_tmp["回访时间"].count()

                    # 3.当期有效回访率=②当期有效回访量/①当期回访量
                    if cur_visit != 0:
                        cur_valie_visit_rate = round(cur_valie_visit / cur_visit, 4)

                    # 4.当期回访满意工单量
                    data_cur_visit_satisfy_tmp = data_cur_visit_satisfy[data_cur_visit_satisfy[cal_group_name] == group]
                    cur_visit_satisfy = data_cur_visit_satisfy_tmp['回访时间'].count()

                    # 当期回访满意度
                    if cur_valie_visit != 0:
                        cur_visit_satisfy_rate = round(cur_visit_satisfy / cur_valie_visit, 4)

                    # 5.当年累计回访量
                    data_cur_visit_year_tmp = data_cur_visit_year[data_cur_visit_year[cal_group_name] == group]
                    cur_visit_year = data_cur_visit_year_tmp['回访时间'].count()

                    # 6.当年累计有效回访量
                    data_cur_valid_visit_year_tmp = data_cur_valid_visit_year[
                        data_cur_valid_visit_year[cal_group_name] == group]
                    cur_valid_visit_year = data_cur_valid_visit_year_tmp['回访时间'].count()

                    # 7.当年累计有效回访率
                    if cur_valid_visit_year != 0:
                        cur_valid_visit_year_rate = round(cur_valid_visit_year / cur_visit_year, 4)

                    # 8.当年累计回访满意工单
                    data_cur_valid_visit_satisfy_year_tmp = data_cur_valid_visit_satisfy_year[
                        data_cur_valid_visit_satisfy_year[cal_group_name] == group]
                    cur_valid_visit_stasify_year = data_cur_valid_visit_satisfy_year_tmp['回访时间'].count()

                    # 9.当年累计回访满意度
                    if cur_valid_visit_year != 0:
                        cur_valid_visit_stasify_year_rate = round(cur_valid_visit_stasify_year / cur_valid_visit_year,4)
                    else:
                        pass

                    group_name.append(group)
                    cur_visit_list.append(cur_visit)
                    cur_valie_visit_list.append(cur_valie_visit)
                    cur_valie_visit_rate_list.append(cur_valie_visit_rate)
                    cur_visit_satisfy_list.append(cur_visit_satisfy)
                    cur_visit_satisfy_rate_list.append(cur_visit_satisfy_rate)
                    cur_visit_year_list.append(cur_visit_year)
                    cur_valid_visit_year_list.append(cur_valid_visit_year)
                    cur_valid_visit_year_rate_list.append(cur_valid_visit_year_rate)
                    cur_valid_visit_satisfy_year_list.append(cur_valid_visit_satisfy_year)
                    cur_valid_visit_stasify_year_rate_list.append(cur_valid_visit_stasify_year_rate)

                    item.append(cur_visit)
                    item.append(cur_valie_visit)
                    item.append(cur_valie_visit_rate)
                    item.append(cur_visit_satisfy)
                    item.append(cur_visit_satisfy_rate)
                    item.append(cur_visit_year)
                    item.append(cur_valid_visit_year)
                    item.append(cur_valid_visit_year_rate)
                    item.append(cur_valid_visit_stasify_year)
                    item.append(cur_valid_visit_stasify_year_rate)

                    items.append(item)

                    for i in range(len(items)):
                        item = items[i]
                        row = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row)
                        for j in range(len(item)):
                            item = QTableWidgetItem(str(items[i][j]))
                            self.tableWidget.setItem(row, j, item)
                    items = []
                    self.pd_dict = {'分组条件': group_name,
                                    '当期回访量': cur_visit_list,
                                    '当期有效回访量': cur_valie_visit_list,
                                    '当期有效回访率': cur_valie_visit_rate_list,
                                    '当期回访满意工单量': cur_visit_satisfy_list,
                                    '当期回访满意度': cur_visit_satisfy_rate_list,
                                    '当年累计有效回访量': cur_valid_visit_year_list,
                                    '当年累计回访量': cur_visit_year_list,
                                    '当年累计有效回访率': cur_valid_visit_year_rate_list,
                                    '当年累计回访满意工单': cur_valid_visit_satisfy_year_list,
                                    '当年累计回访满意度': cur_valid_visit_stasify_year_rate_list
                                    }

                    for i in range(len(items)):
                        item = items[i]
                        row = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row)
                        for j in range(len(item)):
                            item = QTableWidgetItem(str(items[i][j]))
                            self.tableWidget.setItem(row, j, item)

    def selected_tb_text(self, table_view):
        try:
            indexes = table_view.selectedIndexes()  # 获取表格对象中被选中的数据索引列表
            indexes_dict = {}
            for index in indexes:  # 遍历每个单元格
                row, column = index.row(), index.column()  # 获取单元格的行号，列号
                if row in indexes_dict.keys():
                    indexes_dict[row].append(column)
                else:
                    indexes_dict[row] = [column]

            # 将数据表数据用制表符(\t)和换行符(\n)连接，使其可以复制到excel文件中
            text = ''
            for row, columns in indexes_dict.items():
                row_data = ''
                for column in columns:
                    data = table_view.item(row, column).text()
                    if row_data:
                        row_data = row_data + '\t' + data
                    else:
                        row_data = data

                if text:
                    text = text + '\n' + row_data
                else:
                    text = row_data
            return text
        except Exception as e:
            QMessageBox.warning(self, "提示", "error!")

    def keyPressEvent(self, event):  # 重写键盘监听事件
        # 监听 CTRL+C 组合键，实现复制数据到粘贴板
        if (event.key() == Qt.Key_C) and QApplication.keyboardModifiers() == Qt.ControlModifier:
            text = self.selected_tb_text(self.tableWidget)  # 获取当前表格选中的数据
            if text:
                # pyperclip.copy(text)  # 复制数据到粘贴板
                try:
                    clipboard = QApplication.clipboard()
                    clipboard.setText(text)  # 复制到粘贴板
                except BaseException as e:
                    print(e)


    def calTableValue(self):
        if self.path == '':
            QMessageBox.warning(self, "提示", "请导入文件！")
        else:
            # 计算之前先清空表格
            self.tableWidget.setRowCount(0)
            self.tableWidget.clearContents()
            cur_sel = self.comboBox_2.currentText()
            datas = self.data
            if cur_sel == '集中整改期':
                df_list = []
                if self.k == 0:
                    QMessageBox.warning(self, "提示", "集中整改栏目没有提交数据！请提交后再计算~")
                else:
                    for p in range(self.k):
                        if self.tab_3.findChild(QLineEdit, str(p) + "j_lineEdit") is not None:
                            line = self.tab_3.findChild(QLineEdit, str(p) + "j_lineEdit").text().strip()
                            lines = line.split('#')
                            if len(lines) >= 2:
                                if lines[0] == '楼栋':
                                    df = self.data[self.data['楼栋'] == lines[1]]
                                    df_list.append(df)
                                elif lines[0] == '项目分期':
                                    line2 = lines[1].split("&")
                                    if len(line2) == 1:
                                        QMessageBox.warning(self, "提示", "集中整改栏目输入有误，缺少&连接符号，请检查！")
                                    else:
                                        df2 = self.data[(self.data['项目'] == line2[0]) & (self.data['分期'] == line2[1])]
                                        df_list.append(df2)
                    if df_list:
                        datas = pd.concat(df_list)
                        self.cal(datas)
                    else:
                        QMessageBox.warning(self, "提示", "集中整改栏目没有提交数据！请提交后再计算~")

            elif cur_sel == '日常维保期':
                if self.k == 0:
                    pass
                else:
                    for p in range(self.k):
                        if self.tab_3.findChild(QLineEdit, str(p) + "j_lineEdit") is not None:
                            line = self.tab_3.findChild(QLineEdit, str(p) + "j_lineEdit").text()
                            lines = line.split('#')
                            if len(lines) >= 2:
                                if lines[0] == '楼栋':
                                    datas = datas[datas['楼栋'] != lines[1]]
                                elif lines[0] == '项目分期':
                                    if len(lines[1].split("&")) == 1:
                                        QMessageBox.warning(self, "提示", "集中整改栏目输入有误，缺少&连接符号，请检查！")
                                    else:
                                        # 直接在dataframe中添加列，警告：A value is trying to be set on a copy of a slice from a DataFrame
                                        datas = datas.copy()
                                        datas.loc[:, "flag"] = datas.apply(lambda x: True if (x['项目'] + '&' + x['分期']) == lines[1] else False, axis=1)
                                        datas = datas[datas['flag'] == False]
                self.cal(datas)
            elif cur_sel == '全部':
                self.cal(datas)




if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = Window()
    mywindow.setWindowTitle("物业管理计算系统-回访")
    mywindow.show()
    sys.exit(app.exec_())

