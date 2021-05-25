import sys
import second
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QMessageBox, QLineEdit, QComboBox, QPushButton
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon
import pandas as pd
import datetime
import time
import os
import xlrd
import openpyxl
import images

class Window(second.Ui_MainWindow, QtWidgets.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setWindowIcon(QIcon(':/ss.ico'))
        self.i = 1
        self.k = 0
        self.setupUi(self)
        self.path = ''
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
            pd.DataFrame(self.pd_dict).to_excel(savePath)

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

        self.k += 1
        n1 = lineEdit.objectName()
        n2 = pushButton.objectName()
        pushButton.clicked.connect(lambda: self.delete_tab3(n1, n2))

    def save(self):
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

            if select == "楼栋":
                m_cnt += self.data[self.data['楼栋'] == line1]['楼栋'].count()
                self.data['楼栋'] = self.data['楼栋'].apply(lambda x: line2 if x == line1 else x)
            elif select == "项目分期":
                line2s = line2.split("#")
                m_cnt += self.data[(self.data['项目'] + self.data['项目分期']) == line1]['项目分期'].count()
                self.data["项目"] = self.data.apply(lambda x: line2s[0] if (x['项目'] + x['项目分期']) == line1 else x['项目'], axis=1)
                self.data["项目分期"] = self.data.apply(lambda x: line2s[1] if (x['项目'] + x['项目分期']) == line1 else x['项目分期'], axis=1)
        # 集中整改替换
        for l in range(self.k):
            line = self.tab_3.findChild(QLineEdit, str(l) + "j_lineEdit").text()
            lines = line.split('#')
            if lines[0] == '楼栋':
                self.data['维保阶段名称'] = self.data['维保阶段名称'].apply(lambda x: '集中整改期' if x == lines[1] else x)
                n_cnt += self.data[self.data['楼栋'] == lines[1]]['楼栋'].count()
            elif lines[0] == '项目分期':
                self.data['维保阶段名称'] = self.data.apply(lambda x: '集中整改期' if (x['项目'] + '#' + x['项目分期']) == line else x['维保阶段名称'], axis=1)
                n_cnt += self.data[(self.data['项目'] + '#' + self.data['项目分期']) == line]['项目分期'].count()
        QMessageBox.information(self, "信息", "替换数据:" + str(m_cnt) + ";设置集中整改:" + str(n_cnt))



    def fixDele(self):
        self.lineEdit_2.deleteLater()
        self.lineEdit_3.deleteLater()
        self.comboBox.deleteLater()
        self.pushButton_5.deleteLater()


    def delete_tab3(self, n1, n2):
        qts1 = self.tab_3.findChild(QLineEdit, n1)
        qts2 = self.tab_3.findChild(QPushButton, n2)
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
        print(month)
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
        dateStart = str(self.dateStart.dateTime().toPyDateTime())
        dateEnd = str(self.dateEnd.dateTime().toPyDateTime())
        if dateStart > dateEnd:
            QMessageBox.warning(self, "提示", "选取时间有误！！！")
        else:
            select = self.comboBox_2.currentText()
            cal_group_name = self.comboBox_3.currentText()
            # cal_group = np.unique(data[cal_group_name].values).tolist()
            cal_group = set(data[cal_group_name].values)
            print("cal_group:", cal_group)
            print(dateEnd)

            # 当期未关闭：根据“时间段”判断“列AA-报事时间”早于“结束时间”， 且“列L-当前工单状态”为“方案已批准”、“方案制定中”、“施工完成”、“施工中”、“已响应”的excel行数（即工单数）
            data_cur = data[(data["报事时间"] <= dateEnd) & ((data["当前工单状态"] == "方案已批准") |
                                                         (data["当前工单状态"] == "方案制定中") |
                                                         (data["当前工单状态"] == "施工完成") |
                                                         (data["当前工单状态"] == "施工中") |
                                                         (data["当前工单状态"] == "已响应") |
                                                         (data["当前工单状态"] == "已分派") |
                                                         (data["当前工单状态"] == "已上门"))]
            print(data_cur["当前工单状态"].count())

            # 当期新增：根据“时间段”判断“列AA-报事时间”包含在“开始时间”及“结束时间”之间，且“列L-当前工单状态”为“方案已批准”、“方案制定中”、“施工完成”、“施工中”、“已响应”、“非正常关闭”、“强制关闭”、“已关闭”、“已评价”的excel行数（即工单数）
            data_increase = data[(data["报事时间"] >= dateStart) & (data["报事时间"] <= dateEnd) & ((data["当前工单状态"] == "方案已批准") |
                                                                                            (data["当前工单状态"] == "方案制定中") |
                                                                                            (data["当前工单状态"] == "施工完成") |
                                                                                            (data["当前工单状态"] == "施工中") |
                                                                                            (data["当前工单状态"] == "已响应") |
                                                                                            (data["当前工单状态"] == "非正常关闭") |
                                                                                            (data["当前工单状态"] == "强制关闭") |
                                                                                            (data["当前工单状态"] == "已关闭") |
                                                                                            (data["当前工单状态"] == "已评价") |
                                                                                            (data["当前工单状态"] == "已分派") |
                                                                                            (data["当前工单状态"] == "已上门"))]

            print("data_increase", data_increase["当前工单状态"].count())

            # 当期关闭：根据“时间段”判断“列AU-业主关闭时间”或“列AV-非正常关闭时间”或“列AW-强制关闭时间”（三列只可能有一列存在时间数据）包含在“开始时间”及“结束时间”之间，且“列L-当前工单状态”为“非正常关闭”、“强制关闭”、“已关闭”、“已评价”的excel行数（即工单数）
            data_close = data[(((data["业主关闭时间"] >= dateStart) & (data["业主关闭时间"] <= dateEnd)) |
                               ((data["非正常关闭时间"] >= dateStart) & (data["非正常关闭时间"] <= dateEnd)) |
                               ((data["强制关闭时间"] >= dateStart) & (data["强制关闭时间"] <= dateEnd))) & (
                                          (data["当前工单状态"] == "非正常关闭") |
                                          (data["当前工单状态"] == "强制关闭") |
                                          (data["当前工单状态"] == "已关闭") |
                                          (data["当前工单状态"] == "已评价"))]
            print(data_close["当前工单状态"].count())
            # 当年累计关闭
            dateStartYear = self.string2Year(dateStart)
            dateEndYear = self.string2Year(dateEnd)
            if dateEndYear == dateStartYear:
                dateEndYear = str(int(dateEndYear) + 1)
            elif dateEndYear > dateStartYear:
                pass
            elif dateEndYear < dateStartYear:
                pass
            data_close_year = data[(((data["业主关闭时间"] >= dateStartYear) & (data["业主关闭时间"] <= dateEndYear)) |
                                    ((data["非正常关闭时间"] >= dateStartYear) & (data["非正常关闭时间"] <= dateEndYear)) |
                                    ((data["强制关闭时间"] >= dateStartYear) & (data["强制关闭时间"] <= dateEndYear))) & (
                                               (data["当前工单状态"] == "非正常关闭") |
                                               (data["当前工单状态"] == "强制关闭") |
                                               (data["当前工单状态"] == "已关闭") |
                                               (data["当前工单状态"] == "已评价"))]

            # 总体累计关闭 判断“列L-当前工单状态”为“非正常关闭”、“强制关闭”、“已关闭”、“已评价”的excel行数（即工单数）
            data_all_close = data[((data["当前工单状态"] == "非正常关闭") |
                                   (data["当前工单状态"] == "强制关闭") |
                                   (data["当前工单状态"] == "已关闭") |
                                   (data["当前工单状态"] == "已评价"))]

            print("data_all_close", data_all_close['当前工单状态'].count())
            # 当期响应及时工单：根据“时间段”判断“列AC-响应时间”包含在“开始时间”及“结束时间”之间，且“列AD-受理至响应间隔时长”的值<0.51,的excel行数（即工单数）
            data_response = data[(data["响应时间"] >= dateStart) & (data["响应时间"] <= dateEnd) &
                                 (data["受理至响应间隔时长(小时)\n响应时间 - 受理时间"].apply(lambda x: 0.0 if x == '' else float(x)) < 0.51)]

            data_all = data[(data["响应时间"] >= dateStart) & (data["响应时间"] <= dateEnd)]
            print("data_response:", data_response['当前工单状态'].count())
            print("data_all:", data_all['当前工单状态'].count())
            # 当期上门及时工单: 根据“时间段”判断“列AH - 预约上门时间”包含在“开始时间”及“结束时间”之间，且“列AK - 上门超时”的值 < 0.01, 的excel行数（即工单数）
            data_indoor = data[
                (data["预约上门时间"] >= dateStart) & (data["预约上门时间"] <= dateEnd) & (data["上门超时（小时）\n实际上门时间 - 预约上门时间"]
                                                                               .apply(
                    lambda x: 0.0 if x == '' else float(x)) < 0.01)]

            print("data_indoor", data_indoor['当前工单状态'].count())
            # 当期上门及时率=⑫当期上门及时工单/当期预约上门工单（逻辑：根据“时间段”判断“列AH-预约上门时间”包含在“开始时间”及“结束时间”之间的excel行数（即工单数））
            data_indoor_all = data[(data["预约上门时间"] >= dateStart) & (data["预约上门时间"] <= dateEnd)]

            print("data_indoor_all", data_indoor_all['当前工单状态'].count())
            # 当期施工及时完成工单 = 根据“时间段”判断“列AM - 实际完成时间”包含在“开始时间”及“结束时间”之间，且“列AM - 实际完成时间的值”减去“列AL - 预计完成时间的值”=值＜0.1, 的excel行数（即工单数）
            data_finish = data[(data["实际完成时间"] >= dateStart) & (data["实际完成时间"] <= dateEnd) &
                               ((data[data["实际完成时间"] != '']["实际完成时间"].apply(self.date2Timestamp) -
                                 data[data["预计完成时间"] != '']["预计完成时间"].apply(self.date2Timestamp)) < 0.1)]

            # 当期施工完成及时率 =⑭当期施工及时完成工单 /（根据“时间段”判断“列AM - 实际完成时间”包含在“开始时间”及“结束时间”之间的excel行数（即工单数））
            data_finish_all = data[(data["实际完成时间"] >= dateStart) & (data["实际完成时间"] <= dateEnd)]

            # 当期维修关闭总时长
            data_time_notnull = data[((data["业主关闭时间"] != '') | (data["非正常关闭时间"] != '') | (data["强制关闭时间"] != '')) &
                                     (data["报事时间"] != '')]

            group_name = []
            cur_need_do_list = []
            cur_increase_list = []
            cur_close_list = []
            cur_not_close_list = []
            cur_close_rate_list = []
            cur_year_close_rate_list = []
            all_close_rate_list = []
            indoor_rate_list = []
            finish_rate_list = []
            dur_lists = []
            response_rate_list = []
            response_cnt_list = []
            indoor_cnt_list = []
            finish_cnt_list = []

            items = []
            for group in cal_group:
                cur_need_do = 0
                cur_increase = 0
                cur_close = 0
                cur_not_close = 0
                cur_close_rate = 0.0
                cur_year_close_rate = 0.0
                all_close_rate = 0.0
                indoor_rate = 0.0
                finish_rate = 0.0
                dur = 0.0
                response_rate = 0.0
                response_cnt = 0
                indoor_cnt = 0
                finish_cnt = 0
                print("group:", group)
                if cal_group_name == "项目分期":
                    pre_group = set(data[data[cal_group_name] == group]['项目'].values)
                    for pre in pre_group:
                        item = []
                        item.append(pre + group)
                        # 1.当前未关闭
                        data_cur_tmp = data_cur[(data_cur[cal_group_name] == group) & (data_cur['项目'] == pre)]
                        cur_not_close = data_cur_tmp["当前工单状态"].count()
                        # print("cur_not_close:", cur_not_close)
                        # 2.当前新增
                        data_increase_tmp = data_increase[
                            (data_increase[cal_group_name] == group) & (data_increase['项目'] == pre)]
                        cur_increase = data_increase_tmp["当前工单状态"].count()

                        # 3.当前关闭
                        data_close_tmp = data_close[(data_close[cal_group_name] == group) & (data_close['项目'] == pre)]
                        cur_close = data_close_tmp["当前工单状态"].count()

                        # print(cur_close)
                        # 4.当前需处理
                        cur_need_do = cur_not_close + cur_close
                        # print("cur_need_do:", cur_need_do)
                        # 5.当期关闭率 = 当期关闭/当期需处理
                        if cur_need_do != 0:
                            cur_close_rate = round(cur_close / cur_need_do, 4)
                        else:
                            pass

                        # 6.当年累计关闭
                        data_close_year_tmp = data_close_year[
                            (data_close_year[cal_group_name] == group) & (data_close_year['项目'] == pre)]
                        cur_year_close = data_close_year_tmp["当前工单状态"].count()

                        # 7.当年累计关闭率 = ⑥当年累计关闭/（⑥当年累计关闭+①当期未关闭
                        if (cur_year_close + cur_not_close) == 0:
                            cur_year_close_rate = 0
                        else:
                            cur_year_close_rate = round(cur_year_close / (cur_year_close + cur_not_close), 4)

                        # 8.总体累计关闭
                        data_all_close_tmp = data_all_close[
                            (data_all_close[cal_group_name] == group) & (data_all_close['项目'] == pre)]
                        all_close = data_all_close_tmp["当前工单状态"].count()

                        # 9.总体累计关闭率=⑧总体累计关闭/（⑧总体累计关闭+①当期未关闭）
                        if (all_close + cur_not_close) == 0:
                            all_close_rate = 0
                        else:
                            all_close_rate = round(all_close / (all_close + cur_not_close), 4)

                        # 10.当期响应及时工单
                        data_response_tmp = data_response[
                            (data_response[cal_group_name] == group) & (data_response['项目'] == pre)]
                        response_cnt = data_response_tmp["当前工单状态"].count()

                        # 11.当期响应及时率=⑩当期响应及时工单/当期所有工单
                        data_all_tmp = data_all[(data_all[cal_group_name] == group) & (data_all['项目'] == pre)]
                        all_cnt = data_all_tmp["当前工单状态"].count()
                        if all_cnt == 0:
                            pass
                        else:
                            response_rate = round(response_cnt / all_cnt, 4)

                        # 12.当期上门及时工单
                        data_indoor_tmp = data_indoor[(data_indoor[cal_group_name] == group) & (data_indoor['项目'] == pre)]
                        indoor_cnt = data_indoor_tmp["当前工单状态"].count()

                        # 13.当期上门及时率=⑫当期上门及时工单/当期预约上门工单
                        data_indoor_all_tmp = data_indoor_all[
                            (data_indoor_all[cal_group_name] == group) & (data_indoor_all['项目'] == pre)]
                        indoor_all_cnt = data_indoor_all_tmp["当前工单状态"].count()
                        if indoor_all_cnt == 0:
                            pass
                        else:
                            indoor_rate = round(indoor_cnt / indoor_all_cnt, 4)

                        # 14.当期施工及时完成工单
                        data_finish_tmp = data_finish[(data_finish[cal_group_name] == group) & (data_finish['项目'] == pre)]
                        finish_cnt = data_finish_tmp["当前工单状态"].count()

                        # 15.当期施工完成及时率
                        data_finish_all_tmp = data_finish_all[
                            (data_finish_all[cal_group_name] == group) & (data_finish_all['项目'] == pre)]
                        finish_cnt_all = data_finish_all_tmp["当前工单状态"].count()
                        finish_rate = 0
                        if finish_cnt_all != 0:
                            finish_rate = round(finish_cnt / finish_cnt_all, 4)
                        else:
                            pass

                        # print("finish_cnt:", finish_cnt, "finish_cnt_all", finish_cnt_all)

                        # 16.当期维修关闭总时长 平均时长
                        data_time_notnull_tmp = data_time_notnull[
                            (data_time_notnull[cal_group_name] == group) & (data_time_notnull['项目'] == pre)]
                        data_dur = data_time_notnull_tmp["业主关闭时间"].apply(
                            lambda x: 0 if x == '' else self.date2Timestamp(x)) + \
                                   data_time_notnull_tmp["非正常关闭时间"].apply(
                                       lambda x: 0 if x == '' else self.date2Timestamp(x)) + \
                                   data_time_notnull_tmp["强制关闭时间"].apply(
                                       lambda x: 0 if x == '' else self.date2Timestamp(x)) - \
                                   data_time_notnull_tmp["报事时间"].apply(lambda x: 0 if x == '' else self.date2Timestamp(x))
                        dur_list = data_dur.tolist()
                        # print("dur_list", dur_list)
                        dur = 0
                        if len(dur_list) != 0:
                            dur = round(sum(dur_list) / len(dur_list) / 1000 / 60 / 60, 2)
                        else:
                            pass

                        group_name.append(pre + group)
                        cur_need_do_list.append(cur_need_do)
                        cur_increase_list.append(cur_increase)
                        cur_close_list.append(cur_close)
                        cur_not_close_list.append(cur_not_close)
                        cur_close_rate_list.append(cur_close_rate)
                        cur_year_close_rate_list.append(cur_year_close_rate)
                        all_close_rate_list.append(all_close_rate)
                        indoor_rate_list.append(indoor_rate)
                        finish_rate_list.append(finish_rate)
                        dur_lists.append(dur)
                        response_rate_list.append(response_rate)
                        response_cnt_list.append(response_cnt)
                        indoor_cnt_list.append(indoor_cnt)
                        finish_cnt_list.append(finish_cnt)

                        item.append(cur_need_do)
                        item.append(cur_increase)
                        item.append(cur_close)
                        item.append(cur_not_close)
                        item.append(cur_close_rate)
                        item.append(cur_year_close_rate)
                        item.append(all_close_rate)
                        item.append(indoor_rate)
                        item.append(finish_rate)
                        item.append(dur)
                        item.append(response_rate)
                        item.append(response_cnt)
                        item.append(indoor_cnt)
                        item.append(finish_cnt)

                        items.append(item)

                        self.pd_dict = {'分组名称': group_name,
                                        '当前需处理(条)': cur_need_do_list,
                                        '当前新增(条)': cur_increase_list,
                                        '当前关闭(条)': cur_close_list,
                                        '当前未关闭(条)': cur_not_close_list,
                                        '当前关闭率': cur_close_rate_list,
                                        '当年累计关闭率': cur_year_close_rate_list,
                                        '总体累计关闭率': all_close_rate_list,
                                        '当期上门及时率': indoor_rate_list,
                                        '当期施工完成及时率': finish_rate_list,
                                        '平均关单时长(小时)': dur_lists,
                                        '当期响应及时率': response_rate_list
                                        }

                else:
                    item = []
                    print('group:', group)
                    item.append(group)
                    # 1.当前未关闭
                    data_cur_tmp = data_cur[data_cur[cal_group_name] == group]
                    cur_not_close = data_cur_tmp["当前工单状态"].count()

                    # 2.当前新增
                    data_increase_tmp = data_increase[data_increase[cal_group_name] == group]
                    cur_increase = data_increase_tmp["当前工单状态"].count()

                    # 3.当前关闭
                    data_close_tmp = data_close[data_close[cal_group_name] == group]
                    cur_close = data_close_tmp["当前工单状态"].count()

                    # print(cur_close)
                    # 4.当前需处理
                    cur_need_do = cur_not_close + cur_close

                    # 5.当期关闭率 = 当期关闭/当期需处理
                    if cur_need_do != 0:
                        cur_close_rate = round(cur_close / cur_need_do, 4)
                    else:
                        pass

                    # 6.当年累计关闭
                    data_close_year_tmp = data_close_year[data_close_year[cal_group_name] == group]
                    cur_year_close = data_close_year_tmp["当前工单状态"].count()

                    # 7.当年累计关闭率 = ⑥当年累计关闭/（⑥当年累计关闭+①当期未关闭
                    if (cur_year_close + cur_not_close) != 0:
                        cur_year_close_rate = round(cur_year_close / (cur_year_close + cur_not_close), 4)
                    else:
                        pass

                    # 8.总体累计关闭
                    data_all_close_tmp = data_all_close[data_all_close[cal_group_name] == group]
                    all_close = data_all_close_tmp["当前工单状态"].count()

                    # 9.总体累计关闭率=⑧总体累计关闭/（⑧总体累计关闭+①当期未关闭）
                    if (all_close + cur_not_close) != 0:
                        all_close_rate = round(all_close / (all_close + cur_not_close), 4)
                    else:
                        pass

                    # 10.当期响应及时工单
                    data_response_tmp = data_response[data_response[cal_group_name] == group]
                    response_cnt = data_response_tmp["当前工单状态"].count()

                    # 11.当期响应及时率=⑩当期响应及时工单/当期所有工单
                    data_all_tmp = data_all[data_all[cal_group_name] == group]
                    all_cnt = data_all_tmp["当前工单状态"].count()
                    if all_cnt != 0:
                        response_rate = round(response_cnt / all_cnt, 4)
                    else:
                        pass

                    # 12.当期上门及时工单
                    data_indoor_tmp = data_indoor[data_indoor[cal_group_name] == group]
                    indoor_cnt = data_indoor_tmp["当前工单状态"].count()

                    # 13.当期上门及时率=⑫当期上门及时工单/当期预约上门工单
                    data_indoor_all_tmp = data_indoor_all[data_indoor_all[cal_group_name] == group]
                    indoor_all_cnt = data_indoor_all_tmp["当前工单状态"].count()
                    if indoor_all_cnt != 0:
                        indoor_rate = round(indoor_cnt / indoor_all_cnt, 4)
                    else:
                        pass
                    print("indoor_cnt:", indoor_cnt, "indoor_all_cnt", indoor_all_cnt)

                    # 14.当期施工及时完成工单
                    data_finish_tmp = data_finish[data_finish[cal_group_name] == group]
                    finish_cnt = data_finish_tmp["当前工单状态"].count()

                    # 15.当期施工完成及时率
                    data_finish_all_tmp = data_finish_all[data_finish_all[cal_group_name] == group]
                    finish_cnt_all = data_finish_all_tmp["当前工单状态"].count()
                    finish_rate = 0
                    if finish_cnt_all != 0:
                        finish_rate = round(finish_cnt / finish_cnt_all, 4)
                    else:
                        pass

                    print("finish_cnt:", finish_cnt, "finish_cnt_all", finish_cnt_all)

                    # 16.当期维修关闭总时长 平均时长
                    data_time_notnull_tmp = data_time_notnull[data_time_notnull[cal_group_name] == group]
                    data_dur = data_time_notnull_tmp["业主关闭时间"].apply(lambda x: 0 if x == '' else self.date2Timestamp(x)) + \
                               data_time_notnull_tmp["非正常关闭时间"].apply(lambda x: 0 if x == '' else self.date2Timestamp(x)) + \
                               data_time_notnull_tmp["强制关闭时间"].apply(lambda x: 0 if x == '' else self.date2Timestamp(x)) - \
                               data_time_notnull_tmp["报事时间"].apply(lambda x: 0 if x == '' else self.date2Timestamp(x))
                    dur_list = data_dur.tolist()
                    # print("dur_list", dur_list)
                    dur = 0
                    if len(dur_list) != 0:
                        dur = round(sum(dur_list) / len(dur_list) / 1000 / 60 / 60, 2)
                    else:
                        pass

                    group_name.append(group)
                    cur_need_do_list.append(cur_need_do)
                    cur_increase_list.append(cur_increase)
                    cur_close_list.append(cur_close)
                    cur_not_close_list.append(cur_not_close)
                    cur_close_rate_list.append(cur_close_rate)
                    cur_year_close_rate_list.append(cur_year_close_rate)
                    all_close_rate_list.append(all_close_rate)
                    indoor_rate_list.append(indoor_rate)
                    finish_rate_list.append(finish_rate)
                    dur_lists.append(dur)
                    response_rate_list.append(response_rate)
                    response_cnt_list.append(response_cnt)
                    indoor_cnt_list.append(indoor_cnt)
                    finish_cnt_list.append(finish_cnt)

                    item.append(cur_need_do)
                    item.append(cur_increase)
                    item.append(cur_close)
                    item.append(cur_not_close)
                    item.append(cur_close_rate)
                    item.append(cur_year_close_rate)
                    item.append(all_close_rate)
                    item.append(indoor_rate)
                    item.append(finish_rate)
                    item.append(dur)
                    item.append(response_rate)
                    item.append(response_cnt)
                    item.append(indoor_cnt)
                    item.append(finish_cnt)

                    items.append(item)

                    self.pd_dict = {'分组名称': group_name,
                                    '当前需处理(条)': cur_need_do_list,
                                    '当前新增(条)': cur_increase_list,
                                    '当前关闭(条)': cur_close_list,
                                    '当前未关闭(条)': cur_not_close_list,
                                    '当前关闭率': cur_close_rate_list,
                                    '当年累计关闭率': cur_year_close_rate_list,
                                    '总体累计关闭率': all_close_rate_list,
                                    '当期上门及时率': indoor_rate_list,
                                    '当期施工完成及时率': finish_rate_list,
                                    '平均关单时长(小时)': dur_lists,
                                    '当期响应及时率': response_rate_list,
                                    '当期响应及时工单': response_cnt_list,
                                    '当期上门及时工单': indoor_cnt_list,
                                    '当期施工及时完成工单': finish_cnt_list
                                    }

            for i in range(len(items)):
                item = items[i]
                row = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row)
                for j in range(len(item)):
                    item = QTableWidgetItem(str(items[i][j]))
                    self.tableWidget.setItem(row, j, item)


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
                            line = self.tab_3.findChild(QLineEdit, str(p) + "j_lineEdit").text()
                            lines = line.split('#')
                            if lines[1] != '':
                                if lines[0] == '楼栋':
                                    df = self.data[self.data['楼栋'] == lines[1]]
                                    df_list.append(df)
                                elif lines[0] == '项目分期':
                                    line2 = lines[1].split("&")
                                    df2 = self.data[(self.data['项目'] == line2[0]) & (self.data['项目分期'] == line2[1])]
                                    df_list.append(df2)
                    if df_list:
                        datas = pd.concat(df_list)
                        self.cal(datas, cur_sel)
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
                            if lines[1] != '':
                                if lines[0] == '楼栋':
                                    datas = self.data[self.data['楼栋'] != lines[1]]
                                elif lines[0] == '项目分期':
                                    line2 = lines[1].split("&")
                                    datas = self.data[(self.data['项目'] != line2[0]) & (self.data['项目分期'] != line2[1])]
                self.cal(datas)




if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = Window()
    mywindow.setWindowTitle("物业管理计算系统")
    mywindow.show()
    sys.exit(app.exec_())

