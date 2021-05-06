import sys
import fist
from PyQt5.QtWidgets import QApplication, QMainWindow, QHeaderView, QTableWidgetItem
from PyQt5 import QtWidgets
from PyQt5.QtGui import QStandardItemModel
from PyQt5 import QtCore, QtGui
import pandas as pd
import datetime
import numpy as np
import time

class Window(fist.Ui_MainWindow, QtWidgets.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.msg)#一个按钮的点击事件，响应函数为 def msg(self):
        self.pushButton_2.clicked.connect(self.calTableValue)
        self.lineEdit.setAcceptDrops(True)


    def msg(self):
        filePath, filetype = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", "./", "*.*")
        self.path = filePath
        self.lineEdit.setText(filePath)

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


    def calTableValue(self):
        self.tableWidget.setRowCount(0)
        self.tableWidget.clearContents()
        # keep_default_na=False 防止出现nan值
        data = pd.read_excel(self.path, keep_default_na=False)
        dateStart = str(self.dateStart.dateTime().toPyDateTime())
        dateEnd = str(self.dateEnd.dateTime().toPyDateTime())
        cal_group_name = self.comboBox.currentText()
        cal_group = np.unique(data[cal_group_name].values).tolist()
        print(cal_group)
        print(dateEnd)
        # 当期未关闭：根据“时间段”判断“列AA-报事时间”早于“结束时间”， 且“列L-当前工单状态”为“方案已批准”、“方案制定中”、“施工完成”、“施工中”、“已响应”的excel行数（即工单数）
        data_cur = data[(data["报事时间"] <= dateEnd) & ((data["当前工单状态"] == "方案已批准") |
                                                       (data["当前工单状态"] == "方案制定中") |
                                                       (data["当前工单状态"] == "施工完成") |
                                                       (data["当前工单状态"] == "施工中") |
                                                       (data["当前工单状态"] == "已响应"))]
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
                                                       (data["当前工单状态"] == "已评价"))]

        print("data_increase", data_increase["当前工单状态"].count())

        # 当期关闭：根据“时间段”判断“列AU-业主关闭时间”或“列AV-非正常关闭时间”或“列AW-强制关闭时间”（三列只可能有一列存在时间数据）包含在“开始时间”及“结束时间”之间，且“列L-当前工单状态”为“非正常关闭”、“强制关闭”、“已关闭”、“已评价”的excel行数（即工单数）
        data_close = data[(((data["业主关闭时间"] >= dateStart) & (data["业主关闭时间"] <= dateEnd)) |
                    ((data["非正常关闭时间"] >= dateStart) & (data["非正常关闭时间"] <= dateEnd)) |
                    ((data["强制关闭时间"] >= dateStart) & (data["强制关闭时间"] <= dateEnd))) & ((data["当前工单状态"] == "非正常关闭") |
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
                    ((data["强制关闭时间"] >= dateStartYear) & (data["强制关闭时间"] <= dateEndYear))) & ((data["当前工单状态"] == "非正常关闭") |
                                                                                      (data["当前工单状态"] == "强制关闭") |
                                                                                      (data["当前工单状态"] == "已关闭") |
                                                                                      (data["当前工单状态"] == "已评价"))]

        # 总体累计关闭 判断“列L-当前工单状态”为“非正常关闭”、“强制关闭”、“已关闭”、“已评价”的excel行数（即工单数）
        data_all_close = data[((data["当前工单状态"] == "非正常关闭") |
                                (data["当前工单状态"] == "强制关闭") |
                                (data["当前工单状态"] == "已关闭") |
                                (data["当前工单状态"] == "已评价"))]

        print("data_all_close", data_all_close['当前工单状态']. count())
        # 当期响应及时工单：根据“时间段”判断“列AC-响应时间”包含在“开始时间”及“结束时间”之间，且“列AD-受理至响应间隔时长”的值<0.51,的excel行数（即工单数）
        data_response = data[(data["响应时间"] >= dateStart) & (data["响应时间"] <= dateEnd) &
                             (data["受理至响应间隔时长(小时)\n响应时间 - 受理时间"].apply(lambda x: 0.0 if x == '' else float(x)) < 0.51)]

        data_all = data[(data["响应时间"] >= dateStart) & (data["响应时间"] <= dateEnd)]

        # 当期上门及时工单: 根据“时间段”判断“列AH - 预约上门时间”包含在“开始时间”及“结束时间”之间，且“列AK - 上门超时”的值 < 0.01, 的excel行数（即工单数）
        data_indoor = data[(data["预约上门时间"] >= dateStart) & (data["预约上门时间"] <= dateEnd) & (data["上门超时（小时）\n实际上门时间 - 预约上门时间"]
                           .apply(lambda x: 0.0 if x == '' else float(x)) < 0.01)]

        # 当期上门及时率=⑫当期上门及时工单/当期预约上门工单（逻辑：根据“时间段”判断“列AH-预约上门时间”包含在“开始时间”及“结束时间”之间的excel行数（即工单数））
        data_indoor_all = data[(data["预约上门时间"] >= dateStart) & (data["预约上门时间"] <= dateEnd)]

        print("data_indoor_all", data_indoor_all['当前工单状态'].count())
        # 当期施工及时完成工单 = 根据“时间段”判断“列AM - 实际完成时间”包含在“开始时间”及“结束时间”之间，且“列AM - 实际完成时间的值”减去“列AL - 预计完成时间的值”=值＜0.1, 的excel行数（即工单数）
        data_finish = data[(data["实际完成时间"] >= dateStart) & (data["实际完成时间"] <= dateEnd) &
                           ((data[data["实际完成时间"] != '']["实际完成时间"] .apply(self.date2Timestamp) -
                             data[data["预计完成时间"] != '']["预计完成时间"] .apply(self.date2Timestamp)) < 0.1)]

        # 当期施工完成及时率 =⑭当期施工及时完成工单 /（根据“时间段”判断“列AM - 实际完成时间”包含在“开始时间”及“结束时间”之间的excel行数（即工单数））
        data_finish_all = data[(data["实际完成时间"] >= dateStart) & (data["实际完成时间"] <= dateEnd)]

        # 当期维修关闭总时长
        data_time_notnull = data[((data["业主关闭时间"] != '') | (data["非正常关闭时间"] != '') | (data["强制关闭时间"] != '')) &
             (data["报事时间"] != '')]

        items = []
        for group in cal_group:
            item = []
            print(group)
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
            cur_close_rate = round(cur_close / cur_need_do, 4)

            # 6.当年累计关闭
            data_close_year_tmp = data_close_year[data_close_year[cal_group_name] == group]
            cur_year_close = data_close_year_tmp["当前工单状态"].count()

            # 7.当年累计关闭率 = ⑥当年累计关闭/（⑥当年累计关闭+①当期未关闭
            cur_year_close_rate = round(cur_year_close/(cur_year_close + cur_not_close), 4)

            # 8.总体累计关闭
            data_all_close_tmp = data_all_close[data_all_close[cal_group_name] == group]
            all_close = data_all_close_tmp["当前工单状态"].count()

            # 9.总体累计关闭率=⑧总体累计关闭/（⑧总体累计关闭+①当期未关闭）
            all_close_rate = round(all_close /(all_close + cur_not_close), 4)

            # 10.当期响应及时工单
            data_response_tmp = data_response[data_response[cal_group_name] == group]
            response_cnt = data_response_tmp["当前工单状态"].count()

            # 11.当期响应及时率=⑩当期响应及时工单/当期所有工单
            data_all_tmp = data_all[data_all[cal_group_name] == group]
            all_cnt = data_all_tmp["当前工单状态"].count()
            response_rate = round(response_cnt/all_cnt, 4)

            # 12.当期上门及时工单
            data_indoor_tmp = data_indoor[data_indoor[cal_group_name] == group]
            indoor_cnt = data_indoor_tmp["当前工单状态"].count()

            # 13.当期上门及时率=⑫当期上门及时工单/当期预约上门工单
            data_indoor_all_tmp = data_indoor_all[data_indoor_all[cal_group_name] == group]
            indoor_all_cnt = data_indoor_all_tmp["当前工单状态"].count()
            indoor_rate = round(indoor_cnt/indoor_all_cnt, 4)
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
            print("dur_list", dur_list)
            dur = 0
            if len(dur_list) != 0:
                dur = round(sum(dur_list)/len(dur_list)/1000/60, 2)
            else:
                pass


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
            items.append(item)

        for i in range(len(items)):
            item = items[i]
            row = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row)
            for j in range(len(item)):
                item = QTableWidgetItem(str(items[i][j]))
                self.tableWidget.setItem(row, j, item)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mywindow = Window()
    mywindow.show()
    sys.exit(app.exec_())

