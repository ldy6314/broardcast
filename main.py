import os

import win32event,  win32api
from winerror import ERROR_ALREADY_EXISTS
from maiwindow import Ui_MainWindow
from newtaskdig1 import Ui_Dialog
from PyQt5.QtWidgets import QApplication, QHeaderView, QMainWindow, QDialog, QFileDialog, QAbstractItemView, \
    QTableWidgetItem, QMessageBox, QMenu, QToolTip, QSystemTrayIcon, QAction, qApp
from PyQt5.QtCore import Qt, QTimer, QUrl
from timeparse import *
from apscheduler.schedulers.qt import QtScheduler
from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR, EVENT_JOB_MISSED
from sys import argv, exit, platform
from datetime import datetime
from PyQt5.QtSql import QSqlDatabase, QSqlQuery
from PyQt5 import QtMultimedia
from PyQt5.QtGui import QCursor, QIcon, QPixmap
from hashlib import md5
from os import mkdir, getcwd
from PyQt5 import QtGui
import logging
from logging.handlers import RotatingFileHandler

os.environ["PBR_VERSION"] = "4.0.2"

logger = logging.getLogger()
logger.setLevel(logging.INFO)
fh = RotatingFileHandler(
    filename='log.txt',
    maxBytes=1024*1024,
    backupCount=3
)
fmt = logging.Formatter('[%(asctime)s] [%(levelname)s]%(message)s', '%Y-%m-%d %H:%M:%S')
fh.setFormatter(fmt)
logger.addHandler(fh)


class Dialog(QDialog, Ui_Dialog):
    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        self.setupUi(self)
        self.setFixedSize(self.width(), self.height())
        self.chk_list = [self.checkBox, self.checkBox_2, self.checkBox_3, self.checkBox_4, self.checkBox_5,
                         self.checkBox_6, self.checkBox_7]
        self.checkBox_8.stateChanged.connect(lambda: self.multi_select(self.checkBox_8, (0, 7)))
        self.checkBox_9.stateChanged.connect(lambda: self.multi_select(self.checkBox_9, (0, 5)))
        self.toolButton.clicked.connect(self.set_file)

    def get_value(self):
        bell_name = self.lineEdit_2.text()
        bell_type = self.comboBox.currentText()
        file = self.lineEdit.text()
        weekdays_value = 0
        time_value = self.spinBox.value() * 3600 + self.spinBox_2.value() * 60 + self.spinBox_3.value()
        for i in zip(self.chk_list, range(7)):
            if i[0].checkState():
                weekdays_value += 1 << i[1]
        return bell_name, bell_type, file, weekdays_value, time_value

    def multi_select(self, checkbox, rg):
        state = checkbox.isChecked()
        for i in range(rg[0], rg[1]):
            self.chk_list[i].setChecked(state)

    def set_value(self, value=("", "????????????", "", 0, 0)):
        # bell_name, bell_type, file, weekdays_value, time_value
        self.lineEdit_2.setText(value[0])
        self.comboBox.setCurrentText(value[1])
        self.lineEdit.setText(value[2])
        weekdays = parse_to_weekdays(value[3])
        h, m, s = parse_to_time(value[4])
        for checkbox in self.chk_list:
            checkbox.setChecked(False)
        for i in weekdays:
            self.chk_list[i - 1].setChecked(True)
        self.spinBox.setValue(h)
        self.spinBox_2.setValue(m)
        self.spinBox_3.setValue(s)
        self.checkBox_9.setChecked(False)
        self.checkBox_8.setChecked(False)

    def set_file(self):
        file_name, tp = QFileDialog.getOpenFileName(directory=os.getcwd() + '/music', filter="Music Files (*.mp3)")
        self.lineEdit.setText(file_name)


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.dialog = Dialog(self)
        self.tray_icon = TrayIcon(self)
        try:
            mkdir("./music")
        except FileExistsError:
            pass
        # ??????????????????
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.setMouseTracking(True)
        # tabWidget?????????0
        self.textBrowser_2.setStyleSheet("color:blue")
        self.groupBox_4.hide()
        # ???????????????
        self.db = QSqlDatabase.addDatabase("QSQLITE")
        self.db.setDatabaseName("./task2.db")
        self.db.open()
        self.q = QSqlQuery()
        # ######????????????????????????scheduler???????????????????????????????????????
        self.q.exec_("create table schedules(name TEXT,type TEXT, file TEXT, weekdays INT, time INT PRIMARY KEY, "
                     "enable INT)")
        self.q.exec_("create table config (id int primary key, weather_mode int, rest_mode int, weekend int, "
                     "weekday int)")
        self.config_load()
        # ??????????????????
        job_defaults = {
            'coalesce': True,
            'misfire_grace_time': 10
        }
        self.sch = QtScheduler(timezone='Asia/Shanghai', job_defaults=job_defaults)
        self.sch.add_listener(self.job_listener, EVENT_JOB_ERROR | EVENT_JOB_MISSED | EVENT_JOB_EXECUTED)
        self.sch.start()
        self.load_schedules()
        # ??????????????????
        self.timer1 = QTimer()
        self.timer1.start(1000)
        self.timer1.timeout.connect(self.show_current_time)
        # ??????????????????
        self.player = QtMultimedia.QMediaPlayer()


        # ?????????????????????
        self.pushButton.clicked.connect(lambda: self.show_config_task_dialog())
        self.pushButton_1.clicked.connect(self.remove_selected_task)
        self.pushButton_2.clicked.connect(self.play_selected_music)
        self.pushButton_3.clicked.connect(lambda: self.show_config_task_dialog(1))
        self.pushButton_4.clicked.connect(self.multi_change_music)
        self.pushButton_5.clicked.connect(self.stop_music)
        self.pushButton_6.clicked.connect(self.import_data)
        self.pushButton_7.clicked.connect(self.export_data)
        self.pushButton_8.clicked.connect(self.clear_database)
        self.pushButton_9.clicked.connect(self.change_music_path)
        self.pushButton_10.clicked.connect(self.pause_resume_task)
        self.toolButton.clicked.connect(lambda: self.choose_music())
        self.toolButton_2.clicked.connect(lambda: self.choose_music(1))
        self.tableWidget.customContextMenuRequested.connect(self.generate_menu)
        self.tableWidget.cellEntered.connect(self.show_cell_in_tooltip)
        self.checkBox.stateChanged.connect(self.change_weather_mode)
        self.checkBox_2.stateChanged.connect(self.change_rest_mode)
        self.tray_icon.activated[QSystemTrayIcon.ActivationReason].connect(self.tray_icon.on_icon_clicked)
        self.player.durationChanged.connect(self.update_tot_time)
        self.player.positionChanged.connect(self.print_info)
        self.player.stateChanged.connect(self.show_hide_control)
        self.pause_play_button.clicked.connect(self.pause_play_music)
        self.music_Slider.sliderReleased.connect(self.drag_music)
        self.tabWidget.currentChanged.connect(self.update_log)

    def get_qss(self):
        pass

    def load_log(self, filename):
        with open(filename) as f:
            txt = f.read()
            self.textBrowser_2.setText(txt)
            self.textBrowser_2.moveCursor(self.textBrowser_2.textCursor().End)

    def update_log(self, idx):
        if idx == 2:
            self.load_log('log.txt')

    def job_listener(self, event):
        job = self.sch.get_job(event.job_id)
        if not event.exception:
            logger.info("{} has executed successfully".format(job.name))
        else:
            logger.error("{} executed error={} traceback={}".format(job.name, event.exception, event.traceback))

    def move_scroll_top(self, idx):
        if idx == 3:
            self.textBrowser_2.moveCursor(self.textBrowser_2.textCursor().Start)

    def load_html(self, file):
        with open(file, 'r', encoding='gbk') as fin:
            txt = fin.read()
            self.textBrowser_2.insertHtml(txt)

    def pause_play_music(self):
        if self.player.state() == 1:
            self.player.pause()
            icon = QIcon()
            icon.addPixmap(QPixmap(":/pictures/play.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            self.pause_play_button.setIcon(icon)

        else:
            icon = QIcon()
            icon.addPixmap(QPixmap(":/pictures/pause.png"))
            self.pause_play_button.setIcon(icon)
            self.player.play()

    def show_hide_control(self, state):
        if state == 0:
            self.groupBox_4.hide()

        elif state == 1:
            self.groupBox_4.show()

    def drag_music(self):
        value = self.music_Slider.value()
        self.player.setPosition(value)

    def print_info(self):
        cur = self.player.position()
        m = cur // 1000 // 60
        s = cur // 1000 % 60
        ms = cur % 1000
        self.current_time_label.setText('{}:{:02}.{}'.format(m, s, ms))
        self.music_Slider.setValue(self.player.position())

    def update_tot_time(self, dur):
        if dur:
            tot = dur
            m = tot // 1000 // 60
            s = tot // 1000 % 60
            ms = tot % 1000
            # print(m, s, ms)
            self.totle_time_label.setText('{}:{:02}.{}'.format(m, s, ms))
            self.music_Slider.setMaximum(tot)

    def get_config(self):
        self.q.exec_("select * from config")
        res = []
        while self.q.next():
            lst = []
            for i in range(5):
                lst.append(self.q.value(i))
            res.append(tuple(lst))
        return res

    def config_load(self):
        res = self.get_config()
        if not res:
            self.q.exec_("insert into config values(0,0,0,0,0)")

        else:
            weather, rest, wend, weekday = res[0][1:]
            self.checkBox.setChecked(weather != 0)
            self.checkBox_2.setChecked(rest != 0)
            self.comboBox.setCurrentIndex(wend)
            self.comboBox_2.setCurrentIndex(weekday)
            if rest:
                self.comboBox.setEnabled(False)
                self.comboBox_2.setEnabled(False)

    def show_config_task_dialog(self, flag=0):
        """
        ?????????????????????????????? ????????????????????????????????????????????????
        :param flag: flag = ???0 ??????????????? 1 ???????????????
        :return:
        """
        if not flag:
            self.add_task()
        else:
            self.alter_task()

    def show_current_time(self):
        """
        :???????????????????????????
        """
        now = datetime.now()
        current_time = now.strftime('%Y-%m-%d %H:%M:%S')
        self.label_3.setText(current_time)

    def show_schedules_on_table(self):
        """
        ?????????????????????????????????????????????
        """
        self.tableWidget.setRowCount(0)
        self.q.exec_("select * from schedules where enable=1 order by time")
        while self.q.next():
            lst = []
            for i in range(5):
                lst.append(self.q.value(i))
            lst[3] = str(parse_to_weekdays(lst[3])).strip('(').strip(')')
            tm = parse_to_time(lst[4])
            lst[4] = "{:02}:{:02}:{:02}".format(tm[0], tm[1], tm[2])
            lst = [lst[0], lst[4], lst[3], lst[1], lst[2]]
            row_cnt = self.tableWidget.rowCount()
            self.tableWidget.setRowCount(row_cnt + 1)
            for i in range(5):
                item = QTableWidgetItem(lst[i])
                self.tableWidget.setItem(row_cnt, i, item)

    def generate_menu(self, pos):
        row_num = self.get_selected_row()
        if row_num < self.tableWidget.rowCount() and row_num != -1:
            menu = QMenu()
            item1 = menu.addAction(u"????????????")
            item2 = menu.addAction(u"????????????")
            item3 = menu.addAction(u"????????????")
            item4 = menu.addAction(u"????????????")
            item5 = menu.addAction(u"????????????")
            item6 = menu.addAction(u"??????")
            item7 = menu.addAction(u"??????")
            item8 = menu.addAction(u"?????????????????????")
            action = menu.exec_(self.tableWidget.mapToGlobal(pos))
            if action == item1:
                self.remove_selected_task()
            elif action == item2:
                self.add_task()
            elif action == item3:
                self.alter_task()
            elif action == item4:
                self.play_selected_music()
            elif action == item5:
                self.stop_music()
            elif action == item6:
                self.import_data()
            elif action == item7:
                self.export_data()
            elif action == item8:
                try:
                    self.open_music_folder(self)
                except Exception as e:
                    print(e)
            else:
                return

    @staticmethod
    def open_music_folder(self):
        path = os.getcwd() + '\\music'
        if platform == 'win32':
            os.startfile(path)
        else:
            cmd = 'nautilus'
            os.system("{} {}".format(cmd, path))

    def add_task(self):
        """
        ??????????????????????????????
        :param ???value ???????????????????????????????????????
        :return:
        """
        self.dialog.setWindowTitle('????????????')
        self.dialog.set_value()
        if self.dialog.exec_():
            value = self.dialog.get_value()
            self.q.exec_("select * from schedules where time = {}".format(value[4]))
            exist = self.q.next()
            if exist:
                QMessageBox.information(self, "??????", "?????????????????????????????????????????????????????????????????????", QMessageBox.Yes)
                logger.info('????????????{}??????'.format(value))
                return False
            else:
                enable = 1
                try:

                    res = self.get_config()[0]
                    # ????????????,???????????????????????????
                    if not res[1] and value[1] == '????????????' or res[1] and value[1] == '????????????' or \
                            res[2] and value[1] == '????????????':
                        enable = 0
                except IndexError:
                    pass

                self.q.exec_("insert into schedules values('{}','{}','{}',{},{},{})".format(value[0], value[1], value[2]
                                                                                            , value[3], value[4],
                                                                                            enable))

            logger.info('????????????{}??????'.format(value))
            self.load_schedules()
            return True

    def alter_task(self):
        """
        ?????????????????????????????????
        :return:
        """
        self.dialog.setWindowTitle('????????????')
        row = self.get_selected_row()
        if row == -1:
            return
        name, tm, wk, tp, fl = [self.tableWidget.item(row, i).text() for i in range(5)]
        old_value = [name, tp, fl, weekdays_to_number(wk), time_to_number(tm)]
        self.dialog.set_value(old_value)
        if self.dialog.exec_():
            new_value = self.dialog.get_value()
            self.q.exec_("select time from schedules where time={}".format(new_value[4]))
            self.q.next()
            res = self.q.value(0)
            if res and res != old_value[4]:
                QMessageBox.information(self, "??????", "?????????????????????????????????????????????????????????????????????", QMessageBox.Yes)
                logger.info("????????????{}->{}??????".format(old_value, new_value))
                return False
            else:
                res = self.get_config()[0]
                enable = 1
                if not res[1] and new_value[1] == '????????????' or res[1] and new_value[1] == '????????????' or \
                        res[2] and new_value[1] == '????????????':
                    enable = 0
                self.q.exec_("delete from schedules where time = {}".format(old_value[4]))
                self.q.exec_("insert into schedules values('{}','{}','{}',{},{},{})".format(new_value[0], new_value[1],
                                                                                            new_value[2], new_value[3],
                                                                                            new_value[4], enable))
            logger.info("????????????{}->{}??????".format(old_value, new_value))
            self.load_schedules()
            return True
        return False

    def remove_selected_task(self):
        reply = QMessageBox.question(self, "??????", "????????????????????????", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return
        row = self.get_selected_row()
        if row == -1:
            return
        tm = self.tableWidget.item(row, 1).text()
        n = time_to_number(tm)
        self.q.exec_("delete  from schedules where time={}".format(n))
        logger.info('???????????????{}???????????????'.format(tm))
        self.load_schedules()

    def load_schedules(self):
        """
        ????????????????????????????????????????????????
        """
        self.sch.remove_all_jobs()
        records = self.get_all_schedules()
        for record in records:
            h, m, s = parse_to_time(record[4])
            dow = parse_to_weekdays(record[3])
            dow1 = tuple(map(lambda x: x - 1, list(dow)))
            dow = str(dow1)[1:-1].strip(',')
            enable = record[5]
            if dow and enable:
                self.sch.add_job(self.play_music_by_schedules, 'cron', hour=h, minute=m, second=s, day_of_week=dow,
                                 id=str(record[4]))

        # self.sch.print_jobs()
        self.show_schedules_on_table()
        self.set_next_run_time()

    def get_selected_row(self):
        """
        :return:??????????????????????????????
        """
        num = -1
        for i in self.tableWidget.selectionModel().selection().indexes():
            num = i.row()

        return num

    def play_music(self, filename):
        url = QUrl.fromLocalFile(filename)
        name = filename.split('/')[-1]
        icon = QIcon()
        icon.addPixmap(QPixmap(":/pictures/pause.png"))
        self.pause_play_button.setIcon(icon)
        self.musisc_name_label.setText('????????????: ' + name)
        content = QtMultimedia.QMediaContent(url)
        self.player.setMedia(content)
        self.player.play()
        self.set_next_run_time()

    def play_selected_music(self):
        row = self.get_selected_row()
        if row == -1:
            return
        filename = self.tableWidget.item(row, 4).text()
        self.play_music(filename)

    def play_music_by_schedules(self):
        current_time, wk = current_time_str()
        # ??????????????????????????????????????????????????????????????????????????????current_time??????????????????
        res = self.tableWidget.findItems(current_time, Qt.MatchExactly)
        try:
            row = res[0].row()
        except IndexError:
            # ????????????????????????????????????????????????????????????
            time_int = time_to_number(current_time)
            self.q.exec_("select max(time) from schedules where time < {} and weekdays&{}<>0 and  enable=1".format(
                time_int, 1 << wk))
            while self.q.next():
                tm = parse_to_time(self.q.value(0))
                tm = "{:02}:{:02}:{:02}".format(*tm)
            res = self.tableWidget.findItems(tm, Qt.MatchExactly)
            row = res[0].row()

        self.tableWidget.verticalScrollBar().setSliderPosition(row)
        self.tableWidget.selectRow(row)
        filename = self.tableWidget.item(row, 4).text()
        self.play_music(filename)
        # self.statusbar.showMessage((20*' ').join(self.status_info()))

    def get_all_schedules(self):
        res = []
        self.q.exec_("select * from schedules")
        while self.q.next():
            line = []
            for i in range(6):
                line.append(self.q.value(i))
            res.append(tuple(line))
        return res

    def stop_music(self):
        if self.player.state():
            self.player.stop()

    def set_next_run_time(self):
        status, tm, name = self.status_info()
        self.label_4.setText(status)
        self.next_time_label.setText(tm)
        self.next_name_label.setText(name)
        self.statusbar.showMessage("?????????" + status + + 10 * ' ' + "?????????????????????" + tm + 10 * ' ' + "???????????????" + name)

    def show_cell_in_tooltip(self, x, y):
        text = self.tableWidget.item(x, y).text()
        if y == 4:
            QToolTip.showText(QCursor.pos(), text)

    def change_weather_mode(self):
        n = 30
        if self.checkBox_2.checkState():
            if self.comboBox.currentIndex() == 0:
                n += 32
            else:
                n += 64
        if self.checkBox.checkState():
            # ####?????????????????????????????????
            self.q.exec_("update schedules set enable=0  where type='????????????'")
            self.q.exec_("update schedules set enable=1  where type='????????????'")
            # ??????????????????
            self.q.exec_("update config set weather_mode = 1")
            logger.info('????????????????????????')

        else:
            self.q.exec_("update schedules set enable=1  where type='????????????'")
            self.q.exec_("update schedules set enable=0 where type='????????????'")
            # ??????????????????
            self.q.exec_("update config set weather_mode = 0")
            logger.info('????????????????????????????????????')

        self.load_schedules()
        # ??????????????????

    def change_rest_mode(self):
        n = self.comboBox.currentIndex()
        rest_day = 1 << n + 5
        m = self.comboBox_2.currentIndex()
        work_day = 1 << m
        if self.checkBox_2.checkState():
            self.comboBox.setEnabled(False)
            self.comboBox_2.setEnabled(False)
            # ???????????????????????????????????????
            # ######  ??????????????????????????????????????????????????????????????? ??????????????????????????? 0
            self.q.exec_("update schedules set enable=0 where type='????????????'")
            # ????????????????????????weekday & m ????????? 0??????, ??????????????? weekday ^ res
            self.q.exec_("update schedules set weekdays=(weekdays+{}) where  (weekdays&{}=0 and weekdays&{}<>0) "
                         "and (type='????????????'or ""type='????????????' or type='????????????') ".format(rest_day, rest_day, work_day))
            self.q.exec_("update config set rest_mode=1,weekend={}, weekday={}".format(n, m))
            logger.info('?????????????????????')

        else:
            # ??????????????????????????? ??????????????????????????? 1??? ????????????????????????????????????????????????????????????
            # #######
            self.q.exec_("update schedules set enable=1 where type='????????????'")
            self.q.exec_("update schedules set weekdays=(weekdays-{}) where (weekdays&{}!=0 and  weekdays&{}<>0) and "
                         "(type='????????????'or ""type='????????????' or type='????????????')".format(rest_day, rest_day, work_day))
            self.comboBox.setEnabled(True)
            self.comboBox_2.setEnabled(True)
            self.q.exec_("update config set rest_mode=0,weekend=0, weekday=0")
            logger.info('????????????????????????????????????')

        self.load_schedules()

    def clear_database(self):
        reply = QMessageBox.question(self, "??????", "?????????????????????????????????????????????????????????",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return False
        self.q.exec_("delete from schedules")
        self.q.exec_("delete from config")
        self.checkBox.setChecked(False)
        self.checkBox_2.setChecked(False)
        self.comboBox.setCurrentIndex(0)
        self.comboBox_2.setCurrentIndex(0)
        logger.info("???????????????")
        self.load_schedules()
        return True

    def import_data(self):
        filename, ok = QFileDialog.getOpenFileName(self, "????????????", './', filter="Ldy Files(*.ldy)")
        if not ok:
            return
        with open(filename) as fin:
            k = fin.readlines()
            info = k[:-1]
            content = ''.join(info).strip('\n')
            md = md5()
            md.update(content.encode('gbk'))
            hs = k[-1]
            config = eval(info[1].strip('\n')[1:-1])
            if md.hexdigest() == hs:
                if self.clear_database():
                    self.q.exec_("insert into config values({},{},{},{},{})".format(config[0], config[1], config[2],
                                                                                    config[3], config[4]))
                    for data in info[2:]:
                        try:
                            data = eval(data.strip('\n'))
                            print(data)
                            self.q.exec_("insert into schedules values('{}','{}','{}',{},{},{})".format(data[0], data[1],
                                                                                                     data[2], data[3],
                                                                                                     data[4],data[5]))
                        except Exception as e:
                            print(e)
                            raise
                    QMessageBox.information(self, "??????", "???????????????????????????????????????????????????music????????????\n?????????????????????????????????music???????????????"
                                                        "???????????????????????????", QMessageBox.Yes)
                    self.config_load()
                    self.load_schedules()
                    QMessageBox.information(self, "??????", "???????????????", QMessageBox.Yes)
                    self.config_load()
                    logger.info('????????????{}??????'.format(filename))
                else:
                    logger.info('????????????????????????'.format(filename))
                    return
            else:
                logger.info('????????????????????????????????????{}??????'.format(filename))
                QMessageBox.information(self, "??????", "?????????????????????????????????????????????????????????", QMessageBox.Yes)

    def export_data(self):
        filename, ok = QFileDialog.getSaveFileName(self, "????????????", './', filter="Ldy Files(*.ldy)")
        if not ok:
            return
        if filename[-3:] != 'ldy':
            filename += '.ldy'
        with open(filename, 'w+') as fo:
            fo.write('*' * 20 + "????????????????????????" + '*' * 20 + '\n')
            lst = self.get_config()
            fo.write(str(lst) + '\n')
            lst = self.get_all_schedules()
            lst = list(map(str, lst))
            fo.write('\n'.join(lst))
            fo.seek(0)
            data = fo.readlines()
            data = ''.join(data).encode('gbk')
            m = md5()
            m.update(data)
            md = m.hexdigest()
            fo.seek(0, 2)
            fo.write('\n' + md)
            logger.info('???????????????{}'.format(filename))
            QMessageBox.information(self, "??????", "????????????", QMessageBox.Yes)

    def change_music_path(self):
        self.q.exec_("select file from schedules")
        res = []
        while self.q.next():
            info = self.q.value(0).replace('\\', '/')
            res.append(info.split('/')[-1])
        for i in res:
            self.q.exec_("update schedules set file = '{}' where file like '%{}'".
                         format(getcwd().replace("\\", "/") + '/music/' + i, i))
        self.load_schedules()
        QMessageBox.information(self, "??????", "??????????????????????????????????????????????????????????????????music?????????")

    def pause_resume_task(self):
        if self.pushButton_10.text() == '????????????':
            self.sch.pause()
            self.pushButton_10.setText("????????????")
        else:
            self.pushButton_10.setText("????????????")
            self.sch.resume()
            self.set_next_run_time()
        self.set_next_run_time()

    def choose_music(self, flag=0):
        filename, ok = QFileDialog.getOpenFileName(self, "????????????", './music', filter="MusicFile(*.mp3)")
        if ok:
            if not flag:
                self.lineEdit.setText(filename)
            else:
                self.lineEdit_2.setText(filename)

    def multi_change_music(self):
        src, dst = self.lineEdit.text(), self.lineEdit_2.text()
        if src and dst:
            self.q.exec_("update schedules set file = '{}' where file='{}'".format(dst, src))
        self.load_schedules()
        QMessageBox.information(self, "??????", "??????????????????????????????")

    def status_info(self):
        running = ["?????????", "??????"][self.sch.state - 1]
        tm = self.next_run_time()
        next_name = "???"
        if tm != "?????????":
            self.q.exec_("select name from schedules where time={}".format(time_to_number(tm.split()[1])))
            if self.q.next():
                next_name = self.q.value(0)
        if self.sch.state == 2:
            tm = "???"
            next_name = '???'

        return running, tm, next_name

    def next_run_time(self):
        jobs = self.sch.get_jobs()
        next_run_time = "?????????"
        if jobs:
            data = str(jobs[0])
            next_run_time = data.split("at:")[-1].strip().strip(" CST)")
        return next_run_time

    def changeEvent(self, event):
        if self.windowState() & Qt.WindowMinimized:
            self.hide()
            self.tray_icon.show()

    def closeEvent(self, event):
        self.sch.shutdown()
        self.db.close()

    def load_qss(self, qss_file):
        with open(qss_file, "r") as fin:
            qss = fin.read()
            self.setStyleSheet(qss)


class TrayIcon(QSystemTrayIcon):
    def __init__(self, ui, parent=None):
        super(TrayIcon, self).__init__(parent)
        self.ui = ui
        self.menu = QMenu()
        self.create_menu()
        icon = QIcon(QPixmap(":/pictures/icon.ico"))
        self.setIcon(icon)
        self.icon = self.MessageIcon()

    def create_menu(self):
        show_action2 = QAction("????????????", self, triggered=self.show_msg)
        show_action3 = QAction("????????????", self, triggered=self.ui.stop_music)
        show_quit = QAction("??????", self, triggered=self.quit)
        self.menu.addAction(show_action2)
        self.menu.addAction(show_action3)
        self.menu.addAction(show_quit)
        self.setContextMenu(self.menu)

    def show_msg(self):
        self.showMessage("??????????????????", "\n".join(self.ui.status_info()), self.icon)

    def show_window(self):
        self.ui.showNormal()
        self.ui.activateWindow()

    @staticmethod
    def quit(self):
        qApp.quit()

    def on_icon_clicked(self, reason):
        if reason == 2 or reason == 3:
            if self.ui.isHidden():
                self.ui.showNormal()
                self.ui.activateWindow()
                self.ui.show()
            else:
                self.ui.hide()


if __name__ == '__main__':
    app = QApplication(argv)
    win = MainWindow()
    mutex_name = "ldy6314"
    mutex = win32event.CreateMutex(None, 1, mutex_name)
    if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
        QMessageBox.warning(win, '??????', '?????????????????????')
        exit(0)
    cwd = os.getcwd()
    qss_dir = cwd + '/qss'
    qss_list = list(filter(lambda s: s.endswith('.qss'), os.listdir(qss_dir)))
    import random
    qss = random.choice(qss_list)
    print(qss)
    win.load_qss(qss_dir+'/' + qss)
    win.show()
    exit(app.exec_())

