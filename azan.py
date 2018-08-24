#!/usr/bin/python3
# -*- coding: utf-8 -*-

import datetime
import logging
import sys
import time

import pygame
from PyQt5.QtCore import QCoreApplication, QTimer, QRect
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QPushButton, QDesktopWidget, QLabel, QGridLayout, QApplication
from openpyxl import load_workbook

logger = logging.getLogger("main")
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler("info.log")
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

logger.debug("debug message")
logger.info("info message")
logger.error("error message")


# todo: два ядра cpu загружаются на 100%
# todo: добавить систему логирования
# todo: в случае не корректной даты - времени --- игнорировать и записать в лог
# todo: добавить модуль для выявления ошибок в excel фаиле


main_window_title = 'Азан'
main_window_icon = './img/web.png'
excel_file_path = './excel/excel.xlsx'
main_audio_file_path = './audio/1.mp3'
morning_audio_file_path = './audio/beep-06.mp3'
date_format = '%d.%m.%Y'
time_format = '%H:%M'
data_time_format = date_format + ' ' + time_format
prayer_time_names = ['Фаджр', 'Зухр', 'Аср', 'Магриб', 'Иша']
excel_column = ['B', 'D', 'E', 'F', 'G']
excel_date_column = 'A'
sheet_ranges_name = 'Лист1'
today_label = 'Сегодня: '
exit_button_label = 'Выход'
main_button_label = 'play main'
morning_button_label = 'play morning'
main_window_geometry = QRect(100, 100, 600, 300)
range_n = range(5)


class MainWindow(QWidget):
    def __init__(self):
        logger = logging.getLogger("main.MainWindow")
        logger.info("init")
        super().__init__()
        self.btn_exit = QPushButton(exit_button_label, self)
        self.btn_main_play_azan = QPushButton(main_button_label, self)
        self.btn_morning_play_azan = QPushButton(morning_button_label, self)
        self.labels = [QLabel(str(i)) for i in range_n]
        self.init_ui()

    def init_ui(self):
        grid = QGridLayout()
        self.setLayout(grid)
        now = datetime.datetime.now()
        today = today_label + now.strftime(date_format)
        lbl1 = QLabel(today, self)
        for i in range_n:
            grid.addWidget(self.labels[i], 1, i)
        self.btn_exit.clicked.connect(QCoreApplication.instance().quit)
        self.btn_main_play_azan.clicked.connect(main_play_azan)
        self.btn_morning_play_azan.clicked.connect(morning_play_azan)
        grid.addWidget(self.btn_main_play_azan, 0, 0)
        grid.addWidget(self.btn_morning_play_azan, 0, 1)
        grid.addWidget(lbl1, 2, 0)
        grid.addWidget(self.btn_exit, 2, 1)
        self.setGeometry(main_window_geometry)
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        self.setWindowTitle(main_window_title)
        self.setWindowIcon(QIcon(main_window_icon))
        self.show()

    def set_labels(self, prayer_times):
        for i in range_n:
            self.labels[i].setText(prayer_time_names[i] + ': ' + prayer_times['prayer_time'][i][11:])


def pygame_init():
    pygame.init()
    pygame.mixer.init()


def main_play_azan():
    pygame.mixer.music.stop()
    pygame.mixer.music.load(main_audio_file_path)
    pygame.mixer.music.set_volume(1.0)
    pygame.mixer.music.play()


def morning_play_azan():
    pygame.mixer.music.stop()
    pygame.mixer.music.load(morning_audio_file_path)
    pygame.mixer.music.set_volume(1.0)
    pygame.mixer.music.play()


def set_timer(timer, in_function):
    datetime_object = int(datetime.datetime.strptime(timer, data_time_format).timestamp())
    time_now = int(time.time())
    delta = (datetime_object - time_now) * 1000
    t = QTimer()
    t.timeout.connect(in_function)
    t.singleShot(delta, in_function)


def set_timers(timers):
    set_timer(timers['tomorrow'], daily_init)
    set_timer(timers['prayer_time'][0], morning_play_azan)
    for timer in timers['prayer_time'][1:]:
        set_timer(timer, main_play_azan)


def to_day_prayer_time(path_to_file):
    wb = load_workbook(filename=path_to_file)
    sheet_ranges = wb[sheet_ranges_name]
    today = datetime.datetime.now()
    tomorrow = (today + datetime.timedelta(days=1)).strftime(date_format)
    tomorrow = tomorrow + ' 00:01'
    tomorrow = datetime.datetime.strptime(tomorrow, data_time_format).strftime(data_time_format)
    today = today.strftime(date_format)
    i = 2
    c = 0
    prayer_time = []
    while True:
        if sheet_ranges[excel_date_column + str(i)].value is None:
            break
        if today == sheet_ranges[excel_date_column + str(i)].value.strftime(date_format):
            c = i
            prayer_time = [
                today + ' ' + sheet_ranges[excel_column[n] + str(c)].value
                for n in range_n
            ]
            break
        i += 1
    return {
        'status': 'ok' if c != 0 else 'error',
        'prayer_time': prayer_time,
        'tomorrow': tomorrow,
    }


# pygame_init()
# app = QApplication(sys.argv)
# ex = MainWindow()
#
#
# def daily_init():
#     prayer_times = to_day_prayer_time(excel_file_path)
#     set_timers(prayer_times)
#     ex.set_labels(prayer_times)
#
#
# daily_init()
# sys.exit(app.exec_())

from playsound import playsound
# playsound(main_audio_file_path)
playsound(morning_audio_file_path)