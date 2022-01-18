from math import nan
import sys
import json
from turtle import right
import pandas as pd
import numpy as np

from qtpy import QtWidgets, QtGui, QtCore

def read_excel(file_name):
    df = pd.read_excel(file_name, sheet_name=None, header=None)
    return df

def reconnect(signal, newhandler=None, oldhandler=None):
        try:
            if oldhandler is not None:
                while True:
                    signal.disconnect(oldhandler)
            else:
                signal.disconnect()
        except TypeError as e:
            print(e)
        if newhandler is not None:
            signal.connect(newhandler)

class ExcelCheckerFrame(QtWidgets.QWidget):
    def __init__(self, application=None):
        super(ExcelCheckerFrame, self).__init__()
        self.app = application
        self.ui_designer()
        self.events_register()

    def ui_designer(self):
        with open(r'qtframe\qui_config.json', 'r', encoding='utf-8') as f:
            self.qui_configs = json.load(f)
        
        with open(r'qtframe\qss.json', 'r', encoding='utf-8') as f:
            self.qss_configs = json.load(f)

        if type(self.app) is QtWidgets.QApplication:
            w,h = self.app.desktop().screenGeometry().width(), self.app.desktop().screenGeometry().height()
            window_sr = self.qui_configs['window_scale_ratio']
            center_pos = 0.5 * (1-window_sr)
            self.setGeometry(w*center_pos, h*center_pos, w*window_sr, h*window_sr)
        else:
            self.setGeometry(300, 300, 300, 200)

        self.setFixedSize(self.width(), self.height())

        self.setWindowTitle(self.qui_configs['window_title'])
        self.setWindowIcon(QtGui.QIcon(self.qui_configs['window_icon']))

        fixed_left_width = self.qui_configs['fixed_left_width']
        left_bounds_width = self.qui_configs['left_bounds_width']

        self.excel_shown_place = QtWidgets.QTextEdit(self)
        self.excel_shown_place.setGeometry(left_bounds_width, 10, fixed_left_width, 200)
        self.excel_shown_place.setPlaceholderText(self.qui_configs['excel_shown_textbrowser_desc'])

        self.sheet_chosen = QtWidgets.QComboBox(self)
        self.sheet_chosen.setGeometry(left_bounds_width, 220, fixed_left_width, 30)
        self.sheet_chosen.setMaxVisibleItems(5)
        self.sheet_chosen.setStyleSheet(self.qss_configs['qss_chosen_combox'])
        self.sheet_chosen.setView(QtWidgets.QListView())

        self.search_text = QtWidgets.QLineEdit(self)
        self.search_text.setGeometry(left_bounds_width, 260, fixed_left_width, 30)
        self.search_text.setPlaceholderText(self.qui_configs['search_placeholder_desc'])

        self.split_text = QtWidgets.QLineEdit(self)
        self.split_text.setGeometry(left_bounds_width, 300, fixed_left_width, 30)
        self.split_text.setPlaceholderText(self.qui_configs['split_placeholder_desc'])

        self.submit_button = QtWidgets.QPushButton(self.qui_configs['submit_btn_desc'], self)
        self.cancel_button = QtWidgets.QPushButton(self.qui_configs['cancel_btn_desc'], self)
        
        self.horizontal_box = QtWidgets.QHBoxLayout()
        self.horizontal_box.addWidget(self.submit_button)
        self.horizontal_box.addWidget(self.cancel_button)
        self.horizontal_box.setGeometry(QtCore.QRect(left_bounds_width, 340, fixed_left_width, 30))

        space = self.qui_configs['center_space_width']
        right_bounds_width = self.qui_configs['right_bounds_width']
        
        surplus_w = self.width()-(fixed_left_width + space + right_bounds_width)

        self.data_browser = QtWidgets.QTextBrowser(self)
        self.data_browser.setPlaceholderText(self.qui_configs['data_browser_desc'])
        self.data_browser.setStyleSheet(self.qss_configs['qss_textbrowser'])
        # self.data_browser.setGeometry(fixed_left_width + space, 10, surplus_w, self.height()-20)

        self.check_result_place = QtWidgets.QTextBrowser(self)
        self.check_result_place.setPlaceholderText(self.qui_configs['result_checker_desc'])
        self.check_result_place.setStyleSheet(self.qss_configs['qss_textbrowser'])
        # self.check_result_place.setGeometry(fixed_left_width + space, 10, surplus_w, self.height()-20)

        self.right_vertical_box = QtWidgets.QVBoxLayout()
        self.right_vertical_box.addWidget(self.data_browser)
        self.right_vertical_box.addWidget(self.check_result_place)
        self.right_vertical_box.setSpacing(self.qui_configs['right_vertical_spacing'])
        self.right_vertical_box.setGeometry(QtCore.QRect(fixed_left_width + space, 10, surplus_w, self.height()-20))

    def events_register(self):
        self.excel_shown_place.dragEnterEvent = self.ondrag_enter
        self.excel_shown_place.dropEvent = self.ondrop_excel_shown_place_drop

        self.submit_button.clicked.connect(self.on_submit_button_clicked)
        self.cancel_button.clicked.connect(self.on_cancel_button_clicked)

    def on_submit_button_clicked(self):
        search_text = self.search_text.text()
        split_text = self.split_text.text()

        slist = []
        if search_text != '':
            slist.append([search_text])
            for c in split_text:
                length = len(slist)
                for i in range(0, length):
                    for j in slist[i]:
                        slist.append(j.split(c))
                slist = slist[length:]
        
        if not hasattr(self, 'data_type') or self.data_type is None or self.df is None:
            self.data_browser.setText('{}:{}\n{}:{}'.format(self.qui_configs['search_str_desc'], search_text, self.qui_configs['split_str_desc'], ' '.join(split_text[:])))
            self.check_result_place.setText('\n'.join(['{}{}: {}'.format(self.qui_configs['group_desc'], i, ' '.join(slist[i])) for i in range(0, len(slist))]))
        else:
            if self.data_type == 'excel':
                self.data_browser.setText(str('\n'.join([str(elem) for elem in self.df.values()])))
                """
                这段代码可以进行行优化，但是没有时间了，先使用列表解析方式搜索数据
                """

                text_list = []
                for key,value in self.df.items():
                    text_list.extend('\n'.join([key, ''.join(['[{}行,{}列]:匹配值{}\n'.format(each[0], elem, each[1]) for i in enumerate(slist) for j in i[1]  for elem in value for each in enumerate(value[elem]) if str(each[1]) == j])]))

                self.check_result_place.setText(''.join(text_list))
            elif self.data_type == 'json':
                self.data_browser.setText(json.dumps(self.df, indent=4, ensure_ascii=False))

    def on_cancel_button_clicked(self):
        self.search_text.setText('')
        self.split_text.setText('')

    def ondrag_enter(self, event):
        if event.mimeData().hasText():
            event.accept()
        else:
            event.ignore()

    def ondrop_excel_shown_place_drop(self, event):
        data = event.mimeData().text()        
        if data.endswith('.xlsm') or data.endswith('.xlsx') or data.endswith('.xls'):
            df = read_excel(data)
            text_list = ['Excel Struct:']
            text_list.extend([k for k,v in df.items() if not v.empty])

            self.set_excel_checker_data('excel', df)
            self.excel_shown_place.setText('\n'.join(text_list))
            self.data_browser.setText(str(self.df))
            reconnect(self.sheet_chosen.currentIndexChanged, self.on_excelsheet_chosen_index_changed)
            self.sheet_chosen.clear()
            self.sheet_chosen.addItem('FileData')
            self.sheet_chosen.addItems(df.keys())
        elif data.endswith('.json'):
            path = data
            if path.startswith('file:///'):
                path = path.replace('file:///', '')
            with open(path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)

            self.set_excel_checker_data('json', json_data)
            self.data_type = 'json'
            self.df = json_data
            self.data_browser.setText(json.dumps(json_data, indent=4, ensure_ascii=False))
            reconnect(self.sheet_chosen.currentIndexChanged, self.on_json_chosen_index_changed)
            self.sheet_chosen.clear()
            self.sheet_chosen.addItem('FileData')
            self.sheet_chosen.addItems(json_data.keys())
        else:
            event.ignore()
            self.set_excel_checker_data(None, None)
            reconnect(self.sheet_chosen.currentIndexChanged)

    def set_excel_checker_data(self, datatype, data):
        self.data_type = datatype
        self.df = data

    def on_json_chosen_index_changed(self, index):
        if index == -1:
            return
        if index == 0:
            self.data_browser.setText(json.dumps(self.df, indent=4, ensure_ascii=False))
        else:
            self.data_browser.setText('"{}": {}'.format(self.sheet_chosen.currentText(), json.dumps(self.df[self.sheet_chosen.currentText()], indent=4, ensure_ascii=False)))

    def on_excelsheet_chosen_index_changed(self, index):
        if index == -1:
            return
        if index == 0:
            self.data_browser.setText(str('\n'.join([str(elem) for elem in self.df.values()])))
        else:
            self.data_browser.setText(str(self.df[self.sheet_chosen.currentText()]))

class ExcelCheckerWindow():
    def __init__(self):
        self.app = QtWidgets.QApplication(['Excel Checker'])
        self.window = ExcelCheckerFrame(self.app)

    def run(self):
        self.window.show()
        sys.exit(self.app.exec_())
