import os
import sys
import pandas as pd
from datetime import datetime, timedelta
from telegram import Bot
import asyncio
import pytz
import nest_asyncio
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QTextEdit,
    QVBoxLayout, QWidget, QFileDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QTabWidget, QLabel,
    QCheckBox, QLineEdit, QHBoxLayout, QComboBox, QDateTimeEdit
)
from PyQt6.QtCore import Qt, QTimer, QDateTime
import warnings

# 忽略 openpyxl 的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 應用 nest_asyncio 以允許嵌套事件循環
nest_asyncio.apply()

# 設定時區為 GMT+8 (Asia/Taipei)
timezone = pytz.timezone('Asia/Taipei')

# 配置文件路徑改為用戶目錄
CONFIG_DIR = os.path.join(os.path.expanduser("~"), "CFD_Reminder")
if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR, exist_ok=True)

CONFIG_FILE = os.path.join(CONFIG_DIR, "config.txt")
STOCK_CONFIG_FILE = os.path.join(CONFIG_DIR, "stock_config.txt")
TIME_CONFIG_FILE = os.path.join(CONFIG_DIR, "time_config.txt")
TIMER_CONFIG_FILE = os.path.join(CONFIG_DIR, "timer_config.txt")
CHAT_ID_CONFIG_FILE = os.path.join(CONFIG_DIR, "chat_id_config.txt")
EVENTS_CONFIG_FILE = os.path.join(CONFIG_DIR, "events_config.txt")
WEEKLY_CONFIG_FILE = os.path.join(CONFIG_DIR, "weekly_config.txt")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CFD 到期提醒工具")
        self.setGeometry(100, 100, 1200, 600)

        # 創建主佈局
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # 創建 Tab 控件
        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)

        # 結算總覽頁面
        self.summary_tab = QWidget()
        self.summary_layout = QVBoxLayout(self.summary_tab)
        self.tabs.addTab(self.summary_tab, "結算總覽")

        # 通知頁面
        self.notification_tab = QWidget()
        self.notification_layout = QVBoxLayout(self.notification_tab)
        self.tabs.addTab(self.notification_tab, "通知")

        # 導入 Excel 頁面 (CFD到期結算)
        self.excel_tab = QWidget()
        self.excel_layout = QVBoxLayout(self.excel_tab)
        self.tabs.addTab(self.excel_tab, "CFD到期結算")

        # 股票頁面
        self.stock_tab = QWidget()
        self.stock_layout = QVBoxLayout(self.stock_tab)
        self.tabs.addTab(self.stock_tab, "股票結算")

        # 自定義提醒頁面
        self.custom_tab = QWidget()
        self.custom_layout = QVBoxLayout(self.custom_tab)
        self.tabs.addTab(self.custom_tab, "其他事項提醒")

        # 結算總覽頁面控件
        self.cfd_table = QTableWidget()
        self.us_stock_table = QTableWidget()
        self.hk_stock_table = QTableWidget()

        self.tables_layout = QHBoxLayout()
        self.cfd_group_layout = QVBoxLayout()
        self.cfd_group_layout.addWidget(QLabel("CFD 結算通知"))
        self.cfd_table.setMinimumWidth(300)
        self.cfd_group_layout.addWidget(self.cfd_table)
        self.tables_layout.addLayout(self.cfd_group_layout)

        self.us_stock_group_layout = QVBoxLayout()
        self.us_stock_group_layout.addWidget(QLabel("美股結算通知"))
        self.us_stock_table.setMinimumWidth(300)
        self.us_stock_group_layout.addWidget(self.us_stock_table)
        self.tables_layout.addLayout(self.us_stock_group_layout)

        self.hk_stock_group_layout = QVBoxLayout()
        self.hk_stock_group_layout.addWidget(QLabel("港股結算通知"))
        self.hk_stock_table.setMinimumWidth(300)
        self.hk_stock_group_layout.addWidget(self.hk_stock_table)
        self.tables_layout.addLayout(self.hk_stock_group_layout)

        self.summary_layout.addLayout(self.tables_layout)

        self.timer_layout = QHBoxLayout()
        self.timer_label = QLabel("每日更新時間(hh:mm:ss):")
        self.timer_input = QDateTimeEdit()
        self.timer_input.setCalendarPopup(True)
        self.timer_input.setDisplayFormat("HH:mm:ss")
        self.timer_input.setDateTime(QDateTime.currentDateTime())
        self.timer_input.timeChanged.connect(self.save_timer_setting)
        self.timer_layout.addWidget(self.timer_label)
        self.timer_layout.addWidget(self.timer_input)
        self.timer_layout.addStretch()
        self.summary_layout.addLayout(self.timer_layout)

        self.chat_id_layout = QHBoxLayout()
        self.chat_id_label = QLabel("Telegram CHAT ID:")
        self.chat_id_input = QLineEdit()
        self.chat_id_input.setPlaceholderText("輸入 Telegram CHAT ID")
        self.chat_id_input.setText('1052625342')
        self.chat_id_layout.addWidget(self.chat_id_label)
        self.chat_id_layout.addWidget(self.chat_id_input)
        self.summary_layout.addLayout(self.chat_id_layout)

        self.tg_bot_start_button = QPushButton("TG Bot Start")
        self.tg_bot_start_button.clicked.connect(self.start_tg_bot)
        self.summary_layout.addWidget(self.tg_bot_start_button)

        self.tg_bot_stop_button = QPushButton("TG Bot Stop")
        self.tg_bot_stop_button.clicked.connect(self.stop_tg_bot)
        self.summary_layout.addWidget(self.tg_bot_stop_button)

        # 通知頁面控件
        self.check_button = QPushButton("檢查並發送通知")
        self.check_button.clicked.connect(self.run_check)
        self.notification_layout.addWidget(self.check_button)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.notification_layout.addWidget(self.log_text)

        # CFD到期結算頁面控件
        self.import_button = QPushButton("導入 Excel 文件")
        self.import_button.clicked.connect(self.import_excel)
        self.excel_layout.addWidget(self.import_button)

        self.cfd_path_layout = QHBoxLayout()
        self.cfd_path_input = QLineEdit()
        self.cfd_path_input.setPlaceholderText("輸入 CFD Excel 文件路徑")
        self.cfd_save_path_button = QPushButton("儲存")
        self.cfd_save_path_button.clicked.connect(self.save_and_apply_cfd_path)
        self.cfd_path_layout.addWidget(self.cfd_path_input)
        self.cfd_path_layout.addWidget(self.cfd_save_path_button)
        self.excel_layout.addLayout(self.cfd_path_layout)

        self.excel_table = QTableWidget()
        self.excel_table.itemChanged.connect(self.on_table_item_changed)
        self.excel_layout.addWidget(self.excel_table)

        # 股票頁面控件
        self.stock_import_button = QPushButton("導入股票 Excel 文件")
        self.stock_import_button.clicked.connect(self.import_stock_excel)
        self.stock_layout.addWidget(self.stock_import_button)

        self.stock_path_layout = QHBoxLayout()
        self.stock_path_input = QLineEdit()
        self.stock_path_input.setPlaceholderText("輸入股票 Excel 文件路徑")
        self.save_path_button = QPushButton("儲存")
        self.save_path_button.clicked.connect(self.save_and_apply_stock_path)
        self.stock_path_layout.addWidget(self.stock_path_input)
        self.stock_path_layout.addWidget(self.save_path_button)
        self.stock_layout.addLayout(self.stock_path_layout)

        self.summer_checkbox = QCheckBox("夏令")
        self.winter_checkbox = QCheckBox("冬令")
        self.apply_button = QPushButton("Apply")
        self.apply_button.clicked.connect(self.apply_time_setting)
        self.apply_button.setMaximumWidth(80)
        self.stock_layout.addWidget(self.summer_checkbox)
        self.stock_layout.addWidget(self.winter_checkbox)
        self.stock_layout.addWidget(self.apply_button)

        self.summer_checkbox.stateChanged.connect(self.on_summer_changed)
        self.winter_checkbox.stateChanged.connect(self.on_winter_changed)

        self.us_stockcfd_table = QTableWidget()
        self.hk_stockcfd_table = QTableWidget()
        self.stock_layout.addWidget(QLabel("STOCKCFDUSCS"))
        self.stock_layout.addWidget(self.us_stockcfd_table)
        self.stock_layout.addWidget(QLabel("STOCKCFDHKCS"))
        self.stock_layout.addWidget(self.hk_stockcfd_table)

        self.us_stockcfd_table.itemChanged.connect(self.on_stock_table_item_changed)
        self.hk_stockcfd_table.itemChanged.connect(self.on_stock_table_item_changed)

        # 自定義提醒頁面控件
        self.events_table = QTableWidget()
        self.events_table.setColumnCount(2)
        self.events_table.setHorizontalHeaderLabels(["事項", "時間"])
        self.events_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.custom_layout.addWidget(QLabel("自定義事件提醒"))
        self.custom_layout.addWidget(self.events_table)

        self.event_buttons_layout = QHBoxLayout()
        self.add_event_button = QPushButton("新增一行")
        self.add_event_button.clicked.connect(self.add_event_row)
        self.add_event_button.setMaximumWidth(100)
        self.remove_event_button = QPushButton("刪除一行")
        self.remove_event_button.clicked.connect(self.remove_event_row)
        self.remove_event_button.setMaximumWidth(100)
        self.event_buttons_layout.addWidget(self.add_event_button)
        self.event_buttons_layout.addWidget(self.remove_event_button)
        self.custom_layout.addLayout(self.event_buttons_layout)

        self.confirm_events_button = QPushButton("確認事件")
        self.confirm_events_button.clicked.connect(self.confirm_events)
        self.confirm_events_button.setMaximumWidth(100)
        self.custom_layout.addWidget(self.confirm_events_button)

        self.weekly_table = QTableWidget()
        self.weekly_table.setColumnCount(3)
        self.weekly_table.setHorizontalHeaderLabels(["事項", "星期", "時間"])
        self.weekly_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.custom_layout.addWidget(QLabel("每週提醒"))
        self.custom_layout.addWidget(self.weekly_table)

        self.weekly_buttons_layout = QHBoxLayout()
        self.add_weekly_button = QPushButton("新增一行")
        self.add_weekly_button.clicked.connect(self.add_weekly_row)
        self.add_weekly_button.setMaximumWidth(100)
        self.remove_weekly_button = QPushButton("刪除一行")
        self.remove_weekly_button.clicked.connect(self.remove_weekly_row)
        self.remove_weekly_button.setMaximumWidth(100)
        self.weekly_buttons_layout.addWidget(self.add_weekly_button)
        self.weekly_buttons_layout.addWidget(self.remove_weekly_button)
        self.custom_layout.addLayout(self.weekly_buttons_layout)

        self.confirm_weekly_button = QPushButton("確認每週提醒")
        self.confirm_weekly_button.clicked.connect(self.confirm_weekly)
        self.confirm_weekly_button.setMaximumWidth(100)
        self.custom_layout.addWidget(self.confirm_weekly_button)

        self.custom_tg_start_button = QPushButton("TG Bot Start (自定義)")
        self.custom_tg_start_button.clicked.connect(self.start_custom_tg_bot)
        self.custom_tg_start_button.setMaximumWidth(150)
        self.custom_layout.addWidget(self.custom_tg_start_button)

        self.custom_tg_stop_button = QPushButton("TG Bot Stop (自定義)")
        self.custom_tg_stop_button.clicked.connect(self.stop_custom_tg_bot)
        self.custom_tg_stop_button.setMaximumWidth(150)
        self.custom_layout.addWidget(self.custom_tg_stop_button)

        # Telegram 配置
        self.TOKEN = '8129106629:AAHIU4wptkdrJmFovsP-T5G0VY9G-CuRWvc'
        self.CHAT_ID = '1052625342'

        # 文件路徑和數據框架
        self.file_path = None
        self.df = None
        self.stock_file_path = None
        self.us_stockcfd_df = None
        self.hk_stockcfd_df = None

        # TG Bot 相關變量
        self.tg_timer = None
        self.notification_times = []
        self.last_update_time = None

        # 自定義提醒相關變量
        self.custom_tg_timer = None
        self.custom_notification_times = []
        self.events_confirmed = False
        self.weekly_confirmed = False

        # 確保配置文件存在
        self.ensure_config_files_exist()

        # 初始化表格
        self.load_events_table()
        self.load_weekly_table()

        # 自動加載保存的路徑和時間設定
        self.load_saved_path()
        self.load_saved_stock_path()
        self.load_time_setting()
        self.load_timer_setting()
        self.load_chat_id_setting()
        self.update_summary_tables()

        # 設置每日更新計時器
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_timer)
        self.timer.start(1000)

    def closeEvent(self, event):
        """
        當窗口關閉時自動保存每週提醒和事件表格的內容
        """
        self.confirm_events()  # 保存事件表格
        self.confirm_weekly()  # 保存每週提醒表格
        self.log("窗口關閉，已自動保存事件和每週提醒")
        event.accept()

    def log(self, message):
        self.log_text.append(f"{datetime.now(timezone).strftime('%Y-%m-%d %H:%M:%S')}: {message}")

    def ensure_config_files_exist(self):
        """檢查並確保配置文件存在，若不存在則創建空的"""
        for config_file in [EVENTS_CONFIG_FILE, WEEKLY_CONFIG_FILE]:
            if not os.path.exists(config_file):
                try:
                    with open(config_file, 'w', encoding='utf-8') as f:
                        f.write("")
                    self.log(f"未找到 {config_file}，已自動創建於 {CONFIG_DIR}")
                except Exception as e:
                    self.log(f"創建 {config_file} 失敗: {e}")

    def save_path(self):
        if self.file_path:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                f.write(self.file_path)
            self.log(f"已保存 CFD 文件路徑: {self.file_path}")

    def save_stock_path(self):
        if self.stock_file_path:
            with open(STOCK_CONFIG_FILE, 'w', encoding='utf-8') as f:
                f.write(self.stock_file_path)
            self.log(f"已保存股票文件路徑: {self.stock_file_path}")

    def save_and_apply_cfd_path(self):
        self.file_path = self.cfd_path_input.text().strip()
        if self.file_path and os.path.exists(self.file_path):
            self.save_path()
            self.load_cfd_excel_from_path()
        else:
            self.log(f"CFD 路徑無效或文件不存在: {self.file_path}")

    def save_and_apply_stock_path(self):
        self.stock_file_path = self.stock_path_input.text().strip()
        if self.stock_file_path and os.path.exists(self.stock_file_path):
            self.save_stock_path()
            self.load_stock_excel_from_path()
        else:
            self.log(f"股票路徑無效或文件不存在: {self.stock_file_path}")

    def save_time_setting(self):
        setting = "summer" if self.summer_checkbox.isChecked() else "winter" if self.winter_checkbox.isChecked() else "none"
        with open(TIME_CONFIG_FILE, 'w', encoding='utf-8') as f:
            f.write(setting)
        self.log(f"已保存時間設定: {setting}")

    def save_timer_setting(self):
        timer_value = self.timer_input.time().toString("HH:mm:ss")
        if timer_value:
            with open(TIMER_CONFIG_FILE, 'w', encoding='utf-8') as f:
                f.write(timer_value)
            self.log(f"已保存計時器設定: {timer_value}")

    def save_chat_id_setting(self):
        chat_id = self.chat_id_input.text().strip()
        if chat_id:
            with open(CHAT_ID_CONFIG_FILE, 'w', encoding='utf-8') as f:
                f.write(chat_id)
            self.CHAT_ID = chat_id
            self.log(f"已保存 CHAT ID: {chat_id}")

    def load_time_setting(self):
        if os.path.exists(TIME_CONFIG_FILE):
            with open(TIME_CONFIG_FILE, 'r', encoding='utf-8') as f:
                setting = f.read().strip()
                if setting == "summer":
                    self.summer_checkbox.setChecked(True)
                    self.winter_checkbox.setChecked(False)
                elif setting == "winter":
                    self.winter_checkbox.setChecked(True)
                    self.summer_checkbox.setChecked(False)
                else:
                    self.summer_checkbox.setChecked(False)
                    self.winter_checkbox.setChecked(False)
            self.log(f"已加載時間設定: {setting}")
        else:
            self.summer_checkbox.setChecked(False)
            self.winter_checkbox.setChecked(False)

    def load_timer_setting(self):
        if os.path.exists(TIMER_CONFIG_FILE):
            with open(TIMER_CONFIG_FILE, 'r', encoding='utf-8') as f:
                timer_value = f.read().strip()
                dt = QDateTime.fromString(timer_value, "HH:mm:ss")
                if dt.isValid():
                    self.timer_input.setTime(dt.time())
                self.log(f"已加載計時器設定: {timer_value}")

    def load_chat_id_setting(self):
        if os.path.exists(CHAT_ID_CONFIG_FILE):
            with open(CHAT_ID_CONFIG_FILE, 'r', encoding='utf-8') as f:
                chat_id = f.read().strip()
                self.CHAT_ID = chat_id
                self.chat_id_input.setText(chat_id)
            self.log(f"已加載 CHAT ID: {chat_id}")
        else:
            self.chat_id_input.setText(self.CHAT_ID)

    def on_summer_changed(self, state):
        if state == Qt.CheckState.Checked.value:
            self.winter_checkbox.setChecked(False)

    def on_winter_changed(self, state):
        if state == Qt.CheckState.Checked.value:
            self.summer_checkbox.setChecked(False)

    def apply_time_setting(self):
        self.save_time_setting()
        self.update_summary_tables()
        self.log("時間設定已應用")

    def load_saved_path(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                self.file_path = f.read().strip()
                self.cfd_path_input.setText(self.file_path)
            if os.path.exists(self.file_path):
                self.load_cfd_excel_from_path()
            else:
                self.log(f"保存的 CFD 路徑不存在: {self.file_path}")
                self.file_path = None

    def load_saved_stock_path(self):
        if os.path.exists(STOCK_CONFIG_FILE):
            with open(STOCK_CONFIG_FILE, 'r', encoding='utf-8') as f:
                self.stock_file_path = f.read().strip()
                self.stock_path_input.setText(self.stock_file_path)
            if os.path.exists(self.stock_file_path):
                self.load_stock_excel_from_path()
            else:
                self.log(f"保存的股票路徑不存在: {self.stock_file_path}")
                self.stock_file_path = None

    def load_cfd_excel_from_path(self):
        if self.file_path:
            try:
                xls = pd.ExcelFile(self.file_path)
                if 'Summary' not in xls.sheet_names:
                    raise ValueError("缺少 'Summary' 工作表")
                self.df = pd.read_excel(self.file_path, sheet_name='Summary')
                self.log(f"已從路徑加載 CFD 文件: {self.file_path}")
                self.load_table(self.excel_table, self.df)
                self.update_summary_tables()
            except Exception as e:
                self.log(f"從路徑加載 CFD 文件失敗: {e}")

    def load_stock_excel_from_path(self):
        if self.stock_file_path:
            try:
                xls = pd.ExcelFile(self.stock_file_path)
                if 'STOCKCFDUSCS' not in xls.sheet_names or 'STOCKCFDHKCS' not in xls.sheet_names:
                    raise ValueError("缺少必要的股票工作表")
                self.us_stockcfd_df = pd.read_excel(self.stock_file_path, sheet_name='STOCKCFDUSCS')
                self.hk_stockcfd_df = pd.read_excel(self.stock_file_path, sheet_name='STOCKCFDHKCS')
                self.log(f"已從路徑加載股票文件: {self.stock_file_path}")
                self.load_table(self.us_stockcfd_table, self.us_stockcfd_df)
                self.load_table(self.hk_stockcfd_table, self.hk_stockcfd_df)
                self.update_summary_tables()
            except Exception as e:
                self.log(f"從路徑加載股票文件失敗: {e}")

    def load_table(self, table, df):
        table.blockSignals(True)
        table.clear()
        table.setRowCount(df.shape[0])
        table.setColumnCount(df.shape[1])
        table.setHorizontalHeaderLabels(df.columns)

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                value = df.iloc[row, col]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                table.setItem(row, col, item)

        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        for col in range(table.columnCount()):
            if table.columnWidth(col) < 100:
                table.setColumnWidth(col, 100)
        table.blockSignals(False)

    def update_summary_tables(self):
        now = datetime.now(timezone)
        current_date = now.date()
        yesterday = current_date - timedelta(days=1)
        tomorrow = current_date + timedelta(days=1)

        if self.df is not None:
            current_month = now.month
            tomorrow_date = current_date + timedelta(days=1)
            day_after_tomorrow = current_date + timedelta(days=2)
            target_column = f'2025年{current_month}月'
            if target_column in self.df.columns:
                filtered_data = self.df[self.df[target_column].notna()]
                today_tomorrow_data = filtered_data[
                    (filtered_data[target_column].dt.date == current_date) |
                    (filtered_data[target_column].dt.date == tomorrow_date) |
                    ((filtered_data[target_column].dt.date == day_after_tomorrow) &
                     (filtered_data['Unnamed: 0'].str.contains('HK50|China300', case=False, regex=True)))
                    ]

                if not today_tomorrow_data.empty:
                    today_tomorrow_data_sorted = today_tomorrow_data.sort_values(by=target_column)
                    self.cfd_table.clear()
                    self.cfd_table.setColumnCount(1)
                    self.cfd_table.setHorizontalHeaderLabels(["CFD 結算通知"])
                    self.cfd_table.setRowCount(len(today_tomorrow_data_sorted))

                    for row, (_, data) in enumerate(today_tomorrow_data_sorted.iterrows()):
                        formatted_time = data[target_column].strftime('%Y-%m-%d %H:%M:%S')
                        product_name = data['Unnamed: 0']
                        is_special_product = any(x in product_name for x in ['HK50', 'China300'])
                        is_tomorrow_or_day_after = data[target_column].date() in [tomorrow_date, day_after_tomorrow]
                        if is_special_product and is_tomorrow_or_day_after:
                            text = f"{product_name}: {formatted_time} (提前提醒: DAY結算時間前後調整)"
                        else:
                            text = f"{product_name}: {formatted_time}"
                        self.cfd_table.setItem(row, 0, QTableWidgetItem(text))

                    self.cfd_table.resizeColumnsToContents()
                    self.cfd_table.resizeRowsToContents()
                    if self.cfd_table.columnWidth(0) < 300:
                        self.cfd_table.setColumnWidth(0, 300)

        if self.us_stockcfd_df is not None:
            close_column = '收市平倉交易日'
            us_stock_column = 'US STOCK'
            if close_column in self.us_stockcfd_df.columns:
                filtered_us_data = self.us_stockcfd_df[
                    self.us_stockcfd_df[close_column].notna() &
                    (pd.to_datetime(self.us_stockcfd_df[close_column]).dt.date.isin(
                        [yesterday, current_date, tomorrow]))
                    ]
                if not filtered_us_data.empty:
                    self.us_stock_table.clear()
                    self.us_stock_table.setColumnCount(1)
                    self.us_stock_table.setHorizontalHeaderLabels(["美股結算通知"])
                    self.us_stock_table.setRowCount(len(filtered_us_data))

                    for row, (_, data) in enumerate(filtered_us_data.iterrows()):
                        us_stock_name = f"{data[us_stock_column]}.US"
                        close_date = pd.Timestamp(data[close_column])
                        settlement_date = close_date + timedelta(days=1)
                        if self.summer_checkbox.isChecked():
                            settlement_time = settlement_date + timedelta(hours=8, minutes=30)
                        elif self.winter_checkbox.isChecked():
                            settlement_time = settlement_date + timedelta(hours=9, minutes=30)
                        else:
                            settlement_time = settlement_date
                        formatted_time = settlement_time.strftime('%Y-%m-%d %H:%M:%S')
                        text = f"{us_stock_name}: {formatted_time}"
                        self.us_stock_table.setItem(row, 0, QTableWidgetItem(text))

                    self.us_stock_table.resizeColumnsToContents()
                    self.us_stock_table.resizeRowsToContents()
                    if self.us_stock_table.columnWidth(0) < 300:
                        self.us_stock_table.setColumnWidth(0, 300)

        if self.hk_stockcfd_df is not None:
            close_column = '收市平倉交易日'
            hk_stock_column = 'HK STOCK'
            if close_column in self.hk_stockcfd_df.columns:
                filtered_hk_data = self.hk_stockcfd_df[
                    self.hk_stockcfd_df[close_column].notna() &
                    (pd.to_datetime(self.hk_stockcfd_df[close_column]).dt.date.isin(
                        [yesterday, current_date, tomorrow]))
                    ]
                if not filtered_hk_data.empty:
                    self.hk_stock_table.clear()
                    self.hk_stock_table.setColumnCount(1)
                    self.hk_stock_table.setHorizontalHeaderLabels(["香港股票結算通知"])
                    self.hk_stock_table.setRowCount(len(filtered_hk_data))

                    for row, (_, data) in enumerate(filtered_hk_data.iterrows()):
                        hk_stock_name = f"{data[hk_stock_column]}"
                        close_date = pd.Timestamp(data[close_column])
                        settlement_time = close_date + timedelta(hours=18, minutes=30)
                        formatted_time = settlement_time.strftime('%Y-%m-%d %H:%M:%S')
                        text = f"{hk_stock_name}: {formatted_time}"
                        self.hk_stock_table.setItem(row, 0, QTableWidgetItem(text))

                    self.hk_stock_table.resizeColumnsToContents()
                    self.hk_stock_table.resizeRowsToContents()
                    if self.hk_stock_table.columnWidth(0) < 300:
                        self.hk_stock_table.setColumnWidth(0, 300)

        if self.tg_timer is not None:
            self.notification_times = self.get_notification_times()

    def check_timer(self):
        now = datetime.now(timezone)
        current_time = now.time()
        timer_value = self.timer_input.time()

        if not timer_value.isValid():
            if not hasattr(self, '_timer_not_set_logged'):
                self.log("更新時間未設定，跳過檢查")
                self._timer_not_set_logged = True
            return

        if hasattr(self, '_timer_not_set_logged'):
            delattr(self, '_timer_not_set_logged')

        if (timer_value.hour() == current_time.hour and
                timer_value.minute() == current_time.minute and
                timer_value.second() == current_time.second):
            current_time_str = now.strftime('%H:%M:%S')
            if self.last_update_time != current_time_str:
                self.last_update_time = current_time_str
                self.log(f"計時器觸發: {timer_value.toString('HH:mm:ss')}，開始更新資料")
                self.load_cfd_excel_from_path()
                self.load_stock_excel_from_path()
                self.update_summary_tables()

    def start_tg_bot(self):
        self.save_chat_id_setting()
        if self.tg_timer is None:
            self.notification_times = self.get_notification_times()
            if not self.notification_times:
                self.log("警告: 無有效的通知時間，請檢查表格數據")
            self.tg_timer = QTimer(self)
            self.tg_timer.timeout.connect(self.check_tg_notifications)
            self.tg_timer.start(1000)
            self.log(f"TG Bot 已啟動，將根據表格時間提前 30 分鐘發送通知至 CHAT ID: {self.CHAT_ID}")
        else:
            self.log("TG Bot 已在運行")

    def stop_tg_bot(self):
        if self.tg_timer is not None:
            self.tg_timer.stop()
            self.tg_timer = None
            self.notification_times = []
            self.log("TG Bot 已停止")

    def get_notification_times(self):
        notification_times = []
        for row in range(self.cfd_table.rowCount()):
            item = self.cfd_table.item(row, 0)
            if item:
                text = item.text()
                try:
                    time_str = text.split(": ", 1)[1].split(" (")[0]
                    settlement_time = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                    settlement_time = timezone.localize(settlement_time)
                    notify_time = settlement_time - timedelta(minutes=30)
                    notification_times.append((notify_time, f"CFD 結算提醒: {text}"))
                except (IndexError, ValueError) as e:
                    self.log(f"無法解析 CFD 表格時間: {text}, 錯誤: {e}")

        for row in range(self.us_stock_table.rowCount()):
            item = self.us_stock_table.item(row, 0)
            if item:
                text = item.text()
                try:
                    time_str = text.split(": ", 1)[1]
                    settlement_time = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                    settlement_time = timezone.localize(settlement_time)
                    notify_time = settlement_time - timedelta(minutes=30)
                    notification_times.append((notify_time, f"美股結算提醒: {text}"))
                except (IndexError, ValueError) as e:
                    self.log(f"無法解析美股表格時間: {text}, 錯誤: {e}")

        for row in range(self.hk_stock_table.rowCount()):
            item = self.hk_stock_table.item(row, 0)
            if item:
                text = item.text()
                try:
                    time_str = text.split(": ", 1)[1]
                    settlement_time = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                    settlement_time = timezone.localize(settlement_time)
                    notify_time = settlement_time - timedelta(minutes=30)
                    notification_times.append((notify_time, f"港股結算提醒: {text}"))
                except (IndexError, ValueError) as e:
                    self.log(f"無法解析港股表格時間: {text}, 錯誤: {e}")

        return notification_times

    def check_tg_notifications(self):
        now = datetime.now(timezone)
        current_time_str = now.strftime('%Y-%m-%d %H:%M:%S')

        for notify_time, message in self.notification_times[:]:
            notify_time_str = notify_time.strftime('%Y-%m-%d %H:%M:%S')
            if current_time_str == notify_time_str:
                self.log(f"觸發 TG 通知: {message}")
                try:
                    loop = asyncio.get_event_loop()
                    loop.run_until_complete(self.send_tg_message(message))
                    self.notification_times.remove((notify_time, message))
                except Exception as e:
                    self.log(f"發送通知失敗: {e}")

    async def send_tg_message(self, message):
        chat_id = self.chat_id_input.text().strip() or self.CHAT_ID
        for attempt in range(3):
            try:
                bot = Bot(token=self.TOKEN)
                await bot.send_message(chat_id=chat_id, text=message)
                self.log(f"已發送 TG 訊息至 {chat_id}: {message}")
                break
            except Exception as e:
                self.log(f"發送 TG 訊息失敗 (嘗試 {attempt+1}/3): {str(e)}")
                if attempt < 2:
                    await asyncio.sleep(2)
                else:
                    self.log(f"最終發送失敗，使用 CHAT_ID: {chat_id}")

    def on_table_item_changed(self, item):
        if self.df is not None and self.file_path:
            row, col = item.row(), item.column()
            new_value = item.text()
            column_name = self.excel_table.horizontalHeaderItem(col).text()

            try:
                if "年" in column_name and "月" in column_name:
                    new_value = pd.to_datetime(new_value)
                self.df.at[row, column_name] = new_value
            except Exception as e:
                self.log(f"無法更新值: {e}")
                return

            try:
                self.df.to_excel(self.file_path, sheet_name='Summary', index=False)
                self.log(f"已更新並保存 CFD 文件: {self.file_path}")
                self.update_summary_tables()
            except Exception as e:
                self.log(f"保存 CFD 文件失敗: {e}")

    def on_stock_table_item_changed(self, item):
        if self.stock_file_path:
            table = item.tableWidget()
            df = self.us_stockcfd_df if table == self.us_stockcfd_table else self.hk_stockcfd_df
            sheet_name = 'STOCKCFDUSCS' if table == self.us_stockcfd_table else 'STOCKCFDHKCS'

            row, col = item.row(), item.column()
            new_value = item.text()
            column_name = table.horizontalHeaderItem(col).text()

            try:
                if column_name == '收市平倉交易日':
                    new_value = pd.to_datetime(new_value)
                df.at[row, column_name] = new_value
            except Exception as e:
                self.log(f"無法更新股票值: {e}")
                return

            try:
                temp_file = self.stock_file_path + '.tmp'
                with pd.ExcelWriter(temp_file) as writer:
                    self.us_stockcfd_df.to_excel(writer, sheet_name='STOCKCFDUSCS', index=False)
                    self.hk_stockcfd_df.to_excel(writer, sheet_name='STOCKCFDHKCS', index=False)
                os.replace(temp_file, self.stock_file_path)
                self.log(f"已更新並保存股票文件: {self.stock_file_path}")
                self.update_summary_tables()
            except Exception as e:
                self.log(f"保存股票文件失敗: {e}")

    def import_excel(self):
        self.file_path = QFileDialog.getOpenFileName(self, "請選擇 Excel 文件", "", "Excel Files (*.xlsx)")[0]
        if not self.file_path:
            self.log("未選擇 CFD 文件，退出操作。")
            return

        try:
            self.df = pd.read_excel(self.file_path, sheet_name='Summary')
            self.log(f"已導入 CFD 文件: {self.file_path}")
            self.load_table(self.excel_table, self.df)
            self.cfd_path_input.setText(self.file_path)
            self.save_path()
            self.update_summary_tables()
        except Exception as e:
            self.log(f"導入 CFD 文件失敗: {e}")

    def import_stock_excel(self):
        self.stock_file_path = QFileDialog.getOpenFileName(self, "請選擇股票 Excel 文件", "", "Excel Files (*.xlsx)")[0]
        if not self.stock_file_path:
            self.log("未選擇股票 文件，退出操作。")
            return

        try:
            self.us_stockcfd_df = pd.read_excel(self.stock_file_path, sheet_name='STOCKCFDUSCS')
            self.hk_stockcfd_df = pd.read_excel(self.stock_file_path, sheet_name='STOCKCFDHKCS')
            self.log(f"已導入股票文件: {self.stock_file_path}")
            self.load_table(self.us_stockcfd_table, self.us_stockcfd_df)
            self.load_table(self.hk_stockcfd_table, self.hk_stockcfd_df)
            self.stock_path_input.setText(self.stock_file_path)
            self.save_stock_path()
            self.update_summary_tables()
        except Exception as e:
            self.log(f"導入股票文件失敗: {e}")

    def run_check(self):
        self.log("開始檢查...")
        if not self.file_path:
            self.log("未設置 CFD 文件路徑，請先在 'CFD到期結算' 頁面導入或設置路徑")
            return

        try:
            self.df = pd.read_excel(self.file_path, sheet_name='Summary')
            now = datetime.now(timezone)
            current_month = now.month
            current_date = now.date()
            tomorrow_date = current_date + timedelta(days=1)
            day_after_tomorrow = current_date + timedelta(days=2)

            target_column = f'2025年{current_month}月'
            if target_column in self.df.columns:
                filtered_data = self.df[self.df[target_column].notna()]
                today_tomorrow_data = filtered_data[
                    (filtered_data[target_column].dt.date == current_date) |
                    (filtered_data[target_column].dt.date == tomorrow_date) |
                    ((filtered_data[target_column].dt.date == day_after_tomorrow) &
                     (filtered_data['Unnamed: 0'].str.contains('HK50|China300', case=False, regex=True)))
                    ]

                bot = Bot(token=self.TOKEN)

                async def send_message(chat_id, text):
                    await bot.send_message(chat_id=chat_id, text=text)

                if not today_tomorrow_data.empty:
                    today_tomorrow_data_sorted = today_tomorrow_data.sort_values(by=target_column)
                    message = f"{target_column} 到期產品提醒:\n"
                    for index, row in today_tomorrow_data_sorted.iterrows():
                        formatted_time = row[target_column].strftime('%Y-%m-%d %H:%M:%S')
                        product_name = row['Unnamed: 0']
                        is_special_product = any(x in product_name for x in ['HK50', 'China300'])
                        is_tomorrow_or_day_after = row[target_column].date() in [tomorrow_date, day_after_tomorrow]
                        if is_special_product and is_tomorrow_or_day_after:
                            message += f"{product_name}: {formatted_time} (提前提醒: DAY結算時間前後調整)\n"
                        else:
                            message += f"{product_name}: {formatted_time}\n"
                    self.log(message)
                    chat_id = self.chat_id_input.text().strip() or self.CHAT_ID
                    asyncio.run(send_message(chat_id, message))
                else:
                    message = "今天明天沒有產品結算"
                    self.log(message)
                    chat_id = self.chat_id_input.text().strip() or self.CHAT_ID
                    asyncio.run(send_message(chat_id, message))

                self.update_summary_tables()

            if now.weekday() == 0:
                asyncio.run(self.send_weekly_reminder())
        except Exception as e:
            self.log(f"發生錯誤: {str(e)}")

    async def send_weekly_reminder(self):
        message = 'BBG NON FARM FORCAST'
        chat_id = self.chat_id_input.text().strip() or self.CHAT_ID
        bot = Bot(token=self.TOKEN)
        await bot.send_message(chat_id=chat_id, text=message)
        self.log(f"已發送週一提醒至 {chat_id}: {message}")

    def load_events_table(self):
        """
        加載事件表格，從文件中讀取保存的事件數據
        若文件不存在或格式錯誤，初始化為單行空白數據
        """
        self.ensure_config_files_exist()
        try:
            with open(EVENTS_CONFIG_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                self.events_table.setRowCount(0)  # 清空表格
                if lines:
                    for row, line in enumerate(lines):
                        try:
                            event, time_str = line.strip().split(',', 1)  # 使用逗號分隔
                            self.events_table.insertRow(row)
                            self.events_table.setItem(row, 0, QTableWidgetItem(event))
                            dt_edit = QDateTimeEdit()
                            dt_edit.setCalendarPopup(True)
                            dt_edit.setDisplayFormat("yyyy:MM:dd HH:mm:ss")
                            dt = QDateTime.fromString(time_str, "yyyy:MM:dd HH:mm:ss")
                            if not dt.isValid():
                                raise ValueError(f"無效的時間格式: {time_str}")
                            dt_edit.setDateTime(dt)
                            self.events_table.setCellWidget(row, 1, dt_edit)
                        except ValueError as e:
                            self.log(f"解析 {EVENTS_CONFIG_FILE} 第 {row+1} 行失敗: {e}")
                            continue
                    self.events_confirmed = True
                else:
                    # 文件存在但為空，初始化單行
                    self.events_table.setRowCount(1)
                    self.events_table.setItem(0, 0, QTableWidgetItem(""))
                    dt_edit = QDateTimeEdit()
                    dt_edit.setCalendarPopup(True)
                    dt_edit.setDisplayFormat("yyyy:MM:dd HH:mm:ss")
                    dt_edit.setDateTime(QDateTime.currentDateTime())
                    self.events_table.setCellWidget(0, 1, dt_edit)
                    self.events_confirmed = False
        except Exception as e:
            self.log(f"加載 {EVENTS_CONFIG_FILE} 失敗: {e}")
            # 初始化單行，預設事件名稱為「財務報告」
            self.events_table.setRowCount(1)
            self.events_table.setItem(0, 0, QTableWidgetItem("財務報告"))
            dt_edit = QDateTimeEdit()
            dt_edit.setCalendarPopup(True)
            dt_edit.setDisplayFormat("yyyy:MM:dd HH:mm:ss")
            dt_edit.setDateTime(QDateTime.currentDateTime())
            self.events_table.setCellWidget(0, 1, dt_edit)
            self.events_confirmed = False

    def load_weekly_table(self):
        """
        加載每週提醒表格，從文件中讀取保存的每週提醒數據
        若文件不存在或格式錯誤，初始化為單行空白數據
        """
        self.ensure_config_files_exist()
        try:
            with open(WEEKLY_CONFIG_FILE, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                self.weekly_table.setRowCount(0)  # 清空表格
                if lines:
                    for row, line in enumerate(lines):
                        try:
                            parts = line.strip().split(',', 2)  # 使用逗號分隔
                            if len(parts) == 2:  # 舊格式
                                day, time_str = parts
                                event = ""
                            elif len(parts) == 3:  # 新格式
                                event, day, time_str = parts
                            else:
                                raise ValueError("格式錯誤，應包含 2 或 3 個字段")
                            self.weekly_table.insertRow(row)
                            self.weekly_table.setItem(row, 0, QTableWidgetItem(event))
                            combo = QComboBox()
                            combo.addItems(["每日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"])
                            if day in ["每日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]:
                                combo.setCurrentText(day)
                            else:
                                raise ValueError(f"無效的星期: {day}")
                            self.weekly_table.setCellWidget(row, 1, combo)
                            dt_edit = QDateTimeEdit()
                            dt_edit.setCalendarPopup(True)
                            dt_edit.setDisplayFormat("HH:mm:ss")
                            dt = QDateTime.fromString(time_str, "HH:mm:ss")
                            if not dt.isValid():
                                raise ValueError(f"無效的時間格式: {time_str}")
                            dt_edit.setDateTime(dt)
                            self.weekly_table.setCellWidget(row, 2, dt_edit)
                        except ValueError as e:
                            self.log(f"解析 {WEEKLY_CONFIG_FILE} 第 {row+1} 行失敗: {e}")
                            continue
                    self.weekly_confirmed = True
                else:
                    # 文件存在但為空，初始化單行
                    self.weekly_table.setRowCount(1)
                    self.weekly_table.setItem(0, 0, QTableWidgetItem(""))
                    combo = QComboBox()
                    combo.addItems(["每日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"])
                    self.weekly_table.setCellWidget(0, 1, combo)
                    dt_edit = QDateTimeEdit()
                    dt_edit.setCalendarPopup(True)
                    dt_edit.setDisplayFormat("HH:mm:ss")
                    dt_edit.setDateTime(QDateTime.currentDateTime())
                    self.weekly_table.setCellWidget(0, 2, dt_edit)
                    self.weekly_confirmed = False
        except Exception as e:
            self.log(f"加載 {WEEKLY_CONFIG_FILE} 失敗: {e}")
            # 初始化單行
            self.weekly_table.setRowCount(1)
            self.weekly_table.setItem(0, 0, QTableWidgetItem(""))
            combo = QComboBox()
            combo.addItems(["每日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"])
            self.weekly_table.setCellWidget(0, 1, combo)
            dt_edit = QDateTimeEdit()
            dt_edit.setCalendarPopup(True)
            dt_edit.setDisplayFormat("HH:mm:ss")
            dt_edit.setDateTime(QDateTime.currentDateTime())
            self.weekly_table.setCellWidget(0, 2, dt_edit)
            self.weekly_confirmed = False

    def add_event_row(self):
        row_count = self.events_table.rowCount()
        self.events_table.insertRow(row_count)
        self.events_table.setItem(row_count, 0, QTableWidgetItem(""))
        dt_edit = QDateTimeEdit()
        dt_edit.setCalendarPopup(True)
        dt_edit.setDisplayFormat("yyyy:MM:dd HH:mm:ss")
        dt_edit.setDateTime(QDateTime.currentDateTime())
        self.events_table.setCellWidget(row_count, 1, dt_edit)
        self.events_confirmed = False

    def remove_event_row(self):
        row_count = self.events_table.rowCount()
        if row_count > 1:
            self.events_table.removeRow(row_count - 1)
            self.events_confirmed = False

    def add_weekly_row(self):
        row_count = self.weekly_table.rowCount()
        self.weekly_table.insertRow(row_count)
        self.weekly_table.setItem(row_count, 0, QTableWidgetItem(""))
        combo = QComboBox()
        combo.addItems(["每日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"])
        self.weekly_table.setCellWidget(row_count, 1, combo)
        dt_edit = QDateTimeEdit()
        dt_edit.setCalendarPopup(True)
        dt_edit.setDisplayFormat("HH:mm:ss")
        dt_edit.setDateTime(QDateTime.currentDateTime())
        self.weekly_table.setCellWidget(row_count, 2, dt_edit)
        self.weekly_confirmed = False

    def remove_weekly_row(self):
        row_count = self.weekly_table.rowCount()
        if row_count > 1:
            self.weekly_table.removeRow(row_count - 1)
            self.weekly_confirmed = False

    def confirm_events(self):
        """
        確認事件並保存到文件，使用逗號分隔
        """
        events = []
        for row in range(self.events_table.rowCount()):
            event_item = self.events_table.item(row, 0)
            time_widget = self.events_table.cellWidget(row, 1)
            if (event_item and time_widget and
                    event_item.text().strip() and isinstance(time_widget, QDateTimeEdit)):
                time_str = time_widget.dateTime().toString("yyyy:MM:dd HH:mm:ss")
                events.append(f"{event_item.text().strip()},{time_str}")
        try:
            with open(EVENTS_CONFIG_FILE, 'w', encoding='utf-8') as f:
                if events:
                    f.write('\n'.join(events))
                else:
                    f.write("")
            self.events_confirmed = True
            self.log(f"事件已確認並保存到 {EVENTS_CONFIG_FILE}")
            if self.custom_tg_timer is not None:
                self.custom_notification_times = self.get_custom_notification_times()
        except Exception as e:
            self.log(f"保存 {EVENTS_CONFIG_FILE} 失敗: {e}")

    def confirm_weekly(self):
        """
        確認每週提醒並保存到文件，使用逗號分隔
        """
        weekly = []
        for row in range(self.weekly_table.rowCount()):
            event_item = self.weekly_table.item(row, 0)
            day_widget = self.weekly_table.cellWidget(row, 1)
            time_widget = self.weekly_table.cellWidget(row, 2)
            if (event_item and day_widget and time_widget and
                    event_item.text().strip() and isinstance(day_widget, QComboBox) and
                    isinstance(time_widget, QDateTimeEdit)):
                time_str = time_widget.dateTime().toString("HH:mm:ss")
                weekly.append(f"{event_item.text().strip()},{day_widget.currentText()},{time_str}")
        try:
            with open(WEEKLY_CONFIG_FILE, 'w', encoding='utf-8') as f:
                if weekly:
                    f.write('\n'.join(weekly))
                else:
                    f.write("")
            self.weekly_confirmed = True
            self.log(f"每週提醒已確認並保存到 {WEEKLY_CONFIG_FILE}")
            if self.custom_tg_timer is not None:
                self.custom_notification_times = self.get_custom_notification_times()
        except Exception as e:
            self.log(f"保存 {WEEKLY_CONFIG_FILE} 失敗: {e}")

    def start_custom_tg_bot(self):
        if self.custom_tg_timer is not None:
            self.stop_custom_tg_bot()

        self.custom_notification_times = self.get_custom_notification_times()
        if not self.custom_notification_times:
            self.log("警告: 無有效的自定義通知時間")
            return

        self.custom_tg_timer = QTimer(self)
        self.custom_tg_timer.timeout.connect(self.check_custom_tg_notifications)
        self.custom_tg_timer.start(1000)
        self.log(f"自定義 TG Bot 已啟動，將根據自定義時間發送通知至 CHAT ID: {self.CHAT_ID}")

    def stop_custom_tg_bot(self):
        if self.custom_tg_timer is not None:
            self.custom_tg_timer.stop()
            self.custom_tg_timer = None
            self.custom_notification_times = []
            self.log("自定義 TG Bot 已停止")

    def get_custom_notification_times(self):
        notification_times = []
        now = datetime.now(timezone)

        # 處理事件表格
        for row in range(self.events_table.rowCount()):
            event_item = self.events_table.item(row, 0)
            time_widget = self.events_table.cellWidget(row, 1)
            if (event_item and time_widget and
                    event_item.text().strip() and isinstance(time_widget, QDateTimeEdit)):
                notify_time = time_widget.dateTime().toPyDateTime()
                notify_time = timezone.localize(notify_time)
                if notify_time >= now:
                    time_str = time_widget.dateTime().toString("yyyy:MM:dd HH:mm:ss")
                    notification_times.append((notify_time, f"事件提醒: {event_item.text()} at {time_str}", False))

        # 處理每週提醒表格
        for row in range(self.weekly_table.rowCount()):
            event_item = self.weekly_table.item(row, 0)
            day_widget = self.weekly_table.cellWidget(row, 1)
            time_widget = self.weekly_table.cellWidget(row, 2)
            if (event_item and day_widget and time_widget and
                    event_item.text().strip() and isinstance(day_widget, QComboBox) and
                    isinstance(time_widget, QDateTimeEdit)):
                time_str = time_widget.dateTime().toString("HH:mm:ss")
                time_obj = datetime.strptime(time_str, '%H:%M:%S').time()
                day_text = day_widget.currentText()

                if day_text == "每日":
                    notify_date = now.date()
                    notify_time = timezone.localize(datetime.combine(notify_date, time_obj))
                    if notify_time <= now:
                        notify_time += timedelta(days=1)
                    notification_times.append((
                        notify_time,
                        f"每日提醒: {event_item.text()} at {time_str}",
                        True
                    ))
                else:
                    day_map = {
                        "星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3,
                        "星期五": 4, "星期六": 5, "星期日": 6
                    }
                    target_day = day_map[day_text]
                    current_weekday = now.weekday()
                    days_ahead = (target_day - current_weekday) % 7
                    if days_ahead == 0 and now.time() > time_obj:
                        days_ahead = 7
                    notify_date = now.date() + timedelta(days=days_ahead)
                    notify_time = timezone.localize(datetime.combine(notify_date, time_obj))
                    notification_times.append((
                        notify_time,
                        f"每週提醒: {event_item.text()} on {day_text} at {time_str}",
                        True
                    ))

        return notification_times

    def check_custom_tg_notifications(self):
        now = datetime.now(timezone)
        current_time_str = now.strftime('%Y-%m-%d %H:%M:%S')
        for notify_time, message, is_weekly in self.custom_notification_times[:]:
            notify_time_str = notify_time.strftime('%Y-%m-%d %H:%M:%S')
            if current_time_str == notify_time_str:
                self.log(f"觸發自定義 TG 通知: {message}")
                try:
                    loop = asyncio.get_event_loop()
                    loop.run_until_complete(self.send_tg_message(message))
                    self.custom_notification_times.remove((notify_time, message, is_weekly))
                    if is_weekly:
                        if "每日提醒" in message:
                            next_notify_time = notify_time + timedelta(days=1)
                        else:
                            next_notify_time = notify_time + timedelta(days=7)
                        self.custom_notification_times.append((next_notify_time, message, True))
                        self.log(f"已為{'每日' if '每日提醒' in message else '每週'}提醒重新安排下次時間: {next_notify_time.strftime('%Y-%m-%d %H:%M:%S')}")
                except Exception as e:
                    self.log(f"發送自定義通知失敗: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

    aaa