# -*- coding: utf-8 -*-
from __future__ import print_function
import sys

import datetime
import os.path
import json
import time
from threading import Thread, Event

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor, QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QSizePolicy, QLabel, QHBoxLayout, \
    QSystemTrayIcon, QAction, QMenu, QFormLayout, QLineEdit, QPushButton

from win32com.client import Dispatch

from client_credentials import client_credentials


# Locale for different languages
# import locale
# locale.setlocale(locale.LC_TIME, 'pl_pl')
class Settings:
    def __init__(self):
        self.filename = 'settings.json'
        self.SETTINGS = {
            'x': 0,
            'y': 0,
            'width': 400,
            'height': 1000,
        }

    def save(self):
        print('saving')
        with open(self.filename, 'w') as f:
            json.dump(self.SETTINGS, f)

    def load(self):
        print('try loading', os.listdir('.'), os.getcwd())
        try:
            with open(self.filename, 'r') as settings_file:
                self.SETTINGS = json.load(settings_file)
        except (FileNotFoundError, json.JSONDecodeError):
            self.save()
        finally:
            print('finally loading', os.listdir('.'), os.getcwd())
            with open(self.filename, 'r') as settings_file:
                self.SETTINGS = json.load(settings_file)

    @staticmethod
    def get_target_path_exe():
        current_path = os.getcwd()
        return os.path.join(current_path, r"CalendarWidget.exe")

    @staticmethod
    def get_new_shortcut_path():
        home_directory = os.path.expanduser("~")
        startup_folder = os.path.join(
            home_directory, 'AppData', 'Roaming', 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        shortcut_name = 'CalendarWidget.lnk'
        return os.path.join(startup_folder, shortcut_name)

    def enable_startup(self):
        target = self.get_target_path_exe()
        path = self.get_new_shortcut_path()
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.save()

    def disable_startup(self):
        print('exist?', os.path.exists(self.get_new_shortcut_path()))
        if os.path.exists(self.get_new_shortcut_path()):
            os.remove(self.get_new_shortcut_path())

    @property
    def is_startup(self):
        shortcut_dir = os.path.dirname(self.get_new_shortcut_path())
        files = os.listdir(shortcut_dir)
        return 'CalendarWidget.lnk' in files


class CalendarManagerGoogle:
    def __init__(self):
        self.scopes = ['https://www.googleapis.com/auth/calendar.readonly']
        self.event_count = 10
        self.token_path = 'token.json'

    def get_creds(self):
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                # flow = InstalledAppFlow.from_client_secrets_file(
                #     'credentials/credentials.json', self.scopes)
                flow = InstalledAppFlow.from_client_config(client_credentials, self.scopes)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return creds

    def print_events(self, events):
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])

    def get_events(self):
        creds = self.get_creds()

        try:
            service = build('calendar', 'v3', credentials=creds)
            now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
            print('Getting the upcoming 10 events')
            events_result = service.events().list(calendarId='primary', timeMin=now,
                                                  maxResults=self.event_count, singleEvents=True,
                                                  orderBy='startTime').execute()
            events = events_result.get('items', [])
            if not events:
                print('No upcoming events found.')
                return

            self.print_events(events)
            return events

        except HttpError as error:
            print(error)
            return None


class EventBlock(QWidget):
    def __init__(self):
        super().__init__()
        self.stylesheet = None
        self.font = self.init_font()
        self.stylesheet_default = self.get_stylesheet_default()
        self.summary_label = None
        self.date_label = None
        self.initUI()

    def choose_color(self, start_date, date_now):
        if start_date.date() == date_now:
            return QColor(120, 128, 128, 200)
        elif start_date.date() <= date_now + datetime.timedelta(days=1):
            return QColor(255, 128, 0, 200)
        elif start_date.date() <= date_now + datetime.timedelta(days=7):
            return QColor(0, 128, 255, 200)
        else:
            return QColor(0, 0, 0, 200)

    def create_date_label(self):
        self.date_label = QLabel()
        self.date_label.setFont(self.font)
        self.date_label.setStyleSheet(self.stylesheet)
        self.date_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.date_label.setAlignment(Qt.AlignCenter)
        return self.date_label

    def update_date_label(self, start_date):
        start_date_readable = start_date.strftime("%d-%m-%Y \n %H:%M \n %A")
        self.date_label.setText(start_date_readable)
        time.sleep(0.1)
        self.date_label.setStyleSheet(self.stylesheet)
        print(self.date_label.styleSheet())
        return self.date_label

    def create_summary_label(self):
        self.summary_label = QLabel()
        self.summary_label.setFont(self.font)
        self.summary_label.setAlignment(Qt.AlignCenter)
        return self.summary_label

    def update_summary_label(self, summary):
        self.summary_label.setText(summary)
        self.summary_label.setFont(self.font)
        time.sleep(0.1)
        self.summary_label.setStyleSheet(self.stylesheet)
        self.summary_label.setAlignment(Qt.AlignCenter)
        return self.summary_label

    def init_font(self):
        self.font = QFont("Montserrat")
        self.font.setPointSize(11)
        self.font.setBold(True)
        return self.font

    def get_stylesheet(self, color):
        print(color.alpha())
        return f"background-color: rgba({color.red()}, {color.green()}, {color.blue()}, {color.alpha()}); color: white;"

    def get_stylesheet_default(self):
        return f"background-color: rgba(0, 0, 0, 0); color: transparent;"

    def initUI(self):
        layout = QHBoxLayout()

        self.date_label = self.create_date_label()
        layout.addWidget(self.date_label)

        self.summary_label = self.create_summary_label()
        layout.addWidget(self.summary_label)

        self.setLayout(layout)

    def updateUI(self, start_date_str, summary):
        start_date = datetime.datetime.fromisoformat(start_date_str)
        date_now = datetime.datetime.now().date()
        color = self.choose_color(start_date, date_now)
        self.stylesheet = self.get_stylesheet(color)
        self.date_label = self.update_date_label(start_date)
        self.summary_label = self.update_summary_label(summary)

    def updateEmpty(self):
        self.date_label.setText("")
        self.summary_label.setText("")
        self.stylesheet = self.get_stylesheet_default()
        self.summary_label.setStyleSheet(self.stylesheet)
        self.date_label.setStyleSheet(self.stylesheet)


class EventViewer(QMainWindow):
    def __init__(self, app, calendar_manager, settings):
        super().__init__()
        self.app = app
        self.settings = settings
        self.calendar_manager = calendar_manager
        self.layout = None
        self.central_widget = None
        self.initUI()
        self.blocks = self.load_events()

    def initUI(self):
        self.setAttribute(Qt.WA_TranslucentBackground)  # Make the background transparent
        self.setGeometry(self.settings.SETTINGS['x'], self.settings.SETTINGS['y'], self.settings.SETTINGS['width'],
                         self.settings.SETTINGS['height'])
        self.setWindowFlags(Qt.WindowStaysOnBottomHint | Qt.FramelessWindowHint | Qt.Tool)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

    def load_events(self):
        events = self.calendar_manager.get_events()
        blocks = []
        if events:
            for i in range(10):
                event_block = EventBlock()
                self.layout.addWidget(event_block)
                blocks.append(event_block)

        return blocks

    def update_events(self):
        print('update events')
        events = self.calendar_manager.get_events()
        time.sleep(2)

        for i, _ in enumerate(self.blocks):
            block = self.blocks[i]
            if i < len(events):
                event = events[i]
                summary = event['summary']
                start_date = event['start'].get('dateTime', event['start'].get('date'))
                block.updateUI(start_date, summary)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            new_position = event.globalPos() - self.drag_start_position
            self.move(new_position)
            self.settings.SETTINGS['x'] = new_position.x()
            self.settings.SETTINGS['y'] = new_position.y()
            self.settings.save()

    def update_geometry(self):
        self.setGeometry(self.settings.SETTINGS['x'], self.settings.SETTINGS['y'], self.settings.SETTINGS['width'],
                         self.settings.SETTINGS['height'])


class Tray(QSystemTrayIcon):
    def __init__(self, app, event_viewer, updating_thread, settings):
        super().__init__()
        self.settings = settings
        self.event_viewer = event_viewer
        self.updating_thread = updating_thread
        self.app = app
        self.startup_button_name = "Not initialized"
        icon = QIcon("img/cal.png")
        self.setIcon(icon)
        self.setVisible(True)

        self.menu = QMenu()

        self.option2 = QAction("Settings")
        self.option2.triggered.connect(self.openSettingsWindow)
        self.menu.addAction(self.option2)

        self.quit = QAction("Quit")
        self.quit.triggered.connect(self.closeApp)
        self.menu.addAction(self.quit)
        self.setContextMenu(self.menu)

    def closeApp(self):
        self.updating_thread.close()
        self.app.quit()

    def toogle_startup(self):
        if not self.settings.is_startup:
            self.settings.enable_startup()
            print('enable startup')
        else:
            self.settings.disable_startup()
            print('disable startup')
        self.set_button_toggle_startup_name()

    def set_button_toggle_startup_name(self):
        if self.settings.is_startup:
            self.startup_button_name = "Disable startup"
        else:
            self.startup_button_name = "Enable Startup"
        self.toogle_startup_button.setText(self.startup_button_name)

    def openSettingsWindow(self):
        self.settings_window = QMainWindow()
        self.settings_window.setGeometry(self.settings.SETTINGS['x'] - 200, self.settings.SETTINGS['y'] + 200, 400,
                                         200)
        self.settings_window.setWindowTitle("Settings")

        self.settings_widget = QWidget()
        self.form_layout = QFormLayout()

        for key, value in self.settings.SETTINGS.items():
            line_edit = QLineEdit()
            line_edit.setText(str(value))
            self.form_layout.addRow(key, line_edit)

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(lambda: self.save_settings())
        self.form_layout.addRow("", self.save_button)

        self.toogle_startup_button = QPushButton("")
        self.set_button_toggle_startup_name()
        self.toogle_startup_button.clicked.connect(lambda: self.toogle_startup())
        self.form_layout.addRow("", self.toogle_startup_button)

        self.logout_button = QPushButton("Logout")
        self.logout_button.clicked.connect(lambda: self.logout())
        self.form_layout.addRow("", self.logout_button)

        self.update_events_button = QPushButton("Update events")
        self.update_events_button.clicked.connect(lambda: self.event_viewer.update_events())
        self.form_layout.addRow("", self.update_events_button)

        self.settings_widget.setLayout(self.form_layout)
        self.settings_window.setCentralWidget(self.settings_widget)
        self.settings_window.show()

    def logout(self):
        os.remove('token.json')
        print('end deletening loading events')
        self.event_viewer.load_events()
        print('done loadning')

    def save_settings(self):
        form_layout = self.settings_window.findChild(QFormLayout)

        if form_layout:
            for i in range(form_layout.rowCount()):
                item = form_layout.itemAt(i, QFormLayout.LabelRole)
                if item:
                    key = item.widget().text()
                    item = form_layout.itemAt(i, QFormLayout.FieldRole)
                    value = item.widget().text()
                    self.settings.SETTINGS[key] = int(value)
        print('Saved settings')
        self.settings.save()
        self.event_viewer.update_geometry()


class RepeatThread(Thread):
    def __init__(self, event, f, repeat=60 * 60):
        Thread.__init__(self)
        self.stopped = event
        self.f = f
        self.repeat = repeat

    def run(self):
        while not self.stopped.is_set():
            self.f()
            self.stopped.wait(self.repeat)

    def close(self):
        self.stopped.set()
        self.join()


def main():
    settings = Settings()
    settings.load()

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    calendar_manager = CalendarManagerGoogle()
    event_viewer = EventViewer(app, calendar_manager, settings)
    event_viewer.show()

    stop_event = Event()
    updating_thread = RepeatThread(stop_event, event_viewer.update_events)
    updating_thread.start()

    tray = Tray(app, event_viewer, updating_thread, settings)
    updating_thread.close()

    sys.exit(app.exec_())


if __name__ == '__main__':
    # Autostart starts at this directory
    if os.getcwd() == 'C:\windows\system32':
        shortcut_path = Settings.get_new_shortcut_path()
        shortcut_dir = os.path.dirname(shortcut_path)
        os.chdir(shortcut_dir)

    # If you open shortcut file, find where exe is located. There should be all the files needed.
    files = os.listdir('.')
    if 'CalendarWidget.lnk' in files:
        shell = Dispatch("WScript.Shell")
        shortcut_path = shell.CreateShortCut('CalendarWidget.lnk')
        executable_directory = os.path.dirname(shortcut_path.TargetPath)
        print('Changing executable dir: ', executable_directory)
        os.chdir(executable_directory)

    main()
