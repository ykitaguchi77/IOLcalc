import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QLineEdit
from PyQt5.QtCore import QTimer
from AppKit import NSWorkspace
import pyautogui
import time
import pandas as pd
import pyperclip
import re

class ActivateChromeWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IOL autocalculator")
        screen_geometry = QApplication.desktop().screenGeometry()
        window_width = 400
        window_height = 700
        x = screen_geometry.width() - window_width - 20
        y = 20
        self.setGeometry(x, y, window_width, window_height)

        self.link_label = QLabel("<a href='https://iolcalculator.escrs.org/'>https://iolcalculator.escrs.org/</a>", self)
        self.link_label.setGeometry(50, 20, 300, 30)
        self.link_label.setOpenExternalLinks(True)

        self.id_label = QLabel("整理番号を入力してください:", self)
        self.id_label.setGeometry(50, 50, 300, 30)

        self.id_input = QLineEdit(self)
        self.id_input.setGeometry(50, 90, 200, 30)

        # Minimumボタンを追加
        self.minimum_button = QPushButton("Minimum", self)
        self.minimum_button.setGeometry(250, 90, 100, 30)
        self.minimum_button.clicked.connect(self.select_minimum_id)

        self.show_button = QPushButton("Calculate", self)
        self.show_button.setGeometry(50, 140, 200, 50)
        self.show_button.clicked.connect(self.input_data)

        self.edit_button = QPushButton("Edit", self)
        self.edit_button.setGeometry(50, 210, 200, 50)
        self.edit_button.clicked.connect(self.edit_data)

        self.capture_button = QPushButton("Capture", self)
        self.capture_button.setGeometry(50, 280, 200, 50)
        self.capture_button.clicked.connect(self.capture_result)

        self.result_label = QLabel(self)
        self.result_label.setGeometry(50, 340, 300, 360)
        self.result_label.setStyleSheet("font-size: 12px;")

    def activate_chrome(self):
        chrome_app = "Google Chrome"
        workspace = NSWorkspace.sharedWorkspace()
        workspace.launchApplication_(chrome_app)
        print("Google Chrome is now active.")
        time.sleep(1)

    def input_data(self):
        self.show_result()
        self.activate_chrome()
        pyautogui.press('tab', presses=15)
        pyautogui.typewrite(str(self.surgeon))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.initial))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.id))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.Age))
        pyautogui.press('tab')
        pyautogui.press('space')
        pyautogui.press('down')
        pyautogui.press('space')
        time.sleep(1)
        pyautogui.press('tab', presses=2)
        pyautogui.typewrite(str(self.AL))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.ACD))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.LT))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.CCT))
        pyautogui.press('tab', presses=2)
        pyautogui.typewrite(str(self.K1))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.K2))
        pyautogui.press('tab', presses=2)
        pyautogui.typewrite(str(self.Target))
        pyautogui.press('tab')
        pyautogui.press('space')
        pyautogui.press('down', presses=5)
        pyautogui.press('space')
        time.sleep(1)

        pyautogui.press('tab')
        pyautogui.press('enter')
        time.sleep(1)

        pyautogui.press('enter')
        time.sleep(1)

        pyautogui.press('tab', presses=32)
        pyautogui.press('space')

    def edit_data(self):
        self.show_result()
        self.activate_chrome()
        pyautogui.press('tab', presses=15)
        pyautogui.typewrite(str(self.surgeon))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.initial))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.id))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.Age))
        pyautogui.press('tab')
        pyautogui.press('tab', presses=2)
        pyautogui.typewrite(str(self.AL))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.ACD))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.LT))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.CCT))
        pyautogui.press('tab', presses=2)
        pyautogui.typewrite(str(self.K1))
        pyautogui.press('tab')
        pyautogui.typewrite(str(self.K2))
        pyautogui.press('tab', presses=2)
        pyautogui.typewrite(str(self.Target))
        pyautogui.press('tab', presses=34)
        pyautogui.press('space')

    def activate_window(self):
        self.activateWindow()
        self.raise_()

    def show_result(self):
        id_number = int(self.id_input.text())

        try:
            df = pd.read_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', encoding='shift_jis')
        except UnicodeDecodeError:
            df = pd.read_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', encoding='cp932')
            print("cp932")

        if id_number in df['整理番号'].values:
            row = df.loc[df['整理番号'] == id_number, :].squeeze()
            print(f"row: {row}")

            IOLdata = [int(row['年齢']), float(row['AL']), float(row['OA術前ACD(Epi)']), float(row['OA術前LT']),
                     int(row['OA術前角膜厚']), float(row['K(弱)OA']), float(row['K(強)OA']), float(row['予測屈折OA2000']), float(row['IOLﾊﾟﾜｰ'])]

            self.surgeon = "Goto"
            self.initial = "GG"
            self.id = id_number
            self.Age = IOLdata[0]
            self.Gender = "Male"
            self.AL = IOLdata[1]
            self.ACD = IOLdata[2]
            self.LT = IOLdata[3]
            self.CCT = IOLdata[4]
            self.K1 = IOLdata[5]
            self.K2 = IOLdata[6]
            self.Target = IOLdata[7]
            self.IOLPW = IOLdata[8]
            self.Manufacturer = "Alcon"
            self.IOL = "AcrySof AU00T0"

            result_text = f"Number: {id_number}\nSurgeon: {self.surgeon}\nInitial: {self.initial}\nID: {self.id}\nAge: {self.Age}\nGender: {self.Gender}\nAL: {self.AL}\nACD: {self.ACD}\nLT: {self.LT}\nCCT: {self.CCT}\nK1: {self.K1}\nK2: {self.K2}\nTarget: {self.Target}\nManufacturer: {self.Manufacturer}\nIOL: {self.IOL}\n\nIOL_Power: {self.IOLPW}"
            self.result_label.setText(result_text)
        else:
            self.result_label.setText('番号が見つかりませんでした')

    def capture_result(self):
        self.show_result()
        print("Capture button clicked")
        self.activate_chrome()
        time.sleep(10)
        self.capture_screen()
        QTimer.singleShot(1000, self.capture_screen)
        numbers = self.extract_numbers()
        if numbers:
            id_number = int(self.id_input.text())

            try:
                df = pd.read_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', encoding='shift_jis')
            except UnicodeDecodeError:
                df = pd.read_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', encoding='cp932')

            if id_number in df['整理番号'].values:
                index = df[df['整理番号'] == id_number].index[0]

                df.loc[index, 'Barrett狙い'] = float(numbers[0])
                df.loc[index, 'EVO狙い'] = float(numbers[1])
                df.loc[index, 'Hill狙い'] = float(numbers[2])
                df.loc[index, 'HofferQST狙い'] = float(numbers[3])
                df.loc[index, 'Kane狙い'] = float(numbers[4])
                df.loc[index, 'PearlDGS狙い'] = float(numbers[5])

                df.to_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', index=False, encoding='shift_jis')
                
                result_text = "抽出された数字:\n"
                for number in numbers:
                    result_text += f"{number}\n"
                self.result_label.setText(result_text)
            else:
                self.result_label.setText('番号が見つかりませんでした')
        else:
            self.result_label.setText("数字が見つかりませんでした。")
        QTimer.singleShot(500, self.activate_window)

    def capture_screen(self):
        pyautogui.hotkey('command', 'a')
        pyautogui.hotkey('command', 'c')
        print("キャプチャされました")
        numbers = self.extract_numbers()
        if numbers:
            result_text = "抽出された数字:\n"
            for number in numbers:
                result_text += f"{number}\n"
            self.result_label.setText(result_text)
        else:
            self.result_label.setText("数字が見つかりませんでした。")
        QTimer.singleShot(500, self.activate_window)

    def extract_numbers(self):
        clipboard_text = pyperclip.paste()
        lines = clipboard_text.split('\n')
        
        iolpw = self.IOLPW
        
        pattern = rf"^{iolpw:g}\s+([-]?\d+\.\d+)\s+([-]?\d+\.\d+)\s+([-]?\d+\.\d+)\s+([-]?\d+\.\d+)\s+([-]?\d+\.\d+)\s+([-]?\d+\.\d+)"
        match = re.search(pattern, clipboard_text, re.MULTILINE)
        
        if match:
            numbers = match.groups()
            return numbers
        else:
            return None
        
    def select_minimum_id(self):
        try:
            df = pd.read_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', encoding='shift_jis')
        except UnicodeDecodeError:
            df = pd.read_csv('MYT Clareon Oformula 北口先生データシート匿名ID.csv', encoding='cp932')

        blank_rows = df[df['HofferQST狙い'].isna()]

        if not blank_rows.empty:
            min_id = blank_rows['整理番号'].min()
            self.id_input.setText(str(min_id))
        else:
            self.result_label.setText('空白の整理番号が見つかりませんでした')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ActivateChromeWindow()
    window.show()
    sys.exit(app.exec_())