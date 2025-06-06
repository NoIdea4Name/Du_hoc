from PyQt6.uic import loadUi
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal
from openpyxl import load_workbook, Workbook
from modules.Bot import Bot
from modules.ChangePasswordBot import ChangePasswordBot
import sys
import os
import requests
import re
import random

class WorkerThread(QThread):
    finished = pyqtSignal()

    def __init__(self, config, log_callback):
        super().__init__()
        self.config = config
        self.add_log = log_callback
        self.is_running = False
        self.headers = ["Email", "Student ID", "Password", "Full Name", "Gender", "Birthdate", "Street", "City", "State", "Zipcode", "SSN", "Password changed"]

    def stop(self):
        self.is_running = False

    def run(self):
        self.is_running = True

        data_file = self.config['data_file']
        backup_file = self.config['backup_file']

        if os.path.exists(data_file):
            workbook = load_workbook(data_file)
            sheet = workbook.active

            if self.config['backup_enabled']:
                workbook.save(backup_file)
                self.add_log(f"🔒 Đã tạo file backup: {backup_file}")
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Data"
            sheet.append(self.headers)

        if self.config['mail_enabled'] and self.config['password_enabled']:
            self.run_create_and_change_password(sheet, workbook)
        elif self.config['mail_enabled']:
            print("Chạy tạo email")
            self.run_create_only(sheet, workbook)
        elif self.config['password_enabled']:
            self.run_change_existing_passwords(sheet, workbook)

        self.finished.emit()

    def run_create_and_change_password(self, sheet, workbook):
        from threading import Thread, Lock
        num_threads = self.config.get('num_threads', 1)
        bot = Bot(token=self.config['token'], headless_mode=self.config['headless_mode'], wait_sec=self.config['wait_sec_get_mail'])
        pw_bot = ChangePasswordBot(token=self.config['token'], wait_sec=self.config['wait_sec_get_mail'])
        success_count = 0
        lock = Lock()

        def worker(start_idx, step):
            nonlocal success_count
                
            for i in range(start_idx, self.config['tries'] + 1, step):
                if not self.is_running:
                    self.add_log(f"💀 [Bot {start_idx}] Chương trình đã bị dừng bởi người dùng.")
                    break
    
                result = bot.run(domain=self.config["domain"])
                if not result:
                    self.add_log(f"❌ [Bot {start_idx}] Có lỗi xảy ra khi tạo tài khoản {i}")
                    continue
    
                self.add_log(f"✅ [Bot {start_idx}] Tài khoản {i} tạo thành công: {result.get('Student ID')}")
    
                try:
                    change_result = pw_bot.run(
                        email=result["Email"],
                        sfid=result["Student ID"],
                        password=pw_bot.generate_password() if self.config['create_new_password'] else result.get("Password"),
                        wait_sec_get_mail=self.config['wait_sec_get_mail']
                    )
    
                    if change_result:
                        self.add_log(f"✅ [Bot {start_idx}] Đổi thành công mật khẩu cho {result['Student ID']}")
                        result["Password"] = change_result
                        result["Password changed"] = "y"
    
                except Exception as e:
                    self.add_log(f"❌ [Bot {start_idx}] Có lỗi xảy ra khi đổi mật khẩu cho id {result['Student ID']}: {e}")
    
                with lock:
                    sheet.append([result.get(h, "") for h in self.headers])
                    workbook.save(self.config['data_file'])
                    success_count += 1
                    self.add_log(f"🚩 Đã tạo {success_count}/{self.config['tries']} tài khoản")
    
        threads = []
        for idx in range(num_threads):
            t = Thread(target=worker, args=(idx+1, num_threads))
            threads.append(t)
            t.start()
        for t in threads:
            t.join()

    def run_create_only(self, sheet, workbook):
        from threading import Thread, Lock
        num_threads = self.config.get('num_threads', 1)
        bot = Bot(token=self.config['token'], wait_sec=self.config['wait_sec_get_mail'])
        success_count = 0
        lock = Lock()
    
        def worker(start_idx, step):
            nonlocal success_count
                
            for i in range(start_idx, self.config['tries'] + 1, step):
                if not self.is_running:
                    self.add_log(f"💀 [Bot {start_idx}] Chương trình đã bị dừng bởi người dùng.")
                    break
                
                result = bot.run(domain=self.config["domain"])
                if not result:
                    self.add_log(f"❌ [Bot {start_idx}] Đã có lỗi xảy ra khi tạo tài khoản {i}")
                    continue
                with lock:
                    sheet.append([result.get(h, "") for h in self.headers])
                    workbook.save(self.config['data_file'])
                    success_count += 1
                    self.add_log(f"✅ [Bot {start_idx}] Tài khoản {i} tạo thành công: {result.get('Student ID')}")
                    self.add_log(f"🚩 Đã tạo {success_count}/{self.config['tries']} tài khoản")
    
        threads = []
        for idx in range(num_threads):
            t = Thread(target=worker, args=(idx+1, num_threads))
            threads.append(t)
            t.start()
        for t in threads:
            t.join()
            
    def run_change_existing_passwords(self, sheet, workbook):
        from threading import Thread, Lock
        num_threads = self.config.get('num_threads', 1)
        data_file = self.config['data_file']
        lock = Lock()
        success_count = 0
    
        # Đọc lại toàn bộ dữ liệu để chia cho các luồng
        workbook = load_workbook(data_file)
        sheet = workbook.active
        rows = list(sheet.iter_rows(min_row=2, values_only=False))

        change_pw_bot = ChangePasswordBot(token=self.config['token'], wait_sec=self.config['wait_sec_get_mail'])
        def worker(thread_idx, step):
            nonlocal success_count
            for idx in range(thread_idx, len(rows), step):
                row = rows[idx]
                if not self.is_running:
                    self.add_log(f"💀 [Thread {thread_idx+1}] Chương trình đã bị dừng bởi người dùng.")
                    break
                if row[11].value == "n":
                    email, sfid, current_pw = row[0].value, row[1].value, row[2].value
                    try:
                        change_result = change_pw_bot.run(
                            email=email,
                            sfid=sfid,
                            password=change_pw_bot.generate_password() if self.config['create_new_password'] else current_pw,
                            wait_sec_get_mail=self.config['wait_sec_get_mail']
                        )
                        if change_result:
                            with lock:
                                row[2].value = change_result
                                row[11].value = "y"
                                workbook.save(self.config['data_file'])
                                success_count += 1
                                self.add_log(f"✅ [Thread {thread_idx+1}] Đã đổi mật khẩu cho id {sfid}")
                    except Exception as e:
                        self.add_log(f"❌ [Thread {thread_idx+1}] Có lỗi xảy ra khi đổi mật khẩu cho {sfid}: {e}")
    
        threads = []
        for idx in range(num_threads):
            t = Thread(target=worker, args=(idx, num_threads))
            threads.append(t)
            t.start()
        for t in threads:
            t.join()
    
        self.add_log(f"🚩 Đã đổi thành công mật khẩu cho {success_count} tài khoản.")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi('ui/main.ui', self)
        self.token_file = "token.txt"
        self.thread = None

        self.run.clicked.connect(self.start_thread)
        self.stop.clicked.connect(self.stop_thread)

        self.mail.toggled.connect(self.update_ui_state)
        self.password.toggled.connect(self.update_ui_state)
        
        self.saveToken.clicked.connect(self.save_token_to_file)
        self.refreshDomain.clicked.connect(self.fetch_domain)

        self.load_token()
        self.update_ui_state()
        self.fetch_domain()
    
    def add_log(self, text):
        self.log.addItem(text)
        self.log.scrollToBottom()

    def fetch_domain(self):
        resp = requests.get("https://tempmail.id.vn/en").text
        if not resp:
            self.add_log("Không thể lấy danh sách domain")
            return
        
        options = re.findall(r'<option[^>]*>(.*?)</option>', resp)
        if options:
            self.domainList.clear()
            self.domainList.addItem("Ngẫu nhiên")
            for item in options:
                self.domainList.addItem(item)
        else:
            print("Không thể lấy danh sách domain")
            
    def save_token_to_file(self):
        with open(self.token_file, 'w') as file:
            file.write(self.tempMailToken.text())
            self.add_log(f"🎉 Đã lưu token vào file: {self.token_file}")
    
    def load_token(self):
        if os.path.exists(self.token_file):
            with open(self.token_file, "r") as file:
                self.tempMailToken.setText(file.read())

    def update_ui_state(self):
        can_run = self.mail.isChecked() or self.password.isChecked()
        self.run.setEnabled(can_run and not self.thread)
        self.stop.setEnabled(bool(self.thread))

    def start_thread(self):
        token = self.tempMailToken.text()
        data_file = self.dataFile.text()
        backup_file = self.backupFile.text()

        if not token:
            self.show_message("Nhập token của TempMail trước khi chạy.")
            return

        if not data_file:
            self.show_message("Nhập tên file dữ liệu trước khi chạy.")
            return

        if self.backupFile.isEnabled() and not backup_file:
            self.show_message("Nhập tên file backup trước khi chạy.")
            return
        
        domain = self.domainList.currentText()
        if domain == "Ngẫu nhiên":
            items = [self.domainList.itemText(i) for i in range(1, self.domainList.count())]
            domain = random.choice(items)

        config = {
            'token': token,
            'domain': domain,
            'tries': self.quantity.value(),
            'headless_mode': self.headlessMode.isChecked(),
            'data_file': data_file,
            'backup_file': backup_file,
            'backup_enabled': self.isBackup.isChecked(),
            'create_new_password': self.newPass.isChecked(),
            'wait_sec_get_mail': self.waitSecGetMail.value(),
            'mail_enabled': self.mail.isChecked(),
            'password_enabled': self.password.isChecked(),
            'num_threads': self.numThreads.value(),
        }

        self.thread = WorkerThread(config, self.add_log)
        self.thread.finished.connect(self.thread_finished)
        self.thread.start()
        self.update_ui_state()

    def thread_finished(self):
        self.thread = None
        self.update_ui_state()

    def stop_thread(self):
        if self.thread:
            self.thread.stop()

    def show_message(self, text, title="Warning"):
        QMessageBox.warning(self, title, text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("WindowsVista")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())