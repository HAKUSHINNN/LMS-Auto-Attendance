import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import json

import customtkinter as ctk
from tkinter import messagebox

from io import BytesIO
from PIL import Image

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# 各webページ
lms_url = "https://lms-tokyo.iput.ac.jp/"
login_url = f"{lms_url}login/index.php"
prof_url = f"{lms_url}user/profile.php"
cal_url = f"https://lms-tokyo.iput.ac.jp/calendar/view.php?view=day"

options = Options().add_argument("--headless")

# config.jsonのインポートとエクスポート
def load_config():
    with open("config.json", "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(config):
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

config = load_config()

# ユーザー名、パスワード、出席タイトルを格納
username = config["username"]
password = config["password"]
attendance_titles = config["attendance_titles"]

# セッションを作成
session = requests.Session()

# ログイン処理
def login(session, username, password):
    login_data = {
        "username": username,
        "password": password,
        "logintoken": "",
    }
    response = session.post(login_url, data=login_data)
    
    if response.url == lms_url:
        print("ログインに成功しました！")
        config['login_token'] = session.cookies.get('MoodleSession')
        config['token_expiration'] = (datetime.now() + timedelta(hours=1)).isoformat()
        save_config(config)

        return True
    else:
        print("ログインに失敗しました...config.jsonファイルの中身を確認して下さい。")
        return False

# カレンダーページにアクセス
def fetch_calendar_page(session):
    return session.get(cal_url)

# カードエレメントを取得
def fetch_course_cards(calendar_page):
    soup = BeautifulSoup(calendar_page.content, "html.parser")
    return soup.find_all("div", class_="card rounded")

# ユーザー名とアイコンの取得
def get_user_info():
    page = session.get(prof_url)
    soup = BeautifulSoup(page.content, "html.parser")
    
    user_name = soup.find("h1", class_="h2").text.split("（")[0]
    user_icon_url = soup.find("img", class_="userpicture").get("src")
    user_icon = session.get(user_icon_url).content

    return user_name, user_icon

# 講義情報を取得
def get_course_info(card):
    # コース名を取得
    course_name = None
    for link in card.find_all("a", href=True):
        if "course/view.php?id=" in link['href']:
            course_name = link.text.strip()
            break

    if not course_name:
        course_name = "不明なコース"

    # 開始・終了時刻を取得
    time_text = ""
    if card.find(class_="dimmed_text"):
        time_text = card.find(class_="dimmed_text").text.strip()
    elif card.find(class_="col-11"):
        time_text = card.find(class_="col-11").text.strip()
    elif card.find(class_="test"):
        time_text = card.find(class_="test").text.strip()

    times = time_text.split(" » ")

    if len(times) >= 2:
        current_date = datetime.now().date()
        start_time = datetime.strptime(times[0].replace("本日, ", "").strip(), "%H:%M").time()
        end_time = datetime.strptime(times[1].strip(), "%H:%M").time()

        start_datetime = datetime.combine(current_date, start_time)
        end_datetime = datetime.combine(current_date, end_time)
    else:
        start_datetime = end_datetime = None

    return course_name, start_datetime, end_datetime

# カードから講義情報を処理する
def process_course_cards(app, cards):
    courses_check = True
    
    for card in cards:
        try:
            course_name, start_datetime, end_datetime = get_course_info(card)
            current_datetime = datetime.now().replace(microsecond=0)
            attendance_management = card.find("h3", class_="name d-inline-block").text.strip()

            if start_datetime <= current_datetime <= end_datetime and any(title in attendance_management for title in config["attendance_titles"]):
                courses_check = False
                app.after(0, lambda: app.update_course_info(course_name, start_datetime, end_datetime))

                course_link = card.find("a", class_="card-link").get("href")
                app.current_course_url = course_link

                course_page = session.get(course_link)
                course_soup = BeautifulSoup(course_page.content, "html.parser")
                attendance_link = course_soup.find("a", string="出欠を送信する")

                if attendance_link:
                    app.current_attendance_url = attendance_link.get("href")
                else:
                    app.current_attendance_url = None
                break  # 最初に見つかった該当コースで処理を終了
        except Exception as e:
            print(f"エラーが発生しました: {e}")

    if courses_check:
        app.after(0, app.no_courses)

FONT = "源ノ角ゴシック Code JP H"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.user_name = ""
        self.user_icon = None
        self.current_attendance_url = None
        self.current_course_url = None
        self.setup_form()
        self.driver = None

    def setup_form(self):
        # デザイン
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
    
        # ウインドウ，アイコン，タイトルの設定
        self.geometry("700x400")
        self.minsize(700, 400)
        self.title("LMS-Auto-Attendance")

        # ウィンドウ全体のグリッド設定
        self.grid_columnconfigure(0, weight=7)  # コース情報
        self.grid_columnconfigure(1, weight=3)  # 出席入力時刻
        self.grid_rowconfigure(0, weight=1)     # ユーザー情報
        self.grid_rowconfigure(1, weight=1)     # コース情報と出席入力時刻
        self.grid_rowconfigure(2, weight=1)     # パスワード入力
    
        # --メニューバーの作成--
        
        # 設定

        # --フレームの作成--

        # ユーザー情報
        self.user_info_frame = ctk.CTkFrame(self, corner_radius=10)
        self.user_info_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.user_info_frame.grid_columnconfigure(0, weight=1)
        self.user_info_frame.grid_columnconfigure(1, weight=1)
        self.user_info_frame.grid_columnconfigure(2, weight=1)
        self.user_info_label = ctk.CTkLabel(self.user_info_frame, text="ユーザー情報", font=(FONT, 15))
        self.user_info_label.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # アイコン
        self.user_icon_label = ctk.CTkLabel(self.user_info_frame, text="")
        self.user_icon_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")
    
        # ユーザーのおなまえ
        self.user_name_label = ctk.CTkLabel(self.user_info_frame, text="")
        self.user_name_label.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
    
        # ボタン
        self.action_button = ctk.CTkButton(self.user_info_frame, text="設定", font=(FONT, 15), command=self.on_button_click)
        self.action_button.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # コース情報
        self.course_info_frame = ctk.CTkFrame(self, corner_radius=10)
        self.course_info_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.course_info_label = ctk.CTkLabel(self.course_info_frame, text="コース情報", font=(FONT, 15))
        self.course_info_label.pack(side="top", padx=10, pady=5)

        # 開始時間情報
        self.course_time_frame = ctk.CTkFrame(self, corner_radius=10)
        self.course_time_frame.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
        self.course_time_label_frame = ctk.CTkLabel(self.course_time_frame, text="出席入力可能時刻", font=(FONT, 15))
        self.course_time_label_frame.pack(side="top", padx=10, pady=5)
    
        # パスワード入力
        self.password_frame = ctk.CTkFrame(self, corner_radius=10)
        self.password_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.password_label = ctk.CTkLabel(self.password_frame, text="パスワード入力", font=(FONT, 15))
        self.password_label.pack(side="top", padx=10, pady=5)

        self.password_entry = ctk.CTkEntry(self.password_frame, placeholder_text="出席パスワードを入力してください", font=(FONT, 15), width=300, height=40)
        self.password_entry.pack(side="top", padx=10, pady=10)

        # 出席ボタン
        self.attendance_button = ctk.CTkButton(self.password_frame, text="出席", font=(FONT, 15), command=self.submit_attendance)
        self.attendance_button.pack(side="top", padx=10, pady=10)

        # コース名
        self.course_name_label = ctk.CTkLabel(self.course_info_frame, text="", font=(FONT, 15))
        self.course_name_label.pack(side="top", padx=10, pady=10, anchor="center")

        # 開始時間
        self.course_time_label = ctk.CTkLabel(self.course_time_frame, text="", font=(FONT, 15))
        self.course_time_label.pack(side="top", padx=10, pady=10, anchor="center")

    def update_window_size(self):
        self.update_idletasks()  # ウィジェットのサイズを更新
        width = self.winfo_reqwidth()
        height = self.winfo_reqheight()
        self.geometry(f"{width}x{height}")

    def open_settings(self):
        new_window = ctk.CTkToplevel(self)
        new_window.title("Settings")
        new_window.geometry("300x200")
        new_label = ctk.CTkLabel(new_window, text="なんにもないぴょーんwwww")
        new_label.pack(padx=20, pady=20)

    # ボタン
    def on_button_click(self):
        self.open_settings()

    # 出席登録ボタン
    def submit_attendance(self):
        student_password = self.password_entry.get()
        if not student_password:
            messagebox.showwarning("登録エラー", "パスワードが入力されていない状態での登録はできません．")
            return

        if self.current_attendance_url:
            self.process_attendance(self.current_attendance_url, student_password)
        else:
            messagebox.showerror("登録エラー", "出席登録のリンクが見つかりませんでした．\n登録ページが現在閉じられている可能性があります．")

    # ログイン情報の表示
    def login_and_fetch_info(self):
        if login(session, username, password):
            # get_user_info関数でユーザー名とアイコンを取得する
            self.user_name, user_icon = get_user_info()
            self.user_number = config["username"]
            
            # ユーザー名を表示
            self.user_name_label.configure(text=f"{self.user_name}({self.user_number})", font=(FONT, 20))

            # アイコンをGUI上に表示
            user_icon_image = Image.open(BytesIO(user_icon))
            user_icon_ctk_image = ctk.CTkImage(light_image=user_icon_image, size=(50, 50))
            self.user_icon_label.configure(image=user_icon_ctk_image)
            self.user_icon_label.image = user_icon_ctk_image

        else:
            self.user_name_label.configure(text="Failed to login.", font=(FONT, 15))

    def update_course_info(self, course_name, start_datetime, end_datetime):
        formatted_starttime = start_datetime.strftime("%H:%M") if start_datetime else "??"
        formatted_endtime = end_datetime.strftime("%H:%M") if start_datetime else "??"
        self.course_name_label.configure(text=f"{course_name}", font=(FONT, 20))
        self.course_time_label.configure(text=f"{formatted_starttime}~{formatted_endtime}", font=(FONT, 20))
        self.update_window_size()

    def clear_course_info(self):
        self.course_name_label.configure(text="")
        self.course_time_label.configure(text="")
        self.update_window_size()

    def no_courses(self):
        self.course_name_label.configure(text="出席できるコースがありません", font=(FONT, 20))
        self.course_time_label.configure(text="ありません", font=(FONT, 20))
        self.update_window_size()

    def update_window_size(self):
        self.update_idletasks()  # ウィジェットのサイズを更新
        width = self.course_info_frame.winfo_reqwidth() + self.course_time_frame.winfo_reqwidth() + 40
        height = self.user_info_frame.winfo_reqheight() + self.course_info_frame.winfo_reqheight() + 40
        self.geometry(f"{width}x{height}")

    # 出席の処理を行う関数
    def process_attendance(self, url, student_password):
        # Seleniumドライバーのセットアップ
        if not self.driver:
            self.driver = webdriver.Chrome(options=options)

        # 無理やりログイン(あとでなおす)
        self.driver.get(login_url)
        
        self.driver.find_element(By.ID, "username").send_keys(username)
        self.driver.find_element(By.ID, "password").send_keys(password)

        # ログインボタンをクリックする
        self.driver.find_element(By.ID, "loginbtn").click()

        self.driver.get(url)
        
        # ラジオボタンの選択
        # self.driver.find_element(By.XPATH, "//input[@type='radio' and @value='出席']")
        radio_button = self.driver.find_element(By.CLASS_NAME, "statusdesc")
        if radio_button:
            radio_button.click()
        
        # パスワードの入力
        password_input = self.driver.find_element(By.NAME, "studentpassword")
        if password_input:
            password_input.send_keys(student_password)
        
        # 変更を保存するボタンのクリック
        save_button = self.driver.find_element(By.NAME, "submitbutton")
        if save_button:
            save_button.click()
        
        if self.driver.current_url == self.current_course_url:
            # 登録完了メッセージの表示
            messagebox.showinfo("登録完了", "出席登録が完了しました！")
            self.driver.quit()
        
        else:
            messagebox.showerror("登録失敗", "出席登録に失敗しました...\nパスワードが間違っているか，出席ページが閉じられた可能性があります．")


if __name__ == "__main__":
    app = App()
    app.login_and_fetch_info()

    calendar_page = fetch_calendar_page(session)
    cards = fetch_course_cards(calendar_page)
    process_course_cards(app, cards)

    app.mainloop()
