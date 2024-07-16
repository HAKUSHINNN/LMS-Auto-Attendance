import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import re
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

options = Options()
options.add_argument("--headless")

# config.jsonのインポートとエクスポート
def load_config():
    with open("config.json", "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(config):
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

config = load_config()

# 時間割
timetable = [
    "09:20", # 1限目
    "10:55", # 2限目
    "13:20", # 3限目
    "3PM",   # 4限目
    "15:00", # 4限目
    "16:40", # 5限目
    "18:20"  # 6限目
]

# ユーザー名、パスワード、出席タイトルを格納
username = config["username"]
password = config["password"]
attendance_titles = config["attendance_titles"]

# セッションを作成
session = requests.Session()

# ログイン処理
def login(session, username, password):
    response = session.get(login_url)
    soup = BeautifulSoup(response.content, "html.parser")
    login_token_input = soup.find("input", {"name": "logintoken"})
    
    if not login_token_input:
        print("ログイントークンが見つかりません。ログインページの構造が変更された可能性があります。")
        login_token = "0"
    else:
        login_token = login_token_input["value"]
    
    login_data = {
        "username": username,
        "password": password,
        "logintoken": login_token
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
def parse_time_string(time_string):
    # AM/PM形式の時間文字列を24時間形式に変換する
    return datetime.strptime(time_string, "%I%p").time()

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
        
        # 開始時間を解析
        start_time_text = times[0].replace("本日, ", "").strip()
        if re.match(r'\d{1,2}[APap][Mm]', start_time_text):
            start_time = parse_time_string(start_time_text)
        else:
            start_time = datetime.strptime(start_time_text, "%H:%M").time()
        
        # 終了時間を解析
        end_time_text = times[1].strip()
        if re.match(r'\d{1,2}[APap][Mm]', end_time_text):
            end_time = parse_time_string(end_time_text)
        else:
            end_time = datetime.strptime(end_time_text, "%H:%M").time()

        # 時間割と一致した場合は開始時刻を5分前にする
        if start_time.strftime("%H:%M") in timetable:
            start_datetime = datetime.combine(current_date, start_time) - timedelta(minutes=5)
        else:
            start_datetime = datetime.combine(current_date, start_time)
        
        end_datetime = datetime.combine(current_date, end_time)
    else:
        start_datetime = end_datetime = None

    return course_name, start_datetime, end_datetime

# 出席登録が完了していたら次の講義情報を流す
def check_attendance(attendance_url):
    response = session.get(attendance_url)
    soup = BeautifulSoup(response.content, "html.parser")
    attendance_status = soup.find("td", class_="statuscol cell c2", string="出席")
    
    return attendance_status is not None 

# カードから講義情報を処理する
def process_course_cards(app, cards):
    current_course_info = None
    next_course_info = None
    current_datetime = datetime.now().replace(microsecond=0)
    
    for card in cards:
        course_name, start_datetime, end_datetime = get_course_info(card)
        attendance_management = card.find("h3", class_="name d-inline-block").text.strip()

        if start_datetime and end_datetime:
            if start_datetime <= current_datetime <= end_datetime and any(title in attendance_management for title in config["attendance_titles"]):
                course_link = card.find("a", class_="card-link").get("href")
                course_page = session.get(course_link)
                course_soup = BeautifulSoup(course_page.content, "html.parser")
                attendance_link = course_soup.find("a", string="出欠を送信する")

                if attendance_link:
                    attendance_url = attendance_link.get("href")
                    if not check_attendance(attendance_url):
                        current_course_info = (course_name, start_datetime, end_datetime)
                        app.current_attendance_url = attendance_url
                        app.current_course_url = course_link
                        break  # 出席登録可能な講義が見つかったら処理を終了
                else:
                    app.current_attendance_url = None
            elif start_datetime > current_datetime:
                if not next_course_info or start_datetime < next_course_info[1]:
                    next_course_info = (course_name, start_datetime, end_datetime)

    if current_course_info:
        app.after(0, lambda: app.update_course_info(*current_course_info))
    elif next_course_info:
        app.after(0, lambda: app.update_course_info(*next_course_info))
    else:
        app.after(0, app.no_courses)


FONT = "Arial"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.user_name = ""
        self.user_icon = None
        self.current_attendance_url = None
        self.current_course_url = None
        self.driver = None
        self.settings_window = None  # 設定ウインドウの参照を保持
        self.setup_form()

    def setup_form(self):
        # デザイン
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
    
        # ウインドウ，アイコン，タイトルの設定
        self.geometry("700x400")
        self.title("LMS-Auto-Attendance")

        # ウィンドウ全体のグリッド設定
        self.grid_columnconfigure(0, weight=7)  # コース情報
        self.grid_columnconfigure(1, weight=3)  # 出席入力時刻
        self.grid_rowconfigure(0, weight=1)     # ユーザー情報
        self.grid_rowconfigure(1, weight=1)     # コース情報と出席入力時刻
        self.grid_rowconfigure(2, weight=1)     # パスワード入力
    
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
        self.action_button = ctk.CTkButton(self.user_info_frame, text="設定", font=(FONT, 15), command=self.open_settings)
        self.action_button.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # コース情報
        self.course_info_frame = ctk.CTkFrame(self, corner_radius=10)
        self.course_info_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.course_info_label = ctk.CTkLabel(self.course_info_frame, text="コース情報", font=(FONT, 15))
        self.course_info_label.pack(side="top", padx=10, pady=5)

        # 開始時間情報
        self.course_time_frame = ctk.CTkFrame(self, corner_radius=10)
        self.course_time_frame.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
        self.course_time_label_frame = ctk.CTkLabel(self.course_time_frame, text="出席入力可能時刻", font=(FONT, 15))
        self.course_time_label_frame.pack(side="top", padx=10, pady=5)
    
        # パスワード入力
        self.password_frame = ctk.CTkFrame(self, corner_radius=10)
        self.password_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.password_label = ctk.CTkLabel(self.password_frame, text="パスワード入力", font=(FONT, 15))
        self.password_label.pack(side="top", padx=10, pady=5)

        self.password_entry = ctk.CTkEntry(self.password_frame, placeholder_text="出席パスワードを入力してください", font=(FONT, 20), width=400, height=60, justify="center")
        self.password_entry.pack(side="top", padx=10, pady=10)

        # 出席ボタン
        self.attendance_button = ctk.CTkButton(self.password_frame, text="出席", font=(FONT, 30), width=200, height=60, command=self.submit_attendance)
        self.attendance_button.pack(side="top", padx=10, pady=10)

        # コース名
        self.course_name_label = ctk.CTkLabel(self.course_info_frame, text="", font=(FONT, 15))
        self.course_name_label.pack(side="top", padx=10, pady=10, anchor="center")

        # 開始時間
        self.course_time_label = ctk.CTkLabel(self.course_time_frame, text="", font=(FONT, 15))
        self.course_time_label.pack(side="top", padx=10, pady=10, anchor="center")

        # ここでウィンドウサイズを更新し、最小サイズを設定する
        width, height = self.update_window_size()
        width += 300 # 調整
        self.minsize(width, height)

    def update_window_size(self):
        self.update_idletasks()  # ウィジェットのサイズを更新
        width = self.winfo_reqwidth() + 20 # ウィンドウ全体の要求幅を取得
        height = self.winfo_reqheight() + 20  # ウィンドウ全体の要求高さを取得
        self.geometry(f"{width}x{height}")
        return width, height
    
    def open_settings(self):
        self.settings_window = ctk.CTkToplevel(self)
        self.settings_window.title("Settings")
        self.settings_window.geometry("300x300")

        username_label = ctk.CTkLabel(self.settings_window, text="ユーザー名(tkxxxxxx)", font=(FONT, 12))
        username_label.pack(padx=10, pady=5)

        self.username_entry = ctk.CTkEntry(self.settings_window, placeholder_text="ユーザー名を入力して下さい", font=(FONT, 12), width=250, justify="center")
        self.username_entry.pack(padx=10, pady=5)
        self.username_entry.insert(0, username)

        password_label = ctk.CTkLabel(self.settings_window, text="パスワード", font=(FONT, 12))
        password_label.pack(padx=10, pady=5)

        self.password_entry = ctk.CTkEntry(self.settings_window, placeholder_text="Enter password", font=(FONT, 12), width=250, show='*', justify="center")
        self.password_entry.pack(padx=10, pady=5)
        self.password_entry.insert(0, password)

        save_button = ctk.CTkButton(self.settings_window, text="保存", font=(FONT, 12), command=self.save_login_info)
        save_button.pack(padx=10, pady=20)

        self.auto_quit_var = ctk.BooleanVar(value=config.get("auto_quit", False))
        self.auto_quit = ctk.CTkCheckBox(self.settings_window, text="登録が完了したら自動でツールを閉じる", font=(FONT, 12), variable=self.auto_quit_var)
        self.auto_quit.pack(padx=10, pady=5)

    def save_login_info(self):
        global username, password
        username = self.username_entry.get()
        password = self.password_entry.get()
        auto_quit = self.auto_quit_var.get()


        config["username"] = username
        config["password"] = password
        config["font"] = FONT
        config["auto_quit"] = auto_quit

        save_config(config)

        if login(session, username, password):
            messagebox.showinfo("ログイン成功", f"{username}としてログインに成功しました！")
            self.settings_window.destroy()  # 設定ウィンドウを閉じる
            self.settings_window = None  # 設定ウィンドウの参照をリセット

            self.login_and_fetch_info()
        else:
            messagebox.showerror("ログイン失敗", "ログインに失敗しました...\nユーザー名 or パスワードが間違っています．")

    def submit_attendance(self):
        student_password = self.password_entry.get()
        if not student_password:
            messagebox.showwarning("登録エラー", "パスワードが入力されていない状態での登録はできません．")
            return

        if self.current_attendance_url:
            self.process_attendance(self.current_attendance_url, student_password)
            # 出席が完了した後に次の講義を探す
            calendar_page = fetch_calendar_page(session)
            cards = fetch_course_cards(calendar_page)
            process_course_cards(self, cards)
        else:
            messagebox.showerror("登録エラー", "出席登録のリンクが見つかりませんでした．\n登録ページが現在閉じられている可能性があります．")

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

            # カレンダーページを取得して講義情報を処理する
            calendar_page = fetch_calendar_page(session)
            cards = fetch_course_cards(calendar_page)
            process_course_cards(self, cards)
        else:
            self.user_name_label.configure(text="ログインに失敗しました．右の設定ボタンからログインして下さい．", font=(FONT, 20))
        
    def update_course_info(self, course_name, start_datetime, end_datetime):
        formatted_starttime = start_datetime.strftime("%H:%M") if start_datetime else "??"
        formatted_endtime = end_datetime.strftime("%H:%M") if start_datetime else "??"
        self.course_name_label.configure(text=f"{course_name}", font=(FONT, 20))
        self.course_time_label.configure(text=f"{formatted_starttime}~{formatted_endtime}", font=(FONT, 20))
        self.update_window_size()

    def no_courses(self):
        self.course_name_label.configure(text="出席できるコースがありません", font=(FONT, 20))
        self.course_time_label.configure(text="ありません", font=(FONT, 20))
        self.update_window_size()

    # 出席の処理を行う関数
    def process_attendance(self, url, student_password):
        # Seleniumドライバーのセットアップ
        options = Options()
        options.add_argument("--headless")

        self.driver = webdriver.Chrome(options=options)

        # ログインページにアクセス
        self.driver.get(login_url)

        # ユーザー名とパスワードを入力
        self.driver.find_element(By.ID, "username").send_keys(username)
        self.driver.find_element(By.ID, "password").send_keys(password)

        # ログインボタンをクリックする
        self.driver.find_element(By.ID, "loginbtn").click()

        self.driver.get(url)
    
        # ラジオボタンの選択
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
            self.password_entry.delete(0, ctk.END)
            
        if config["auto_quit"] == True:
            self.quit()
            
        else:
            messagebox.showerror("登録失敗", "出席登録に失敗しました...\nパスワードが間違っているか，出席ページが閉じられた可能性があります。")

if __name__ == "__main__":
    app = App()
    app.login_and_fetch_info()

    calendar_page = fetch_calendar_page(session)
    cards = fetch_course_cards(calendar_page)
    process_course_cards(app, cards)

    app.mainloop()
