import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


def manual_login():
    chrome_options = Options()

    # Используем ту же папку профиля, что и в основном скрипте
    script_dir = os.path.dirname(os.path.abspath(__file__))
    profile_path = os.path.join(script_dir, "selenium_profile")
    chrome_options.add_argument(f"--user-data-dir={profile_path}")
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        # Открываем главную страницу
        driver.get("https://bankrotbaza.ru/")

        print("\n" + "=" * 50)
        print("БРАУЗЕР ОТКРЫТ ДЛЯ РЕГИСТРАЦИИ/ВХОДА")
        print("1. В окне браузера нажми 'Войти' или 'Регистрация'.")
        print("2. Пройди все этапы (введи почту, пароль, подтверди капчу).")
        print("3. Когда увидишь, что ты вошел в личный кабинет — вернись сюда.")
        print("4. НЕ ЗАКРЫВАЙ БРАУЗЕР ВРУЧНУЮ!")
        print("=" * 50 + "\n")

        input("Когда закончишь вход, нажми ENTER здесь, в консоли PyCharm, чтобы сохранить данные и закрыть окно...")

    finally:
        driver.quit()
        print("Данные сохранены в папку 'selenium_profile'. Теперь можно запускать основной main.py!")


if __name__ == "__main__":
    manual_login()