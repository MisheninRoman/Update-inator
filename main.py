import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import *
import threading
import time
import os
import json


class BrowserApp:
    def __init__(self, root):
        self.root = root
        self.driver = None
        self.is_running = False
        self.browser_launched = False

        # файл где будем хранить настройки
        self.settings_file = "browser_settings.json"
        self.settings = self.load_settings()

        # делаем кнопки

        # поле для пути к браузеру
        tk.Label(root, text="Путь к Яндекс.раузеру:").pack(pady=5)
        browser_frame = tk.Frame(root)
        browser_frame.pack(pady=5, fill=tk.X, padx=10)

        self.browser_path_var = tk.StringVar(
            value=self.settings.get("browser_path", "")
        )
        self.browser_entry = tk.Entry(
            browser_frame, textvariable=self.browser_path_var, width=35
        )
        self.browser_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_browse = tk.Button(
            browser_frame, text="Обзор", command=self.browse_browser_path
        )
        self.btn_browse.pack(side=tk.RIGHT, padx=(5, 0))

        # кнопка для автоматического поиска браузера
        self.btn_auto_find = tk.Button(
            root, text="Авто-поиск браузера", command=self.auto_find_browser
        )
        self.btn_auto_find.pack(pady=5)

        self.btn_launch = tk.Button(
            root, text="Запустить браузер", command=self.launch_browser
        )
        self.btn_launch.pack(pady=5)

        self.btn_close_browser = tk.Button(
            root, text="Закрыть браузер", command=self.close_browser, state=tk.DISABLED
        )
        self.btn_close_browser.pack(pady=5)

        # поле для селектора если автоматически не находит
        tk.Label(root, text="CSS-селектор (если автоматически не находит):").pack()
        self.selector_var = tk.StringVar(
            value=self.settings.get("selector", "[aria-label='Обновить']")
        )
        self.selector_entry = tk.Entry(root, textvariable=self.selector_var, width=40)
        self.selector_entry.pack(pady=5)

        self.btn_find = tk.Button(
            root, text="Найти кнопку", command=self.find_button, state=tk.DISABLED
        )
        self.btn_find.pack(pady=5)

        # поле для ввода интервала
        tk.Label(root, text="Интервал (секунды):").pack()
        self.interval_var = tk.StringVar(value="5")
        self.interval_entry = tk.Entry(root, textvariable=self.interval_var)
        self.interval_entry.pack(pady=5)

        self.btn_start = tk.Button(
            root,
            text="Запустить авто-клик",
            command=self.start_clicking,
            state=tk.DISABLED,
        )
        self.btn_start.pack(pady=5)

        # кнопка остановки
        self.btn_stop = tk.Button(
            root, text="Остановить", command=self.stop_clicking, state=tk.DISABLED
        )
        self.btn_stop.pack(pady=5)

        # статус
        self.status = tk.Label(root, text="Браузер не запущен")
        self.status.pack(pady=5)

        # поле для урла
        tk.Label(root, text="URL страницы:").pack()
        self.url_var = tk.StringVar(
            value=self.settings.get("url", "https://hpsm.emias.mos.ru/sm/index.do")
        )
        self.url_entry = tk.Entry(root, textvariable=self.url_var, width=40)
        self.url_entry.pack(pady=5)

    def load_settings(self):
        # пробуем загрузить настройки из файла
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Ошибка когда загружали настройки: {e}")
        return {}

    def save_settings(self):
        # сохраняем настройки в файл
        try:
            settings = {
                "browser_path": self.browser_path_var.get(),
                "url": self.url_var.get(),
                "selector": self.selector_var.get(),
            }
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка когда сохраняли настройки: {e}")

    def browse_browser_path(self):
        # выбираем файл браузера через диалог
        file_path = filedialog.askopenfilename(
            title="Выберите исполняемый файл браузера",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")],
        )
        if file_path:
            self.browser_path_var.set(file_path)
            self.save_settings()

    def auto_find_browser(self):
        # ищем яндекс браузер в компе
        possible_paths = [
            "C:/Users/User/AppData/Local/Yandex/YandexBrowser/Application/browser.exe",
            "C:/Program Files (x86)/Yandex/YandexBrowser/Application/browser.exe",
            "C:/Program Files/Yandex/YandexBrowser/Application/browser.exe",
            os.path.expanduser(
                "~/AppData/Local/Yandex/YandexBrowser/Application/browser.exe"
            ),
        ]

        for path in possible_paths:
            if os.path.exists(path):
                self.browser_path_var.set(path)
                self.status.config(text=f"Браузер найден!")
                self.save_settings()
                return

        self.status.config(text="Браузер не найден автоматически")

    def launch_browser(self):
        # проверяем если браузер уже запущен
        if self.browser_launched and self.driver:
            self.status.config(text="Браузер уже запущен!")
            return

        try:
            options = webdriver.ChromeOptions()

            # берем путь из поля ввода
            browser_path = self.browser_path_var.get().strip()

            if browser_path and os.path.exists(browser_path):
                options.binary_location = browser_path
                self.status.config(text="Используется указанный путь к браузеру...")
            else:
                # если путь не указан пробуем найти сами
                self.auto_find_browser()
                browser_path = self.browser_path_var.get().strip()

                if browser_path and os.path.exists(browser_path):
                    options.binary_location = browser_path
                    self.status.config(
                        text="Используется автоматически найденный браузер..."
                    )
                else:
                    self.status.config(text="Браузер не найден. Укажите путь вручную.")
                    return

            # путь к драйверу
            driver_path = "./yandexdriver.exe"

            # проверяем есть ли драйвер
            if not os.path.exists(driver_path):
                self.status.config(
                    text="Драйвер не найден! Убедитесь, что yandexdriver.exe в папке со скриптом."
                )
                return

            # запускаем драйвер
            service = ChromeService(executable_path=driver_path)
            self.driver = webdriver.Chrome(service=service, options=options)

            # открываем урл
            url = self.url_var.get()
            self.driver.get(url)

            # сохраняем настройки
            self.save_settings()

            # обновляем статус
            self.browser_launched = True
            self.status.config(text="Браузер запущен")

            # меняем кнопки
            self.btn_launch.config(state=tk.DISABLED)
            self.btn_close_browser.config(state=tk.NORMAL)
            self.btn_find.config(state=tk.NORMAL)
            self.btn_start.config(state=tk.DISABLED)

        except Exception as e:
            self.status.config(text=f"Ошибка: {str(e)}")
            print("Ошибка при запуске браузера:", e)

    def close_browser(self):
        # закрываем браузер
        if self.driver:
            # останавливаем кликер если работает
            self.stop_clicking()

            # закрываем
            self.driver.quit()
            self.driver = None
            self.browser_launched = False

            # убираем селектор если был
            if hasattr(self, "successful_selector"):
                del self.successful_selector

            # обновляем кнопки
            self.btn_launch.config(state=tk.NORMAL)
            self.btn_close_browser.config(state=tk.DISABLED)
            self.btn_find.config(state=tk.DISABLED)
            self.btn_start.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.DISABLED)

            self.status.config(text="Браузер закрыт")

    def find_button(self):
        # ищем кнопку на странице
        if not self.driver:
            self.status.config(text="Сначала запустите браузер!")
            return

        try:
            self.status.config(text="Ищу кнопку...")
            wait = WebDriverWait(self.driver, 10)

            # разные способы найти кнопку
            selectors = [
                "[aria-label='Обновить']",
                ".x-btn-text[aria-label='Обновить']",
                "button[aria-label='Обновить']",
                "[aria-label='Обновить'][style*='trefresh.png']",
                ".x-btn-text[style*='trefresh.png']",
                ".x-btn-text",
                "//button[@aria-label='Обновить']",
                "//*[contains(@class, 'x-btn-text') and @aria-label='Обновить']",
                "//*[@aria-label='Обновить' and contains(@style, 'trefresh.png')]",
                "//*[contains(@class, 'x-btn-text') and contains(@style, 'trefresh.png')]",
                "//button[text()='Обновить']",
            ]

            # добавляем свой селектор если написали
            custom_selector = self.selector_var.get().strip()
            if custom_selector and custom_selector not in selectors:
                selectors.insert(0, custom_selector)

            found_selector = None

            for selector in selectors:
                try:
                    # пробуем CSS селектор
                    if not selector.startswith("//"):
                        element = wait.until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    else:
                        # пробуем XPath
                        element = wait.until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    found_selector = selector
                    print(f"Кнопка найдена с селектором: {selector}")

                    # подсвечиваем кнопку красным
                    self.driver.execute_script(
                        "arguments[0].style.border='3px solid red'", element
                    )
                    break
                except Exception as e:
                    print(f"Селектор {selector} не сработал: {e}")
                    continue

            if found_selector:
                self.successful_selector = found_selector
                self.status.config(text=f"Кнопка найдена! Селектор: {found_selector}")

                # сохраняем настройки
                self.save_settings()

                # включаем кнопку старта
                self.btn_start.config(state=tk.NORMAL)
            else:
                self.status.config(text="Кнопка не найдена! Попробуйте другой селектор")

        except Exception as e:
            self.status.config(text=f"Ошибка: {str(e)}")
            print("Ошибка при поиске кнопки:", e)

    def find_and_click_button(self):
        # ищем и кликаем кнопку
        if not hasattr(self, "successful_selector") or not self.successful_selector:
            return False

        try:
            wait = WebDriverWait(self.driver, 5)
            selector = self.successful_selector

            # пробуем найти элемент
            if not selector.startswith("//"):
                element = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
            else:
                element = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))

            # прокручиваем чтобы было видно
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", element
            )

            # кликаем
            element.click()
            return True

        except StaleElementReferenceException:
            # элемент устарел, пробуем еще раз
            print("Элемент устарел, пытаюсь найти заново...")
            return self.find_and_click_button()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Не удалось найти элемент: {e}")
            return False
        except Exception as e:
            print(f"Ошибка при клике: {e}")
            return False

    def start_clicking(self):
        # запускаем авто-клик
        if not hasattr(self, "successful_selector"):
            self.status.config(text="Сначала найдите кнопку!")
            return

        # проверяем интервал
        try:
            interval = int(self.interval_var.get())
            if interval <= 0:
                self.status.config(text="Интервал должен быть положительным числом!")
                return
        except ValueError:
            self.status.config(text="Введите корректный интервал (число)!")
            return

        self.is_running = True
        self.status.config(text="Авто-клик запущен")

        # меняем кнопки
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_find.config(state=tk.DISABLED)

        # запускаем в отдельном потоке
        threading.Thread(target=self.click_loop, daemon=True).start()

    def click_loop(self):
        # цикл кликов
        interval = int(self.interval_var.get())
        click_count = 0

        while self.is_running and self.driver and self.browser_launched:
            try:
                # пробуем кликнуть
                if self.find_and_click_button():
                    click_count += 1
                    current_time = time.strftime("%H:%M:%S")
                    status_text = f"Клик #{click_count} выполнен в {current_time}"
                    self.status.config(text=status_text)
                    print(status_text)
                else:
                    self.status.config(text="Не удалось найти кнопку для клика")
                    self.is_running = False
                    break

                # ждем
                for i in range(interval):
                    if not self.is_running:
                        break
                    time.sleep(1)

            except Exception as e:
                error_msg = f"Ошибка клика: {str(e)}"
                print(error_msg)
                self.status.config(text=error_msg)
                self.is_running = False
                # обновляем кнопки при ошибке
                self.btn_start.config(state=tk.NORMAL)
                self.btn_stop.config(state=tk.DISABLED)
                self.btn_find.config(state=tk.NORMAL)

    def stop_clicking(self):
        # останавливаем авто-клик
        self.is_running = False
        self.status.config(text="Авто-клик остановлен")

        # обновляем кнопки
        if self.browser_launched:
            self.btn_start.config(state=tk.NORMAL)
            self.btn_stop.config(state=tk.DISABLED)
            self.btn_find.config(state=tk.NORMAL)

    def __del__(self):
        # закрываем при завершении
        self.is_running = False
        if self.driver:
            self.driver.quit()


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Обнавлятинатор @Mishenin")
    root.geometry("300x500")
    root.configure(bg="#1e8687")
    app = BrowserApp(root)

    # при закрытии окна
    def on_closing():
        app.stop_clicking()
        if app.driver:
            app.driver.quit()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
