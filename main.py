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
        self.selected_browser = None

        # Настройки
        self.settings_file = "browser_settings.json"
        self.settings = self.load_settings()
        
        # Фиксированный URL
        self.target_url = "https://hpsm.emias.mos.ru/sm/index.do"

        self.create_widgets()

    def create_widgets(self):
        """Создание всех элементов интерфейса"""
        # Выбор браузера
        tk.Label(self.root, text="Выберите браузер:").pack(pady=5)
        
        # Фрейм для кнопок браузеров
        browser_frame = tk.Frame(self.root)
        browser_frame.pack(pady=5)
        
        self.btn_yandex = tk.Button(
            browser_frame, 
            text="Яндекс.Браузер", 
            command=lambda: self.select_browser("yandex"),
            width=15
        )
        self.btn_yandex.pack(side=tk.LEFT, padx=5)
        
        self.btn_chrome = tk.Button(
            browser_frame, 
            text="Chrome", 
            command=lambda: self.select_browser("chrome"),
            width=15
        )
        self.btn_chrome.pack(side=tk.LEFT, padx=5)

        # Поле для пути к браузеру
        tk.Label(self.root, text="Путь к браузеру:").pack(pady=5)
        
        path_frame = tk.Frame(self.root)
        path_frame.pack(pady=5, fill=tk.X, padx=10)

        self.browser_path_var = tk.StringVar(value=self.settings.get("browser_path", ""))
        self.browser_entry = tk.Entry(path_frame, textvariable=self.browser_path_var, width=35)
        self.browser_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_browse = tk.Button(path_frame, text="Обзор", command=self.browse_browser_path)
        self.btn_browse.pack(side=tk.RIGHT, padx=(5, 0))

        # Автопоиск браузера
        self.btn_auto_find = tk.Button(self.root, text="Авто-поиск браузера", command=self.auto_find_browser)
        self.btn_auto_find.pack(pady=5)

        # Управление браузером
        self.btn_launch = tk.Button(self.root, text="Запустить браузер", command=self.launch_browser)
        self.btn_launch.pack(pady=5)

        self.btn_close_browser = tk.Button(
            self.root, text="Закрыть браузер", command=self.close_browser, state=tk.DISABLED
        )
        self.btn_close_browser.pack(pady=5)

        # Селектор кнопки
        tk.Label(self.root, text="CSS-селектор (если автоматически не находит):").pack()
        self.selector_var = tk.StringVar(value=self.settings.get("selector", "[aria-label='Обновить']"))
        self.selector_entry = tk.Entry(self.root, textvariable=self.selector_var, width=40)
        self.selector_entry.pack(pady=5)

        self.btn_find = tk.Button(self.root, text="Найти кнопку", command=self.find_button, state=tk.DISABLED)
        self.btn_find.pack(pady=5)

        # Интервал и управление авто-кликом
        tk.Label(self.root, text="Интервал (секунды):").pack()
        self.interval_var = tk.StringVar(value="5")
        self.interval_entry = tk.Entry(self.root, textvariable=self.interval_var, width=10)
        self.interval_entry.pack(pady=5)

        control_frame = tk.Frame(self.root)
        control_frame.pack(pady=5)

        self.btn_start = tk.Button(
            control_frame, text="Запустить авто-клик", command=self.start_clicking, state=tk.DISABLED
        )
        self.btn_start.pack(side=tk.LEFT, padx=5)

        self.btn_stop = tk.Button(
            control_frame, text="Остановить", command=self.stop_clicking, state=tk.DISABLED
        )
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        # Статус
        self.status = tk.Label(self.root, text="Браузер не запущен", font=("Arial", 10, "bold"), fg="red")
        self.status.pack(pady=5)

    def select_browser(self, browser_type):
        """Выбор браузера и автоматический поиск пути"""
        self.selected_browser = browser_type
        
        if browser_type == "yandex":
            self.auto_find_yandex()
        elif browser_type == "chrome":
            self.auto_find_chrome()

    def auto_find_yandex(self):
        """Автопоиск Яндекс.Браузера"""
        possible_paths = [
            "C:/Users/User/AppData/Local/Yandex/YandexBrowser/Application/browser.exe",
            "C:/Program Files (x86)/Yandex/YandexBrowser/Application/browser.exe",
            "C:/Program Files/Yandex/YandexBrowser/Application/browser.exe",
            os.path.expanduser("~/AppData/Local/Yandex/YandexBrowser/Application/browser.exe"),
        ]

        for path in possible_paths:
            if os.path.exists(path):
                self.browser_path_var.set(path)
                self.status.config(text="Яндекс.Браузер найден!", fg="green")
                self.save_settings()
                return

        self.status.config(text="Яндекс.Браузер не найден автоматически", fg="red")

    def auto_find_chrome(self):
        """Автопоиск Google Chrome"""
        possible_paths = [
            "C:/Program Files/Google/Chrome/Application/chrome.exe",
            "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
            os.path.expanduser("~/AppData/Local/Google/Chrome/Application/chrome.exe"),
        ]

        for path in possible_paths:
            if os.path.exists(path):
                self.browser_path_var.set(path)
                self.status.config(text="Google Chrome найден!", fg="green")
                self.save_settings()
                return

        self.status.config(text="Google Chrome не найден автоматически", fg="red")

    def auto_find_browser(self):
        """Автопоиск браузера в зависимости от выбранного типа"""
        if self.selected_browser == "yandex":
            self.auto_find_yandex()
        elif self.selected_browser == "chrome":
            self.auto_find_chrome()
        else:
            self.status.config(text="Сначала выберите тип браузера!", fg="orange")

    def load_settings(self):
        """Загрузка настроек из файла"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
        return {}

    def save_settings(self):
        """Сохранение настроек в файл"""
        try:
            settings = {
                "browser_path": self.browser_path_var.get(),
                "selector": self.selector_var.get(),
                "selected_browser": self.selected_browser
            }
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")

    def browse_browser_path(self):
        """Выбор файла браузера через диалог"""
        file_path = filedialog.askopenfilename(
            title="Выберите исполняемый файл браузера",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")],
        )
        if file_path:
            self.browser_path_var.set(file_path)
            self.save_settings()

    def launch_browser(self):
        """Запуск браузера"""
        if self.browser_launched and self.driver:
            self.status.config(text="Браузер уже запущен!", fg="orange")
            return

        try:
            options = webdriver.ChromeOptions()
            browser_path = self.browser_path_var.get().strip()

            # Проверка пути к браузеру
            if browser_path and os.path.exists(browser_path):
                options.binary_location = browser_path
            else:
                self.auto_find_browser()
                browser_path = self.browser_path_var.get().strip()
                
                if not browser_path or not os.path.exists(browser_path):
                    self.status.config(text="Браузер не найден. Укажите путь вручную.", fg="red")
                    return

            # Определение драйвера
            driver_path = self.get_driver_path()
            if not driver_path:
                return

            # Запуск драйвера
            service = ChromeService(executable_path=driver_path)
            self.driver = webdriver.Chrome(service=service, options=options)

            # Открытие целевого URL
            self.driver.get(self.target_url)
            self.save_settings()

            # Обновление интерфейса
            self.browser_launched = True
            self.status.config(text="Браузер запущен", fg="green")
            self.update_buttons_state()

        except Exception as e:
            self.status.config(text=f"Ошибка: {str(e)}", fg="red")

    def get_driver_path(self):
        """Получение пути к драйверу"""
        if self.selected_browser == "yandex":
            driver_path = "./yandexdriver.exe"
        elif self.selected_browser == "chrome":
            driver_path = "./chromedriver.exe"
        else:
            driver_path = "./yandexdriver.exe"

        if not os.path.exists(driver_path):
            self.status.config(
                text=f"Драйвер не найден! Убедитесь, что {driver_path} в папке со скриптом.",
                fg="red",
            )
            return None
        return driver_path

    def update_buttons_state(self):
        """Обновление состояния кнопок"""
        if self.browser_launched:
            self.btn_launch.config(state=tk.DISABLED)
            self.btn_close_browser.config(state=tk.NORMAL)
            self.btn_find.config(state=tk.NORMAL)
            self.btn_start.config(state=tk.DISABLED if not hasattr(self, 'successful_selector') else tk.NORMAL)
        else:
            self.btn_launch.config(state=tk.NORMAL)
            self.btn_close_browser.config(state=tk.DISABLED)
            self.btn_find.config(state=tk.DISABLED)
            self.btn_start.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.DISABLED)

    def close_browser(self):
        """Закрытие браузера"""
        self.stop_clicking()
        
        if self.driver:
            self.driver.quit()
            self.driver = None
        
        self.browser_launched = False
        if hasattr(self, "successful_selector"):
            del self.successful_selector
            
        self.status.config(text="Браузер закрыт", fg="red")
        self.update_buttons_state()

    def find_button(self):
        """Поиск кнопки на странице"""
        if not self.driver:
            self.status.config(text="Сначала запустите браузер!", fg="orange")
            return

        try:
            self.status.config(text="Ищу кнопку...", fg="orange")
            wait = WebDriverWait(self.driver, 10)

            # Список селекторов для поиска
            selectors = self.get_selectors_list()
            found_selector = None

            for selector in selectors:
                try:
                    element = self.find_element_by_selector(wait, selector)
                    if element:
                        found_selector = selector
                        self.highlight_element(element)
                        break
                except Exception:
                    continue

            if found_selector:
                self.successful_selector = found_selector
                self.status.config(text=f"Кнопка найдена! Селектор: {found_selector}", fg="green")
                self.save_settings()
                self.btn_start.config(state=tk.NORMAL)
            else:
                self.status.config(text="Кнопка не найдена! Попробуйте другой селектор", fg="red")

        except Exception as e:
            self.status.config(text=f"Ошибка: {str(e)}", fg="red")

    def get_selectors_list(self):
        """Получение списка селекторов для поиска"""
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

        # Добавление пользовательского селектора
        custom_selector = self.selector_var.get().strip()
        if custom_selector and custom_selector not in selectors:
            selectors.insert(0, custom_selector)

        return selectors

    def find_element_by_selector(self, wait, selector):
        """Поиск элемента по селектору"""
        if not selector.startswith("//"):
            return wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
        else:
            return wait.until(EC.element_to_be_clickable((By.XPATH, selector)))

    def highlight_element(self, element):
        """Подсветка элемента"""
        self.driver.execute_script("arguments[0].style.border='3px solid red'", element)

    def find_and_click_button(self):
        """Поиск и клик по кнопке"""
        if not hasattr(self, "successful_selector") or not self.successful_selector:
            return False

        try:
            wait = WebDriverWait(self.driver, 5)
            selector = self.successful_selector

            element = self.find_element_by_selector(wait, selector)
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            element.click()
            return True

        except StaleElementReferenceException:
            return self.find_and_click_button()
        except (NoSuchElementException, TimeoutException, Exception):
            return False

    def start_clicking(self):
        """Запуск авто-клика"""
        if not hasattr(self, "successful_selector"):
            self.status.config(text="Сначала найдите кнопку!", fg="red")
            return

        try:
            interval = int(self.interval_var.get())
            if interval <= 0:
                self.status.config(text="Интервал должен быть положительным числом!", fg="red")
                return
        except ValueError:
            self.status.config(text="Введите корректный интервал (число)!", fg="red")
            return

        self.is_running = True
        self.status.config(text="Авто-клик запущен", fg="green")
        self.update_control_buttons()

        threading.Thread(target=self.click_loop, daemon=True).start()

    def update_control_buttons(self):
        """Обновление кнопок управления"""
        if self.is_running:
            self.btn_start.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.NORMAL)
            self.btn_find.config(state=tk.DISABLED)
        else:
            self.btn_start.config(state=tk.NORMAL)
            self.btn_stop.config(state=tk.DISABLED)
            self.btn_find.config(state=tk.NORMAL)

    def click_loop(self):
        """Цикл авто-кликов"""
        interval = int(self.interval_var.get())
        click_count = 0

        while self.is_running and self.driver and self.browser_launched:
            try:
                if self.find_and_click_button():
                    click_count += 1
                    current_time = time.strftime("%H:%M:%S")
                    status_text = f"Клик #{click_count} выполнен в {current_time}"
                    self.status.config(text=status_text)
                else:
                    self.status.config(text="Не удалось найти кнопку для клика", fg="red")
                    self.stop_clicking()
                    break

                # Ожидание с проверкой остановки
                for _ in range(interval):
                    if not self.is_running:
                        break
                    time.sleep(1)

            except Exception as e:
                self.status.config(text=f"Ошибка клика: {str(e)}")
                self.stop_clicking()
                break

    def stop_clicking(self):
        """Остановка авто-клика"""
        self.is_running = False
        self.status.config(text="Авто-клик остановлен", fg="orange")
        self.update_control_buttons()


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Обнавлятинатор @Mishenin")
    root.geometry("300x450")
    root.configure(bg="#1e8687")
    app = BrowserApp(root)
    
    def on_closing():
        app.stop_clicking()
        if app.driver:
            app.driver.quit()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()