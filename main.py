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

        # Цвета
        self.bg_color = "#1e8687"  # Основной фон
        self.button_color = "#1a7576"  # Темнее основного фона
        self.accent_color = "#166667"  # Еще темнее для акцентов
        self.text_color = "#ffffff"  # Белый текст
        self.entry_bg = "#2a9d9e"  # Фон полей ввода
        self.entry_fg = "#ffffff"  # Текст в полях ввода

        #сохранение настроек в джисон
        self.settings_file = "browser_settings.json"
        self.settings = self.load_settings()
        
        #см
        self.target_url = "https://hpsm.emias.mos.ru/sm/index.do"

        self.configure_styles()
        self.create_widgets()

    def configure_styles(self):
        """Настройка стилей виджетов"""
        # для кнопок
        self.button_style = {
            "bg": self.button_color,
            "fg": self.text_color,
            "font": ("Arial", 9),
            "relief": "flat",
            "bd": 0,
            "padx": 15,
            "pady": 8,
            "activebackground": self.accent_color,
            "activeforeground": self.text_color
        }

        #для полей ввода
        self.entry_style = {
            "bg": self.entry_bg,
            "fg": self.entry_fg,
            "font": ("Arial", 8),
            "relief": "flat",
            "bd": 2,
            "highlightbackground": self.accent_color,
            "highlightcolor": self.accent_color,
            "highlightthickness": 1,
            "insertbackground": self.text_color
        }

        # для меток
        self.label_style = {
            "bg": self.bg_color,
            "fg": self.text_color,
            "font": ("Arial", 8)
        }

        # для заголовков
        self.title_style = {
            "bg": self.bg_color,
            "fg": self.text_color,
            "font": ("Arial", 8, "bold")
        }

    def create_widgets(self):
        """Создание всех элементов интерфейса"""
        #фон основного окна
        self.root.configure(bg=self.bg_color)

        # шапка (кринж)
        #title_label = tk.Label(
        #    self.root, 
        #    text="Обнавлятинатор by Mishenin", 
        #    **{**self.title_style, "font": ("Arial", 10, "bold")}
        #)
        #title_label.pack(pady=(10, 5))

        #разделитель
        #self.create_separator()

        # ыбор браузера
        tk.Label(self.root, text="Выберите браузер:", **self.title_style).pack(pady=(10, 5))
        
        # фрейм для кнопок браузеров
        browser_frame = tk.Frame(self.root, bg=self.bg_color)
        browser_frame.pack(pady=5)
        
        self.btn_yandex = tk.Button(
            browser_frame, 
            text="Яндекс.Браузер", 
            command=lambda: self.select_browser("yandex"),
            **self.button_style,
            width=15
        )
        self.btn_yandex.pack(side=tk.LEFT, padx=5)
        
        self.btn_chrome = tk.Button(
            browser_frame, 
            text="Chrome", 
            command=lambda: self.select_browser("chrome"),
            **self.button_style,
            width=15
        )
        self.btn_chrome.pack(side=tk.LEFT, padx=5)

        #разделитель
        self.create_separator()

        # Путь к браузеру
        tk.Label(self.root, text="Путь к браузеру:", **self.title_style).pack(pady=(10, 5))
        
        path_frame = tk.Frame(self.root, bg=self.bg_color)
        path_frame.pack(pady=5, fill=tk.X, padx=20)

        self.browser_path_var = tk.StringVar(value=self.settings.get("browser_path", ""))
        self.browser_entry = tk.Entry(
            path_frame, 
            textvariable=self.browser_path_var, 
            **self.entry_style,
            width=35
        )
        self.browser_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_browse = tk.Button(
            path_frame, 
            text="Обзор", 
            command=self.browse_browser_path,
            **self.button_style
        )
        self.btn_browse.pack(side=tk.RIGHT, padx=(5, 0))

        # Автопоиск браузера
        self.btn_auto_find = tk.Button(
            self.root, 
            text="Авто-поиск браузера", 
            command=self.auto_find_browser,
            **self.button_style
        )
        self.btn_auto_find.pack(pady=5)

        #разделитель
        self.create_separator()

        # запуск\закрытие браузера
        control_frame = tk.Frame(self.root, bg=self.bg_color)
        control_frame.pack(pady=10)

        self.btn_launch = tk.Button(
            control_frame, 
            text="Запустить браузер", 
            command=self.launch_browser,
            **self.button_style
        )
        self.btn_launch.pack(side=tk.LEFT, padx=5)

        self.btn_close_browser = tk.Button(
            control_frame, 
            text="Закрыть браузер", 
            command=self.close_browser, 
            state=tk.DISABLED,
            **self.button_style
        )
        self.btn_close_browser.pack(side=tk.LEFT, padx=5)

        #разделитель
        self.create_separator()

        # поиск кнопки
        tk.Label(
            self.root, 
            text="CSS-селектор (если автоматически не находит):", 
            **self.label_style
        ).pack(pady=(10, 5))
        
        self.selector_var = tk.StringVar(value=self.settings.get("selector", "[aria-label='Обновить']"))
        self.selector_entry = tk.Entry(
            self.root, 
            textvariable=self.selector_var, 
            **self.entry_style,
            width=40
        )
        self.selector_entry.pack(pady=5)

        self.btn_find = tk.Button(
            self.root, 
            text="Найти кнопку", 
            command=self.find_button, 
            state=tk.DISABLED,
            **self.button_style
        )
        self.btn_find.pack(pady=5)

        #разделитель
        self.create_separator()

        # интервал и управление авто-кликом
        interval_frame = tk.Frame(self.root, bg=self.bg_color)
        interval_frame.pack(pady=5)

        tk.Label(interval_frame, text="Интервал (секунды):", **self.label_style).pack(side=tk.LEFT)
        
        self.interval_var = tk.StringVar(value="5")
        self.interval_entry = tk.Entry(
            interval_frame, 
            textvariable=self.interval_var, 
            **self.entry_style,
            width=10
        )
        self.interval_entry.pack(side=tk.LEFT, padx=5)

        control_frame = tk.Frame(self.root, bg=self.bg_color)
        control_frame.pack(pady=10)

        self.btn_start = tk.Button(
            control_frame, 
            text="Запустить авто-клик", 
            command=self.start_clicking, 
            state=tk.DISABLED,
            **self.button_style
        )
        self.btn_start.pack(side=tk.LEFT, padx=5)

        self.btn_stop = tk.Button(
            control_frame, 
            text="Остановить", 
            command=self.stop_clicking, 
            state=tk.DISABLED,
            **self.button_style
        )
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        #разделитель
        self.create_separator()

        # Статус
        self.status = tk.Label(
            self.root, 
            text="Браузер не запущен", 
            font=("Arial", 8, "bold"), 
            fg="#ff6b6b", 
            bg=self.bg_color
        )
        self.status.pack(pady=10)

    def create_separator(self):
        """Создание разделительной линии"""
        separator = tk.Frame(self.root, height=2, bg=self.accent_color)
        separator.pack(fill=tk.X, padx=20, pady=5)

    def select_browser(self, browser_type):
        """Выбор браузера и автоматический поиск пути"""
        self.selected_browser = browser_type
        
        # Сброс стилей кнопок
        self.btn_yandex.config(bg=self.button_color)
        self.btn_chrome.config(bg=self.button_color)
        
        # изменение выбранной кнопки
        if browser_type == "yandex":
            self.btn_yandex.config(bg=self.accent_color)
            self.auto_find_yandex()
        elif browser_type == "chrome":
            self.btn_chrome.config(bg=self.accent_color)
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
                self.status.config(text="Яндекс.Браузер найден!", fg="#51e3a1")  
                self.save_settings()
                return

        self.status.config(text="Яндекс.Браузер не найден автоматически", fg="#ff6b6b")  

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
                self.status.config(text="Google Chrome найден!", fg="#51e3a1") 
                self.save_settings()
                return

        self.status.config(text="Google Chrome не найден автоматически", fg="#ff6b6b")  

    def auto_find_browser(self):
        """Автопоиск браузера в зависимости от выбранного типа"""
        if self.selected_browser == "yandex":
            self.auto_find_yandex()
        elif self.selected_browser == "chrome":
            self.auto_find_chrome()
        else:
            self.status.config(text="Сначала выберите тип браузера!", fg="#ffd166")  

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
            self.status.config(text="Браузер уже запущен!", fg="#ffd166")  
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
                    self.status.config(text="Браузер не найден. Укажите путь вручную.", fg="#ff6b6b")
                    return

            # Определение драйвера
            driver_path = self.get_driver_path()
            if not driver_path:
                return

            # Запуск драйвера
            service = ChromeService(executable_path=driver_path)
            self.driver = webdriver.Chrome(service=service, options=options)

            # Открытие сайта
            self.driver.get(self.target_url)
            self.save_settings()

            # Обновление интерфейса
            self.browser_launched = True
            self.status.config(text="Браузер запущен", fg="#51e3a1")  # Зеленый
            self.update_buttons_state()

        except Exception as e:
            self.status.config(text=f"Ошибка: {str(e)}", fg="#ff6b6b")

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
                fg="#ff6b6b",
            )
            return None
        return driver_path

    def update_buttons_state(self):
        """Обновление состояния кнопок"""
        if self.browser_launched:
            self.btn_launch.config(state=tk.DISABLED, bg=self.accent_color)
            self.btn_close_browser.config(state=tk.NORMAL, bg=self.button_color)
            self.btn_find.config(state=tk.NORMAL, bg=self.button_color)
            self.btn_start.config(state=tk.DISABLED if not hasattr(self, 'successful_selector') else tk.NORMAL, 
                                bg=self.button_color if hasattr(self, 'successful_selector') else self.accent_color)
        else:
            self.btn_launch.config(state=tk.NORMAL, bg=self.button_color)
            self.btn_close_browser.config(state=tk.DISABLED, bg=self.accent_color)
            self.btn_find.config(state=tk.DISABLED, bg=self.accent_color)
            self.btn_start.config(state=tk.DISABLED, bg=self.accent_color)
            self.btn_stop.config(state=tk.DISABLED, bg=self.accent_color)

    def close_browser(self):
        """Закрытие браузера"""
        self.stop_clicking()
        
        if self.driver:
            self.driver.quit()
            self.driver = None
        
        self.browser_launched = False
        if hasattr(self, "successful_selector"):
            del self.successful_selector
            
        self.status.config(text="Браузер закрыт", fg="#51e3a1")
        self.update_buttons_state()

    def find_button(self):
        """Поиск кнопки на странице"""
        if not self.driver:
            self.status.config(text="Сначала запустите браузер!", fg="#ffd166")
            return

        try:
            self.status.config(text="Ищу кнопку...", fg="#ffd166")
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
                self.status.config(text=f"Кнопка найдена! Селектор: {found_selector}", fg="#51e3a1")
                self.save_settings()
                self.btn_start.config(state=tk.NORMAL, bg=self.button_color)
            else:
                self.status.config(text="Кнопка не найдена! Попробуйте другой селектор", fg="#ff6b6b")

        except Exception as e:
            self.status.config(text=f"Ошибка: {str(e)}", fg="#ff6b6b")

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

        # ввод селектора
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
            self.status.config(text="Сначала найдите кнопку!", fg="#ff6b6b")
            return

        try:
            interval = int(self.interval_var.get())
            if interval <= 0:
                self.status.config(text="Интервал должен быть положительным числом!", fg="#ff6b6b")
                return
        except ValueError:
            self.status.config(text="Введите корректный интервал (число)!", fg="#ff6b6b")
            return

        self.is_running = True
        self.status.config(text="Авто-клик запущен", fg="#51e3a1")
        self.update_control_buttons()

        threading.Thread(target=self.click_loop, daemon=True).start()

    def update_control_buttons(self):
        """Обновление кнопок управления"""
        if self.is_running:
            self.btn_start.config(state=tk.DISABLED, bg=self.accent_color)
            self.btn_stop.config(state=tk.NORMAL, bg=self.button_color)
            self.btn_find.config(state=tk.DISABLED, bg=self.accent_color)
        else:
            self.btn_start.config(state=tk.NORMAL, bg=self.button_color)
            self.btn_stop.config(state=tk.DISABLED, bg=self.accent_color)
            self.btn_find.config(state=tk.NORMAL, bg=self.button_color)

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
                    self.status.config(text="Не удалось найти кнопку для клика", fg="#ff6b6b")
                    self.stop_clicking()
                    break

                # Ожидание
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
        self.status.config(text="Авто-клик остановлен", fg="#ffd166")
        self.update_control_buttons()


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Обнавлятинатор by Mishenin")
    root.geometry("400x600")
    root.configure(bg="#1e8687")
    
    # Не забыть подобрать иконку
    try:
        root.iconbitmap("icon.ico") 
    except:
        pass
    
    app = BrowserApp(root)
    
    def on_closing():
        app.stop_clicking()
        if app.driver:
            app.driver.quit()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()