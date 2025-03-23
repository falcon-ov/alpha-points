import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from datetime import datetime
import subprocess
import psutil
import random

# Словарь предметов: код предмета -> название
subjects = {
    "142": "CAMP: 3D Modeling in Tinkercad",
    "135": "CAMP: Robotics",
    "116": "CAMP: Start in IT",
    "134": "CAMP: TikTok blogging",
    "133": "CAMP: Web Design",
    "141": "Course June: Dance",
    "140": "Course june: English",
    "125": "Course June: General",
    "126": "Course June: Musical",
    "136": "COURSE:",
    "118": "COURSE: 3D Modeling in Tinkercad",
    "129": "COURSE: Advanced explorers",
    "121": "COURSE: AI and Chatbots",
    "120": "COURSE: AI and Machine Learning",
    "113": "COURSE: Animation and AI",
    "111": "COURSE: Arduino",
    "128": "COURSE: Build and Program",
    "119": "COURSE: Design and Animation",
    "110": "COURSE: Design and Branding",
    "106": "COURSE: Design and WebFactory",
    "108": "COURSE: Design Thinking",
    "66": "COURSE: Dream work",
    "127": "COURSE: First steps in Lego",
    "122": "COURSE: Games on Java - GreenFoot",
    "105": "COURSE: GDevelop",
    "107": "COURSE: Hello IT World",
    "124": "COURSE: Lego-Laboratory",
    "115": "COURSE: Marketing and Branding",
    "138": "COURSE: Microsoft Excel",
    "137": "COURSE: Microsoft Word",
    "109": "COURSE: Minecraft",
    "139": "COURSE: Power Point",
    "112": "COURSE: PyGame",
    "123": "COURSE: Start in IT",
    "104": "COURSE: Thunkable",
    "114": "COURSE: Web design and Promotion",
    "132": "OH-ALL-Временная мера",
    "130": "Online: Разработка Игр TEST",
    "131": "Open Hours Intern",
    "86": "CAMP: 2D Animation & Game Development",
    "92": "CAMP: Adventures in Roblox Studio",
    "94": "CAMP: Blender 3D",
    "100": "CAMP: Construct 3",
    "85": "CAMP: Content Creators",
    "98": "CAMP: Design and Animation",
    "93": "CAMP: Future Jobs in IT",
    "80": "CAMP: Games in GDevelop",
    "89": "CAMP: Hachathon",
    "91": "CAMP: Kodu Game Lab",
    "79": "CAMP: Lego Laboratory",
    "82": "CAMP: Motion-Design",
    "81": "CAMP: Programming in Minecraft",
    "99": "CAMP: Python on Codesters",
    "103": "CAMP: The Sandbox",
    "101": "CAMP: Pygame",
    "95": "CAMP: Tinkercad - Minecraft models",
    "96": "CAMP: Unity",
    "78": "CAMP: Unity & VR",
    "97": "CAMP: Unreal Engine",
    "84": "CAMP: 3D Modeling & VR",
    "88": "CAMP: Web Development",
    "6": "COURSE: 3D Modeling",
    "44": "COURSE: Blender 3D",
    "65": "COURSE: Backend",
    "75": "CAMP: YouTube",
    "71": "CAMP: Programming in Tynker",
    "23": "COURSE: VEXcode VR",
    "76": "CAMP: Desktop on C#",
    "72": "CAMP: Games in Roblox Studio",
    "69": "COURSE: Chat Bot",
    "3": "COURSE: Coding",
    "45": "COURSE: Java",
    "33": "COURSE: App Development",
    "15": "COURSE: Kodu Game Lab",
    "27": "COURSE: Construct 3",
    "25": "COURSE: Blockbench",
    "35": "COURSE: Game Design",
    "53": "COURSE: Data Base",
    "19": "COURSE: Web Factory",
    "68": "COURSE: Data Science",
    "13": "COURSE: Art&Design",
    "37": "COURSE: Web Factory (Tilda HTML CSS)",
    "63": "COURSE: Frontend",
    "21": "COURSE: Roblox",
    "17": "COURSE: Motion Design",
    "51": "COURSE: C#",
    "41": "COURSE: JavaScript",
    "31": "COURSE: WoofJS",
    "39": "COURSE: Python",
    "11": "COURSE: Visual Programming",
    "67": "COURSE: Games on Python",
    "2": "COURSE: Robotics",
    "47": "COURSE: QA Testing",
    "30": "CAMP: Thunkable"
}

class PointsAdderApp:
    def __init__(self, root):
        # Инициализация основного окна приложения
        self.root = root
        self.root.title("Добавление баллов ученикам")
        self.root.geometry("575x650")
        self.root.configure(bg="#e8ecef")  # Мягкий серо-голубой фон

        # Поле ввода URL группы
        tk.Label(root, text="URL страницы со списком учеников:\n(Ctrl+V работает на англ. раскладке)",
                 bg="#e8ecef", fg="#000000", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=10, pady=5,
                                                                                  sticky="w")
        self.url_entry = tk.Entry(root, width=50, bg="white", fg="#000000", relief="flat",
                                  highlightthickness=1, highlightbackground="#cbd5e0")
        self.url_entry.grid(row=0, column=1, padx=10, pady=5)

        # Поле ввода даты
        tk.Label(root, text="Дата (дд.мм.гггг, по умолчанию сегодня):",
                 bg="#e8ecef", fg="#000000", font=("Helvetica", 10, "bold")).grid(row=1, column=0, padx=10, pady=5,
                                                                                  sticky="w")
        self.date_entry = tk.Entry(root, bg="white", fg="#000000", relief="flat",
                                   highlightthickness=1, highlightbackground="#cbd5e0")
        self.date_entry.grid(row=1, column=1, padx=10, pady=5)

        # Выпадающий список предметов с текстом "(по умолчанию с сайта)"
        tk.Label(root, text="Предмет (по умолчанию с сайта):",
                 bg="#e8ecef", fg="#000000", font=("Helvetica", 10, "bold")).grid(row=2, column=0, padx=10, pady=5,
                                                                                  sticky="w")
        self.subject_combobox = ttk.Combobox(root, values=list(subjects.values()), width=47)
        self.subject_combobox.grid(row=2, column=1, padx=10, pady=5)
        self.subject_combobox.set("Выберите предмет")
        self.subject_combobox.configure(foreground="#4a5568")

        # Кнопки
        button_frame = tk.Frame(root, bg="#e8ecef")
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)

        self.alpha_button = tk.Button(button_frame, text="Alpha", command=self.open_alpha, width=10,
                                      bg="#f7fafc", fg="#000000", relief="flat", font=("Helvetica", 10, "bold"),
                                      activebackground="#e2e8f0")
        self.alpha_button.pack(side="left", padx=5)

        self.start_button = tk.Button(button_frame, text="Запустить", command=self.start_process, state="disabled",
                                      bg="#f7fafc", fg="#000000", relief="flat", font=("Helvetica", 10, "bold"),
                                      activebackground="#e2e8f0")
        self.start_button.pack(side="left", padx=5)

        self.get_data_button = tk.Button(button_frame, text="Получить данные", command=self.get_students_data,
                                         bg="#f7fafc", fg="#000000", relief="flat", font=("Helvetica", 10, "bold"),
                                         activebackground="#e2e8f0")
        self.get_data_button.pack(side="right", padx=5)

        # Текст инструкции для пользователя
        instructions_text = (
            "Как пользоваться программой:\n"
            "1. Нажмите кнопку 'Alpha', чтобы открыть Microsoft Edge и перейти на страницу календаря Alpha.\n"
            "2. Введите URL страницы с группой студентов.\n"
            "3. Укажите дату в формате дд.мм.гггг (или оставьте пустым).\n"
            "4. Выберите предмет (по умолчанию берется с сайта).\n"
            "5. Нажмите 'Получить данные' для загрузки списка студентов.\n"
            "6. Отметьте студентов, укажите баллы и примечания.\n"
            "7. Нажмите 'Запустить' для добавления баллов."
        )
        instructions_label = tk.Label(root, text=instructions_text, justify=tk.LEFT, wraplength=550,
                                      bg="#e8ecef", fg="#000000", font=("Helvetica", 9))
        instructions_label.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

        # Область для списка студентов с прокруткой
        self.students_frame = tk.Frame(root, bg="white", bd=1, relief="solid")
        self.students_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.canvas = tk.Canvas(self.students_frame, bg="white")
        self.scrollbar = tk.Scrollbar(self.students_frame, orient="vertical", command=self.canvas.yview,
                                      bg="#e2e8f0", troughcolor="#edf2f7")
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Словари для хранения данных студентов
        self.student_vars = {}
        self.student_urls = {}
        self.student_points = {}
        self.student_comments = {}

        # Настройка веса строк и столбцов для адаптивности
        root.grid_rowconfigure(4, weight=1)
        root.grid_columnconfigure(1, weight=1)

    def get_subject_code(self, subject_name):
        for code, name in subjects.items():
            if name == subject_name:
                return code
        return None

    def close_edge_processes(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'msedge.exe':
                try:
                    proc.terminate()
                    proc.wait(timeout=3)
                except psutil.TimeoutExpired:
                    proc.kill()
                except:
                    pass

    def open_alpha(self):
        edge_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
        try:
            self.close_edge_processes()
            subprocess.Popen([edge_path, "--remote-debugging-port=9222", "https://impactacademies.s20.online/teacher/1/calendar/index"])
            messagebox.showinfo("Успех", "Edge открыт на странице Alpha. Войдите в аккаунт, если нужно.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть Edge: {str(e)}")

    def ensure_edge_running(self, url):
        edge_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
        edge_running = False
        for proc in psutil.process_iter(['name', 'cmdline']):
            if proc.info['name'] == 'msedge.exe' and '--remote-debugging-port=9222' in (proc.info['cmdline'] or []):
                edge_running = True
                break

        if not edge_running:
            try:
                self.close_edge_processes()
                subprocess.Popen([edge_path, "--remote-debugging-port=9222", url])
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось запустить Edge автоматически: {str(e)}")
                return False
        return True

    def get_students_data(self):
        url = self.url_entry.get()
        if not url:
            messagebox.showerror("Ошибка", "Введите URL страницы со списком учеников!")
            return

        if not self.ensure_edge_running(url):
            return

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.student_vars.clear()
        self.student_urls.clear()
        self.student_points.clear()
        self.student_comments.clear()

        edge_options = Options()
        edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Edge(service=webdriver.edge.service.Service(EdgeChromiumDriverManager().install()), options=edge_options)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к Edge: {str(e)}")
            return

        try:
            driver.get(url)
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "list-unstyled"))
            )

            # Попытка найти название курса
            try:
                course_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'col-xs-12 text-muted')]/small[contains(., 'COURSE:')]"))
                )
                course_text = course_element.text.strip()
                course_name = course_text.split("COURSE:")[1].strip()
                full_course_name = f"COURSE: {course_name}"
                if full_course_name in subjects.values():
                    self.subject_combobox.set(full_course_name)
                else:
                    self.subject_combobox.set("Выберите предмет")
                    messagebox.showwarning("Предупреждение", f"Курс '{full_course_name}' не найден в списке.")
            except Exception as e:
                self.subject_combobox.set("Выберите предмет")
                messagebox.showwarning("Предупреждение", f"Не удалось определить курс: {str(e)}")

            students = driver.find_elements(By.XPATH,
                                           "//ul[@class='list-unstyled m-t-xs m-b-xs']/li[not(contains(@class, 'archive'))]/a[@title='Открыть карточку']")
            for i, student in enumerate(students):
                student_name = student.text.strip()
                student_url = student.get_attribute("href")
                if student_name and student_url:
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(self.scrollable_frame, text=student_name, variable=var,
                                        bg="white", fg="#2d3748", selectcolor="#e2e8f0")
                    cb.grid(row=i, column=0, sticky="w", padx=5, pady=2)

                    points_default = random.choices(["4", "5"], weights=[3, 1], k=1)[0]
                    points_entry = tk.Entry(self.scrollable_frame, width=5, bg="white", fg="#2d3748",
                                            relief="flat", highlightthickness=1, highlightbackground="#cbd5e0")
                    points_entry.insert(0, points_default)
                    points_entry.grid(row=i, column=1, padx=5, pady=2)

                    comment_entry = tk.Entry(self.scrollable_frame, width=30, bg="white", fg="#2d3748",
                                             relief="flat", highlightthickness=1, highlightbackground="#cbd5e0")
                    comment_entry.grid(row=i, column=2, padx=5, pady=2)

                    self.student_vars[student_name] = var
                    self.student_urls[student_name] = student_url
                    self.student_points[student_name] = points_entry
                    self.student_comments[student_name] = comment_entry

            if not self.student_vars:
                messagebox.showwarning("Предупреждение", "Студенты не найдены на указанной странице.")
            else:
                self.start_button.config(state="normal")  # Активация кнопки "Запустить"
                messagebox.showinfo("Успех", "Список студентов загружен.")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")
        finally:
            pass

    def start_process(self):
        url = self.url_entry.get()
        date = self.date_entry.get() or datetime.now().strftime("%d.%m.%Y")
        subject_name = self.subject_combobox.get()
        subject = self.get_subject_code(subject_name)

        if not url or not subject or subject_name == "Выберите предмет":
            messagebox.showerror("Ошибка", "Введите URL и выберите предмет!")
            return

        selected_students = [name for name, var in self.student_vars.items() if var.get()]
        if not selected_students:
            messagebox.showerror("Ошибка", "Выберите хотя бы одного студента!")
            return

        self.close_edge_processes()
        edge_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
        try:
            subprocess.Popen([edge_path, "--remote-debugging-port=9222", url])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось запустить Edge: {str(e)}")
            return

        try:
            self.add_points(url, date, subject, selected_students)
            messagebox.showinfo("Успех", "Баллы успешно добавлены!")
            self.start_button.config(state="disabled")  # Блокировка кнопки "Запустить" после выполнения
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

    def add_points(self, url, date, subject, selected_students):
        edge_options = Options()
        edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Edge(service=webdriver.edge.service.Service(EdgeChromiumDriverManager().install()), options=edge_options)

        try:
            driver.get(url)
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "list-unstyled"))
            )

            for student_name in selected_students:
                student_url = self.student_urls.get(student_name)
                points = self.student_points[student_name].get()
                comment = self.student_comments[student_name].get()
                if not student_url:
                    continue

                driver.get(student_url)
                results_tab = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@href='#testresults']"))
                )
                results_tab.click()

                try:
                    add_result_button = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn btn-sm btn-w-m btn-white crm-modal-btn m-t-sm') and contains(., 'Результат')]"))
                    )
                    add_result_button.click()
                except:
                    add_button = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'crm-dashed-link') and text()='Добавить']"))
                    )
                    add_button.click()

                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "testresult-date"))
                )

                date_field = driver.find_element(By.ID, "testresult-date")
                date_field.clear()
                date_field.send_keys(date)

                subject_select = driver.find_element(By.ID, "testresult-subject_id")
                driver.execute_script(f"document.getElementById('testresult-subject_id').value = '{subject}';")
                driver.execute_script("$('#testresult-subject_id').trigger('chosen:updated');")

                grade_field = driver.find_element(By.ID, "testresult-grade")
                grade_field.clear()
                grade_field.send_keys(points)

                if comment:
                    comment_field = driver.find_element(By.ID, "testresult-comment")
                    comment_field.clear()
                    comment_field.send_keys(comment)

                save_button = driver.find_element(By.XPATH, "//button[@type='submit']")
                save_button.click()

                WebDriverWait(driver, 2).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, "modal-body"))
                )

        finally:
            pass

if __name__ == "__main__":
    root = tk.Tk()
    app = PointsAdderApp(root)
    root.mainloop()