import json
import os
import sys

import flet as ft

from datetime import datetime, timedelta

from src.docx import replace_placeholders


class OfferGeneratorApp:
    _name = ft.TextField(label="Имя")
    _title = ft.TextField(label="Название")
    _patronymic = ft.TextField(label="Отчество")
    _surname = ft.TextField(label="Фамилия")
    _salary = ft.TextField(label="Зарплата", input_filter=ft.NumbersOnlyInputFilter(), suffix_text="руб.")
    _position = ft.TextField(label="Должность")
    _email = ft.TextField(label="Email")
    _phone = ft.TextField(label="Телефон")

    def __init__(self, page: ft.Page):
        self.__positions = [self._title, self._salary]
        self.__employee = [self._name, self._patronymic, self._surname, self._position, self._email, self._phone]
        self.__company = [self._title]
        self.__departments = [self._title]
        self.get_resource_path()
        self.data_path = str(os.path.join("data", "data.json"))
        self.page = page
        self.dialog = None
        self.setup_page()
        self.init_data()
        self.create_controls()
        self.__offer_fields = [
            self.position_salary,
            self.contact_person_dropdown["dropdown"],
            self.signing_person_dropdown["dropdown"],
            self.company_dropdown["dropdown"],
            self.position_dropdown["dropdown"],
            self.accept_days_dropdown,
            self.contract_term_dropdown,
            self.name_field,
            self.department_dropdown["dropdown"],
            self.responsibilities
        ]
        self.setup_ui()

    def get_resource_path(self):
        if hasattr(sys, '_MEIPASS'):
            self.res_path = str(sys._MEIPASS)
        else:
            self.res_path = str(os.path.abspath("."))

    def setup_page(self):
        self.page.title = "Генератор оффера"
        self.page.vertical_alignment = ft.MainAxisAlignment.CENTER
        self.page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.page.padding = 20
        self.page.scroll = True

    def init_data(self):
        with open(self.data_path, "r", encoding="utf-8") as f:
            self.__data = json.load(f)
        self.positions = self.__data["positions"]
        self.positions.sort(key=lambda dct: dct["salary"])
        self.departments = []
        self.companies = self.__data["companies"]
        self.companies.sort(key=lambda dct: dct["title"])
        self.contact_persons = self.__data["employees"]
        self.contact_persons.sort(key=lambda dct: dct["surname"])
        self.signing_persons = self.__data["employees"]
        self.signing_persons.sort(key=lambda dct: dct["surname"])
        self.contract_terms = self.__data["contract_terms"]
        self.accept_days = self.__data["accept_days"]

    def __reset_settings(self):
        path = str(os.path.join("data", "settings_tmplt.json"))
        with open(path, "r", encoding="utf-8") as f:
            self.__data = json.load(f)
        with open(self.data_path, "w+", encoding="utf-8") as f:
            data = json.dumps(self.__data, ensure_ascii=False, indent=2)
            f.write(data)

        self.page.clean()
        self.setup_page()
        self.init_data()
        self.create_controls()
        self.setup_ui()

    def create_controls(self):
        self.text_info = ft.Text("DEVELOPMENT BY VLADISLAV POROSHIN", size=8, weight=ft.FontWeight.BOLD, )

        self.name_field = ft.TextField(label="ИО кандидата", width=400)
        self.male_dropdown = self.create_dropdown_with_add(
            "Пол", [{"title": "М"}, {"title": "Ж"}], base_value="М", only=True, width=100)

        self.position_dropdown = self.create_dropdown_with_add(
            "Должность", self.positions, self.update_salary, btn_flds=self.__positions,
            save_func=self.__add_position
        )
        self.position_salary = ft.TextField(
            label="Зарплата", width=180, input_filter=ft.NumbersOnlyInputFilter(), suffix_text="руб.")
        self.contract_term_dropdown = self.create_dropdown_with_add(
            "Срок контракта", self.contract_terms, only=True, width=150, base_value="3 месяца")
        self.accept_days_dropdown = self.create_dropdown_with_add(
            "Срок принятия", self.accept_days, only=True, width=150, base_value="7")
        self.department_dropdown = self.create_dropdown_with_add(
            "Отдел", self.departments, btn_flds=self.__departments, save_func=self.__add_department, btn_actv=True)
        self.company_dropdown = self.create_dropdown_with_add(
            "Название компании", self.companies, self.update_departmets,
            btn_flds=self.__company, save_func=self.__add_company
        )
        self.contact_person_dropdown = self.create_dropdown_with_add(
            "Контактный специалист", self.contact_persons,
            btn_flds=self.__employee, save_func=self.__add_employee
        )
        self.signing_person_dropdown = self.create_dropdown_with_add(
            "Подписывающий специалист", self.signing_persons,
            btn_flds=self.__employee, save_func=self.__add_employee
        )
        self.responsibilities = ft.TextField(label="Обязанности", width=500, multiline=True,
                                             helper_text='Вводим без -. Каждая новая строчка через нажатие "enter"')

    def create_dropdown_with_add(self,
                                 label,
                                 options,
                                 func=None,
                                 only=False,
                                 width=440,
                                 btn_flds=None,
                                 save_func=None,
                                 btn_actv=False,
                                 base_value=None):
        btn_flds = btn_flds or []
        dropdown = ft.Dropdown(
            label=label,
            options=[ft.dropdown.Option(opt["title"]) for opt in options],
            on_change=lambda e: func(e.control.value) if func else None,
            width=width,
            value=base_value,
        )
        if only:
            return dropdown
        add_button = ft.IconButton(
            icon=ft.Icons.ADD,
            on_click=lambda e, flds=btn_flds, func=save_func: self.show_add_dialog(flds, save_func),
            tooltip="Добавить новое значение",
            disabled=btn_actv,
        )

        return {"dropdown": dropdown, "options": options, "add_button": add_button}

    @staticmethod
    def search_by_title(lst, value):
        for i in lst:
            if i["title"] == value:
                return i

    def update_departmets(self, value):
        self.departments = self.search_by_title(self.companies, value)["departments"]
        self.departments.sort(key=lambda dct: dct["title"])
        values = [ft.dropdown.Option(opt["title"]) for opt in self.departments]
        self.department_dropdown["dropdown"].options = values
        self.department_dropdown["add_button"].disabled = False
        self.page.update()

    def update_salary(self, value):
        self.position_salary.value = self.search_by_title(self.positions, value)["salary"]
        self.page.update()

    def setup_ui(self):
        self.page.add(
            ft.Column(
                controls=[
                    ft.Row([
                        ft.Text("Данные оффера", size=20, weight=ft.FontWeight.BOLD, expand=True),
                        ft.GestureDetector(content=self.text_info, on_tap=lambda e: self.show_info_dialog()),
                        ft.IconButton(icon=ft.Icons.SETTINGS, on_click=lambda e: self.show_settings_dialog()),
                    ]),
                    ft.Row([self.name_field, self.male_dropdown]),
                    ft.Row([self.company_dropdown["dropdown"], self.company_dropdown["add_button"]]),
                    ft.Row([self.department_dropdown["dropdown"], self.department_dropdown["add_button"]]),
                    ft.Row([self.position_dropdown["dropdown"], self.position_dropdown["add_button"]]),
                    ft.Row([self.position_salary, self.contract_term_dropdown, self.accept_days_dropdown]),
                    self.responsibilities,
                    ft.Divider(height=20),
                    ft.Text("Сотрудники выдающие оффер", size=20, weight=ft.FontWeight.BOLD),
                    ft.Row([self.contact_person_dropdown["dropdown"], self.contact_person_dropdown["add_button"]]),
                    ft.Row([self.signing_person_dropdown["dropdown"], self.signing_person_dropdown["add_button"]]),
                    ft.Divider(height=20),
                    ft.ElevatedButton(
                        text="Сгенерировать оффер",
                        on_click=self.generate_offer,
                        width=500,
                        height=50,
                    ),
                ],
                spacing=10,
                alignment=ft.MainAxisAlignment.CENTER,
            )
        )

    def __add_employee(self):
        employee = {
            "title": f"{self._surname.value} {self._name.value[0]}. {self._patronymic.value[0]}. ({self._position.value})",
            "name": self._name.value,
            "patronymic": self._patronymic.value,
            "surname": self._surname.value,
            "position": self._position.value,
            "email": self._email.value,
            "phone": self._phone.value,
        }
        self.__data["employees"].append(employee)

        with open(self.data_path, "w+", encoding="utf-8") as f:
            data = json.dumps(self.__data, ensure_ascii=False, indent=2)
            f.write(data)
        self.contact_persons.sort(key=lambda dct: dct["surname"])
        self.signing_persons.sort(key=lambda dct: dct["surname"])
        self.contact_person_dropdown["dropdown"].options = [
            ft.dropdown.Option(opt["title"]) for opt in self.contact_persons
        ]
        self.signing_person_dropdown["dropdown"].options = [
            ft.dropdown.Option(opt["title"]) for opt in self.signing_persons
        ]

    def __add_company(self):
        company = {
            "title": self._title.value,
            "departments": [],
        }
        self.__data["companies"].append(company)
        with open(self.data_path, "w+", encoding="utf-8") as f:
            data = json.dumps(self.__data, ensure_ascii=False, indent=2)
            f.write(data)
        self.companies.sort(key=lambda dct: dct["title"])
        self.company_dropdown["dropdown"].options = [ft.dropdown.Option(opt["title"]) for opt in self.companies]

    def __add_position(self):
        position = {
            "title": f"{self._title.value} ({self._salary.value})",
            "name": self._title.value,
            "salary": int(self._salary.value),
        }
        self.__data["positions"].append(position)
        with open(self.data_path, "w+", encoding="utf-8") as f:
            data = json.dumps(self.__data, ensure_ascii=False, indent=2)
            f.write(data)
        self.positions.sort(key=lambda dct: dct["salary"])
        self.position_dropdown["dropdown"].options = [ft.dropdown.Option(opt["title"]) for opt in self.positions]

    def __add_department(self):
        department = {
            "title": self._title.value,
        }
        company = self.search_by_title(self.companies, self.company_dropdown["dropdown"].value)
        company["departments"].append(department)
        # self.__data["positions"].append(department)
        with open(self.data_path, "w+", encoding="utf-8") as f:
            data = json.dumps(self.__data, ensure_ascii=False, indent=2)
            f.write(data)
        self.departments = company["departments"]
        self.departments.sort(key=lambda dct: dct["title"])
        self.department_dropdown["dropdown"].options = [ft.dropdown.Option(opt["title"]) for opt in self.departments]

    def show_settings_dialog(self):
        self.dialog = ft.AlertDialog(
            title=ft.Text(f"Настройки"),
            content=ft.TextButton("Сбросить настройки", on_click=lambda e: self.__reset_settings()),
            actions=[
                ft.TextButton("Закрыть", on_click=lambda e: self.close_dialog()),
            ],
        )
        self.page.add(self.dialog)

        self.page.dialog = self.dialog
        self.page.dialog.open = True
        self.page.update()

    def show_info_dialog(self):

        self.dialog = ft.AlertDialog(
            title=ft.Text(f"Спасибо что используете."),
            content=ft.Column([
                ft.Text(
                    "Данная программа была разработана мной для упрощения генерации офферов.\n"
                    "Я очень рад, что она оказалась вам полезна.\n\n"
                    "В папке с программой можно найти файл data.json, в котором хранятся все добавленные вами "
                    "данные. Вы всегда можете их отредактировать открыв файл в блокноте. Но будьте осторожны, так "
                    "как JSON формат имеет определенное форматирование. Перед сохранением советую скопировать всё и"
                    " проверить через любой валидатор JSON, которых очень много можнно найти в интернете.\n\n"

                    "Так-же рядом лежит файл template.docx, вы всеггда можете его отредактировать, если нужно "
                    "изменить формулировку или текст оффера. Но обратите внимание, что все параметры заключенные в "
                    "{{}} должны быть сохранены.\n\n"

                    "Разработчик: Порошин Владислав\n"
                    "GitHub: https://github.com/NewterraV\n"
                    "Email: vld.poroshin@gmail.com")
            ]),
            actions=[
                ft.TextButton("ОК", on_click=lambda e: self.close_dialog()),
            ],
        )
        self.page.add(self.dialog)

        self.page.dialog = self.dialog
        self.page.dialog.open = True
        self.page.update()

    def check_values(self, flds):
        result = True
        for fld in flds:
            if not fld.value:
                fld.error_text = "Введите значение!"
                result = False
            else:
                fld.error_text = None
            fld.update()
        return result

    def show_add_dialog(self, fields, save_func):

        def save_dialog(fields, func):
            if all([bool(i.value) for i in fields]):
                func()
                for field in fields:
                    field.value = None
                    field.error_text = None
                self.init_data()
                self.page.dialog.open = False
                self.page.update()
                return
            for field in self.__positions:
                if not field.value:
                    field.error_text = "Введите значение!"
                else:
                    field.error_text = None
                field.update()

        self.dialog = ft.AlertDialog(
            title=ft.Text(f"Добавить новое значение"),
            content=ft.Column(controls=fields),
            actions=[
                ft.TextButton("Отмена", on_click=lambda e, flds=fields: self.close_dialog(flds)),
                ft.TextButton(
                    "Сохранить", on_click=lambda e, flds=fields, func=save_func: save_dialog(flds, save_func)),
            ],

        )
        self.page.add(self.dialog)

        self.page.dialog = self.dialog
        self.page.dialog.open = True
        self.page.update()

    def close_dialog(self, fields=None):
        if fields:
            for field in fields:
                field.value = None
                field.error_text = None
        self.page.dialog.open = False
        self.page.update()

    def add_new_option(self, field_name, new_value):
        dropdowns = {
            "Должность": self.position_dropdown,
            "Отдел": self.department_dropdown,
            "Название компании": self.company_dropdown,
            "Контактный специалист": self.contact_person_dropdown,
            "Подписывающий специалист": self.signing_person_dropdown,
        }

        if field_name in dropdowns and new_value not in dropdowns[field_name]["options"]:
            dropdowns[field_name]["options"].append(new_value)
            dropdowns[field_name]["dropdown"].options.append(ft.dropdown.Option(new_value))
            dropdowns[field_name]["dropdown"].update()

    def calculate_salary(self, value: int):

        def formate_number(number):
            return str("{:,}".format(number).replace(",", " "))

        tax = self.__data["tax"]
        base_slr = value * self.__data['base_salary']
        bonus_slr = value * self.__data['bonus_salary']
        kpi_slr = value * self.__data['kpi_salary']
        data = {
            'salary': formate_number(value),
            'base_salary': formate_number(int(base_slr)),
            'tax_base_salary': formate_number(int(base_slr * tax)),
            'bonus_salary': formate_number(int(bonus_slr)),
            'tax_bonus_salary': formate_number(int(bonus_slr * tax)),
            'kpi_salary': formate_number(int(kpi_slr)),
            'tax_kpi_salary': formate_number(int(kpi_slr * tax)),
        }
        return data

    def generate_offer(self, e):
        if not self.check_values(self.__offer_fields):
            return

        salary_data = self.calculate_salary(int(self.position_salary.value))
        employee = self.search_by_title(self.contact_persons, self.contact_person_dropdown["dropdown"].value)
        certifying_hr = self.search_by_title(self.signing_persons, self.signing_person_dropdown["dropdown"].value)
        position = self.search_by_title(self.positions, self.position_dropdown["dropdown"].value)
        date = datetime.now()
        end_date = (date + timedelta(days=int(self.accept_days_dropdown.value))).date()
        data_to_insert = {
            "dear": "Уважаемый" if self.male_dropdown.value == "М" else "Уважаемая",
            'contract_term': self.contract_term_dropdown.value,
            'candidate_name': self.name_field.value,
            'company': self.company_dropdown["dropdown"].value,
            'speciality': position["name"],
            'department': self.department_dropdown["dropdown"].value,
            'responsibilities': self.responsibilities.value.split("\n"),
            'current_date': date.date().strftime("%d.%m.%Y"),
            'end_date': end_date.strftime("%d.%m.%Y"),
            'email': employee["email"],
            'hr_name': f'{employee["surname"]} {employee["name"]}',
            'hr_phone': employee["phone"],
            'job_title': employee["position"],
            'certifying_hr': f'{certifying_hr["name"][0]}. {certifying_hr["patronymic"][0]}. {certifying_hr["surname"]}',
        }
        data_to_insert.update(salary_data)
        output_file = f'{self.name_field.value.replace(" ", "_")}_{date.strftime("%d_%m_%Y_%H%M%S")}.docx'
        os.makedirs("results", exist_ok=True)
        replace_placeholders(
            str(os.path.join('data', 'template.docx')),
            data_to_insert,
            str(os.path.join("results", output_file))
        )

    def show_snackbar(self, message):
        self.page.snack_bar = ft.SnackBar(ft.Text(message))
        self.page.snack_bar.open = True
        self.page.update()
