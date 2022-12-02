import csv
import openpyxl
from openpyxl.styles import Font, Color, Border, Side
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from openpyxl import Workbook
import matplotlib.pyplot as plt
import numpy as np
import pdfkit
from jinja2 import Environment, FileSystemLoader


"""
vacancies.csv
Аналитик
uwu
"""

class Report:
    def __init__(self, profession):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "Статистика по годам"
        self.ws1 = self.wb.create_sheet("Статистика по городам")
        self.border = Border(
            left=Side(border_style="thin", color="000000"),
            right=Side(border_style="thin", color="000000"),
            top=Side(border_style="thin", color="000000"),
            bottom=Side(border_style="thin", color="000000")
        )
        self.fig, self.ax = plt.subplots(nrows=2, ncols=2, figsize=(10, 10))
        self.profession = profession
        self.env = Environment(loader=FileSystemLoader('.'))
        self.html = self.env.get_template("main.html")
        self.a = 1

    def set_column_width(self, name_column, sheet, count_rows, start):
        for i in range(len(name_column)):
            max_len = 0
            for j in range(1, count_rows + 1):
                if max_len < len(str(sheet[chr(i + ord(start)) + str(j)].value)):
                    max_len = len(str(sheet[chr(i + ord(start)) + str(j)].value)) + 2
            sheet.column_dimensions[chr(i + ord(start))].width = max_len

    def fill_header(self, boxes, name_column, is_bold):
        i = 0
        for column in boxes:
            for box in column:
                box.value = name_column[i]
                box.font = Font(bold=is_bold)
                box.border = self.border
                i += 1

    def generate_excel(self, list_name, data, header, start, end, type=""):
        self.ws = self.wb[list_name]
        self.fill_header(list(self.ws[start + '1': end + '1']), header, True)
        for i in range(len(data)):
            keys = list(data[i].keys())
            for j in range(len(keys)):
                self.ws[start + str(j + 2)].value = keys[j]
                self.ws[start + str(j + 2)].border = self.border
                self.ws[chr(i + ord(start) + 1) + str(j + 2)].value = data[i][keys[j]]
                self.ws[chr(i + ord(start) + 1) + str(j + 2)].border = self.border
                if type:
                    self.ws[chr(i + ord(start) + 1) + str(j + 2)].number_format = FORMAT_PERCENTAGE_00
        self.set_column_width(header, self.ws, len(list(data[0].keys())) + 1, start)

    def save_excel(self):
        self.wb.save("report.xlsx")

    def generate_image(self, salary_by_year, salary_by_year_for_profession, count_by_year, count_by_year_for_profession, sum_salary_by_year_for_city, fraction_by_city):
        width = 0.35
        x = np.array(list(salary_by_year.keys()))
        self.ax[0, 0].bar(x - width / 2, list(salary_by_year.values()), width, label='Средняя з/п')
        self.ax[0, 0].bar(x + width / 2, list(salary_by_year_for_profession.values()), width, label=f'з/п {profession}')
        self.ax[0, 0].set_title('Уровень зарплат по годам')
        self.ax[0, 0].set_xticks(x)
        self.ax[0, 0].set_xticklabels(list(salary_by_year.keys()), rotation=90, fontsize=8)
        self.ax[0, 0].legend(prop={'size': 8})
        self.ax[0, 0].grid(axis='y')

        self.ax[0, 1].bar(x - width / 2, list(count_by_year.values()), width, label='Количество вакансий')
        self.ax[0, 1].bar(x + width / 2, list(count_by_year_for_profession.values()), width, label=f'Количество вакансий {profession}')
        self.ax[0, 1].set_title('Количество вакансий по годам')
        self.ax[0, 1].set_xticks(x)
        self.ax[0, 1].set_xticklabels(list(salary_by_year.keys()), rotation=90, fontsize=8)
        self.ax[0, 1].legend(prop={'size': 8})
        self.ax[0, 1].grid(axis='y')

        label = []
        for item in list(sum_salary_by_year_for_city.keys()):
            if "-" in item:
                label.append(item.replace("-", "-\n"))
            elif " " in item:
                label.append(item.replace(" ", "\n"))
            else:
                label.append(item)

        y = np.array(label)
        self.ax[1, 0].barh(y, list(sum_salary_by_year_for_city.values()))
        self.ax[1, 0].set_yticks(y, labels=label, fontsize=6)
        self.ax[1, 0].invert_yaxis()
        self.ax[1, 0].set_title('Уровень зарплат по городам')
        self.ax[1, 0].grid(axis='x')

        summ = sum(list(fraction_by_city.values())[10:])

        fraction_by_city = dict(list(fraction_by_city.items())[0:10])
        fraction_by_city["Другие"] = summ

        self.ax[1, 1].pie(list(fraction_by_city.values()), labels=list(fraction_by_city.keys()), textprops={'fontsize': 6})
        self.ax[1, 1].axis("equal")
        self.ax[1, 1].set_title('Доля вакансий по городам')

        plt.tight_layout()
        plt.savefig('graph.png', dpi=100)

    def create_pdf(self):
        options = {
            "enable-local-file-access": None
        }

        xfile = openpyxl.load_workbook("report.xlsx")
        data = xfile['Статистика по годам']
        data2 = xfile['Статистика по городам']

        config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
        pdfkit.from_string(self.html.render({'profession': self.profession, 'first': data, 'second': data2}), 'report.pdf', configuration=config, options=options)



class DataSet:
    dict_count_by_year = {}
    dict_sum_salary_by_year = {}
    dict_count_by_year_for_profession = {}
    dict_sum_salary_by_year_for_profession = {}
    dict_count_by_year_for_city = {}
    dict_sum_salary_by_year_for_city = {}

    def __init__(self, file_name, filtering_parameter):
        self.file_name = file_name
        self.filtering_parameter = filtering_parameter
        self.vacancies_objects = self.csv_reader()



    def csv_reader(self):
        data = list(csv.reader(open(self.file_name, encoding="utf-8-sig")))
        if len(data) == 0:
            print("Пустой файл")
            quit()
        data_header = data[0]
        vacancies = [x for x in data[1:]]
        return self.csv_filer(data_header, vacancies)

    def csv_filer(self, list_naming, reader):
        reader = [x for x in reader if not "" in x and len(x) == len(list_naming)]
        filter_list_vacancies = []
        for i in range(len(reader)):
            dict_vacancy = {}
            for j in range(len(list_naming)):
                vacancies_element = reader[i][j]
                if list_naming[j] == "key_skills":
                    vacancies_element = vacancies_element.split("\n")
                    vacancies_element = [x.strip() for x in vacancies_element]
                dict_vacancy[list_naming[j]] = vacancies_element
            dict_vacancy["salary_from"] = Salary([dict_vacancy["salary_from"], dict_vacancy["salary_to"], dict_vacancy["salary_currency"]])
            dict_vacancy = self.formatter(dict_vacancy)
            filter_list_vacancies.append(Vacancy(dict_vacancy))

        for item in filter_list_vacancies:
            self.set_statistics(item.vacansy_dict)

        if len(filter_list_vacancies) == 0:
            print("Нет данных")
            quit()

        for item in self.dict_sum_salary_by_year:
            self.dict_sum_salary_by_year[item] = self.dict_sum_salary_by_year[item] // self.dict_count_by_year[item]
            if item not in self.dict_sum_salary_by_year_for_profession:
                self.dict_sum_salary_by_year_for_profession[item] = 0
                self.dict_count_by_year_for_profession[item] = 0
                continue
            self.dict_sum_salary_by_year_for_profession[item] = self.dict_sum_salary_by_year_for_profession[item] // self.dict_count_by_year_for_profession[item]

        sorted_dict_sum_salary_by_year_for_city = {}
        sorted_dict_fraction_by_city = {}
        for item in self.dict_sum_salary_by_year_for_city:
            if(int(len(reader) * 0.01) <= self.dict_count_by_year_for_city[item]):
                sorted_dict_sum_salary_by_year_for_city[item] = self.dict_sum_salary_by_year_for_city[item] // self.dict_count_by_year_for_city[item]
                sorted_dict_fraction_by_city[item] = round(self.dict_count_by_year_for_city[item] / len(reader), 4)


        print(f"Динамика уровня зарплат по годам: {self.dict_sum_salary_by_year}")
        print(f"Динамика количества вакансий по годам: {self.dict_count_by_year}")
        print(f"Динамика уровня зарплат по годам для выбранной профессии: {self.dict_sum_salary_by_year_for_profession}")
        print(f"Динамика количества вакансий по годам для выбранной профессии: {self.dict_count_by_year_for_profession}")

        sorted_dict_sum_salary_by_year_for_city = dict(sorted(sorted_dict_sum_salary_by_year_for_city.items(), key=lambda item: item[1], reverse=True)[0:10])
        sorted_dict_fraction_by_city = dict(sorted(sorted_dict_fraction_by_city.items(), key=lambda item: item[1], reverse=True)[0:10])
        print(f"Уровень зарплат по городам (в порядке убывания): {sorted_dict_sum_salary_by_year_for_city}")
        print(f"Доля вакансий по городам (в порядке убывания): {sorted_dict_fraction_by_city}")
        xls = Report(self.filtering_parameter)

        header = ["Год", "Средняя зарплата", f"Средняя зарплата - {self.filtering_parameter}", "Количество вакансий", f"Количество вакансий - {self.filtering_parameter}"]
        if choise == "Вакансии":
            xls.generate_excel("Статистика по годам", [self.dict_sum_salary_by_year, self.dict_count_by_year, self.dict_sum_salary_by_year_for_profession, self.dict_count_by_year_for_profession], header, 'A', 'E')

            header = ["Город", "Уровень зарплат"]
            xls.generate_excel("Статистика по городам", [sorted_dict_sum_salary_by_year_for_city], header, 'A', 'B')

            header = ["Город", "Доля вакансий"]
            xls.generate_excel("Статистика по городам", [sorted_dict_fraction_by_city], header, 'D', 'E', "percent")
            xls.ws.column_dimensions["C"].width = 2

            xls.save_excel()
        elif choise == "Статистика":
            xls.generate_image(self.dict_sum_salary_by_year, self.dict_sum_salary_by_year_for_profession, self.dict_count_by_year, self.dict_count_by_year_for_profession, sorted_dict_sum_salary_by_year_for_city, sorted_dict_fraction_by_city)


    def set_statistics(self, list_vacancy):
        if list_vacancy["published_at"] not in self.dict_count_by_year:
            self.dict_count_by_year[list_vacancy["published_at"]] = 1
            self.dict_sum_salary_by_year[list_vacancy["published_at"]] = list_vacancy["salary_from"]
        else:
            self.dict_count_by_year[list_vacancy["published_at"]] += 1
            self.dict_sum_salary_by_year[list_vacancy["published_at"]] += list_vacancy["salary_from"]

        if list_vacancy["published_at"] not in self.dict_count_by_year_for_profession and self.filtering_parameter in list_vacancy["name"]:
            self.dict_count_by_year_for_profession[list_vacancy["published_at"]] = 1
            self.dict_sum_salary_by_year_for_profession[list_vacancy["published_at"]] = list_vacancy["salary_from"]
        elif self.filtering_parameter in list_vacancy["name"]:
            self.dict_count_by_year_for_profession[list_vacancy["published_at"]] += 1
            self.dict_sum_salary_by_year_for_profession[list_vacancy["published_at"]] += list_vacancy["salary_from"]

        if list_vacancy["area_name"] not in self.dict_count_by_year_for_city:
            self.dict_count_by_year_for_city[list_vacancy["area_name"]] = 1
            self.dict_sum_salary_by_year_for_city[list_vacancy["area_name"]] = list_vacancy["salary_from"]
        else:
            self.dict_count_by_year_for_city[list_vacancy["area_name"]] += 1
            self.dict_sum_salary_by_year_for_city[list_vacancy["area_name"]] += list_vacancy["salary_from"]

    def formatter(self, row):
        date = row["published_at"][:10].split("-")
        row["published_at"] = int(date[0])
        row["salary_from"] = int((float(row["salary_from"].salary_from) + float(row["salary_from"].salary_to)) // 2 * currency_to_rub[row["salary_from"].salary_currency])
        return row

class Vacancy:
    def __init__(self, list):
        self.vacansy_dict = list

class Salary:
    def __init__(self, list):
        self.salary_from, \
        self.salary_to,\
        self.salary_currency = list

currency_to_rub = {
    "AZN": 35.68,
    "BYR": 23.91,
    "EUR": 59.90,
    "GEL": 21.74,
    "KGS": 0.76,
    "KZT": 0.13,
    "RUR": 1,
    "UAH": 1.64,
    "USD": 60.66,
    "UZS": 0.0055,
}

file_name = input("Введите название файла: ")
profession = input("Введите название профессии: ")
choise = input("Введите как отобразить данные: ")

correct_list_vacancies = DataSet(file_name, profession)