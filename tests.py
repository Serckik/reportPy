from unittest import TestCase
from notMain import Salary
from notMain import Vacancy

class SalaryTest(TestCase):
    def test_salary_type(self):
        self.assertEqual(Salary.get_salary_type([40,50,'RUR']), 'Salary')

    def test_salary_from(self):
        self.assertEqual(Salary.get_salary_from([40,50,'RUR']), 40)

    def test_salary_to(self):
        self.assertEqual(Salary.get_salary_to([40,50,'RUR']), 50)

    def test_salary_currency(self):
        self.assertEqual(Salary.get_salary_currency([40,50,'AZN']), 'AZN')

    def test_formatter_salary(self):
        self.assertEqual(Salary.get_salary([40,50,'RUR']), 45)

    def test_formatter_salary_with_other_currency(self):
        self.assertEqual(Salary.get_salary([40, 50, 'UAH']), 73)

class VacancyTest(TestCase):
    vacancy = {'name': 'IT аналитик', 'salary_from': 40000, 'salary_to': '45000.0', 'salary_currency': 'RUR', 'area_name': 'Санкт-Петербург', 'published_at': 2007}
    def test_vacancy_type(self):
        self.assertEqual(Vacancy.get_vacancy_type(self.vacancy),'Vacancy')

    def test_vacancy_objects(self):
        self.assertEqual(Vacancy.get_vacancy_fields(self.vacancy), [self.vacancy['name'], self.vacancy['salary_from'], self.vacancy['salary_to'], self.vacancy['salary_currency'], self.vacancy['area_name'], self.vacancy['published_at']])

    def test_vacancy_one_field(self):
        self.assertEqual(Vacancy.get_vacancy_field(self.vacancy, 'name'),'IT аналитик')