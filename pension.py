#! /usr/bin/env python3

import csv
import json
import xml.etree.ElementTree as ET


##############################################################################
# Класс заполнения зарплаты
##############################################################################
class Zarplata:
    def __init__(self):
        self.month = ['Січень',
                      'Лютий',
                      'Березень',
                      'Квітень',
                      'Травень',
                      'Червень',
                      'Липень',
                      'Серпень',
                      'Вересень',
                      'Жовтень',
                      'Листопад',
                      'Грудень']
        self.salary = []  # Зарплата {}
        with open('./data/Coefficient.csv', 'r') as theFile:
            reader = csv.DictReader(theFile, delimiter=';')
            i = 0
            for line in reader:
                self.salary.append({'nn': i,
                                    'year': int(line['year']),
                                    'month': int(line['month']),
                                    'Coefficient': float(line['Coefficient'].replace(',', '.')),
                                    'zp': 0,
                                    'rc': 0,
                                    'pension': None,
                                    'experience': False})
                i += 1
            theFile.close()

    # Метод добавления зарплаты и расчета коэффициента
    def set_zp(self, year, month, zp, experience=True):
        for ll in range(0, len(self.salary)):
            if self.salary[ll].get('year') == year \
                    and self.salary[ll].get('month') == month \
                    and self.salary[ll].get('Coefficient') > 0:
                if year < 1995:
                    self.salary[ll]['zp'] = round(zp, 5)
                else:
                    self.salary[ll]['zp'] = round(zp, 2)

                self.salary[ll]['rc'] = round(zp / self.salary[ll].get('Coefficient'), 5)
                self.salary[ll]['pension'] = False
                self.salary[ll]['experience'] = experience

    def get_salary(self):
        return self.salary

    def set_pension(self, pens):
        for line in pens:
            self.salary[line.get('nn')]['pension'] = True

    # Очищаю для следующего ввода
    def erase_salary(self):
        for ll in range(len(self.salary)):
            self.salary[ll]['zp'] = 0
            self.salary[ll]['rc'] = 0
            self.salary[ll]['pension'] = None
            self.salary[ll]['experience'] = False

    # Метод добавления зарплаты из джейсона (наверное будет нужен)
    def set_zp_json(self, zp_json):
        for zz in json.loads(zp_json.decode()):
            self.set_zp(zz.get('year'), zz.get('month'), zz.get('zp'))

    # Метод загрузки XML из ПФУ справки ОК5 (можно взять на кабинете ПФУ)
    # Или csv без шапки ("год", "месц(число)", "ЗП")
    # - нужно для отладки, чтоб не набивать данные в ручную
    def import_ok5(self, file_name):
        self.erase_salary()
        ret = []
        if file_name[-3:] == 'xml':
            tree = ET.ElementTree(file=file_name)
            rtree = tree.getroot()
            revenues = rtree.findall('revenues')
            for i in revenues[0]:
                for revenue in i:
                    yyyy = ''
                    if revenue.tag == 'year':
                        yyyy = revenue.text
                    if revenue.tag == 'pensgrn':
                        mm = 1
                        for month in revenue:
                            self.set_zp(int(yyyy), int(mm), float(month.text))
                            ret.append({'year': int(yyyy), 'month': int(mm), 'zp': float(month.text)})
                            mm += 1
        if file_name[-3:] == 'csv':
            with open(file_name) as cf:
                delim = cf.read(5)[-1]
                cf.seek(0)
                in_csv = csv.reader(cf, delimiter=delim)
                for line in in_csv:
                    if line[2]:
                        zp = float(line[2].replace(',', '.'))
                        self.set_zp(int(line[0]), int(line[1]), zp)
                        ret.append({'year': int(line[0]), 'month': int(line[1]), 'zp': zp})
        return ret


##############################################################################
# Класс определения лучшего периода
##############################################################################
class GoodSalary:
    def __init__(self, salary, kz, sp1):
        self.kz = kz  # Коефициент стажу
        self.rsm = int((kz * 1200) - (sp1 * 2) + sp1)  # Расчет месяцев стажа с учетом спецстажа СПИСОК1
        self.salary_old = []   # зарплата {}
        self.salary_new = []   # зарплата {}
        self.start_p = 0  # Определяю начало персонификации

        # Разделяем на до 01.07.2000 и после
        for ii in salary:
            if ii.get('year') == 2000 and ii.get('month') == 7:
                self.start_p = ii.get('nn')
            if self.start_p == 0:
                if ii.get('experience'):
                    self.salary_old.append(ii)
            else:
                if ii.get('experience'):
                    self.salary_new.append(ii)
        self.period_a = []      # Возможные периоды для добавления заполняется методом add_period() из self.salary_old
        self.period_d = []      # Возможные периоды для удаления заполняется методом del_period() из self.salary_good
        self.salary_good = []   # Лучшая зарплата до 01.07.200
        self.salary_clear = []  # масив с чистой зарплаты по которой начислено пенсию
        self.add_p = False      # Стоит ли добавлять период

    # Метод выбора лучшего период до 01.07.2000
    def add_period(self, add=True):
        if len(self.salary_old) <= 60:
            rc = 0
            for rr in self.salary_old:
                rc += rr.get('rc')
            try:
                self.period_a.append([0, len(self.salary_old) - 1, rc, self.salary_old[0]])
            except IndexError:
                pass
        else:
            for m_st in range(len(self.salary_old)-60):
                rc = 0
                for rr in self.salary_old[m_st:m_st+59]:
                    rc += rr.get('rc')
                self.period_a.append([m_st, m_st+59, rc, self.salary_old[m_st]])

        self.period_a.sort(key=lambda x: x[2], reverse=True)
        if add:
            self.salary_good.extend(self.salary_old[self.period_a[0][0]: self.period_a[0][1]])
        self.salary_good.extend(self.salary_new)

    def d_period(self, start, end):
        sg = self.salary_good[:]
        sd = self.salary_good[start:start+end]

        for ll in sd:
            sg.remove(ll)

        rc = 0
        km = len(sg)
        if sd:
            for rs in sg:
                # km += 1
                rc += rs.get('rc')
            self.period_d.append([start, start+end, km, rc, rc/km,
                                  self.salary_good[start],
                                  sd[0].get('nn'), sd[-1].get('nn')])
        else:
            self.period_d.append([start, start+end, 0, 0, 0])

    def del_period(self):
        # отнимать можно 10% но не более 60 месяцев
        dm = 60 if int(self.rsm * 0.1) >= 60 else int(self.rsm * 0.1)
        month = len(self.salary_good)
        if month > 60:
            if int(month - dm) <= 60:  # остатся может не меньше 60
                mm = int(month - 60)
            else:
                mm = dm

            for ii in range(1, mm+1):
                for cc in range(month - ii+1):
                    self.d_period(cc, ii)

            self.period_d.sort(key=lambda x: x[4], reverse=True)
            self.salary_clear.extend(self.salary_good[:self.period_d[0][0]])
            self.salary_clear.extend(self.salary_good[self.period_d[0][1]:])
        else:
            self.salary_clear.extend(self.salary_good)
            self.period_d.append([0, len(self.salary_good)-1, 0, 0, 0])

    def clear(self):
        self.period_a = []
        self.period_d = []
        self.salary_good = []
        self.salary_clear = []

    def calc_kz(self, dd=True):
        self.clear()
        self.add_p = False
        self.add_period(add=self.add_p)
        self.del_period()
        try:
            if self.period_d[0][4] < self.period_a[0][2] / 60 or not dd:
                self.clear()
                self.add_p = True
                self.add_period(add=self.add_p)
                self.del_period()
        except IndexError:
            pass

        return self.add_p
