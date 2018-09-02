import csv
import os

class OutFile:
    def __init__(self, salary, output, start_p):
        self.salary = salary
        self.output = output
        self.start_p = start_p
        if not os.path.isdir('./in'):
            os.mkdir('./in')
        if not os.path.isdir('./report'):
            os.mkdir('./report')
        if not os.path.isfile('./data/clients.csv'):
            cl_file = open('./data/clients.csv', 'w', encoding='cp1251')
            cl = csv.writer(cl_file, delimiter=';')
            cl.writerow(['ПІБ',	'Телефон',	'Область',	'Населений пункт',	'До оптимізації',	'Після оптимізації'])
            cl_file.close()

    def save(self, csv_file=False):
        cl = csv.writer(open('./data/clients.csv', 'a', encoding='cp1251'), delimiter=';')
        cl.writerow([self.output[-1].get('fio'),
                     self.output[-1].get('tel'),
                     self.output[-1].get('stat'),
                     self.output[-1].get('city'),
                     str(self.output[5][1]).replace('.', ','),
                     str(self.output[5][2]).replace('.', ',')])


        f_name = self.output[-1].get('tel').replace(' ', '') if self.output[-1].get('tel') else '9999999999'

        if csv_file:
            salary_csv = csv.writer(open('./in/{}.csv'.format(f_name), 'w'))

        report = open('./report/{}.xls'.format(f_name), 'w', encoding='utf8')

        htmls = '<html><head>'
        htmls += '<meta charset="utf-8">'
        htmls += '</head><body>\n'
        htmls += '<table border="1">\n'
        htmls += '<tr><td colspan="3">ПІБ<td colspan="4">{}\n'.format(self.output[-1].get('fio'))
        htmls += '<tr><td colspan="3">Телефон<td colspan="4">="{}"\n'.format(self.output[-1].get('tel'))
        htmls += '<tr><td colspan="3">Область<td colspan="4">{}\n'.format(self.output[-1].get('stat'))
        htmls += '<tr><td colspan="3">Населений пункт<td colspan="4">{}\n'.format(self.output[-1].get('city'))
        htmls += '</table><br>\n\n'

        htmls += '<table border="1">\n'
        htmls += '<tr><td colspan="3">Оптимізація<td>Було<td>Стало \n'
        for line in self.output[:-3]:
            htmls += '<tr><td colspan="3">{}<td>{}<td>{}\n'.format(line[0],
                                                                   str(line[1]).replace('.', ','),
                                                                   str(line[2]).replace('.', ','))
        htmls += '</table><br>\n\n'

        htmls += '<table border="1">\n'
        for line in self.output[-3:-1]:
            htmls += '<tr><td colspan="3">{}<td>{}<td>{}\n'.format(line[0],
                                                                   str(line[1]).replace('.', ','),
                                                                   str(line[2]).replace('.', ','))
        htmls += '</table>\n\n'
        htmls += '<br><br><b>Відомості по заробітній платі</b>'

        htmls += '''<table border="1"><tr><td>Рік\n<td>None<td>Січень\n<td>Лютий\n<td>Березень\n<td>Квітень\n
                 <td>Травень\n<td>Червень\n<td>Липень\n<td>Серпень\n<td>Вересень\n
                 <td>Жовтень\n<td>Листопад\n<td>Грудень\n<td>Всього\n</tr>\n'''
        out_zp = {}
        out_rc = {}
        out_cv = {}
        for line in self.salary:
            if line.get('experience'):
                if csv_file:
                    salary_csv.writerow([line.get('year'), line.get('month'), line.get('zp')])

                if not out_zp.get(line.get('year')):
                    out_zp[line.get('year')] = [None, None, None, None, None, None,
                                                None, None, None, None, None, None, 0]
                    out_rc[line.get('year')] = [None, None, None, None, None, None,
                                                None, None, None, None, None, None, 0]
                    out_cv[line.get('year')] = [None, None, None, None, None, None,
                                                None, None, None, None, None, None, 0]
                out_zp[line.get('year')][line.get('month')-1] = line.get('zp')
                out_zp[line.get('year')][-1] += line.get('zp')

                out_rc[line.get('year')][line.get('month')-1] = line.get('rc')
                out_rc[line.get('year')][-1] += line.get('rc')

                out_cv[line.get('year')][line.get('month') - 1] = line.get('nn')

        for yyyy in out_zp.keys():
            htmls += '<tr><td rowspan="2" valign="middle" align="center">{}'.format(yyyy)
            h1 = '<td>ЗП'
            h2 = '<td>Коеф. ЗП'
            ii = 0
            for zp, rc, nn in zip(out_zp.get(yyyy), out_rc.get(yyyy), out_cv.get(yyyy)):
                if ii == 12:
                    h1 += '<td>' + str(zp).replace('.', ',')
                    h2 += '<td>' + str(round(rc, 5)).replace('.', ',')
                else:
                    if nn:
                        if nn >= self.output[-2][3] and nn <= self.output[-2][4]:
                            h1 += '<td bgcolor="#a9a9a9">'
                            h2 += '<td bgcolor="#a9a9a9">'
                        else:
                            h1 += '<td>'
                            h2 += '<td>'
                        if len(self.output[-3]) == 5:
                            if (nn >= self.output[-3][3] and nn <= self.output[-3][4]) or nn >= self.start_p:
                                h1 += str(zp).replace('.', ',')+'</i></font>'
                                h2 += str(rc).replace('.', ',')+'</i></font>'
                            else:
                                h1 += '<font color="#ff0000"><i>'+str(zp).replace('.', ',')
                                h2 += '<font color="#ff0000"><i>'+str(rc).replace('.', ',')
                        else:
                            if nn < self.start_p:
                                h1 += '<font color="#ff0000"><i>'+str(zp).replace('.', ',')
                                h2 += '<font color="#ff0000"><i>'+str(rc).replace('.', ',')
                            else:
                                h1 += str(zp).replace('.', ',')
                                h2 += str(rc).replace('.', ',')
                    else:
                        h1 += '<td>' + str(zp).replace('.', ',')
                        h2 += '<td>' + str(rc).replace('.', ',')
                ii += 1
            htmls += h1
            htmls += '</tr><tr>'
            htmls += h2

            htmls += '</tr>\n'
        htmls += '</table><br>\n'
        htmls += 'Відповідно до Закону України «Про захист персональних даних» від 01.06.2010 №2297-VI<br>'
        htmls += 'надаю згоду Політичні партії «Об’єднання «Самопоміч»» на обробку та використання персональних<br>'
        htmls += 'даних про з метою оптимізації обрахунку пенсії та інших дії, необхідних для досягнення'
        htmls += 'вказаної вище мети.<br>'
        htmls += '<br>Підпис _______________<br>'

        htmls += '<br>Я попереджений, що розрахунки є оціночноми і носять рекомендаційний характер.<br>'
        htmls += '<br>Дата _______________ Підпис _______________ ПІБ ______________________________________'
        htmls += '</body></html>'

        report.write(htmls.replace('None', '&nbsp;'))
