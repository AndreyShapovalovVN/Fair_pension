#! /usr/bin/env python3

from tkinter import *
from tkinter import messagebox
from tkinter import filedialog as fd
from sys import platform
import os
import json
import pension
import outreport


class MainForm:
    def __init__(self, top=None):
        self.zp = pension.Zarplata()
        self.zarplata = []
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

        top.geometry("795x480")
        top.title("<<Справедлива пенсія>>")

        self.Frame1 = Frame(top)
        self.Frame1.place(relx=0.0, rely=0.02, relheight=0.97, relwidth=1.0)
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(borderwidth="2")
        self.Frame1.configure(width=795)

        self.Button1 = Button(self.Frame1)
        self.Button1.place(relx=0.35, rely=0.927, height=29, width=209)
        self.Button1.configure(text='''Разрахунок''')
        self.Button1.configure(command=lambda: self.calck(''))
        self.Button1.bind('<Return>', self.calck)
        self.Button1.bind('<KP_Enter>', self.calck)

        self.Entry4 = Entry(self.Frame1)
        self.Entry4.place(relx=0.11, rely=0.04, height=23, relwidth=0.31)
        self.Entry4.configure(background="white")
        self.Entry4.configure(width=246)
        self.Entry4.bind('<Return>', lambda event, f='Entry9': self.next_obj(event, f))
        self.Entry4.bind('<KP_Enter>', lambda event, f='Entry9': self.next_obj(event, f))

        self.Entry9 = Entry(self.Frame1)
        self.Entry9.place(relx=0.63, rely=0.04, height=23, relwidth=0.31)
        self.Entry9.configure(background="white")
        self.Entry9.configure(width=246)
        self.Entry9.bind('<Return>', lambda event, f='Entry10': self.next_obj(event, f))
        self.Entry9.bind('<KP_Enter>', lambda event, f='Entry10': self.next_obj(event, f))

        self.Entry10 = Entry(self.Frame1)
        self.Entry10.place(relx=0.11, rely=0.11, height=23, relwidth=0.31)
        self.Entry10.configure(background="white")
        self.Entry10.configure(width=246)
        self.Entry10.bind('<Return>', lambda event, f='Entry11': self.next_obj(event, f))
        self.Entry10.bind('<KP_Enter>', lambda event, f='Entry11': self.next_obj(event, f))

        self.Entry11 = Entry(self.Frame1)
        self.Entry11.place(relx=0.63, rely=0.11, height=23, relwidth=0.31)
        self.Entry11.configure(background="white")
        self.Entry11.configure(width=246)
        self.Entry11.bind('<Return>', lambda event, f='Entry5': self.next_obj(event, f))
        self.Entry11.bind('<KP_Enter>', lambda event, f='Entry5': self.next_obj(event, f))

        self.Label10 = Label(self.Frame1)
        self.Label10.place(relx=0.03, rely=0.04, height=21, width=59)
        self.Label10.configure(anchor=NE)
        self.Label10.configure(text='''ПІБ''')
        self.Label10.configure(width=59)

        self.Label10_3 = Label(self.Frame1)
        self.Label10_3.place(relx=0.44, rely=0.05, height=21, width=139)
        self.Label10_3.configure(anchor=NE)
        self.Label10_3.configure(text='''Телефон''')
        self.Label10_3.configure(width=139)

        self.Label10_4 = Label(self.Frame1)
        self.Label10_4.place(relx=0.01, rely=0.11, height=21, width=69)
        self.Label10_4.configure(anchor=NE)
        self.Label10_4.configure(text='''Область''')
        self.Label10_4.configure(width=69)

        self.Label10_4 = Label(self.Frame1)
        self.Label10_4.place(relx=0.44, rely=0.11, height=21, width=139)
        self.Label10_4.configure(anchor=NE)
        self.Label10_4.configure(text='''Населений понкт''')

        self.Frame2 = Frame(self.Frame1)
        self.Frame2.place(relx=0.01, rely=0.52, relheight=0.4, relwidth=0.97)
        self.Frame2.configure(relief=GROOVE)
        self.Frame2.configure(borderwidth="2")
        self.Frame2.configure(width=775)
##
        self.Text1 = Text(self.Frame2)
        self.Text1.place(relx=0.49, rely=0.05, relheight=0.89, relwidth=0.5)
        self.Text1.configure(background="white")
        self.Text1.configure(takefocus="0")
        self.Text1.configure(width=386)
        self.Text1.configure(wrap=WORD)

        self.Frame4 = Frame(self.Frame1)
        self.Frame4.place(relx=0.01, rely=0.22, relheight=0.29, relwidth=0.97)
        self.Frame4.configure(relief=GROOVE)
        self.Frame4.configure(borderwidth="2")
        self.Frame4.configure(width=775)

        self.Entry5 = Entry(self.Frame4)
        self.Entry5.place(relx=0.26, rely=0.3, height=23, relwidth=0.21)
        self.Entry5.configure(background="white")
        self.Entry5.bind('<Return>', lambda event, f='Entry6': self.next_obj(event, f))
        self.Entry5.bind('<KP_Enter>', lambda event, f='Entry6': self.next_obj(event, f))
        self.Entry5.bind('<Tab>', lambda event, f='Entry6': self.next_obj(event, f))

        self.Entry6 = Entry(self.Frame4)
        self.Entry6.place(relx=0.26, rely=0.52, height=23, relwidth=0.21)
        self.Entry6.configure(background="white")
        self.Entry6.bind('<Return>', lambda event, f='Entry12': self.next_obj(event, f))
        self.Entry6.bind('<KP_Enter>', lambda event, f='Entry12': self.next_obj(event, f))
        self.Entry6.bind('<Tab>', lambda event, f='Entry12': self.next_obj(event, f))

        self.Entry12 = Entry(self.Frame4)
        self.Entry12.place(relx=0.26, rely=0.74, height=23, relwidth=0.21)
        self.Entry12.configure(background="white")
        self.Entry12.bind('<Return>', lambda event, f='Entry7': self.next_obj(event, f))
        self.Entry12.bind('<KP_Enter>', lambda event, f='Entry7': self.next_obj(event, f))
        self.Entry12.bind('<Tab>', lambda event, f='Entry7': self.next_obj(event, f))

        self.Entry7 = Entry(self.Frame4)
        self.Entry7.place(relx=0.77, rely=0.3, height=23, relwidth=0.21)
        self.Entry7.configure(background="white")
        self.Entry7.bind('<Return>', lambda event, f='Entry8': self.next_obj(event, f))
        self.Entry7.bind('<KP_Enter>', lambda event, f='Entry8': self.next_obj(event, f))
        self.Entry7.bind('<Tab>', lambda event, f='Entry8': self.next_obj(event, f))

        self.Entry8 = Entry(self.Frame4)
        self.Entry8.place(relx=0.77, rely=0.52, height=23, relwidth=0.21)
        self.Entry8.configure(background="white")
        self.Entry8.bind('<Return>', lambda event, f='Entry13': self.next_obj(event, f))
        self.Entry8.bind('<KP_Enter>', lambda event, f='Entry13': self.next_obj(event, f))
        self.Entry8.bind('<Tab>', lambda event, f='Entry13': self.next_obj(event, f))

        self.Entry13 = Entry(self.Frame4)
        self.Entry13.place(relx=0.77, rely=0.74, height=23, relwidth=0.21)
        self.Entry13.configure(background="white")
        self.Entry13.bind('<Return>', lambda event, f='Entry1': self.next_obj(event, f))
        self.Entry13.bind('<KP_Enter>', lambda event, f='Entry1': self.next_obj(event, f))
        self.Entry13.bind('<Tab>', lambda event, f='Entry1': self.next_obj(event, f))

        self.Label1 = Label(self.Frame4)
        self.Label1.place(relx=0.01, rely=0.07, height=21, width=750)
        self.Label1.configure(text='''Дані з Пенсійного Фонду''')

        self.Label6 = Label(self.Frame4)
        self.Label6.place(relx=0.01, rely=0.3, height=21, width=179)
        self.Label6.configure(anchor=NE)
        self.Label6.configure(text='''ЗП для обрахунку пенсії''')

        self.Label7 = Label(self.Frame4)
        self.Label7.place(relx=0.01, rely=0.52, height=21, width=179)
        self.Label7.configure(anchor=NE)
        self.Label7.configure(text='''Коеф. Стажу''')

        self.Label8 = Label(self.Frame4)
        self.Label8.place(relx=0.53, rely=0.3, height=21, width=179)
        self.Label8.configure(anchor=NE)
        self.Label8.configure(text='''Коефіцієнт ЗП''')

        self.Label9 = Label(self.Frame4)
        self.Label9.place(relx=0.49, rely=0.74, height=21, width=213)
        self.Label9.configure(anchor=NE)
        self.Label9.configure(text='''Пенсія за віком''')

        self.Label7_1 = Label(self.Frame4)
        self.Label7_1.place(relx=0.49, rely=0.52, height=21, width=211)
        self.Label7_1.configure(anchor=NE)
        self.Label7_1.configure(text='''Додаткові виплати''')

        self.Label7_2 = Label(self.Frame4)
        self.Label7_2.place(relx=0.01, rely=0.74, height=21, width=179)
        self.Label7_2.configure(anchor=NE)
        self.Label7_2.configure(text='''Стаж за 1 списком''')

        self.Frame3 = Frame(self.Frame2)
        self.Frame3.place(relx=0.01, rely=0.05, relheight=0.89, relwidth=0.47)
        self.Frame3.configure(relief=GROOVE)
        self.Frame3.configure(borderwidth="2")
        self.Frame3.configure(width=365)

        self.Button2 = Button(self.Frame3)
        self.Button2.place(relx=0.03, rely=0.79, height=29, width=109)
        self.Button2.configure(text='''Додати''')
        self.Button2.configure(command=lambda: self.save(''))
        self.Button2.bind('<Return>', self.save)
        self.Button2.bind('<KP_Enter>', self.save)

        self.Button3 = Button(self.Frame3)
        self.Button3.place(relx=0.34, rely=0.79, height=29, width=118)
        self.Button3.configure(text='''Імпорт з файлу''')
        self.Button3.configure(command=lambda: self.insertOK5(''))
        self.Button3.bind('<Return>', self.insertOK5)
        self.Button3.bind('<KP_Enter>', self.insertOK5)

        self.Button4 = Button(self.Frame3)
        self.Button4.place(relx=0.68, rely=0.79, height=29, width=109)
        self.Button4.configure(text='''Видалити''')
        self.Button4.configure(command=lambda: self.clear(''))
        self.Button4.bind('<Return>', self.clear)
        self.Button4.bind('<KP_Enter>', self.clear)

        self.Entry1 = Entry(self.Frame3)
        self.Entry1.place(relx=0.52, rely=0.24, height=23, relwidth=0.45)
        self.Entry1.configure(background="white")
        self.Entry1.bind('<Return>', lambda event, f='Entry2': self.next_obj(event, f))
        self.Entry1.bind('<KP_Enter>', lambda event, f='Entry2': self.next_obj(event, f))
        self.Entry1.bind('<Tab>', lambda event, f='Entry2': self.next_obj(event, f))

        self.Entry2 = Entry(self.Frame3)
        self.Entry2.place(relx=0.52, rely=0.42, height=23, relwidth=0.45)
        self.Entry2.configure(background="white")
        self.Entry2.bind('<Return>', lambda event, f='Entry3', s=True: self.next_obj(event, f, s))
        self.Entry2.bind('<KP_Enter>', lambda event, f='Entry3', s=True: self.next_obj(event, f, s))
        self.Entry2.bind('<Tab>', lambda event, f='Entry3', s=True: self.next_obj(event, f, s))

        self.Entry3 = Entry(self.Frame3)
        self.Entry3.place(relx=0.52, rely=0.61, height=23, relwidth=0.45)
        self.Entry3.configure(background="white")
        self.Entry3.bind('<Return>', lambda event, f='Button2': self.next_obj(event, f))
        self.Entry3.bind('<KP_Enter>', lambda event, f='Button2': self.next_obj(event, f))
        self.Entry3.bind('<Tab>', lambda event, f='Button2': self.next_obj(event, f))

        self.Label2 = Label(self.Frame3)
        self.Label2.place(relx=0.01, rely=0.06, height=21, width=349)
        self.Label2.configure(text='''Дані про заробітну плату''')

        self.Label3 = Label(self.Frame3)
        self.Label3.place(relx=0.01, rely=0.24, height=21, width=179)
        self.Label3.configure(anchor=NE)
        self.Label3.configure(text='''Рік''')

        self.Label4 = Label(self.Frame3)
        self.Label4.place(relx=0.01, rely=0.42, height=21, width=179)
        self.Label4.configure(anchor=NE)
        self.Label4.configure(text='''Місяць''')

        self.Label5 = Label(self.Frame3)
        self.Label5.place(relx=0.01, rely=0.61, height=21, width=179)
        self.Label5.configure(anchor=NE)
        self.Label5.configure(text='''Заробітна плата''')
        ##

        self.Entry4.focus_set()
        self.Entry5.insert(0, 3764.40)

    def next_obj(self, event, obj, select=False):
        eval('self.'+obj+'.focus_set()')
        if select:
            eval('self.' + obj + '.select_range(0, "end")')

    def get_float(self, obj):
        ret = eval('float(self.'+obj+'.get().replace(",", "."))') if eval('self.'+obj+'.get()') else 0.0
        return ret

    def save(self, event):
        # Валидация показателей
        y_begin = self.zp.salary[0].get('year')
        y_end = self.zp.salary[-1].get('year')+1
        yyyy = int(self.Entry1.get()) if self.Entry1.get() else 0
        if yyyy not in [i for i in range(y_begin, y_end)]:
            messagebox.showerror("Error", "Невірний рік!\nПеревірте та повторить")
            self.Entry1.focus_set()
            self.Entry1.select_range(0, "end")
            return False

        mm = int(self.Entry2.get()) if self.Entry2.get() else 0
        if mm not in [i for i in range(1, 13)]:
            messagebox.showerror("Error", "Невірний місяць!\nПеревірте та повторить")
            self.Entry2.focus_set()
            self.Entry2.select_range(0, "end")
            return False

        if not self.Entry3.get():
            vv = messagebox.askyesno("Видалити",
                                     "Ви бажаєте видалити місяць\n{} {} року".format(self.month[mm-1], yyyy))
            if vv:
                self.zp.set_zp(yyyy, mm, self.get_float('Entry3'), False)
                self.Text1.insert(1.0,
                                  '{} {} ({}): {}\n'.format(yyyy, mm, self.month[mm - 1], 'Видалено'))
            else:
                self.Entry3.focus_set()
                return False

        else:
            # Запись информации
            self.zp.set_zp(yyyy, mm, self.get_float('Entry3'), True)
            self.Text1.insert(1.0, '{} {} ({}): {}\n'.format(yyyy, mm, self.month[mm-1], self.get_float('Entry3')))

            # Подготовка к следующей записи
            mm = int(self.Entry2.get())
            self.Entry2.delete(0, 'end')
            if mm < 12:
                self.Entry2.insert(0, str(mm + 1).zfill(2))
            else:
                self.Entry2.insert(0, str(1).zfill(2))
                yy = int(self.Entry1.get())
                self.Entry1.delete(0, 'end')
                self.Entry1.insert(0, str(yy + 1))

            self.Entry1.focus_set()
            return True

    def insertOK5(self, event):
        self.clear('')
        file_name = fd.askopenfilename()
        if file_name[-4:] == 'json':
            self.zp.erase_salary()
            with open(file_name) as f:
                data = json.load(f)

            for ll in data.get('salary'):
                if ll.get('experience'):
                    self.zp.set_zp(ll.get('year'), ll.get('month'), ll.get('zp'))
                    self.Text1.insert(1.0, '{} {} ({}): {}\n'.format(ll.get('year'),
                                                                     ll.get('month'),
                                                                     self.month[ll.get('month') - 1],
                                                                     ll.get('zp')))
            for ff in data.get('form'):
                for key, value in ff.items():
                    eval('{}.delete(first=0, last=1000)'.format(key))
                    eval('{}.insert(0, "{}")'.format(key, value))
        else:
            for ll in self.zp.import_ok5(file_name):
                self.Text1.insert(1.0, '{} {} ({}): {}\n'.format(ll.get('year'),
                                                                 ll.get('month'),
                                                                 self.month[ll.get('month') - 1],
                                                                 ll.get('zp')))

    def clear(self, event):
        self.zp.erase_salary()
        self.Entry1.delete(first=0, last=1000)
        self.Entry2.delete(first=0, last=1000)
        self.Entry3.delete(first=0, last=1000)
        self.Text1.delete("1.0", END)
        self.Entry1.focus_set()

    def calck(self, event):
        # Расчет пенсии
        if self.get_float('Entry6')*1200 < self.get_float('Entry12') / 2 or self.get_float('Entry6')*1200 < 1:
            messagebox.showerror("Error", 'Перевірте:\n "Коеф. Стажу" та\n "Стаж за 1 списком"')
        else:
            gs = pension.GoodSalary(self.zp.get_salary(), self.get_float('Entry6'), self.get_float('Entry12'))
            gs.calc_kz()
            self.zp.set_pension(gs.salary_clear)

            # формирование выходного масива дла печати
            output = []
            output.append(['ЗП для обрахунку пенсії',
                           self.get_float('Entry5'),
                           self.get_float('Entry5')])  # 0
            output.append(['Інд. Коеф. Стажу',
                           self.get_float('Entry6'),
                           self.get_float('Entry6')])  # 1

            # todo добавить спецстаж

            kk = round(gs.period_d[0][4], 5) if gs.period_d[0][4] > 0 else round(gs.period_a[0][2]/60, 5)
            output.append(['Інд. Коеф. ЗП',
                           self.get_float('Entry7'),
                           kk])  # 2
            output.append(['Додаткові виплати (понаднормовий стаж, інфвалідність, тощо)',
                           self.get_float('Entry8'),
                           self.get_float('Entry8')])  # 3
            output.append(['Пенсія за віком',
                           self.get_float('Entry13'),
                           round(output[0][2]*output[1][2]*output[2][2], 2)])  # 4
            output.append(['Пенсія з урахуванням доплат',
                           round(output[3][1] + output[4][1], 2),
                           round(output[3][2] + output[4][2], 2)])  # 5

            if not gs.add_p:
                output.append(['Період для додавання', 'Період', 'не додано'])  # 4
            else:
                output.append(['Період для додавання',
                               '{} {}'.format(self.zp.month[gs.salary_old[gs.period_a[0][0]].get('month') - 1],
                                              gs.salary_old[gs.period_a[0][0]].get('year')),
                               '{} {}'.format(self.zp.month[gs.salary_old[gs.period_a[0][1]].get('month') - 1],
                                              gs.salary_old[gs.period_a[0][1]].get('year')),
                               gs.salary_old[gs.period_a[0][0]].get('nn'),
                               gs.salary_old[gs.period_a[0][1]].get('nn')])  # 4

            output.append(['Період для видалення',
                           '{} {}'.format(self.zp.month[self.zp.salary[gs.period_d[0][-2]].get('month') - 1],
                                          self.zp.salary[gs.period_d[0][-2]].get('year')),
                           '{} {}'.format(self.zp.month[self.zp.salary[gs.period_d[0][-1]].get('month') - 1],
                                          self.zp.salary[gs.period_d[0][-1]].get('year')),
                           gs.period_d[0][-2],
                           gs.period_d[0][-1]])  # 5

            output.append({'fio': self.Entry4.get(),
                           'tel': self.Entry9.get(),
                           'stat': self.Entry10.get(),
                           'city': self.Entry11.get()})  # 6

            mes = messagebox.askyesno("Результат розрахунку", 'Коефіцієнт ЗП - {}\n Зберегти?'.format(output[2][2]))
            if mes:
                # Сохраняю текущее состояние
                savejson = {'salary': self.zp.salary,
                            'form': [{'self.Entry4': self.Entry4.get()},
                                     {'self.Entry5': self.Entry5.get()},
                                     {'self.Entry6': self.Entry6.get()},
                                     {'self.Entry7': self.Entry7.get()},
                                     {'self.Entry8': self.Entry8.get()},
                                     {'self.Entry9': self.Entry9.get()},
                                     {'self.Entry10': self.Entry10.get()},
                                     {'self.Entry11': self.Entry11.get()},
                                     {'self.Entry12': self.Entry12.get()},
                                     {'self.Entry13': self.Entry13.get()}]}
                f_name = output[-1].get('tel').replace(' ', '') if output[-1].get('tel') else '9999999999'
                with open('./in/'+f_name+'.json', 'w', encoding='utf8') as js:
                    json.dump(savejson, js)

                # ФОрмирую печатную форму
                outreport.OutFile(self.zp.salary, output, gs.start_p).save()
                gs.clear()

                if platform == "linux" or platform == "linux2":
                    # linux
                    os.system('libreoffice ./report/'+f_name+'.xls')
                elif platform == "win32":
                    # Windows...
                    os.system('start ./report/{}.xls'.format(output[-1].get('tel')))

                mes1 = messagebox.askyesno('Наступний', 'Наступна людина?')

                if mes1:
                    self.clear('')
                    self.zp.erase_salary()
                    self.Entry4.delete(0, 'end')
                    self.Entry9.delete(0, 'end')
                    self.Entry10.delete(0, 'end')
                    self.Entry11.delete(0, 'end')
                    self.Entry6.delete(0, 'end')
                    self.Entry7.delete(0, 'end')
                    self.Entry8.delete(0, 'end')
                    self.Entry13.delete(0, 'end')
                    self.Entry12.delete(0, 'end')
                    self.Entry4.focus_set()


if __name__ == '__main__':

    root = Tk()
    top = MainForm(root)
    root.mainloop()
