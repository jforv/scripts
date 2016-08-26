#! /usr/bin/python
import os
from datetime import datetime, timedelta, date
import time
from calendar import month_abbr, monthrange

# ******** START PART 2 **** DATA COLLECTION PROCESS
__name__ == 'Josh!'
# __Version__ == '0.1.2'


class StructGen(object):

    def __init__(self, pattern='AZUR'):
        self.pattern = pattern.upper()
        # r = args[1]
        # self.year = set_year(args[2])

        self.current_month = datetime.today().month
        self.current_year = datetime.today().year

        # self.m1, self.m2 = set_range()
        self.operators = ['AIRTEL', 'AZUR', 'GTM', 'GTF', 'MOOV']
        self.services = ['ST1-TO', 'ST2-TN', 'ST3-TNT', 'ST4-ITE', 'ST5-ITS', 'ST6-TR']
        self.header = ['Service', 'Date', 'Calls', 'Minutes', 'Comments', '']
        self.link = [l for l in self.select()]
        # self.filename = 'Voix-{}-{}_{}-{}.xls' # 1 Declarant.-Origin-Destination-Month-year
        self.filename = 'OFFNET-{}_{}-{}.xls'  # 1 Declarant.-Origin-Destination-Month-year
        self.basedir = os.path.join(os.environ['HOME'], 'Matrix-Conciliation/')

    def __str__(self):
        # print self.days
        # print str(self.link)[1:-1].replace("),", ")\n")
        return 'Combinations are: ' + ''.join(str(self.link))[1:-1]

    def set_pattern(self, pattern):
        self.pattern = pattern.upper()
        return self.pattern

    # def set_range(self, r):
    #     rang = r[1][0] if len(r[1]) != 1 else r[1][0], r[1][1] + 1 if len(r[1]) == 2 else r[1][0] + 1
    #     return rang

    def set_filename(self, new_filename):
        self.filename = new_filename
        return self.filename

    def set_basedir(self, trail ='/', user=False):
        if user:
            loc = input('Selectionnez l\'emplacement de votre fichier  ')
            n_dir = os.path.join(loc, os.mkdir('Matrix-Conciliation'), trail)
        else:
            n_dir = os.path.join(self.basedir, trail)
            self.basedir = n_dir
        return self.basedir

    def days(self):
        return [d for d in self.get_monthdays()]

    def months(self, month):
        months_choices = []
        r = month
        m1, m2 = r[0] if len(r) != 1 else r[0], r[1] + 1 if len(r) == 2 else r[0] + 1
        for i in range(m1, m2):
            months_choices.append((date(self.current_year, i, 1).strftime('%b-%Y')))
        return months_choices

    def set_month(self, new_month):
        self.current_month = new_month
        return self.current_month

    def set_year(self, new_year):
        self.current_year = new_year
        return self.current_month

    def set_operators(self, new_operator_list):
        self.operators = new_operator_list
        return self.operators

    def mk_filename(self, op1, op2, m):
        return self.filename.format(op1, op2, m)


    def combine(self):
        from itertools import permutations
        perm = [p for p in permutations(self.operators, 2)]
        return perm

    def get_monthdays(self):
        month = self.current_month
        year = self.current_year
        first_date = time.mktime(time.strptime('{}-{}-01'.format(year, month), '%Y-%m-%d'))
        num_days = monthrange(year, month)[1]
        base = datetime.fromtimestamp(first_date)
        # base = datetime.today()
        dates = [base + timedelta(days=x) for x in range(0, num_days)]

        for i in range(len(dates)):
            yield datetime.strftime(dates[i].date(), '%Y-%m-%d')

    def select(self):
        operator = self.combine()
        for i in range(len(operator)):
            if self.pattern in operator[i]:
                yield operator[i]

    def service(self):
        serve = []
        for v in self.s.services:
            for d in self.days():
                serve.append(v)
                serve.append(d)
                serve.append(0)
                serve.append(0)
                serve.append('')
        return serve


# ******** END PART 1
# ******** START PART 2 **** FILE WRITTING PROCESS **********


class MakeStruct:

    def __init__(self, args):
        self.operator = args[0]
        global s
        s = StructGen(self.operator)

        self.default_folder = self.mk_folder(s.basedir)

        self.month = args[1]
        # self.year = s.set_year(args[2][0]) if isinstance([2], int) else s.set_year(args[2])
        self.year = s.set_year(args[2] if isinstance(args[2], int) else args[2][0])
        # print self.year
        self.location = args[3]
        self.mk_struct()

    def mk_folder(self, dir):
        import os, shutil as sh
        bd = os.path.join(s.basedir, dir)
        try:

            if not os.path.exists(bd):
                # sh.rmtree(bd)
                os.makedirs(bd)
                # print 'Le repertoir existait recree a {} '.format(bd)
                return bd
            else:
            #     print 'Directory exist!!'
            #     # os.makedirs(bd)
            #     # print 'Le repertoir a ete cree sur {}'.format(bd)
                return bd
        except OSError:
            print 'C\'est pas permis de creer ici'

    def make_dic(self):
        import numpy as np
        from collections import OrderedDict
        lst = [s.services] + [s.days()]
        # print lst
        dic = OrderedDict(zip(s.header, lst))
        # arr = np.array(dic)
        # print arr
        # yield dic
        return dic

    def write_file(self, file):
        from xlwt import Workbook, easyxf
        wb = Workbook()
        ws = wb.add_sheet('-')
        ws.col(1).width = 2560
        # style1 = XFStyle()
        style = easyxf('font: bold true; pattern: pattern solid, fore_colour gray25;')
        # print self.make_dic()
        c = 0
        for i in range(len(s.header)):
            # print 0,i, self.header[i]
            ws.write(0, i, s.header[i], style)
            for q, p in enumerate(s.days()):
                ws.write(c+1, 0, s.services[i])
                ws.write(c+1, 1, p)
                ws.write(c+1, 2, 0)
                ws.write(c+1, 3, 0)
                ws.write(c+1, 4, "")
                c += 1
        wb.save(file)

    def mk_struct(self):
        i = 0
        # s.set_pattern(self.pattern)
        month = s.months(self.month)
        # print month
        for p in month:
            print "\t"
            # pr.mk_folder(p)
            num = list(month_abbr).index(p.split('-')[0])
            i += 1
            s.set_month(num)
            # print os.path.join(self.default_folder, str(num)+'-'+p, self.operator)
            chdir = self.default_folder + str(num) + '-' + p + '/' + s.pattern + '/'
            self.mk_folder(chdir)
            print chdir
            for j, o in enumerate(s.link):
                print '\t',
                full_path = chdir + s.mk_filename(o[0], o[1], p)
                if not os.path.isfile(full_path):
                    print full_path
                    # print os.path.isfile(full_path)
                    self.write_file(full_path)
                else:
                    print full_path.lstrip(chdir)+': File exists'



# Default Settings
pr = StructGen()
answer = [pr.pattern, [pr.current_month], pr.current_year, pr.basedir]
questions = ['Operator', 'Month[1-January, 1-3 Jan-Mar]', 'Year[2004+]', 'Location to save [default: HOMEDIR]']
default = dict(zip(questions, answer))


def check_input(question):
    while True:
        try:
            entered = raw_input('Which {} ?: '.format(question))
            if entered in pr.operators and question == 'Operator':
                break
            elif entered not in pr.operators and question == 'Operator':
                print entered
                print 'Incorrect operator Name, please Correct.'
                print pr.operators
            if entered not in range(1, 13) and question != 'Month[1-January, 1-3 Jan-Mar]':
                print 'Still'
            else:
                print 'Month Must be in range 1-12, please Correct.'
                print '1-January, 1-3 January-March'
            # if entered < 2004:
            #     print 'Year must be > 2004, please Correct.'
            #     print '1-January, 1-3 January-March'
            # if entered < 2004:
            #     print 'Year must be > 2004, please Correct.'
            #     print '1-January, 1-3 January-March'
            #     break
            # else:
            #     break
        except ValueError as p:
            print p.__class__.__name__()


def get_parameters():
    # default = []
    res =[]
    i = 0
    for q in questions:
        # txt = raw_input('Which {} ?: '.format(q))
        txt = check_input(q)
        if txt == 0:
            i -= 1
        else:
            res.append(txt)
            i += 1
        
        # if txt.split('-')[0].isdigit() or txt.isdigit() :
        #     txt = [int(t) for t in txt.split('-')]
        #     res.append(txt)
        # elif txt.isdigit():
        #     # print len(txt)
        #     txt = [int(t) for t in txt.split('-')]
        #     res.append(txt)
        # elif len(txt) <= 0:
        #     res.append(default[q])
        # else:
        #     # print txt.split('-')[0].isdigit()
        #     res.append(txt)
        # # print questions
        # # print res
    param = '\nGenerating Structure with the parameters below:' \
          '\nOperator:{}, Month:{}, Year:{}, Saved Directory:{}'\
          .format(res[0],res[1],res[2],res[3])
    print param
    print '\n', '*'* (len(param)/2)
    MakeStruct(res)


if __name__ == '__main__':
    # pass
    get_parameters()
