import csv
import os

from calendar import monthrange
from datetime import datetime


class ReportsGeneratorException(Exception):
    pass


class ReportsGenerator(object):

    MONTHS = {
        'january': 1,
        'february': 2,
        'march': 3,
        'april': 4,
        'may': 5,
        'june': 6,
        'july': 7,
        'august': 8,
        'september': 9,
        'october': 10,
        'november': 11,
        'december': 12
    }

    SHEET_HEADERS = {
        'Minutes': (4, 8),
        'Hours': (5, 9),
        'Total': (7,)
    }

    SHEET_SUBHEADERS = {
        'Date': (0,),
        'Start': (1,),
        'End': (2,),
        'Description': (3,),
        'Worked': (4, 5, 8, 9)
    }

    COLS_NUMBER = \
        9

    MINUTES_CALCULATOR = \
        '=if(C{0},if(C{0}>=B{0},(hour(C{0})*60+minute(C{0}))-' + \
        '(hour(B{0})*60+minute(B{0})),24*60 + (hour(C{0})*60+minute(C{0}))' + \
        '-(hour(B{0})*60+minute(B{0}))),0)'

    HOURS_CALCULATOR = \
        '=E{0}/60'

    def __init__(self, month, empty_rows=0):
        self.month = month
        if self.month.lower() not in self.MONTHS:
            raise ReportsGeneratorException('Undefined month: ' + \
                                            '{0}'.format(self.month))

        if isinstance(empty_rows, str):
            if empty_rows.isdigit():
                empty_rows = int(empty_rows)
            else:
                raise ReportsGeneratorException('Parameter empty_rows ' + \
                                                'should be a number.')

        if not isinstance(empty_rows, int):
            raise ReportsGeneratorException('Parameter empty_rows ' + \
                                            'should be a number.')

        if empty_rows >= 8:
            raise ReportsGeneratorException('Parameter empty_rows ' + \
                                            'should be not greater than 8.')

        self.empty_rows = empty_rows
        self.current_year = datetime.now().year
        self.current_month = self.MONTHS[self.month.lower()]

        self.days_amount = monthrange(self.current_year, self.current_month)[1]
        self.output_file = os.getcwd() + '/{0}_{1}.csv'.format(self.month, \
                                                          self.current_year)

    def template_generator(self):
        with open(self.output_file, 'wb') as f:
            wr = csv.writer(f, quoting=csv.QUOTE_ALL)
            wr.writerow(self._sheet_headers(self.SHEET_HEADERS))
            wr.writerow(self._sheet_headers(self.SHEET_SUBHEADERS))
            wr.writerow(['' for i in range(0, 9)])
            row_number = 0

            for day in range(1, self.days_amount + 1):
                date = '%02d/%02d' % (day, self.current_month)

                if day == 1:
                    wr.writerow([date, '', '', '', \
                                 self.MINUTES_CALCULATOR.format(day + 3), \
                                 self.HOURS_CALCULATOR.format(day + 3), '', '', \
                                 '=SUM(E4:E{0})'.format(self.days_amount + 4), \
                                 '=SUM(F4:F{0})'.format(self.days_amount + 4)])
                else:
                    if row_number != 0:
                        rn = row_number
                    else:
                        rn = day + 3
                    wr.writerow([date, '', '', '', \
                                 self.MINUTES_CALCULATOR.format(rn), \
                                 self.HOURS_CALCULATOR.format(rn), \
                                 '', '', '', ''])

                if self.empty_rows != 0:
                    for i in range(1, self.empty_rows + 1):
                        if row_number != 0:
                            rn = row_number + i
                        else:
                            rn = day + 3 + i
                        wr.writerow(['', '', '', '', \
                                     self.MINUTES_CALCULATOR.format(rn), \
                                     self.HOURS_CALCULATOR.format(rn), \
                                     '', '', '', ''])

                    if row_number != 0:
                        row_number = rn + 1
                    else:
                        row_number = self.empty_rows + 4 + day

    def _sheet_headers(self, headers):
        cols = []
        cols_nums = [e for t in tuple(headers.values()) for e in t]

        for i in range(0, self.COLS_NUMBER + 1):
            if i not in cols_nums:
                cols.append('')
            else:
                cols.append(self._get_item(headers, i))

        if not cols:
            raise ReportGeneratorException('Empty list of headers!')

        return cols

    def _get_item(self, ddict, ind):
        for n, v in ddict.iteritems():
            if ind in v:
                return n

        raise ReportsGeneratorException('No header under {0} ' + \
                                        'index.'.format(ind))

