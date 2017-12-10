import json
import pyexcel as pe
from collections import defaultdict, OrderedDict
from decimal import Decimal


def zero_factory():
    return Decimal(0)


def default():
    # read data
    filename = 'C:/Users/Jason/OneDrive/Documents/Finance/Finance Planning.xlsx'
    book = pe.get_book(file_name=filename)
    sheets = book.to_dict()
    stocks = sheets['Stocks Records']

    summary = defaultdict(zero_factory)
    for row in stocks:
        if row[2] == 'SELL':
            change = Decimal(row[6])
            summary[row[4]] += change

    total_gains = Decimal(0)
    total_loss = Decimal(0)
    for key, value in summary.items():
        if value > 0:
            total_gains += value
        else:
            total_loss += value

    print('-----------------------------------------------\nTOTAL GAIN/LOSS\n-----------------------------------------------')
    ordered = OrderedDict(sorted(summary.items(), key=lambda t: t[1], reverse=True))
    for key, value in ordered.items():
        if value > 0:
            percent = abs(value/total_gains)
        else:
            percent = abs(value/total_loss)
        print('{}\t|  {:.2f}\t|  {:.1f}%'.format(key, value, percent * 100))

if __name__ == "__main__":
    default()
