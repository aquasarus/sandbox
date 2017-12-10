import os
import pyexcel as pe
import openpyxl as op
import csv


ETFS = [
    'VBK',
    'VGT',
    'IBB',
    'IVV',
    'QCLN',
    'ICLN',
    'BOTZ',
    'GAMR'
]


def main():
    # read transaction data
    filename = 'C:/Users/Jason/Downloads/transactions.csv'
    sheet = pe.get_sheet(file_name=filename)

    data = []
    rows = sheet.get_array()

    # skip first and last lines
    for i in range(1, len(rows) - 1):
        row = rows[i]

        if row[4] in ETFS:
            continue # skip ETFs
        elif row[7] < 0 and row[3] < 1:
            continue # skip dividend reinvestments

        data.append({
            'DATE': row[0],
            'QUANTITY': row[3],
            'TICKER': row[4],
            'PRICE': row[5],
            'COMMISSION': row[6],
            'AMOUNT': row[7]
        })

    # remove transactions file
    os.remove(filename)

    # load tracker excel sheet
    filename = 'C:/Users/Jason/OneDrive/Documents/Finance/Allowance Tracker.xlsx'
    wb = op.load_workbook(filename)
    sheet = wb['Transactions']

    row = sheet.max_row + 1 # NOTE: this only works if there is existing data
    # row = 2

    for record in data:
        sheet.cell(row=row, column=1).value = record['DATE']
        sheet.cell(row=row, column=2).value = record['AMOUNT']

        if record['AMOUNT'] > 0:
            sheet.cell(row=row, column=3).value = 'SELL'
        else:
            sheet.cell(row=row, column=3).value = 'BUY'

        sheet.cell(row=row, column=4).value = record['TICKER']
        sheet.cell(row=row, column=5).value = record['QUANTITY']
        sheet.cell(row=row, column=6).value = record['PRICE']
        sheet.cell(row=row, column=7).value = record['COMMISSION']
        row += 1 # increment

    wb.save(filename)

if __name__ == "__main__":
    main()
