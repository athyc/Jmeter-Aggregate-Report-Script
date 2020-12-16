# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import csv
from distutils.util import strtobool
from xlwt import Workbook, Alignment
from xlwt import XFStyle
import glob
from datetime import date
import re


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def letters_to_indexes(letter):
    switcher = {'A': 1,
                'B': 2,
                'C': 3,
                'D': 4,
                'E': 5,
                'F': 6,
                'G': 7,
                'H': 8,
                'I': 9,
                'J': 10,
                'K': 11,
                'L': 12,
                'M': 13,
                'N': 14,
                'O': 15,
                'P': 16,
                'Q': 17,
                'R': 18,
                'S': 19,
                'T': 20,
                'U': 21,
                'V': 22,
                'W': 23,
                'X': 24,
                'Y': 25,
                'Z': 26,
                'AA': 27,
                'AB': 28,
                'AC': 29,
                'AD': 30,
                'AE': 31,
                'AF': 32,
                'AG': 33,
                'AH': 34,
                'AI': 35,
                'AJ': 36,
                'AK': 37,
                'AL': 38,
                'AM': 39,
                'AN': 40,
                'AO': 41,
                'AP': 42,
                'AQ': 43,
                'AR': 44,
                'AS': 45,
                'AT': 46,
                'AU': 47,
                'AV': 48,
                'AW': 49,
                'AX': 50,
                'AY': 51,
                'AZ': 52
                }
    return switcher.get(letter.upper()) - 1


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def invalid(columns):
    if re.split(',', columns).__len__() != 9:
        print(f"{bcolors.FAIL}INVALID NUMBER OF COLUMNS{bcolors.ENDC}")
        return True
    for i in re.split(',', columns):

        if not i.isalpha():
            return True


def excel_writer(filenames):
    style2 = XFStyle()
    style2.num_format_str = '####.##0'
    al = Alignment()
    al.horz = Alignment.HORZ_CENTER
    al.vert = Alignment.VERT_BOTTOM
    style2.alignment = al
    style3 = XFStyle()
    style3.num_format_str = '0.00%'
    style3.alignment = al
    wb = Workbook()
    sheet1 = wb.add_sheet("Sheet 1")
    index = 0
    columns_list = ['c', 'e', 'f', 'g', 'h', 'i', 'k', 'l', 'm']
    if True:
        print(f"{bcolors.WARNING}PLEASE INPUT THE 9 COLUMN LETTERS YOU NEED SEPARATED WITH A COMMA{bcolors.ENDC}")
        print(f"{bcolors.WARNING}Example: D,A,B,C,F,G,E,K,W{bcolors.ENDC}")
        columns = input()
        while invalid(columns):
            print(f"{bcolors.FAIL}INVALID INPUT{bcolors.ENDC}")
            print(f"{bcolors.WARNING}Example: D,A,B,C{bcolors.ENDC}")
            columns = input()
        columns_list = re.split(',', columns)

    sheet1.write(index, letters_to_indexes(columns_list[0]), "Concurrent Users")
    sheet1.write(index, letters_to_indexes(columns_list[1]), "Average Total Execution Time (sec)")
    sheet1.write(index, letters_to_indexes(columns_list[2]), "90% Total Execution Time  (sec)")
    sheet1.write(index, letters_to_indexes(columns_list[3]), "Min Total Execution Time (sec)")
    sheet1.write(index, letters_to_indexes(columns_list[4]), "Max Total Execution Time (sec)")
    sheet1.write(index, letters_to_indexes(columns_list[5]), "Number of Calls")
    sheet1.write(index, letters_to_indexes(columns_list[6]), "Error Rate (%)")
    sheet1.write(index, letters_to_indexes(columns_list[7]), "Date")
    sheet1.write(index, letters_to_indexes(columns_list[8]), "Start Time")

    index = 1
    for file in filenames:
        print(file)
        numtuple = getnumbers(file)
        print(numtuple)
        sheet1.write(index, letters_to_indexes(columns_list[0]), numtuple[6])
        sheet1.write(index, letters_to_indexes(columns_list[1]), numtuple[0], style2)
        sheet1.write(index, letters_to_indexes(columns_list[2]), numtuple[1], style2)
        sheet1.write(index, letters_to_indexes(columns_list[3]), numtuple[2], style2)
        sheet1.write(index, letters_to_indexes(columns_list[4]), numtuple[3], style2)
        sheet1.write(index, letters_to_indexes(columns_list[5]), numtuple[4])
        sheet1.write(index, letters_to_indexes(columns_list[6]), numtuple[5], style3)
        sheet1.write(index, letters_to_indexes(columns_list[7]), str(date.today()))
        sheet1.write(index, letters_to_indexes(columns_list[8]), "4:42:00")

        index += 1

    wb.save("example.xls")


def getnumbers(filename):
    elapsed = list()
    success = list()
    threads = list()
    with open(filename) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                print(f'Column names are {", ".join(row)}')
                line_count += 1
            else:
                elapsed.append(int(row[1]))
                success.append(strtobool(row[7]))
                threads.append(int(row[10]))
                line_count += 1
    elapsed.sort()
    threads.sort()
    samples = len(elapsed)
    ninety_percent = elapsed[int(0.9 * (len(elapsed)) - 1)]
    error_rate = float(([i for i, x in enumerate(success) if x == 0].__len__() * 100) / samples)
    print(samples)
    print(f'Average of list: {int(sum(elapsed) / samples)}')
    print(f'90% Line {ninety_percent}')
    print(f'Min {min(elapsed)}')
    print(f'Max {max(elapsed)}')
    print(f'Error % {error_rate}')
    print(f"{bcolors.WARNING}Threads: {threads[threads.__len__() - 1]}{bcolors.ENDC}")

    return int(sum(elapsed) / samples) / 1000, ninety_percent / 1000, min(elapsed) / 1000, max(
        elapsed) / 1000, samples, error_rate / 100, threads[threads.__len__() - 1]


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    sorter = list()
    temp1 = False
    for filename in glob.glob('C:\\Users\\achatzop\\Downloads\\A12_10_Inline\\A12*'):
        temp = re.split(r'\\', filename)
        for token in re.split("_", temp[temp.__len__() - 1]):
            if token == 'aggr':
                temp1 = True
            if token.isdigit() and int(token) > 9 and temp1:
                print(token)
                sorter.append((int(token), filename))
        temp1 = False
    sorter = sorted(sorter, key=lambda x: x[0])
    excel_writer([i[1] for i in sorter])
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
