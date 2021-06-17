# Importing Libraries
from tabulate import tabulate
import datetime as datetime
from tqdm import tqdm as progressbar # Visual using progress bar
import time
import pyfiglet as art
import openpyxl as excel


name = str(input("\nEnter your name:"))
name = name.capitalize()
print('\nHi,', name.capitalize())

passcode = 0x75AB
count = 0
user_pass = ""
while user_pass != passcode and count < 3:
    try:
        user_pass = int(input('Enter PIN:\n'))
    except:
        print("\nERROR\n•Password can't consists of alphabets.\n•Password must be 5 digits.\n")

    if user_pass == passcode:
        print('Access granted for', name.capitalize() + ".\n")
        break
    else:
        print('Access denied. Try again ' + name.capitalize() + ".")
        count += 1
        print('You have', 3 - count, 'counts left.\n')
        if count == 3:
            print('Tumse na ho payega', name.lower(), ':)')
            print('\033[1m' + 'USER LOCKED OUT!!')
            exit()

print('\033[1m' + 'WELCOME TO BILLING MACHINE, ' + name.upper())
# Use '\033[1m' to make statement bold like you use '*' in Whatsapp


rate = [['Gold', '500ml', 26.18],
     ['Gold', '1L', 51.35],
     ['Shakti', '500ml', 23.20],
     ['Shakti', '1L', 45.40],
     ['Cow', '500ml', 22.20],
     ['Cow', '1L', 43.40],
     ['Taaza', '500ml', 21.18],
     ['Taaza', '1L', 41.40],
     ['Amul Kool (Plastic)', '200ml * 30', 540.00],
     ['Dahi', '200g', 13.86],
     ['Dahi', '400g', 24.63],
     ['Dahi', '1KG', 56.80],
     ['Lassi (Packet)', '180ml', 9.10],
     ['Paneer', '200g', 66.00],
     ['Paneer', '1KG', 315.00]
     ]

ans1 = input('Want to know rates? [Y/N] \n')
if (ans1 == "Y") or (ans1 == "y"):
    print(tabulate(rate, headers=["Item", "Quantity", "Selling Price"], numalign="left"))
    print('\nBahut sasta hai, \nWholesale price hai!!!! \n#Exclusively_On_Softline')
else:
    print('Paisa toh ped pe uggta hai na be,', name, ':)')

ans2 = input('\nWant to start billing machine? [Y/N] \n')
if (ans2 == "Y") or (ans2 == "y"):
    try:
        for i in progressbar(range(10),
                             desc='Opening with superfast speed',
                             ascii=False, ncols=75):
            time.sleep(0.25)
        print("Complete.\n")
    except:
        print("\nERROR ~~ 'tqdm' package not found\nSkipping progress bar visualization\n\n")

    # Let's create art
    try:

        banner = art.figlet_format("WELCOME TO SOFTLINE AMUL BILLING", font="digital")
        print(banner)
    except:
        print("\nERROR ~~ 'pyfiglet' package not found\nSkipping art visualization\n\n")
        print('Opening Billing Machine...')

    # Excel Connection Starts Here!
    print('Start Order Sheet of date:\n')
    print("[1] Today\n"
          "[2] Tomorrow\n"
          "[3] Enter Date Manually\n")
    ans3 = int(input("Enter 1 or 2 or 3:\n"))
    if ans3 == 1:
        dat = today = datetime.date.today().strftime("%d-%m-%Y")
    if ans3 == 2:
        dat = tomorrow = datetime.date.today().strftime("%d-%m-%Y") + datetime.timedelta(days=1).strftime("%d-%m-%Y")
    if ans3 == 3:
        dat = manual_date = input("Enter Date Manually (DD-MM-YYYY):\n")
    try:
        excel_path = 'O-' + dat + '.xls'
    except:
        print("ERROR\n"
              "Can't find Excel sheet with date - " + dat + ".\n"
              "Please Enter Excel File Location Manually.")
        excel_path = input("\n\nEnter Excel file path \n(without quotes)(with double slashes['\']):\n")
    try:
        wb = excel.load_workbook("Excel Files\\"+excel_path)
        print('Connecting with sheets inside Excel...')
        sh1 = wb["A1 KTRA-IND"]
        sh2 = wb["A2 GANJ"]
        sh3 = wb["A3 BIDUPUR"]
        sh4 = wb['A4 MARAI']
        print('File successfully connected.\n')
    except:
        print("File path is invalid!")

    date_order_sheet = sh1["L1"].value
    print("Excel Order Sheet Date is ", date_order_sheet, '\n')

    start = input("\nStart Taking Order? [Y/N]\n")
    if (start == "Y") or (start == "y"):
        print("Let's Start taking orders,", name, "!")
    else:
        print('Bye,', name, '!')
        exit()

    sh = int(input("[1] 'A1 KTRA-IND'\n"
                   "[2] 'A2 GANJ'\n"
                   "[3] 'A3 BIDUPUR'\n"
                   "[4] 'A4 MARAI'\n"
                   "Enter 1 or 2 or 3 or 4:\n"))

    row = 5
    column = 1
    count = 1
    if sh == 1:
        print('Connected to', sh1)
        while row < 33:
            print(count)
            print("Store: ", sh1.cell(row, 2).value)
            print("Call: ", sh1.cell(row, 3).value, "\n")
            count += 1
            row += 1
    if sh == 2:
        print('Connected to', sh2)
        while row < 33:
            print(count)
            print("Store: ", sh2.cell(row, 2).value)
            print("Call: ", sh2.cell(row, 3).value, "\n")
            count += 1
            row += 1
    if sh == 3:
        print('Connected to', sh3)
        while row < 33:
            print(count)
            print("Store: ", sh2.cell(row, 2).value)
            print("Call: ", sh2.cell(row, 3).value, "\n")
            count += 1
            row += 1
    if sh == 4:
        print('Connected to', sh4)
        while row < 33:
            print(count)
            print("Store: ", sh4.cell(row, 2).value)
            print("Call: ", sh4.cell(row, 3).value, "\n")
            count += 1
            row += 1

else:
    print('Gaand mara bsdk ' + name)
    exit()
