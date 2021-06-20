# Declaring Color Values
class Color:
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'


# Importing Libraries
try:
    from tabulate import tabulate
    import datetime as datetime
    from tqdm import tqdm as progressbar  # Visual using progress bar
    import time
    import pyfiglet as art
    import openpyxl as excel
    from sys import exit
    from twilio.rest import Client
except:
    print(Color.RED + Color.BOLD + 'ERROR:\nPython Packages not found!' + Color.END)
    print("\nInstall Python Packages to get started!")
    exit()

name = str(input("\nEnter your name:"))
name = name.capitalize()
print('\nHi,', Color.BOLD + name.capitalize() + Color.END)

# SMS API Integration with Twilio Trial
sendto = "+91" + input('\nEnter your phone number:')
sid = 'ACc514b4800253a9b8c56867d650b2e8d2'
auth_token = '28c895b0ac12186bbdc2c2b95ffd2168'
client = Client(sid, auth_token)

passcode = 0x75AB  # Password in HEX ('0x' - tells python value is in hex)

count = 0  # Count user's invalid attempts
user_pass = ""  # Value received by user

# Password Check
while user_pass != passcode and count < 3:  # Ask password from user thrice
    try:
        user_pass = int(input('Enter PIN:\n'))
    except:
        print(
            "\nERROR\n•Password can't consists of alphabets.\n•Password must be 5 digits.\n")  # If the value isn't numeric

    if user_pass == passcode:
        print('Access granted for', name.capitalize() + ".\n")
        break  # Let's the loop end & starts new block
    else:
        print('Access denied. Try again,' + name.capitalize() + ".")
        count += 1
        print('You have', 3 - count, 'counts left.\n')
        if count == 3:
            print('Let the brain rest,', name.lower(), ':)')
            print(Color.BOLD + 'USER LOCKED OUT!!' + Color.END)
            exit()
            # Throws user out of script

print(Color.BOLD + 'WELCOME TO BILLING MACHINE, ' + name.upper() + Color.END)

# try:
#     resp = client.messages.create(body="Welcome to Sahil's program, " + name.capitalize(), from_='+13126754624',
#                                   to=sendto)
# except:
#     print('Phone Number invalid or not verified!')

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
        ]  # Rate of products

ans1 = input('Want to know rates? [Y/N] \n')
if (ans1.upper() == "Y") or (ans1 == "1"):
    print(tabulate(rate, headers=[Color.BOLD + "Item", "Quantity", "Selling Price" + Color.END], numalign="left"))
    print('\n#Exclusively_On_Softline')
else:
    print('Okay,', name, ':)')

ans2 = input('\nWant to start billing machine? [Y/N] \n')

if (ans2.upper() == "Y") or (ans2 == "1"):

    # Progress Bar for vibes...
    for i in progressbar(range(10),
                         desc='Opening with superfast speed',
                         ascii=False, ncols=75):
        time.sleep(0.25)
    print("Complete.\n")

    # Let's create art
    banner = art.figlet_format("WELCOME TO SOFTLINE AMUL BILLING", font="digital")
    print(Color.BOLD + Color.GREEN + banner + Color.END)

    # Excel Connection Starts Here!
    print('Connect to EXCEL of date:\n')
    print(Color.BOLD + "[1] Today\n"
                       "[2] Yesterday\n"
                       "[3] Enter Date Manually\n"
                       "[4] USE SAMPLE FILE\n" + Color.END)
    ans3 = int(input("Enter 1 or 2 or 3:\n"))
    dat = ''
    if ans3 == 1:
        dat = today = datetime.date.today().strftime("%d-%m-%Y")
    if ans3 == 2:
        yesterday = datetime.date.today() - datetime.timedelta(days=1)
        dat = yesterday.strftime("%d-%m-%Y")
    if ans3 == 3:
        dat = manual_date = str(input("\nEnter Date Manually (DD-MM-YYYY):"))
    if ans3 == 4:
        dat = "15-06-2021"
        print("Sample Excel File Date is: " + dat)
        print("Connecting...\n")
    try:
        excel_path = 'O-' + dat + '.xlsx'
        wb = excel.load_workbook("Excel Files/" + excel_path)
    except:
        print(Color.RED + "ERROR\n"
                          "Can't find Excel sheet with date - " + dat + ".\n"
                                                                        "Please Enter Excel File Location Manually." + Color.END)
        excel_path = input("\n\nEnter Excel File Path \n(without quotes)(with double slashes):\n")
        wb = excel.load_workbook(excel_path)
    try:
        print('Connected with Excel - ' + Color.BOLD + excel_path + Color.END)
        print('Connecting with sheets inside Excel...')
        sh1 = wb["A1 KTRA-IND"]
        sh2 = wb["A2 GANJ"]
        sh3 = wb["A3 BIDUPUR"]
        sh4 = wb['A4 MARAI']
        print(Color.BLUE + Color.BOLD + 'File successfully connected.\n' + Color.END)
    except:
        print(Color.RED + "ERROR:\n File is invalid!" + Color.END)

    # After Excel Connected
    date_order_sheet = sh1["L1"].value
    print("EXCEL ORDER SHEET ", Color.BOLD + date_order_sheet + Color.END, '\n')

    start = input("Start Taking Order? [Y/N]")
    if (start.upper() == "Y") or (start == "1"):
        print("Let's Start taking orders,", name, "!\n")
    else:
        print('Bye,', name, '!')
        exit()

    print(Color.BOLD + "[1] 'A1 KTRA-IND'\n"
                       "[2] 'A2 GANJ'\n"
                       "[3] 'A3 BIDUPUR'\n"
                       "[4] 'A4 MARAI'" + Color.END)
    sh = int(input("Enter 1 or 2 or 3 or 4:\n"))

    row = 5
    column = 1
    count = 1

    if sh == 1:
        print('\nConnected to', sh1)
        sh1_list = []
        while row < 33:
            sh1_list.append([count, sh1.cell(row, 2).value, sh1.cell(row, 3).value])
            row += 1
            count += 1
        print(tabulate(sh1_list, headers=["SN No.", "Store", "Phone No."]))

        count = 0
        while count < 30:
            ans4 = int(input('\nENTER NUMBER: \n')) - 1
            # resp = client.messages.create(
            #     body="\nStore: " + str(sh1_list[ans4][1]) + "\nCall: " + str(sh1_list[ans4][2]),
            #     from_='+13126754624', to=sendto)
            # print('---NOTIFICATION SENT')

            print('\n\n' + Color.BOLD + sh1_list[ans4][1] + '\n' + str(sh1_list[ans4][2]) + Color.END)
            print('\nEnter Orders: \n')
            count2 = 0
            while count2 < len(rate):
                print('[' + str(count2 + 1) + '] ' + str(rate[count2][0]) + ' (' + str(rate[count2][1]) + ') - ' + str(rate[count2][2]))
                count2 += 1

            count += 1
        else:
            print('Attempts Exceeded, Try Again')

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
    print('Great, ' + name)
    exit()
