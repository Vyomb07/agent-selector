import time
from datetime import datetime
from datetime import timedelta
import random
import pandas as pd

data = {}


def status(l_inp):
    if l_inp == 'YES':
        return True
    return False


class Ad:
    def __init__(self, inp):
        self.inp = inp
        if inp == 1:
            Ad.ad_auto()
        elif inp == 2:
            Ad.ad_man()

    # ---------------Function for adding details of agent Automatically-------------#
    @staticmethod
    def ad_auto():
        a = len(data.keys())

        df = pd.read_excel('Data.xlsx')

        for ind in df.index:
            temp = []
            name = df['Name'][ind]
            temp.append(name)

            availability = df['Availibility'][ind]
            temp.append(availability)

            tim = df['temp_time'][ind]
            tim = str(tim)
            hr, mi, ss = tim.split(':')
            c = time.struct_time((2000, 1, 1, int(hr), int(mi), 0, 3, 362, 0))
            tim2 = time.strftime("%I:%M %p", c)
            temp.append(tim2)

            role = df['Role'][ind]
            temp.append(role)

            data[ind + a] = temp

        dis_inp = input('\n\nDo you want to view your data (Yes or no) = ').upper()
        if dis_inp == 'YES':
            Dis_data.dis()
        elif dis_inp == 'NO':
            print('\n\nYour data has been added successfully....Going to Main Menu ')
            main1()

        else:
            print("Wrong Choice............Try Again !!!!!")

    # ------------------------Function for adding details manually----------------------- #

    @staticmethod
    def ad_man():
        global s
        a = len(data.keys())
        n = int(input('Enter the number of agent`s detail you want to add = '))
        for i in range(n):
            temp = []

            name = input('\nEnter Name:').capitalize()
            temp.append(name)
            print()

            s = input('Is agent Available Yes or No = ').upper()
            if status(s):
                temp.append('Y')
            else:
                temp.append('N')
            print()

            tim = input('Enter the time agent is available since (HH:MM) = ')
            hr, mi = tim.split(':')
            c = time.struct_time((2000, 1, 1, int(hr), int(mi), 0, 3, 362, 0))
            tim = time.strftime("%I:%M %p", c)
            temp.append(tim)
            print()

            role = input('Enter the role of agent = ').upper()
            temp.append(role)

            data[i + a] = temp

        dis_inp = input('Do you want to view your data (Yes or no) = ').upper()
        if dis_inp == 'YES':
            Dis_data.dis()
        elif dis_inp == 'NO':
            print('Your data has been added successfully....Going to Main Menu ')
            main1()

        else:
            print("Wrong Choice............Try Again !!!!!")


# ---------------Function for displaying details-------------##

data_time = {}


class Dis_data:
    def __init__(self, data):
        self.data = data

    @staticmethod
    def dis():
        a = bool(data)

        if not a:

            print('No Entries Found')

            c = input('Do You want to add Details(Y/N):')

            if c == 'Y':
                Menu.ad()
            elif c == 'N':
                print('Main menu.....')
                main1()

        elif a:
            print(
                "\n\n{:<15} {:<15} {:<15} {:<15}  {:<15}".format('Sr No.', 'NAME', 'Availability', 'Starting Time',
                                                                 'Role'))

            for key, value in data.items():
                n, r, p, c = value

                print("{:<15} {:<15} {:<15} {:<15} {:<15}".format(key + 1, n, r, p, c))


# -----------------Function for All available mode--------------------------##
class Mode:
    def __init__(self):
        print('\n\nAgent Selection mode are as Follow:\n1)All Available Mode\n2)Least Busy Mode\n3)Random Mode')
        inp_mode = int(input('\n\nEnter Your Choice (eg: for selecting first mode type "1") = '))
        if inp_mode == 1:
            Mode.all_avail()
        elif inp_mode == 2:
            Mode.least_busy()
        elif inp_mode == 3:
            Mode.random_mode()
        else:
            print("Try Again....")
            main1()
        try:
            Mode()
        except Exception:
            print('Invalid Input')
        finally:
            print("Press Enter to continue ...")
            input()
            Mode()

    @staticmethod
    def all_avail():
        limiter = 0
        if Counter == 1:
            print("\n\nAvailable agents are : \n")
            print("\n{:<15} {:<15} {:<15}  {:<15}".format('NAME', 'Availability', 'Starting Time', 'Role'))
            for key, value in data.items():
                name, s, tim, role = value
                if role == first_issue or second_issue:
                    print("\n{:<15} {:<15} {:<15} {:<15}".format(name, s, tim, role))
                    limiter += 1
            if limiter == 0:
                print('No Agent is available regarding your issue........\nTry Again.......')
                main1()
            input('\n\n Press Any key to go to main menu = ')
            main1()

        elif Counter == 2:
            print("\n\nAvailable agents are : \n")
            print("\n{:<15} {:<15} {:<15}  {:<15}".format('NAME', 'Availability', 'Starting Time', 'Role'))
            for key, value in data.items():
                name, stat, tim, role = value
                if role == first_issue:
                    print("\n{:<15} {:<15} {:<15} {:<15}".format(name, stat, tim, role))
                    limiter += 1
            if limiter == 0:
                print('No Agent is available regarding your issue........\nTry Again.......')
                main1()
            input('\n\n Press Any key to go to main menu = ')
            main1()

    # -----------------Function For Least_busy mode ------------------------------##

    @staticmethod
    def least_busy():
        global role
        for key, value in data.items():
            name, s2, tim, role = value

            docformat = '%I:%M %p'
            local_time = time.strftime('%I:%M %p', time.localtime())
            diff = datetime.strptime(local_time, docformat) - datetime.strptime(tim, docformat)
            con_time = timedelta(seconds=diff.seconds)
            data_time[key] = con_time
        data_time1 = data_time.copy()
        tom = 0

        def get_key(val):
            for key, value in data_time.items():
                if val == value:
                    return key

        def max1(issue):
            if tom == 1:
                for i in list(data_time.items()):
                    c = max(data_time.values())
                    temp_key = get_key(c)
                    val = data.get(temp_key)
                    one, two, three, four = val
                    if issue == four:
                        return val
                        break
                    else:
                        del data_time[temp_key]
                        max1(issue)
            elif tom == 2:
                for i in list(data_time1.items()):
                    c = max(data_time1.values())
                    temp_key = get_key(c)
                    val = data.get(temp_key)
                    one, two, three, four = val
                    if issue == four:
                        return val
                        break
                    else:
                        del data_time1[temp_key]
                        max1(issue)

        limiter = 0

        if Counter == 1:
            tom = 1
            value11 = max1(first_issue)
            for r3 in value11[3]:
                if r3 in first_issue:
                    name3, s3, t3, r33 = value11
                    print('\nRandom selected agent available for your issue = \n')
                    print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
                    print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name3, s3, t3, r33))
                    limiter += 1
                    break
                if limiter == 0:
                    print('No agent found for your issue..........\nTry Again..................')
                    main1()
            tom = 2
            value12 = max1(second_issue)
            for r4 in value12[3]:
                if r4 in first_issue:
                    name4, s4, t4, r44 = value12
                    print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name4, s4, t4, r44))
                    limiter += 1
                    break
                if limiter == 0:
                    print('No agent found for your issue..........\nTry Again..................')
                    main1()
            input('\n\n Press Any key to go to main menu = ')
            main1()

        elif Counter == 2:
            tom = 1
            value11 = max1(first_issue)
            name1, s1, t1, r1 = value11
            if role == first_issue:
                print('\nAgent available for your issue = \n')
                print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
                print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name1, s1, t1, r1))
            else:
                print('No agent found for your issue..........\nTry Again..................')
                main1()
            input('\n\n Press Any key to go to main menu = ')
            main1()

    # -----------------Function for All Random mode--------------------------##

    @staticmethod
    def random_mode():
        global val1
        global val2

        def ran_keys(issue1):
            ran_number = random.sample(data.keys(), 1)
            for i in ran_number:
                temp_val = data.get(i)
                temp_name, temp_status, temp_time, temp_role = temp_val
                if issue1 == temp_role:
                    return temp_val
                else:
                    return ran_keys(issue1)

        if Counter == 1:
            limiter = 0
            val1 = ran_keys(first_issue)

            for r3 in val1[3]:
                if r3 in first_issue:
                    name3, s3, t3, r33 = val1
                    print('\nRandom selected agent available for your issue = \n')
                    print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
                    print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name3, s3, t3, r33))
                    limiter += 1
                    break
                if limiter == 0:
                    print('No agent found for your issue..........\nTry Again..................')
                    main1()
            val2 = ran_keys(second_issue)
            limiter = 0
            for r4 in val2[3]:
                if r4 in second_issue:
                    name4, s4, t4, r44 = val2
                    print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name4, s4, t4, r44))
                    limiter += 1
                    break
                if limiter == 0:
                    print('No agent found for your issue..........\nTry Again..................')
                    main1()
            input('\n\n Press Any key to go to main menu = ')
            main1()
        elif Counter == 2:
            vall = ran_keys(first_issue)
            name3, s3, t3, r3 = vall
            if r3 == first_issue:
                print('\nRandom selected agent available for your issue = \n')
                print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
                print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name3, s3, t3, r3))
            else:
                print('No agent found for your issue..........\nTry Again..................')
                main1()
            input('\n\n Press Any key to go to main menu = ')
            main1()


# ---------------Function for adding details of agent menu -------------##
class Menu:
    def __init__(self, inp_menu):
        self.inp_menu = inp_menu
        if inp_menu == 1:
            Menu.ad()
        elif inp_menu == 2:
            Menu.issue()

    @staticmethod
    def ad():
        print('\n\nModes of adding details as follow:\n\n1)Adding details using excel sheet\n2)Adding details Manually')
        ad_inp = int(input('\n\nEnter your choice( eg : for 1st option enter "1") = '))
        if ad_inp == 1:
            Ad(1)
        elif ad_inp == 2:
            Ad(2)
        else:
            print('Wrong Choice')

            defa = input('Do you want to try again(Y or N):').upper()
            if defa == 'Y':
                Menu.ad()
            elif defa == 'N':
                print('\n\nYour data has been added successfully')
                main1()

    # ---------------Function for adding issues -------------##
    @staticmethod
    def issue():
        global first_issue
        global second_issue
        global Counter
        Counter = 0
        inp = int(input('\n\nType "1" For multi role issues \nType "2" For single role issues\n\nEnter Your choice = '))

        if inp == 1:
            print('Issue Categories are as follow :\n1)Sales\n2)Support\n3)Spanish Speaker\n')
            issue_type = input('Enter Your issues ( use ";" for separating multiple issue) = ').upper()
            first_issue, second_issue = issue_type.split(';')
            Counter = 1
        elif inp == 2:
            print('Issue Categories are as follow :\n1)Sales\n2)Support\n3)Spanish Speaker\n')
            first_issue = input('Enter Your issues (Note: you have type the issue from the given list) = ').upper()
            Counter = 2

        Mode()


# ---------------Function for Main menu -------------##
def main1():
    print('\n\n=-----Agent Selector------=')
    print('1)Add Details of Agent')
    print('2)Report Issue ')
    print('3)Exit')
    a = int(input('Enter Your Choice='))
    if a == 1:
        Menu(1)
    elif a == 2:
        Menu(2)
    elif a == 3:
        exit()
    else:
        print('Wrong Choice')
    main1()


if __name__ == '__main__':
    try:
        main1()
    except Exception:
        print('Invalid Input')
    finally:
        print("Press Enter to continue ...")
        input()
        main1()
