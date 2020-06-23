import time
from datetime import datetime
from datetime import timedelta
import random
import pandas as pd
from mainfile import *

data = {}


# ---------------Function for adding details of agent Automatically-------------#
def status(l):
    if l == 'YES':
        return bool(1)
    return bool(0)


def ad_auto():
    a = len(data.keys())

    df = pd.read_excel('C:/Users/Acer/Desktop/Book1.xlsx')

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
        dis()
    elif dis_inp == 'NO':
        print('\n\nYour data has been added successfully....Going to Main Menu ')
        main1()

    else:
        print("Wrong Choice............Try Again !!!!!")


# ---------------Function for adding details of agent Manually-------------##


def ad_man():
    global s
    global time
    a = len(data.keys())
    n = int(input('Enter the number of agent`s detail you want to add = '))
    for i in range(n):
        temp = []

        name = input('\nEnter Name:').capitalize()
        temp.append(name)
        print()

        s = input('Is agent Available Yes or No = ').upper()
        if status(s):
            temp.append('Yes')
        else:
            temp.append('No')
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
        dis()
    elif dis_inp == 'NO':
        print('Your data has been added successfully....Going to Main Menu ')
        main1()

    else:
        print("Wrong Choice............Try Again !!!!!")


# ---------------Function for displaying details-------------##


def dis():
    a = bool(data)

    if not a:

        print('No Entries Found')

        c = input('Do You want to add Details(Y/N):')

        if c == 'Y':
            ad()
        elif c == 'N':
            print('Main menu.....')
            main1()

    elif a:
        print(
            "\n\n{:<15} {:<15} {:<15} {:<15}  {:<15}".format('Sr No.', 'NAME', 'Availability', 'Starting Time', 'Role'))

        for key, value in data.items():
            n, r, p, c = value

            print("{:<15} {:<15} {:<15} {:<15} {:<15}".format(key + 1, n, r, p, c))


# ---------------Function for adding issues -------------##


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

    modes()


# -----------------Function for All available mode--------------------------##

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


# -----------------Function For Least_busy mode ------------------------------##

data_time = {}


def get_key(val):
    for key, value in data_time.items():
        if val == value:
            return key


def least_busy():
    global role
    for key, value in data.items():
        name, s2, tim, role = value

        docformat = '%I:%M %p'
        local_time = time.strftime('%I:%M %p', time.localtime())
        diff = datetime.strptime(local_time, docformat) - datetime.strptime(tim, docformat)
        con_time = timedelta(seconds=diff.seconds)
        data_time[key] = con_time

    temp_local = data_time.copy()

    def max1(issue,issue2=None):
        for key, value in list(temp_local.items()):
            c = max(temp_local.values())
            temp_key = get_key(c)
            val = data.get(temp_key)
            one, two, three, four = val
            if Counter == 1:
                if issue or issue2 == four:
                    return val
                    break
                else:
                    del temp_local[temp_key]
                    max1(first_issue,second_issue)
            elif Counter == 2:
                if issue == four:
                    return val
                    break
                else:
                    del temp_local[temp_key]
                    max1(first_issue)

    if Counter == 1:
        value11 = max1(first_issue,second_issue)
        name1, s1, t1, r1 = value11
        if role == first_issue or second_issue:
            print('\nAgent available for your issue = \n')
            print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
            print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name1, s1, t1, r1))
        else:
            print('No agent found for your issue..........\nTry Again..................')
            main1()

    elif Counter == 2:
        value11 = max1(first_issue)
        name1, s1, t1, r1 = value11
        if role == first_issue:
            print('\nAgent available for your issue = \n')
            print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
            print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name1, s1, t1, r1))
        else:
            print('No agent found for your issue..........\nTry Again..................')
            main1()


# -----------------Function for All Random mode--------------------------##
def ran_keys(issue1,issue2=None):
    ran_number = random.sample(data.keys(), 1)
    for i in ran_number:
        val1 = data.get(i)
    n3, s3, t3, r3 = val1
    if issue1 == r3:
        return val1
    else:
        if Counter == 1:
            return ran_keys(first_issue,second_issue)
        elif Counter == 2:
            return ran_keys(first_issue)


def random_mode():
    if Counter == 1:
        vall = ran_keys(first_issue,second_issue)
        name3, s3, t3, r3 = vall
        if r3 == first_issue or second_issue:
            print('\nRandom selected agent available for your issue = \n')
            print('{:<15}  {:<15}  {:<15}  {:<15}'.format('Name', 'Availability', 'Available Since', 'Role'))
            print('{:<15}  {:<15}  {:<15}  {:<15}'.format(name3, s3, t3, r3))
        else:
            print('No agent found for your issue..........\nTry Again..................')
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


# -----------------Function for modes---------------------------------------##

def modes():
    select_mode = int(input(
        '\n\nAgent Selection Mode are as follow:\n1)All Available\n2)Least Busy\n3)Random\n\nEnter your choice ='))

    if select_mode == 1:
        all_avail()
    elif select_mode == 2:
        least_busy()
    elif select_mode == 3:
        random_mode()
    else:
        print('Wrong Choice!!!!')
        print('Going back to Main menu........')


# ---------------Function for adding details of agent menu -------------##

def ad():
    print(
        '\n\nModes of adding details as follow:\n\n1)Adding details using excel sheet\n2)Adding details Manually')
    ad_inp = int(input('\n\nEnter your choice( eg : for 1st option enter "1") = '))
    if ad_inp == 1:
        ad_auto()
    elif ad_inp == 2:
        ad_man()
    else:
        print('Wrong Choice')

        defa = input('Do you want to try again(Y or N):').upper()
        if defa == 'Y':
            ad()
        elif defa == 'N':
            print('\n\nYour data has been added successfully')
            main1()


# ---------------Function for Main menu -------------##
def main1():
    print('\n\n-----Agent Selection------')
    print('1)Add Details of Agent')
    print('2)Report Issue ')
    print('3)Exit')
    a = int(input('Enter Your Choice='))
    if a == 1:
        ad()
    elif a == 2:
        issue()
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
