import csv
import time
from datetime import datetime, timedelta
import os
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfilenames, askopenfilename, asksaveasfilename
from workbook import *
from sql import *


def tkinter_module():
    root = Tk()
    root.geometry('650x200')
    hour_var = IntVar()
    minute_var = IntVar()
    minute_end_var = IntVar()

    def select_files():
        homedir = os.path.expanduser("~")
        types = [('CSV', '*.csv')]
        global files
        files = askopenfilenames(filetypes=types, defaultextension=types,
                                 initialdir=os.path.join(homedir,
                                                         'C:/Users/asus/PycharmProjects/ATTENDANCE/DATA'),
                                 title="Select files")
        print(files, 'These are the list of files\n')
        # following is wrong assumption as the file of previous hour could be downloaded in next hour
        # todo it's temp remove the following code
        # # taking only last file of the time period by taking consideration of hour and assumptions files are in
        # # ascending time order
        # temp = [[file, datetime.strptime(time.ctime(os.path.getctime(file)), '%a %b %d %H:%M:%S %Y')] for file in files]
        # print(temp)
        # i = 0
        # while i < len(temp):
        #     if i != 0:
        #         c = iter(['%d', '%b', '%Y', '%H'])
        #         for j in range(4):
        #             t = next(c)
        #             if temp[i][1].strftime(t) != temp[i - 1][1].strftime(t):
        #                 break
        #         else:
        #             temp.pop(i - 1)
        #             i -= 1
        #     i += 1
        # temp = [i[0] for i in temp]
        # files = list(temp)
        return files

    def write_where():
        homedir = os.path.expanduser("~")
        types = [('Excel', '*.xlsx')]
        global file
        file = askopenfilename(filetypes=types, defaultextension=types,
                               initialdir=os.path.join(homedir,
                                                       'C:\\Users\\asus\\PycharmProjects\\Attendance Project pp'),
                               title="Select the file")
        print(file, 'file chosen\n')

        return file

    def save():
        homedir = os.path.expanduser("~")
        types = [('Excel', '*.xlsx')]
        global file
        file = asksaveasfilename(filetypes=types, defaultextension=types,
                                 initialdir=os.path.join(homedir, ''),
                                 title="Create Class list")
        print(file, 'CREATED\n')
        if file:
            types = [('CSV', '*.csv')]
            global file1
            file1 = askopenfilename(filetypes=types, defaultextension=types,
                                    initialdir=os.path.join(homedir, 'NEW DATA'),
                                    title='Choose file to upload')
            print(file1, 'file chosen\n')

    def submit():
        if ch1.get() == 1:
            making(doing(file1), file)
        elif ch1.get() == 2:
            minutes = int(minute_entry.get())
            hour = int(hour_entry.get())
            minute_end = int(minute_end_entry.get())

            t = 0
            while t < len(files):
                print(files[t])
                j = doing(files[t], minutes, hour, minute_end)
                write1(j[0], j[1], file)
                t += 1
        # root.destroy()

    def view():
        if ch1.get() == 1:
            savebtn.grid(row=2, column=3)
            info_label.grid(row=1, column=2)
            sub_btn.grid(row=6, column=4)
            write_where_btn.grid_forget()
            minute_entry.grid_forget()
            minute_end_entry.grid_forget()
            minute_label.grid_forget()
            minute_end_label.grid_forget()
            hour_entry.grid_forget()
            hour_label.grid_forget()
            select_file_btn.grid_forget()


        # elif ch1.get() == 3:
        #     global buttons_classlist, classlists
        #     classlists = database_view()
        #     buttons_classlist = {}
        #     classlist_tables = {}
        #     for i in range(len(classlists)):
        #         buttons_classlist['b%s' % (i + 1)] = Radiobutton(root, text=classlists[i][0], variable=ch2,
        #                                                          value=float('3.%s' % (i + 1)), command=view)
        #         buttons_classlist['b%s' % (i + 1)].grid(row=2, column=i + 1)
        #         classlist_tables[classlists[i][0]] = {'value': float('3.%s' % (i + 1)), 'tables':table_list(classlists[i][0])}
        #     print(classlist_tables)
        #     sub_btn.grid(row=6, column=4)
        #     write_where_btn.grid_forget()
        #     minute_entry.grid_forget()
        #     minute_label.grid_forget()
        #     select_file_btn.grid_forget()
        #     savebtn.grid_forget()
        #     info_label.grid_forget()
        elif ch1.get() == 2:
            minute_label.grid(row=1, column=0)
            minute_end_label.grid(row=3, column=0)
            minute_entry.grid(row=1, column=1)
            minute_end_entry.grid(row=3, column=1)
            hour_label.grid(row=1, column=2)
            hour_entry.grid(row=1, column=3)
            write_where_btn.grid(row=4, column=2)
            select_file_btn.grid(row=5, column=2)
            sub_btn.grid(row=6, column=3)
            savebtn.grid_forget()
            info_label.grid_forget()
            # for i in globals().keys():
            #     if i == 'buttons_classlist':
            #         for i in buttons_classlist:
            #             buttons_classlist[i].deselect()
            #             buttons_classlist[i].grid_forget()

    minute_label = Label(root, text='minutes', font=('calibre', 10, 'bold'))
    minute_end_label = Label(root, text='meeting end time pattern', font=('calibre', 10, 'bold'))
    hour_label = Label(root, text='hour', font=('calibre', 10, 'bold'))
    info_label = Label(root, text='''1. You need to create a file in which your classlist would be saved 
2. Then you upload the file from which you want classlist to be created''',
                       font=('Comic Sans MS', 10, 'bold'), foreground='RED', relief='raised',
                       wraplength=250)
    minute_entry = Entry(root, textvariable=minute_var, font=('calibre', 10, 'normal'))
    minute_end_entry = Entry(root, textvariable=minute_end_var, font=('calibre', 10, 'normal'))
    hour_entry = Entry(root, textvariable=hour_var, font=('calibre', 10, 'normal'))
    write_where_btn = Button(root, text='write where', command=write_where)
    savebtn = Button(root, text='create where', command=save)
    global ch1
    ch1 = IntVar()
    # ch2 = DoubleVar()
    Button1 = Radiobutton(root, text='make', variable=ch1, value=1, command=view)
    Button2 = Radiobutton(root, text='write', variable=ch1, value=2, command=view)
    # Button3 = Radiobutton(root, text='view data', variable=ch1, value=3, command=view)
    select_file_btn = Button(root, text='Select files', command=select_files)

    sub_btn = Button(root, text='Submit', command=submit)

    Button1.grid(row=0, column=0)
    Button2.grid(row=0, column=1)
    # Button3.grid(row=0, column=2)
    root.mainloop()


def correcting_data(filename):
    os.rename(filename, filename[:-3] + 'txt')
    with open(filename[:-3] + 'txt', 'r+') as f:
        c = f.readlines()
        data = []
        for i in range(len(c) - 1):  # removing last blank line

            if i == 0:
                c[i] = c[i][2:]
            c[i] = c[i][:-2]
            c[i] = c[i].replace('\x00', '')
            c[i] = c[i].split('\t')
            data.append(c[i])
    os.rename(filename[:-3] + 'txt',
              filename)  # os.remove(filename[:-3] + 'txt') and don't take inside with as file is open there
    with open(filename[:-4] + 'my.csv', 'w+', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(data)
    file_to_upload = filename[:-4] + "my.csv"
    return file_to_upload


def reading_csv(filename):
    # initializing the titles and rows list

    with open(filename, 'r') as csvfile:
        fields = []
        rows = []
        # creating a csv reader object
        csvreader = csv.reader(csvfile)
        # extracting field names through first row
        fields = next(csvreader)

        # extracting each data row one by one
        for row in csvreader:
            rows.append(row)
        # get total number of rows
        # print("Total no. of rows: %d" % csvreader.line_num)
        # printing the field names
        # print('Field names are: ' + ', '.join(field for field in fields))
    return fields, rows


def doing(file_to_upload, minutes=45, hour=0, minute_end=15):
    v = time.ctime(os.path.getctime(file_to_upload))
    print(v)
    file_to_upload = correcting_data(file_to_upload)
    fields, rows = reading_csv(file_to_upload)
    if len(fields) != 3:
        for i in range(len(rows)):
            if 'Full Name' in rows[i]:
                fields = rows.pop(i)
                break #required as after popping len rows is not same
    rows.sort(key=lambda z: z[0])  # sorting on basis of names
    if ch1.get() == 1:
        time_use = list(dict.fromkeys([i[0] for i in rows]))
        return time_use

    # meeting_end_time
    code = '%m/%d/%Y, %I:%M:%S %p'
    k = []
    # correcting errors
    # error type 1
    try:
        if len(rows[
                   0]) == 3:  # due to the statement 'The data could be incomplete ....' in starting and there could be blanks
            for x in rows:
                k.append(datetime.strptime(x[2], code))
        else:
            while len(rows[0]) != 3:
                print(rows.pop(0))
            for x in rows:
                k.append(datetime.strptime(x[2], code))
            print('No error occurred')
    except Exception as e:
        print(e)
        while rows[0] == ['']:
            i = 0
            while i < len(rows):
                if rows[i] == ['']:

                    rows.pop(i)
                else:

                    rows[i][2] = str(rows[i][2])
                i += 1

        for x in rows:
            x[2]=x[2][1:-1]
            k.append(datetime.strptime(x[2], code))
    p = max(k)
    print(p, 'first')
    first = max(k)
    while p.strftime("%M") in ([str(i) for i in range(30, 45)]) and len(k) != 2:
        r = k.pop(k.index(p))
        p = max(k)
        print(r, p)
    second = p
    p = p.replace(minute=minute_end, second=0)
    s = p - timedelta(minutes=minutes, hours=hour)
    print(p, s)
    ctr2 = 0
    for i in rows:
        if i[1] in ('Joined before', 'Joined',):
            if s + timedelta(minutes=20) > datetime.strptime(i[2], code) > s - timedelta(minutes=15):
                ctr2 += 1
        if ctr2 == 3:
            meeting_end_time = p.strftime(code)
            meeting_start_time = s.strftime(code)
            break
    else:
        p = first.replace(minute=minute_end, second=0)
        s = p - timedelta(minutes=minutes, hours=hour)
        for i in rows:#it is for the case if there is no entry after 30 min of actual starting time
            if i[1] in ('Joined before', 'Joined',):
                if s + timedelta(minutes=20) > datetime.strptime(i[2], code) > s - timedelta(minutes=15):
                    ctr2 += 1
            if ctr2 == 3:
                break
        else:
            p = second.replace(minute=minute_end, second=0)+timedelta(hours=1)
            s = p - timedelta(minutes=minutes, hours=hour)
        meeting_end_time = p.strftime(code)
        meeting_start_time = s.strftime(code)
    # this is removing entries in rows that are removed from k in above while loop
    if ctr2 == 3:
        ctr = 0
        while ctr < len(rows):
            if datetime.strptime(rows[ctr][2], code) > p + timedelta(minutes=15):
                print(rows.pop(ctr))
            ctr += 1
    print(meeting_start_time, '--START')
    print(meeting_end_time, '--END')
    # calculating time in the meeting
    time_per_user = timedelta()
    time_use = []
    j = 0
    # print(rows)
    print('reach')
    while j < len(rows):
        if j + 1 != len(rows) and rows[j][1] == 'Joined before' and datetime.strptime(rows[j][2], code) <= s:
            if rows[j][0] == rows[j + 1][0] and rows[j + 1][1] == 'Left':  # bakwas entry as joined and left before starting
                if datetime.strptime(rows[j + 1][2], code) <= s:
                    rows.pop(j)
                    rows.pop(j)
                    if j==len(rows):
                        break
                else:
                    temp = rows.pop(j)
                    rows.insert(j, [temp[0], 'Joined', meeting_start_time])
            else:
                temp = rows.pop(j)
                rows.insert(j, [temp[0], 'Joined', meeting_start_time])
        if j + 1 == len(rows) and rows[j][1] == 'Joined before' and datetime.strptime(rows[j][2], code) <= s:
            temp = rows.pop(j)
            rows.insert(j, [temp[0], 'Joined', meeting_start_time])

        if j + 1 != len(rows) and rows[j][1] == 'Joined' and datetime.strptime(rows[j][2],
                                                                               code) <= s:  # same case as with joined before
            temp = rows.pop(j)
            if rows[j][1] == 'Left' and datetime.strptime(rows[j][2], code) <= s:
                rows.pop(j)
            else:
                rows.insert(j, [temp[0], 'Joined', meeting_start_time])
        if j + 1 == len(rows) and rows[j][1] == 'Joined' and datetime.strptime(rows[j][2], code) <= s:
            temp = rows.pop(j)
            rows.insert(j, [temp[0], 'Joined', meeting_start_time])
        if j + 1 != len(rows) and rows[j][1] == 'Joined' and datetime.strptime(rows[j][2],
                                                                               code) >= p:  # removing joined entries after meeting ending
            temp = rows.pop(j)
            if j!=0 and temp[0] == rows[j - 1][0] and rows[j-1][1] != 'Left':
                rows.insert(j, [rows[j - 1][0], 'Left', meeting_end_time])
        if j + 1 == len(rows) and rows[j][1] == 'Joined' and datetime.strptime(rows[j][2],
                                                                                code) >= p:
            temp = rows.pop(j)
            if temp[0] == rows[j - 1][0] and rows[j-1][1] != 'Left':
                rows.insert(j, [rows[j - 1][0], 'Left', meeting_end_time])
            else:
                break
        # parsing each column of a row
        if j + 1 != len(rows) and j != len(rows) and rows[j][0] != rows[j + 1][0] and rows[j][1] != 'Left':
            rows.insert(j + 1, [rows[j][0], 'Left', meeting_end_time])  # length of rows are increasing thus while
        row = rows[j]
        # checkig last value
        if j + 1 == len(rows) and rows[j][1] != 'Left':
            rows.insert(j + 1, [rows[j][0], 'Left', meeting_end_time])
        # calculating the times
        if j >= 1 and row[1] == 'Left':
            local_time = datetime.strptime(row[2], code) - datetime.strptime(rows[j - 1][2], code)
            time_per_user += local_time
            if j + 1 != len(rows):
                if rows[j + 1][0] != rows[j][0]:
                    time_use.append([row[0], str(time_per_user)])
                    time_per_user = timedelta()
            else:
                time_use.append([row[0], str(time_per_user)])
                time_per_user = timedelta()
        # # printing list of students
        # for i in range(3):
        #     print("{:.^35}".format(row[i]), end="-")
        # print('\n',j)
        j += 1
    # deleting names
    i = 0
    while i < len(time_use):
        if time_use[i][0] in ['Ranjul Rastogi-GU0314111615', 'Suniti Rastogi']:
            time_use.pop(i)
        i += 1
    print('Deleting names')
    # print(time_use)
    # quit()
    print(meeting_end_time, type(meeting_end_time))
    return time_use, meeting_end_time, code


def making(time_use, name_of_file):
    making_classlist(time_use, name_of_file)
    # database_creation(name_of_database(name_of_file))


def write1(time_use, meeting_end_time, name_of_file):
    write(meeting_end_time, time_use, name_of_file)
    # todo only create tables if class match


tkinter_module()
