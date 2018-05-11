from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
import os
import pandas as pd

### For now, though this will be replaced by pandas.

### This code is written to help automate the scheduling process.


### Preset Constants
Hours = ['10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8']
Days = ['Monday','Tuesday','Wednesday','Thursday','Friday']



#Folder = 'C:\\Users\\bgkodalen\\Desktop\\TA\\Tutoring Center\\Cterm2018Schedule'
####


### The actual GUI
root = Tk()
root.title("Tutoring Schedule")
#root.geometry('100x500')

lbl = Label(root, text="This project should help make the scheduling process easier.",font = ("Arial Bold",10))
lbl.grid(row=0,column=0,columnspan=100,sticky=W+E+N+S)

global init
global final
global possibilities
global unscheduled
global Tutorinfo
init = 0

### Close button
def quit(window):
    window.destroy()
btn = Button(root,text = "Close",command = lambda: quit(root))
btn.grid(column = 99, row = 99)


Label(master=root,text="Direct this program to the folder containing the list of tutors and all tutor schedules.\n (The schedules should be in a subfolder of the specified folder.)").grid(row = 2,column = 2)
### Browse for list of tutors
global Slistcheck
Slistcheck = 0
def findtutorlist():
    global Slistcheck
    global Tlist_path
    global Slist_path
    filename = filedialog.askdirectory()
    Tlist_path.set(filename)
    Slist_path.set(filename+'/Tutor Schedules/')
    print(filename)
    Slistcheck = 1

Tlist_path = StringVar()
Slist_path = StringVar()
lbl1 = Label(master = root, textvariable=Tlist_path)
lbl1.grid(row=99,column = 2)
Tlist = Button(text = "Browse", command=findtutorlist)
Tlist.grid(row = 99, column = 1)






def update_pos():
    global Tutorinfo
    oneposs = {day:{hour:[] for hour in Hours} for day in Days}
    for tutor in Tutorinfo:
        if Tutorinfo[tutor]['scheduled'] < Tutorinfo[tutor]['hours']:
            for day in Days:
                for hour in Hours:
                    try:
                        if Tutorinfo[tutor]['pref'][day][hour] == 1:
                            oneposs[day][hour].append(tutor+' ('+Tutorinfo[tutor]['role']+')')
                    except:
                        print(Tutorinfo[tutor])
    return oneposs 

def update_pref(schedule):
    global Tutorinfo
    for tutor in Tutorinfo:
        Tutorinfo[tutor]['scheduled'] = 0
    for day in Days:
        for hour in Hours:
            for shift in [1,2]:
                tutor = schedule[day][hour][shift].get().split()
                if tutor[0]!='None':
                    tutor = tutor[0]+' '+tutor[1]
                    Tutorinfo[tutor]['scheduled'] +=1
    

def Showschedule(tutor):
    global Tutorinfo
    S = Toplevel(root)
    for i,day in enumerate(Days):
        Label(S,text = day).grid(row = 0, column = i+1)
        for j,hour in enumerate(Hours):
            if i==0:
                Label(S,text = hour).grid(row = j+1,column = 0)
            Label(S,text = Tutorinfo[tutor]['pref'][day][hour]).grid(row = j+1,column = i+1)
    quitbtn = Button(S, text = "Close",command = lambda: quit(S))
    quitbtn.grid(row = 99,column = 99)

def update():
    global Tutorinfo
    global final
    global init
    global possibilities
    global unscheduled
    if Slistcheck and (init==0):
        TAlist = pd.read_excel(Tlist_path.get()+"/List of Tutors.xlsm", sheet_name = "Sheet1")['TA']
        PLAlist = pd.read_excel(Tlist_path.get()+"/List of Tutors.xlsm", sheet_name = "Sheet2")['PLA']
        Tutorinfo = {}
        for tutor in TAlist:
            if type(tutor) == str:
                name = tutor.split(', ')
                Tutorinfo[tutor] = {'firstname':name[1],'lastname':name[0],'hours':2,'scheduled':0,'role':'TA'}
        for tutor in PLAlist:
            if type(tutor) == str:
                name = tutor.split(', ')
                Tutorinfo[tutor] = {'firstname':name[1],'lastname':name[0],'hours':1,'scheduled':0,'role':'PLA'}
        for tutor in Tutorinfo:
            info = Tutorinfo[tutor]
            try:
                possiblenames = os.listdir(Slist_path.get())
                for i in range(len(possiblenames)):
                    if info['firstname'].lower()+info['lastname'].lower()+'.xls' == possiblenames[i].lower():
                        index = i
                temppref = pd.read_excel(Slist_path.get()+possiblenames[index])
                Tutorinfo[tutor]['pref'] = {temppref[list(temppref)[i+1]][2]:{Hours[j-3]:temppref[list(temppref)[i+1]][j] for j in range(3,13)} for i in range(5)}
            except:
                Tutorinfo[tutor]['pref'] = 'None'
        possibilities = update_pos()

        ### Current term
        term = Combobox(root)
        term.grid(column = 1,row = 1)
        term['values'] = ('A-term','B-term','C-term','D-term')

        ### The grid:
        schedule = Frame(root)
        schedule.grid(row = 2, column = 2)
        final = {day:{hour:{1:'',2:''} for hour in Hours} for day in Days}
        for i,day in enumerate(Days):
            Label(schedule,text = day).grid(row = 1, column = i+2)
            for j,time in enumerate(Hours):
                
                if i==0:
                    Label(schedule,text = str(time)).grid(row = 2*j+2,column = 1)
                for h in [1,2]:
                    final[day][time][h] = Combobox(schedule)
                    final[day][time][h].grid(row = 2*j+1+h,column = i+2)
                    final[day][time][h]['values'] = tuple(['None']) + tuple(possibilities[day][time])
                    final[day][time][h].current(0)

        unscheduled = Combobox(root)
        unscheduled.grid(row = 1, column = 20)
        unscheduled['values'] = tuple(Tutorinfo)
        unscheduled.current(0)

        show = Button(root, text = "Show Schedule", command = lambda: Showschedule(unscheduled.get()))
        show.grid(row = 1,column = 21)
        init = 1
    elif init:
        update_pref(final)
        possibilities = update_pos()
        for day in Days:
            for time in Hours:
                for h in [1,2]:
                    final[day][time][h]['values'] = tuple(['None']) + tuple(possibilities[day][time])
        temp = unscheduled.get()
        unscheduled['values'] = tuple([tutor for tutor in Tutorinfo if Tutorinfo[tutor]['scheduled']<Tutorinfo[tutor]['hours']])
        if temp not in unscheduled['values']:
            unscheduled.current(0)
    root.after(1000,lambda: update())




root.after(1000,lambda: update())
root.mainloop()


