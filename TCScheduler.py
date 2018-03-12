import re
import pickle
import numpy as np
from tkinter import *
from tkinter.ttk import *
import os
import pandas as pd

### For now, though this will be replaced by pandas.
import xlrd

### This code is written to help automate the scheduling process.


### Preset Constants
Hours = ['10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8']
Days = ['Monday','Tuesday','Wednesday','Thursday','Friday']
Folder = 'C:\\Users\\bgkodalen\\Desktop\\TA\\Tutoring Center\\Cterm2018Schedule'

####
TAlist = pd.read_excel(Folder+'\\List of Tutors.xlsm', sheet_name = "Sheet1")['TA']
PLAlist = pd.read_excel(Folder+'\\List of Tutors.xlsm', sheet_name = "Sheet2")['PLA']

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
        possiblenames = os.listdir(Folder+'\\Tutor Schedules')
        for i in range(len(possiblenames)):
            if info['firstname'].lower()+info['lastname'].lower()+'.xls' == possiblenames[i].lower():
                index = i
        temppref = pd.read_excel(Folder+'\\Tutor Schedules\\'+possiblenames[index])
        Tutorinfo[tutor]['pref'] = {temppref[list(temppref)[i+1]][2]:{Hours[j-3]:temppref[list(temppref)[i+1]][j] for j in range(3,13)} for i in range(5)}
    except:
        Tutorinfo[tutor]['pref'] = 'None'

def update_pos():
    oneposs = {day:{hour:[] for hour in Hours} for day in Days}
    for tutor in Tutorinfo:
        if Tutorinfo[tutor]['scheduled'] < Tutorinfo[tutor]['hours']:
            for day in Days:
                for hour in Hours:
                    if Tutorinfo[tutor]['pref'][day][hour] == 1:
                        oneposs[day][hour].append(tutor+' ('+Tutorinfo[tutor]['role']+')')
    return oneposs 

def update_pref(schedule):
    for tutor in Tutorinfo:
        Tutorinfo[tutor]['scheduled'] = 0
    for day in Days:
        for hour in Hours:
            for shift in [1,2]:
                tutor = schedule[day][hour][shift].get().split()
                if tutor[0]!='None':
                    tutor = tutor[0]+' '+tutor[1]
                    Tutorinfo[tutor]['scheduled'] +=1
    

possibilities = update_pos()


### The actual GUI
root = Tk()
root.title("Tutoring Schedule")
#root.geometry('100x500')

lbl = Label(root, text="This project should help make the scheduling process easier.",font = ("Arial Bold",10))
lbl.grid(row=0,column=0,columnspan=100,sticky=W+E+N+S)

### Close button
def quit(window):
    window.destroy()
btn = Button(root,text = "Close",command = lambda: quit(root))
btn.grid(column = 99, row = 99)

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

def Showschedule(tutor):
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



### Other document functions

# imprim = StringVar()
# imprim.set('primitive')
# Radiobutton(root, text = "primitive", variable = imprim, value = 'primitive').grid(column=2,row=1,sticky=W)
# Radiobutton(root, text = "antipodal", variable = imprim, value = 'antipodal').grid(column=2,row=2,sticky=W)
# Radiobutton(root, text = "bipartite", variable = imprim, value = 'bipartite').grid(column=2,row=3,sticky=W)

# def Matrixfrmt(mat,name,window,r,c,string=0):
#     dim = mat.shape
#     if len(name)>0:
#         Label(window,text = name+'=').grid(row=r+int(dim[0]/2),column=c)
#     if len(dim)>2:
#         for i in range(dim[0]):
#             [t,c] = Matrixfrmt(mat[i],'',window,r,c,string)
#         return [t,c]
#     else:
#         rowpos = r
#         for i in range(dim[0]):
#             colpos = c+2
#             for j in range(dim[1]):
#                 if string == 1:
#                     Label(window,text = mat[i,j]).grid(row=rowpos,column = colpos)
#                 elif (type(mat[i,j]) == bytes) or (type(mat[i,j]) == str):
#                     Label(window,text = mat[i,j]).grid(row=rowpos,column=colpos)
#                 elif round(mat[i,j],2) == int(mat[i,j]):
#                     Label(window,text = int(mat[i,j])).grid(row=rowpos,column=colpos)
#                 else:
#                     Label(window,text = format(float(mat[i,j]),'.2f')).grid(row=rowpos,column=colpos)
#                 colpos+=1         
#             rowpos+=1
        
#         for b in range(dim[0]):
#             Label(window,text = '|').grid(row = r+b,column = c+1)
#             Label(window,text = '|').grid(row = r+b,column = colpos)
#         Label(window,text = ' ').grid(row = rowpos,column = colpos+1)
#         Label(window,text = ' ').grid(row = rowpos,column = colpos+2)
#         return [rowpos,colpos+2]


# ### Various filters on the schemes
# irrat = IntVar()
# irr = Checkbutton(root,text="Irrational schemes", variable = irrat)
# irr.grid(column = 1,row = 4,sticky=W)

# geg = IntVar()
# ge = Checkbutton(root,text="Spherical bound", variable = geg)
# ge.grid(column = 1,row = 5,sticky=W)

# SD = IntVar()
# SDcheck = Checkbutton(root,text="Spherical design", variable = SD)
# SDcheck.grid(column = 2,row = 4,sticky=W)
# numdesign = Combobox(root,width = 2)
# numdesign.grid(column = 3,row = 4,sticky = W)
# numdesign['values'] = (3,4,5,6,7,8,9)
# numdesign.current(0)

# equi = IntVar()
# equicheck = Checkbutton(root, text = "Equiangular lines", variable = equi)
# equicheck.grid(column = 2, row = 5, sticky = W)


# def parameterlist():
#     selectedschemes = schemes[numclasses.get()][imprim.get()]
#     schemelist = [scheme for scheme in selectedschemes]
#     tol=10**(-14)
#     if irrat.get():
#         schemelist = [scheme for scheme in schemelist if selectedschemes[scheme]['irrational'] == 1]
#     if geg.get():
#         schemelist = [scheme for scheme in schemelist if (Qpoly.Gegproj(Qpoly.Lsm(selectedschemes[scheme]['P']))).min()<-tol]
#     if SD.get():
#         schemelist = [scheme for scheme in schemelist if sum(np.absolute(Qpoly.Gegproj(Qpoly.Lsm(selectedschemes[scheme]['P']))[0,1:(int(numdesign.get())+1)]))<tol]
#     if equi.get():
#         schemelist = [scheme for scheme in schemelist if sum(Qpoly.Qm(selectedschemes[scheme]['P'])[0,:])-Qpoly.Qm(selectedschemes[scheme]['P']).max()<140]
#     if len(schemelist) == 0:
#         schemelist = ['None']
#     else:
#         schemelist = [scheme+selectedschemes[scheme]['exists'] for scheme in schemelist]
#     return tuple(schemelist)


# ### Shows the available parameters based on the Radiobutton input.
# params = Combobox(root)
# params.grid(column = 96,row = 1,columnspan=3)
# params['values'] = parameterlist()

# params.current(0)
# temp = params['values'][0]
# def update_comb(temp):
#     params['values'] = parameterlist()
#     if temp != params['values'][0]:
#         params.current(0)
#     temp = params['values'][0]
#     if params['values'][0]!='None':
#         Label(root, text = str(len(params['values']))+' Schemes').grid(row = 2,column = 99)
#     else:
#         Label(root, text = '0 Schemes').grid(row = 2,column = 99)
#     root.after(1000,lambda: update_comb(temp))

# ### The Examine Scheme window.
# def examine():
#     Details = Toplevel(root)
#     scheme = re.findall(r'<[,\d\w;]*>',params.get())[0]
    
#     Data = schemes[numclasses.get()][imprim.get()][scheme]

#     ### Load in the P matrix and calculate the rest
#     P = Data['P']
#     Q = Qpoly.Qm(P)
#     L = Qpoly.Lm(P)
#     Ls = Qpoly.Lsm(P)
#     Geg = Qpoly.Gegproj(Ls,10,0,1)


#     ### Display the various information
#     fp = Frame(Details)
#     fp.grid(row = 2, column = 2,sticky = W)
#     [r,c] = Matrixfrmt(P,'P',fp,2,2)
#     [r,c] = Matrixfrmt(Q,'Q',fp,2,c+1)
#     [r,t] = Matrixfrmt(Geg,'G',fp,2,c+1,1)

#     fl = Frame(Details)
#     fl.grid(row = 4, column = 2,sticky = W)
#     [r,t] = Matrixfrmt(L,'L',fl,r+1,2)

#     fls = Frame(Details)
#     fls.grid(row = 5, column = 2,sticky = W)
#     [r,t] = Matrixfrmt(Ls,'L*',fls,r+1,2)

#     ### Extra information
#     if 'exists' in Data:
#         exists = Data['exists']
#     else:
#         exists = '?'
#     if 'Comments' in Data:
#         comments = Data['Comments']
#     else:
#         comments = ''

#     if exists == '-':
#         color = 'red'
#     elif exists == '?':
#         color = 'yellow'
#     elif exists == 'O':
#         color = 'blue'
#     elif exists == 'P':
#         color = 'yellow'
#     else:
#         color = 'green'
#     Label(Details, text = (scheme+' '+exists),background=color).grid(row = 1,column = 1)
#     Label(Details, text = comments).grid(row = 1,column = 2)
    
#     ### Temporary Labels
#     if Ls[1,1,1]>0:
#         m = Ls[1,0,1]
#         a = Ls[1,1,1]**2
#         b = Ls[1,2,2]/Ls[1,1,1]
#         c = Ls[1,1,2]*Ls[1,2,1]
#         d = 4*m*(2*m-3)/(m+6)
#         Label(Details, text = '%0.2f + (2+%0.2f)*%0.2f > %0.2f' % (a,b,c,d)).grid(row = 99,column = 1,columnspan = 5)
#         Label(Details, text = '%0.2f > %0.2f' % (a+(2+b)*c,d)).grid(row = 100,column = 1,columnspan = 5)
    
#     Label(Details, text = Qpoly.Equiangular(Q)[0]).grid(row = 6, column = 2)

#     exclose = Button(Details,text = "Close",command = lambda: quit(Details))
#     exclose.grid(column = 99, row = 99)

# ex = Button(root, text = "Examine Scheme", command = examine)
# ex.grid(column = 99, row = 1)




root.after(1000,lambda: update())
root.mainloop()


