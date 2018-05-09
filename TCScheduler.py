import re
import pickle
import numpy as np
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
import os
import xlrd

### This code is written to help automate the scheduling process.

Hours = ['10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8']
Days = ['Monday','Tuesday','Wednesday','Thursday','Friday']
Folder = 'Tutor Schedules'
numtas = 19
numplas = 43




Names = [re.sub('\.[a-z]*','',tutor) for tutor in os.listdir(Folder)]







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

### Browse for folder
def browse_button():
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    print(filename)

### Current term
term = Combobox(root)
term.grid(column = 1,row = 1)
term['values'] = ('A-term','B-term','C-term','D-term')



### Check for schedules and possibly load them in.
def GetSchedules(preloaded,window=0):
    if preloaded:
        schedules = pickle.load(open("Schedules.p",'rb'))
        if window!=0:
            quit(window)
    schedules = {}
    Tutorfiles = os.listdir(Folder)
    for tutor in Tutorfiles:
        sheet = xlrd.open_workbook(Folder+'\\'+tutor).sheet_by_index(0)
        schedules[re.sub(r'\.[a-z]*','',tutor)] = np.array([[sheet.row_slice(rowx = r,start_colx = 1,end_colx = 6)] for r in range(4,14)])
    pickle.dump(schedules,open("Schedules.p",'wb'))
    if window != 0:
        quit(window)

def CheckSchedules():
    if "Schedules.p" in os.listdir('.'):
        localwindow = Toplevel(root)
        Label(localwindow,text = "Schedules detected in your current directory.\n Would you like to load these in?\n or would you like to load the schedules from scratch?").grid(row = 0, column = 1, columnspan = 3)
        btn1 = Button(localwindow, text = "Load previous schedules", command = lambda: GetSchedules(1,localwindow))
        btn2 = Button(localwindow, text = "Load from scratch", command = lambda: GetSchedules(0,localwindow))
        btn1.grid(row = 1,column = 2)
        btn2.grid(row = 1,column = 3)
    else:
        GetSchedules(0)
GetS = Button(root, text = "Load in Schedules", command = CheckSchedules)
GetS.grid(column = 1, row = 2)



### The grid:
schedule = Frame(root)
schedule.grid(row = 2, column = 2)
final = {}
for i,day in enumerate(Days):
    Label(schedule,text = day).grid(row = 1, column = i+2)
    for j,time in enumerate(Hours):
        if i==0:
            Label(schedule,text = str(time)).grid(row = 2*j+2,column = 1)
        for h in [1,2]:
            final["{0}{1}{2}".format(day,time,h)] = Combobox(schedule)
            final["{0}{1}{2}".format(day,time,h)].grid(row = 2*j+1+h,column = i+2)
            final["{0}{1}{2}".format(day,time,h)]['values'] = ('None')




def update():
    print('nothing here yet')
    #nothing yet
     


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




root.after(1000,update())
root.mainloop()


