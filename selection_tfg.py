#!/usr/bin/env python3

debug=False

import pandas as pd
import numpy as np
import random
import sys
df = pd.read_excel(r'Solicitud.xlsx')
data = df.to_dict()
# Extract required data
#-------------------------
# Student data
IDs       = list(data['ID'].values())
names     = list(data['Nombre'].values())
mails     = list(data['Correo electrónico'].values())
# Students ranks
marks     = list(data['Nota media\n'].values())
credits   = list(data['Créditos superados\n'].values())
# Selections
options1  = data['Opción 1']
options2  = data['Opción 2\n']
options3  = data['Opción 3\n']
options4  = data['Opción 4\n']
options5  = data['Opción 5\n']
options6  = data['Opción 6\n']
options7  = data['Opción 7\n']
options8  = data['Opción 8\n']
options9  = data['Opción 9\n']
options10 = data['Opción 10\n']
more_options = data['Opciones no priorizadas\n']
# Proofs
urls = data['Captura de pantalla de Sigma con nota media y créditos superados\n']
erasmus = data['Correo UAM del tutor Erasmus\n']

# Some people give mark over 100...
for i,mark in enumerate(marks):
    if mark>10.:
        marks[i] /= 100

#--------------------
# Check mark/credits
#--------------------
verification = dict()
fproofs = open('Verificacion.txt')
for line in fproofs:
    mail,mark,credit,status,url = line.split('&')
    mail = mail.strip()
    verification[mail] = {'mark':float(mark), 
                          'credit':int(credit), 
                          'url':url.strip(), 
                          'status':status.strip()}
fproofs.close()
# Write new data
fproofs = open('Verificacion.txt','w')
for i,mail in enumerate(mails):
    if mail in verification.keys():
        mark   = verification[mail]['mark']
        credit = verification[mail]['credit']
        url    = verification[mail]['url']
        status = verification[mail]['status']
        # If marked as "External" the info in the file overrides the excel sheet
        if status == 'External':
            print(f'NOTE: Entry {mail} has externally updated data')
            marks[i] = mark
            credits[i] = credit
        # Check if entry has changed (and updated status if so)
        if mark != marks[i]:
            print(f'NOTE: Entry {mail} has changed! (Mark: {mark} -> {marks[i]})')
            # Set updated value
            mark = marks[i]
            status = 'ToCheck'
        if credit != credits[i]:
            print(f'NOTE: Entry {mail} has changed! (Credits: {credit} -> {credits[i]})')
            # Set updated value
            credit = credits[i]
            status = 'ToCheck'
        if url != urls[i]:
            print(f'NOTE: Entry {mail} has changed! (URL)')
            # Set updated value
            url = urls[i]
            status = 'ToCheck'
        if str(erasmus[i]) != 'nan':
            print(f'NOTE: Entry {mail} is Erasmus, with contact {erasmus[i]}')
            mark   = marks[i]
            credit = credits[i]
            url    = urls[i]
            status = 'Erasmus'
    else:
        mark   = marks[i]
        credit = credits[i]
        url    = urls[i]
        if str(erasmus[i]) != 'nan':
            print(f'NOTE: Entry {mail} is Erasmus, with contact {erasmus[i]}')
            status = 'Erasmus'
        else
            status = 'ToCheck'
    print(f'{mail:60} & {mark:5.2f} & {credit:3} & {status:10} & {url}', file=fproofs)
fproofs.close()
print()


#----------------------------
# Open list of TFG projects
#----------------------------
ftfg = open('ListaTFGs.dat')
tfgs = []
for line in ftfg:
    tfg = {'title': line.split('&')[0].strip(),
           'dpto' : line.split('&')[1].strip(),
           'ext'  : line.split('&')[2].strip()}
    tfgs.append(tfg)

#-----------------------------
# Make assignments
#-----------------------------
# FIRST ROUND: order students
#------------------------------
credits = np.array(credits)
marks   = np.array(marks)
ndx     = np.arange(len(credits))
# Top layer is ordered 1) by mark and 2) by nr. credits 3) arrival date (TODO)
# 1) Sort by mark
ndx_top = ndx[credits>=175]
marks_top = marks[ndx_top]
credits_top = credits[ndx_top]
ndx_top   = ndx_top[marks_top.argsort()[::-1]]
# 2) Sort by credits
i = 0
ndx_top_reord = []
for mark in sorted(list(set(marks_top)))[::-1]:
    credits_top_sect = credits_top[marks_top==mark]
    ndx_top_sect = ndx[credits_top_sect.argsort()[::-1]]
    ndx_top_reord += list(ndx_top[ndx_top_sect+i])
    i += len(ndx_top_sect)
ndx_top = np.array(ndx_top_reord)
# 3) Sort by index
# i = 0
# ndx_top_reord = []
# marks_credits_top = [marks_top, credits_top]
# IDs_top = #
# for mark_credit in sorted(list(set(marks_cedits_top)))[::-1]:
#     credits_top_sect = credits_top[marks_top==mark]
#     ndx_top_sect = ndx[credits_top_sect.argsort()[::-1]]
#     ndx_top_reord += list(ndx_top[ndx_top_sect+i])
#     i += len(ndx_top_sect)
# ndx_top = np.array(ndx_top_reord)

# Low layer is ordered 1) by nr. credits 2) by marks 3) arrival date (TODO)
ndx_low = ndx[credits<175]
credits_low = credits[ndx_low]
ndx_low   = ndx_low[credits_low.argsort()[::-1]]
marks_low = marks[ndx_low]
i = 0
ndx_low_reord = []
for credit in sorted(list(set(credits_low)))[::-1]:
    marks_low_sect = marks_low[credits_low==credit]
    ndx_low_sect = ndx[marks_low_sect.argsort()[::-1]]
    ndx_low_reord += list(ndx_low[ndx_low_sect+i])
    i += len(ndx_low_sect)
ndx_low = np.array(ndx_low_reord)
# Pack them all
if len(ndx_top) == 0:
    ndx_ordered = ndx_low
elif len(ndx_low) == 0:
    ndx_ordered = ndx_top
else:
    ndx_ordered = np.hstack([ndx_top,ndx_low])

# Print
IDs = np.array(IDs)
prev = [None,None]
print('Orden de prioridad\n----------------------')
for i in ndx_ordered:
    credit = credits[i]
    mark   = marks[i]
    name   = names[i]
    if [mark,credit] == prev:
        print(f'{name:30} {credit:4} {mark:5.2f} *')
    else:
        print(f'{name:30} {credit:4} {mark:5.2f}')
print('')

# Check if there are equivalent students (same 
students = [ [mark,credit] for mark,credit in zip(marks[ndx_ordered],credits[ndx_ordered]) ]
students_equiv = np.array([ students.count(student) for student in students ])
if len(students_equiv[students_equiv != 1]) != 0:
    print('Warning! equivalent students')
    sys.exit()

# SECOND ROUND: get selections
#------------------------------
print('Asignación\n-----------------')
assign = []
for i in ndx_ordered:
    # Handle sumbitter info for mailing
    submitter_name = names[i]
    submitter_mail = mails[i]
    ID = IDs[i]

    # Handle info in the Form
    name  = names[i]
    email = mails[i]  
    options_raw = [ options1[i], options2[i], options3[i], options4[i], options5[i], 
                    options6[i], options7[i], options8[i], options9[i], options10[i] ]
    if str(more_options[i]) != 'nan':
        reserva = (more_options[i].split(';')[:-1])
        random.shuffle(reserva)
        options_raw += reserva
    options = []
    for option in options_raw:
        if option in options:
            option_id = option.split('-')[0]
            if debug:
                print(f'Repeated selection: {option_id}')
        elif str(option) == 'nan':
            pass
        else:
            options.append(option)

    # Make assigment
    selected_tfg = 'Sin selección'
    nsel = 1
    tfgs_titles = [ tfg['title'] for tfg in tfgs ]
    for option in options:
        if option in tfgs_titles:
            i = tfgs_titles.index(option)
            selected_tfg = tfgs.pop(i)
            break
        else:
            if debug:
                option_id = option.split('-')[0]
                print(f'Unassigned {nsel}: {option_id}')
            nsel += 1

    print(f'{name:35} & {selected_tfg["title"][:60]:60}...  ({nsel})')
    assign.append([name,selected_tfg])
print('')

# Print assigment
TFG_IDs = [ item[1]['title'].split('-')[0] for item in assign ]
# Get ordering
ndx = [ TFG_IDs.index(item) for item in sorted(TFG_IDs) ]
selected_tfgs = []
student_names = []
selected_dptos= []
selected_exts = []
for i in ndx:
    student_names.append(assign[i][0])
    selected_tfgs.append(assign[i][1]['title'])
    selected_dptos.append(assign[i][1]['dpto'])
    selected_exts.append(assign[i][1]['ext'])
# Generate excel
df = pd.DataFrame(list(zip(selected_tfgs,selected_dptos,selected_exts,student_names)))
df.to_excel('AsignacionTFGs.xlsx', sheet_name='TFGs')

# Lo más populares
print('Populares\n--------------')
options1 = list(options1.values())
for option in set(options1):
    print(f'{option[:60]:60}...  {options1.count(option)}')
