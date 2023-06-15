#!/usr/bin/env python3

import pandas as pd
import numpy as np
df = pd.read_excel(r'Propuesta.xlsx')
data = df.to_dict()
# Extract required data
names1   = data['Nombre (tutor/a 1)\n']
emails1  = data['Correo electrónico (tutor/a 1):\n']
dptos1   = data['Departamento (tutor/a 1):\n']
names2   = data['Nombre (tutor/a 2)\n']
emails2  = data['Correo electrónico (tutor/a 2):\n']
dptos2   = data['Departamento (tutor/a 2):\n']
dptos_ext= data['Institución externa\n']
titulos  = data['Título']
titles   = data['Título en inglés\n']
tipos    = data['Tipo de trabajo\n']
areas    = data['Área de conocimiento en la que se enmcarca\n']
descrs   = data['Breve descripción del proyecto y tareas a realizar (250 palabras)\n']
caps     = data['Capacidades y conocimientos que son recomendables para que el alumno/a pueda realizar este proyecto\n']
infos    = data['Indique cualquier otra aclaración que considere que pueda ayudar a los estudiantes en la elección del TFG\n']
dires    = data['Dirección del lugar donde se realizará el proyecto\n']
cats1    = data['Categoría profesional (tutor/a 1)\n']
cats2    = data['Categoría profesional (tutor/a 2)\n']
# Data of the submitter
submitter_names = data['Nombre']
submitter_mails = data['Correo electrónico']

ntfg   = len(names1)

# Links
links = {}
with open("links") as f:
    for line in f:
        (key, val) = line.split()
        links[key] = val

# Open list to read entries
flist = open('ListaTFGs.dat','r')
ids_done = []
for line in flist:
    ids_done.append(line.split('-')[0])
flist.close()

# Open list to append new entries
flist = open('ListaTFGs.dat','a')

# Corrections
corrections = {}
with open('Correcciones.txt') as f:
    for line in f:
        (key1, key2, val) = line.split('&')
        if key1.strip() not in corrections.keys():
            corrections[key1.strip()] = {}
        corrections[key1.strip()][key2.strip()] = val.strip()

# Warnings
warnings = []

IDs = []
links_list = []
dptos = []
extern = []
tipos_tfg = []
for i in range(ntfg):
    skip_id = False
    # Set ID and check if already done
    ID = f'{i:03g}'
    if ID in ids_done:
        continue

    # Handle info in the Form
    dpto1  = dptos1[i]
    dpto2  = dptos2[i]
    dpto_ext = dptos_ext[i]
    titulo = titulos[i]
    tipo   = tipos[i]
    area   = areas[i] 
    descr  = descrs[i] 

    # Check if we need to correct something
    if ID in corrections.keys():
        corr = corrections[ID]
        for key in corr.keys():
            if key == 'Título':
                titulo = corr[key]
            elif key == 'Title':
                title = corr[key]
            elif key == 'Departamento1':
                dpto1 = corr[key]
            elif key == 'Departamento2':
                dpto2 = corr[key]
            elif key == 'DepartamentoExt':
                dpto_ext = corr[key]
            elif key == 'Eliminar':
                skip_id = True

    if skip_id:
        continue

    # Check length of description
    words_in_descr = len(descr.split())
    if words_in_descr > 300:
        warnings.append(f'Waning en {ID}: la descripción es muy larga ({words_in_descr}). Contacta con {submitter_mails[i]}')

    # Manage multiple tutors
    if str(dpto2) == 'nan' or dpto2 == 'Externo (especificar)':
        dpto = dpto1
    elif str(dpto1) == str(dpto2):
        dpto = dpto1
    else:
        dpto  = dpto1+'/'+dpto2
    # Manage external as separate entry
    if dpto2 == 'Externo (especificar)':
        dpto_ext = dpto_ext.strip()
    else:
        dpto_ext = 'No'
    link = links[ID]

    # Set some additional lists
    IDs.append(ID+'-'+titulo)
    links_list.append(f'"{link}", "tfg{ID}.pdf"')
    dptos.append(dpto)
    extern.append(dpto_ext)
    tipos_tfg.append(tipo)

    # Add list line
    print(f'{ID}-{titulo} & {dpto} & {dpto_ext} & {tipo} & "{link}", "tfg{ID}.pdf"',file=flist)

flist.close()

# Generate excel
df = pd.DataFrame(list(zip(IDs,dptos,extern,tipos_tfg,links_list)))
#print(list(zip(IDs,dptos,dptos_ext,tipos,links_list)))
df.to_excel('ListaTFGs.xlsx', sheet_name='TFGs')

# Generate tex (with boldface)
f = open('ListaTFGs.tex','w')
print(r'\documentclass{article}', file=f)
print(r'\begin{document}', file=f)
print(r'\begin{tabular}{l}', file=f)
for item in IDs:
    ID = item.split('-')[0]
    title = '-'.join(item.split('-')[1:])
    print(r'\textbf{'+ID+r'}-'+title+r'\\', file=f)

print(r'\end{tabular}', file=f)
print(r'\end{document}', file=f)



for warning in warnings:
    print('\n'+warning+'\n')

