#!/usr/bin/env python3

import pandas as pd
import numpy as np
import sys
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

tfg = []
for i in range(ntfg):
    skip_id = False
    # Set ID and check if already done
    ID = f'{i:03g}'

    # Handle info in the Form
    name1  = names1[i]
    email1 = emails1[i]  
    dpto1  = dptos1[i]
    name2  = names2[i]
    email2 = emails2[i]
    dpto2  = dptos2[i]
    dpto_ext = dptos_ext[i]
    titulo = titulos[i]
    title  = titles[i]
    tipo   = tipos[i]
    area   = areas[i] 
    descr  = descrs[i] 
    cap    = caps[i]
    info   = infos[i]  
    dire   = dires[i] 
    cat1   = cats1[i]
    cat2   = cats2[i]
    if str(cat1) == 'Otro (especificar)':
        cat1 = str(data['Categoría profesional tutor/a 1 (otro)\n'][i])
        if cat1 == 'nan':
            cat1 = 'Otro'
    if str(cat2) == 'Otro (especificar)':
        cat2 = str(data['Categoría profesional tutor/a 2 (otro)\n'][i])
        if cat2 == 'nan':
            cat2 = 'Otro'
    # Backward (when this option was not in the Form)
    cat1 = str(cat1)
    cat2 = str(cat2)

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
            elif key == 'Tutor1_nombre':
                name1 = corr[key]
            elif key == 'Tutor2_nombre':
                name2 = corr[key]
            elif key == 'Tutor1_email':
                email1 = corr[key]
            elif key == 'Tutor2_email':
                email2 = corr[key]
            elif key == 'Tutor1_dpto':
                dpto1 = corr[key]
            elif key == 'Tutor2_dpto':
                dpto2 = corr[key]
            elif key == 'Tutor1_cat':
                cat1 = corr[key]
            elif key == 'Tutor2_cat':
                cat2 = corr[key]
            elif key == 'Descripción':
                descr = corr[key]
            elif key == 'Eliminar':
                skip_id = True

    if skip_id:
        continue


    # Start new item of dict
    tfg.append({})

    # Check length of description
    words_in_descr = len(descr.split())
    tfg[-1]['descr_len'] = words_in_descr
    if words_in_descr > 300:
        warnings.append(f'Waning en {ID}: la descripción es muy larga ({words_in_descr}). Contacta con {submitter_mails[i]}')

    # Store ID
    tfg[-1]['ID'] = ID
    if ID == '018':
        print('Oh no!')
        sys.exit()

    # Store dir
    tfg[-1]['dir'] = dire
    
    # Manage multiple tutors
    if str(name2) == 'nan' or str(name2) == '':
        name  = name1
        email = email1
        dpto  = dpto1
        cat   = cat1
        tfg[-1]['is_ext'] = False
        tfg[-1]['dpto'] = [dpto]
        tfg[-1]['cat'] = [cat1]
        tfg[-1]['mail'] = [email1]
    elif dpto2 == 'Externo (especificar)':
        dpto2 = dpto_ext
        name = name2+' / '+name1
        email = email2+' / '+email1
        dpto  = dpto2+' / '+dpto1
        cat   = cat2+' / '+cat1
        tfg[-1]['is_ext'] = True
        tfg[-1]['dpto'] = [dpto1]
        tfg[-1]['ext'] = dpto_ext
        tfg[-1]['cat'] = [cat1]
        tfg[-1]['mail'] = [email1,email2]
    elif dpto2 == 'Externo (especificar) - no_ext':
        dpto2 = dpto_ext
        name = name1+' / '+name2
        email = email1+' / '+email2
        dpto  = dpto1+' / '+dpto2
        cat   = cat1+' / '+cat2
        tfg[-1]['is_ext'] = True
        tfg[-1]['dpto'] = [dpto1]
        tfg[-1]['ext'] = dpto_ext
        tfg[-1]['cat'] = [cat1]
        tfg[-1]['mail'] = [email1,email2]
    else:
        name = name1+' / '+name2
        email = email1+' / '+email2
        if str(dpto1) == str(dpto2):
            dpto = dpto1
            tfg[-1]['dpto'] = [dpto1]
        else:
            dpto  = dpto1+' / '+dpto2
            tfg[-1]['dpto'] = [dpto1,dpto2]
        cat   = cat1+' / '+cat2
        tfg[-1]['is_ext'] = False
        tfg[-1]['cat'] = [cat1,cat2]
        tfg[-1]['mail'] = [email1,email2]
    # Optional fields
    if str(info) == 'nan':
        info = ''
    if str(cap) == 'nan':
        cap = ''


n_tot = len(tfg)
external = [ item['is_ext'] for item in tfg ]
tfg_ext = [ item for item in tfg if item['is_ext'] ]
tfg_uam = [ item for item in tfg if not item['is_ext'] ]
n_ext = len(tfg_ext)
n_uam = len(tfg_uam)
print(f'Número total de TFGs  : {n_tot:<5}')
print(f' Núm. proyectos UAM   : {n_uam:<5}  ({n_uam/n_tot*100:>4.1f}%)')
print(f' Núm. proyectos ext.  : {n_ext:<5}  ({n_ext/n_tot*100:>4.1f}%)')
print('')

print(f'Proyectos UAM ({n_uam})')
nr_tutors = [ len(item['cat']) for item in tfg_uam ]
n_1tut = nr_tutors.count(1)
n_2tut = nr_tutors.count(2)
print(f' - 1 tutor            : {n_1tut:<5}')
print(f' - 2 tutores          : {n_2tut:<5}')
print('')

# This dict contains all dptos to get the weight
total_tfg = {}

singldpto = [ item['dpto'][0] for item in tfg_uam if len(item['dpto']) == 1 ]
print(f' - 1 departamento     : {len(singldpto):<5}')
# Order in terms of number of projects
uq_dpto = np.array(list(set(singldpto)))
nr_dpto = np.array([ singldpto.count(item) for item in uq_dpto ])
idx = nr_dpto.argsort()[::-1]
uq_dpto = uq_dpto[idx]
for dpto in uq_dpto:
    print(f'    {dpto:<58}: {singldpto.count(dpto)}')
    total_tfg[dpto] = singldpto.count(dpto)
print('')

interdpto = [ sorted(item['dpto']) for item in tfg_uam if len(item['dpto']) == 2 ]
print(f' - 2 departamentos    : {len(interdpto):<5}')
# Order in terms of number of projects
pairsdpto = []
for dptos in interdpto:
    dptos.sort()
    dpto1,dpto2 = dptos
    pairsdpto.append(f'{dpto1:<58} / {dpto2:<58}')
uq_dpto = np.array(list(set(pairsdpto)))
nr_dpto = np.array([ pairsdpto.count(item) for item in uq_dpto ])
idx = nr_dpto.argsort()[::-1]
uq_dpto = uq_dpto[idx]
for dpto in uq_dpto:
    print(f'    {dpto.strip():<100}: {pairsdpto.count(dpto)}')
    dpto1, dpto2=dpto.split('/')
    if dpto1.strip() in total_tfg.keys():
        total_tfg[dpto1.strip()] += pairsdpto.count(dpto) * 0.5
    else:
        total_tfg[dpto1.strip()]  = pairsdpto.count(dpto) * 0.5
    if dpto2.strip() in total_tfg.keys():
        total_tfg[dpto2.strip()] += pairsdpto.count(dpto) * 0.5
    else:
        total_tfg[dpto2.strip()]  = pairsdpto.count(dpto) * 0.5
print('')


print(f'Proyectos externos ({n_ext})')
singldpto = [ item['dpto'][0] for item in tfg_ext if len(item['dpto']) == 1 ]
# Order in terms of number of projects
uq_dpto = np.array(list(set(singldpto)))
nr_dpto = np.array([ singldpto.count(item) for item in uq_dpto ])
idx = nr_dpto.argsort()[::-1]
uq_dpto = uq_dpto[idx]
for dpto in uq_dpto:
    print(f'    {dpto:<58}: {singldpto.count(dpto)}')
    if dpto in total_tfg.keys():
        total_tfg[dpto] += singldpto.count(dpto) 
    else:
        total_tfg[dpto]  = singldpto.count(dpto) 
print('')


print(f'Contribución a los TFG totales ({n_tot})')
uq_dpto = np.array(list(total_tfg.keys()))
nr_dpto = np.array([ total_tfg[dpto] for dpto in uq_dpto ])
idx = nr_dpto.argsort()[::-1]
uq_dpto = uq_dpto[idx]
for dpto in uq_dpto:
    print(f'    {dpto:<58}: {total_tfg[dpto]/n_tot*100:4.1f}%')
print('')


# Average description length
# Averate length
print('Análsis de la extensión de la descripción')
descr_len = np.array([ item['descr_len'] for item in tfg ])
print(f'  Average descr. length: {descr_len.mean():8.1f}')
print(f'  Max     descr. length: {descr_len.max():6}')
print(f'  Min     descr. length: {descr_len.min():6}')
print('')

#sys.exit()

# Check number of jobs per teacher (by email)
print('TFG por tutor')
mails = []
IDsR  = []
for item in tfg:
    mails += [ mail.strip() for mail in item['mail'] ]
    IDsR  += [ item['ID'] for mail in item['mail'] ]
uq_mails = np.array(list(set(mails)))
nr_mails = np.array([ mails.count(item) for item in uq_mails ])
idx = nr_mails.argsort()[::-1]
uq_mails = uq_mails[idx]
mails = np.array(mails)
IDsR  = np.array(IDsR)
for mail in uq_mails:
    if not '@uam.es' in mail:
        continue
    nproj = len(mails[mails==mail])
    projs = list(IDsR[mails==mail])
    print(f'     {mail:<35}: {nproj:<3} {projs}') 

print('      ----CSIC-----')
for mail in uq_mails:
    if not 'csic.es' in mail:
        continue
    nproj = len(mails[mails==mail])
    projs = list(IDsR[mails==mail])
    print(f'     {mail:<35}: {nproj:<3} {projs}') 

print('      ----Otros-----')
for mail in uq_mails:
    if 'csic.es' in mail or '@uam.es' in mail:
        continue
    nproj = len(mails[mails==mail])
    projs = list(IDsR[mails==mail])
    print(f'     {mail:<35}: {nproj:<3} {projs}') 
print('')


#IDs   = [ item['ID'] for item in tfg ]
#mails = [ item['mail'] for item in tfg ]
#for ID,mail in zip(IDs,mails):
#    for m in mail:
#        if 'maryam.derogar' in m:
#            print(ID)


print()
for item in tfg_ext:
    print(item['ID'], item['dir'])
