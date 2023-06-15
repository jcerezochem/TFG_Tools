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

# Open template
template = open('template.tex').read()

# Check MailsToSend.txt
try:
    f = open('MailsToSend.txt')
    content_length = len(f.read())
    f.close()
except:
    content_length = 0
if content_length:
    raise Exception('Algunos correos por enviar. Hazlo e intentaló de nuevo')

# Open for writting
fmails = open('MailsToSend.txt','w')

# Open sent mails to get done proyects
ids_done = list(np.loadtxt('MailsSent.txt', usecols=[0], dtype=str))

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

for i in range(ntfg):
    skip_id = False
    # Set ID and check if already done
    ID = f'{i:03g}'
    if ID in ids_done:
        continue
    # Handle sumbitter info for mailing
    submitter_name = submitter_names[i]
    submitter_mail = submitter_mails[i]
    print(f'{ID:3} {submitter_mail:<20} {submitter_name}',file=fmails)

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

    # Check length of description
    words_in_descr = len(descr.split())
    if words_in_descr > 300:
        warnings.append(f'Waning en {ID}: la descripción es muy larga ({words_in_descr}). Contacta con {submitter_mails[i]}')
    
    # Manage multiple tutors
    if str(name2) == 'nan' or str(name2) == '':
        name  = name1
        email = email1
        dpto  = dpto1
        cat   = cat1
    elif dpto2 == 'Externo (especificar)':
        dpto2 = dpto_ext
        name = name2+' / '+name1
        email = email2+' / '+email1
        dpto  = dpto2+' / '+dpto1
        cat   = cat2+' / '+cat1
    elif dpto2 == 'Externo (especificar) - no_ext':
        dpto2 = dpto_ext
        name = name1+' / '+name2
        email = email1+' / '+email2
        dpto  = dpto1+' / '+dpto2
        cat   = cat1+' / '+cat2
    else:
        name = name1+' / '+name2
        email = email1+' / '+email2
        if str(dpto1) == str(dpto2):
            dpto = dpto1
        else:
            dpto  = dpto1+' / '+dpto2
        cat   = cat1+' / '+cat2
    # Optional fields
    if str(info) == 'nan':
        info = ''
    if str(cap) == 'nan':
        cap = ''

    # Manage spcial character with LaTeX
    replacements = {'<':'$<$',
                    '>':'$>$',
                    '&':r'\&',
                    'ºC':'$^\circ$C',
                    'α':r'$\alpha$',
                    'β':r'$\beta$',
                    'μ':r'$\mu$',
                    'ı́':'í',
                    '‐':'-'
                    }
    for char in replacements.keys():
        dpto   = dpto.replace(char,replacements[char])
        dire   = dire.replace(char,replacements[char])
        descr  = descr.replace(char,replacements[char])
        titulo = titulo.replace(char,replacements[char])
        title  = title.replace(char,replacements[char])
        info   = info.replace(char,replacements[char])
        cap    = cap.replace(char,replacements[char])
    #dpto  = dpto.replace('&',r'\&')
    #dire  = dire.replace('&',r'\&')
    #descr = descr.replace('<','$<$')
    #descr = descr.replace('>','$>$')
    #descr = descr.replace('ºC','$^\circ$C')
    #descr = descr.replace('α',r'$\alpha$')
    #descr = descr.replace('β',r'$\beta$')
    #descr = descr.replace('μ',r'$\mu$')
    #descr = descr.replace('&',r'\&')
    #titulo= titulo.replace('<','$<$')
    #titulo= titulo.replace('>','$>$')
    #titulo= titulo.replace('ºC','$^\circ$C')
    #titulo= titulo.replace('α',r'$\alpha$')
    #titulo= titulo.replace('β',r'$\beta$')
    #titulo= titulo.replace('μ',r'$\mu$')
    #titulo= titulo.replace('&',r'\&')
    #title = title.replace('<','$<$')
    #title = title.replace('>','$>$')
    #title = title.replace('ºC','$^\circ$C')
    #title = title.replace('α',r'$\alpha$')
    #title = title.replace('β',r'$\beta$')
    #title = title.replace('μ',r'$\mu$')
    #title = title.replace('&',r'\&')
    #info  = info.replace('<','$<$')
    #info  = info.replace('>','$>$')
    #info  = info.replace('ºC','$^\circ$C')
    #cap   = cap.replace('<','$<$')
    #cap   = cap.replace('>','$>$')
    #cap   = cap.replace('ºC','$^\circ$C')
    #cap   = cap.replace('&',r'\&')

    # Get proposal from template
    tfg = template.replace('_XXX_',ID)
    tfg = tfg.replace('_NAME_',name)
    tfg = tfg.replace('_DPTO_',dpto)
    tfg = tfg.replace('_CAT_',cat)
    tfg = tfg.replace('_EMAIL_',email)
    tfg = tfg.replace('_TITULO_',titulo)
    tfg = tfg.replace('_TITLE_',title)
    tfg = tfg.replace('_TIPO_',tipo)
    tfg = tfg.replace('_AREA_',area)
    tfg = tfg.replace('_DESCRIP_',descr)
    tfg = tfg.replace('_CAPAC_',cap)
    tfg = tfg.replace('_INFO_',info)
    tfg = tfg.replace('_DIR_',dire)
    
    # Write proposal
    fname = f'tfg{i:03g}.tex'
    f = open(fname,'w')
    f.write(tfg)
    f.close()

fmails.close()

for warning in warnings:
    print('\n'+warning+'\n')
