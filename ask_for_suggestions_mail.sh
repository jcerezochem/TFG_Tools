#!/bin/bash

list_emails=""
while read mail; do
    list_emails="$list_emails,$mail"
done <ListaEmails.txt
echo $list_emails

#cat MAIL_sugerencias | mail -s "Sugerencias de mejora en el proceso de TFG en Química" -r coordinador.tfg.quimica@uam.es -c javier.cerezo@uam.es,mm.montero@uam.es -b $list_emails coordinador.tfg.quimica@uam.es

echo DONE
