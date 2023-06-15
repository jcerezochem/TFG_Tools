#!/bin/bash

while read id mail name; do
    sed "s/NAME/$name/" MAIL | mail -s "Recepción proyecto TFG (ID:$id)" -r coordinador.tfg.quimica@uam.es -c javier.cerezo@uam.es,mm.montero@uam.es -a tfg${id}.pdf $mail
done < MailsToSend.txt

# Pass entries to Sent mail
cat MailsToSend.txt >> MailsSent.txt
rm MailsToSend.txt
touch MailsToSend.txt

