#!/bin/bash

# Requires: conda activate py3
./generate_tfgs.py

if (( $? )); then exit; fi
sleep 1

for file in tfg???.tex; do 
    if [ -f ${file/tex/pdf} ]; then continue; fi
    pdflatex $file; 
done

rm tfg???.{aux,log} -v
