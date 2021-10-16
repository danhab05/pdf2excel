import pandas  as pd
import os
import PyPDF2
import platform
import schedule
import time
from datetime import datetime
import sys
path = sys.argv[1]
"""
Pour lancer le programme utiliser la commande 
python3 excel.py "chemin des pdf" "nom excel (sans extension)"
"""
folder = os.listdir(path)
folder.sort()
class PdfToExcel:
    def __init__(self, outpoutfilename):
        self.outpoutfilename = outpoutfilename
    def createExcel(self):
        Name = []
        firstName = []
        bornDate = []
        Num_secu = []
        zipCode = []
        hourDate = []
        phone = []
        place = []
        for f in folder:
            if True:
                # if '2021-0' in f and int(f.split('-')[2]) <= 28:+
                
                print(f)
                if True:
                    if int(f.split('-')[2]) > 45615461564156416512 :
                        pass
                    else:
                        daate = os.listdir(f'{path}/{f}')
                        for d in daate:
                            if d == 'O.H' or d == 'MH' or d == 'M.H':
                                pass
                            else:
                                print('    -' + d)
                                if 'pdf' in d:
                                    pass
                                elif True:
                                    try:
                                        files = os.listdir(f'{path}/{f}/{d}')
                                        for file in files:
                                            if '.pdf' in file:
                                                pdf = PyPDF2.PdfFileReader(f'{path}/{f}/{d}/{file}').getFields()
                                                Name.append(pdf['Nom de naissance']['/V'])
                                                hourDate.append(pdf['Date']['/V'] + ' ' + pdf['heureTest']['/V'])
                                                firstName.append(pdf['Prénom']['/V'])
                                                bornDate.append(pdf['Date de naissance JJMMAAAA']['/V'])
                                                Num_secu.append(pdf['Numéro de sécurité sociale']['/V'])
                                                zipCode.append(pdf['Code Postal']['/V'])
                                                phone.append(pdf['N téléphone mobile']['/V'])
                                                place.append(pdf['N']['/V'] + ' ' + pdf['Voie']['/V'])
                                    except Exception as e:
                                        pass
                                
        df = pd.DataFrame(
                {
                    'Horodateur': hourDate,
                    'Nom':Name,
                    'Prenom':firstName,
                    'Date_de_naissance':bornDate,
                    'Num_secu':Num_secu,
                    'Numéro de téléphone': phone,
                    'Adresse': place,
                    'Code_postal':zipCode
                }
            )
        df.to_excel('./' + self.outpoutfilename + '.xlsx', index=False)


        print('Done!')


excel = PdfToExcel(outpoutfilename=sys.argv[2])
excel.createExcel()


