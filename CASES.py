import PySimpleGUI as sg 
import sqlite3 
from datetime import datetime
from datetime import timedelta
import getpass
import os
import shutil
import pyzipper


#---function definitions
def refresh():
    c.execute('SELECT oid, * FROM ypotheseis')
    data = c.fetchall()
    win2['-TABLE-'].update(values=data)
    win2['-UPDPROTOCOLINPUT-'].update(value='')
    win2['-UPDADIKIMAINPUT-'].update(value='')
    win2['-UPDKATASTASHINPUT-'].update(value='') 
    win2['-UPDPERAIWSHINPUT-'].update(value='')
    win2['-UPDSXOLIAINPUT-'].update(value='')
    win2['-PROTOCOLINPUT-'].update(value='')
    win2['-UPDALERTINPUT-'].update(value='')

def export2CSV():
    try:
        csv_content = 'ID,ΠΡΩΤΟΚΟΛΛΟ,ΑΔΙΚΗΜΑ,ΚΑΤΑΣΤΑΣΗ,ΧΡΟΝΟΣ_ΠΕΡΑΙΩΣΗΣ,ΣΧΟΛΙΑ,ΥΠΕΝΘΥΜΙΣΗ\n'
        c.execute('SELECT oid, * FROM ypotheseis')
        for row in c.fetchall():
            sx = row[5]
            sx_fin = sx[:sx.find('\n')]
            csv_content += f'{row[0]},{row[1]},{row[2]},{row[3]},{row[4]},{sx_fin},{row[6]}\n' 
        
        with open('csv_report.csv', 'w', encoding='utf-8') as fout:
            fout.write(csv_content)
        sg.PopupOK(f'Το αρχείο CSV δημιουργήθηκε επιτυχώς\nστη διαδρομή{root}\\CASES_files!', title='!')
    except Exception as e:
        print(e)
        sg.PopupOK('Σφάλμα κατά τη δημιουργία του CSV!', title='!')

def peraiwsiAlert():
    try:
        #--- geniko output
        geniko_list = ''
        c.execute('SELECT oid, * FROM ypotheseis')
        for row in c.fetchall():
            if row[4] != '':
                geniko_list+=f'\n{row[1]}: {row[4]}'
        if len(geniko_list) == 0:
            sg.PopupOK('Δεν υπάρχουν υποθέσεις με χρόνο περαίωσης στη βάση', title='!')
        else:
            sg.PopupOK(f'Υποθέσεις με χρόνο περαίωσης:{geniko_list}', title='!')
            #--- eidopoihsh gia 2 evdomades
            protocol_list = []        
            today_dt= datetime.today() # datetime object ths shmerinhs meras
            c.execute('SELECT oid, * FROM ypotheseis')
            for row in c.fetchall():
                if row[4] != '':
                    t_dt = datetime.strptime(row[4], '%d-%m-%Y') # datetime object tou xronou peraiwshs KINDYNOS na exei timh pou den mporei na metatrapei se datetime object
                    difference = (t_dt-today_dt).days
                    if difference>0 and difference<30:
                        protocol_list.append(f'{row[1]}')                     
            if len(protocol_list) == 0:
                sg.PopupOK('Δεν υπάρχουν υποθέσεις με χρόνο περαίωσης που να λήγει εντός μήνα!', title='!')
            elif len(protocol_list) == 1:
                sg.PopupOK(f'H υπόθεση {protocol_list[0]} έχει χρόνο περαίωσης που λήγει εντός μήνα!', title='!')
            else:
                protocol_str = ''
                for prot in protocol_list:
                    protocol_str += f'{prot} ' 
                sg.PopupOK(f'Οι υποθέσεις {protocol_str}έχουν χρόνο περαίωσης που λήγει εντός 2 εβδομάδων!', title='!')
    except: #Exception as e:
        # print(e)
        sg.PopupOK('Σφάλμα ανάγνωσης της Βάσης Δεδομένων\nΑνεπιτυχής έλεγχος για χρόνους περαίωσης!', title='!') #8a skasei an sth sthlh xr_peraiwshs exei mpei otidhpote allo ektos apo keno (NULL) ;h hmeromhnia typoy 03-11-2019

def deleteYpothesh(oid, protocol):
    try:
        c.execute('DELETE FROM ypotheseis WHERE oid == (?)', (oid,))
        conn.commit()
        sg.PopupOK(f'H υπόθεση με αριθμό πρωτοκόλλου {protocol} διαγράφηκε από τη βάση', title='!')
    except:
        sg.PopupOK('Σφάλμα ανάγνωσης της Βάσης Δεδομένων\nΑνεπιτυχής διαγραφή υπόθεσης!', title='Warning')

def updateYpothesh(prot, adik, kat, per, sxol, yp):
    try:
        if sxol == '\n': # den exei grapsei sxolia alla emfanizei to newline to table opote to allazw se '' egw
            c.execute('UPDATE ypotheseis SET protocol = (?), adikhma = (?), katastash = (?), xr_peraiwshs = (?), sxolia = (?), alert = (?) WHERE protocol == (?)', (prot, adik, kat, per, '', yp, prot,))
            conn.commit()
            sg.PopupOK(f'Eπιτυχές update της υπόθεσης {prot}!', title=':)')
        else:
            c.execute('UPDATE ypotheseis SET protocol = (?), adikhma = (?), katastash = (?), xr_peraiwshs = (?), sxolia = (?), alert = (?) WHERE protocol == (?)', (prot, adik, kat, per, sxol, yp, prot,))
            conn.commit()
            sg.PopupOK(f'Eπιτυχές update της υπόθεσης {prot}!', title=':)')
    except: #Exception as e:
        # print(e)
        sg.PopupOK('Σφάλμα ανάγνωσης της Βάσης Δεδομένων\nΑνεπιτυχές update υπόθεσης!', title='Warning')

def insertYpothesh(prot, adik, kat, per, sxol, yp):
    try:
        if sxol == '\n': # den exei grapsei sxolia alla emfanizei to newline to table opote to allazw se '' egw
            c.execute('INSERT INTO ypotheseis VALUES (?,?,?,?,?,?)', (prot, adik, kat, per, '', yp,))
            conn.commit()
            sg.PopupOK(f'Eπιτυχής εισαγωγή της υπόθεσης {prot}!', title=':)')
        else:
            c.execute('INSERT INTO ypotheseis VALUES (?,?,?,?,?,?)', (prot, adik, kat, per, sxol, yp,))
            conn.commit()
            sg.PopupOK(f'Eπιτυχής εισαγωγή της υπόθεσης {prot}!', title=':)')
    except Exception as e:
        print(e)
        sg.PopupOK('Σφάλμα ανάγνωσης της Βάσης Δεδομένων\nΑνεπιτυχές update υπόθεσης!', title='Warning')

LoginPassword = '1234'
LoginImageData = b'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'
win2ImageData1 = b'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'
win2ImageData2 = b'iVBORw0KGgoAAAANSUhEUgAAAEsAAABECAYAAADJGMg/AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAtzSURBVHhe7VtrbBxXFT7z2BnPvr3xM042j9pNohA1UqqqTYoEqAGkokgRINEK+qdpoaKoEhJ/EKjwo4IqSAkoVRAlUlDaSkGoRUhFhrZShQi0pUnUhhJFkNLEcRzX8dqb3fXuzs6D79yZcRw7u95Ndt047Gcf3zvn3rkz99vzuHd2TG200UYbbbTRRhtt1ILkS8OQ/fJ2Q0AIz0+FaBBDluXYoUOH7jhw4ECf39YQbojhWwgBKYKY1atXq9u3b9eHhoYid911VxfKwVQqNRQOhzdANmuaNgjCOsfHx5/p7+//Mc6pQOrGciAruEe2BK4rhmEou3fvDkN6Mem+np6eoWg0Ogj9YCgUWqWqKksP9wf4XK6IkjE9Pf0SSHwU1ZKnqQ+3Ell8L4EwMbKiKOqOHTv0PXv29A8ODq5bsWLFUCQSGQIpG2Al62AlSfRJgAh5PinXIylAqVR6F1b4uUwmk8eh62kXx8KRWovgelwKQvr6+tRdu3aFcfPRjRs39q1fv34TCFnPrqPr+hBbCibcAWIMlApPvhYRteC6Hi+2bU8fOXLki5s2bcoODw+P7N27tyAaFkGzyJo7TlAP3IZLBW6hPPbYY8mtW7f2Av2JRCIN11nPrgNS1sJS0r6VMAQRrRTAQqmOjIzsSqfTr+LYYWUt3AxZ0s6dOztWrVoVRtzQu7u7w/F4PAKLiGLyMUw8grIzFouthaVwPLkDVtIL0lIQzkwMceNcBhIcB5OaL+LCfsmopx7gerrR0dGHMYejqLaELHnbtm0da9euje3bt+/3IIUzjIYbYRdh4fgRSIDZG623Xk/J4HrgXvP18zFfx8fIjI8jFBzCYfPJgo9zTNkPgtYNDAykYT06X3S+MGrpFqvXUzKYqGKxRDPF4ixpjLl9Asw9v0PXCPMgBPmnkE2fg9oWjTWwcMRFcPTo0fSaNWtOoxqG6xHcS1w8EMbc4/m6RurXKxlzdaZp0nNv/Id++97UNWQthju7dTr40GacX/4BYuizUFleS3VcvYM6oMR7Ojbe+8Dap5946PWQqgwkk0lBFu5a3Lj4wYjeZLz61Ta/LgrRaY6Oj/nXL4M21oiKrxc6r92vURlkffuVUfpwqqH1pcDzu/tIreR+vWvP9/aW/n1sDKqaS4ngqvVA6nnk4M5Q97oXNVXqlCVSxDRnb7wG6rxKQFGjKFouOfUb1Sx0RbhxxbScklMu/G7swJe/BbXptS4Ep/W6IRvxNJaKXaYjKyVbpqIt0QyMd1HBh16PFCruDcmNEMUoI0qZjhTCnGKSqt8JVU0+GiKrgcXuckVN026QrP9vNI8szkTLQW7COxoK8H3ffOFRJd77vH98Fa5DbjHLFe/4VoakkGTEueId+3Ar5WMX939pJ6pFT7MQzSHrNkA9ZC1NzIL5O+YMuVbZV1yFa5lCyLl2t+HaFUI699p8uI4923/2vACwbtG/svAazULLyXIdiypTF8gpTJGVvUTWNK/9PHCbNX0R+jGyZzK+FnqzSNbURbILGbSPipJh5y97Y1zxZfK8GIM/hArqDvfHWNynFWg5Wc7MtPB1SdUQLkJQ2NBxfANsCzpV6F3UBWCFdmGSJFkmpSOGNo0cjoewHElGvIHIWoRkPerHHgxTwDVkjKOFxXXYutxKQw9B68ISWBaCPwiQw0lSwp1gTRLWINrgaiJDsQsGZHGSYB368TkyCJM7EmIc0QrS5I4oKZEUqfEeQRIGEmMqkU6hl/WIuGazsQQxi2/av3FYEWYnrIMh4o4gkwXLaZTcjj+eHu6FXQMp0ZSwwKpgq0K7U8yRFOogJdYF68OetclYArKugt1NTfYLi2EEAZpdShDElgYwOUywfWVigTvxXpTdzr4yPtvG1sRu7JRz5JRyQtcKLClZ14CtiK2Jn4qqsAyUAXkcexSDCXVFsL6GMLY820R2Lc6Sy1alxHowpC2SAcfJVmBJyWK34slwKWIQ4gpbCme/uZbFkMMJSKcgws5d9rXMsU0yx6vkSkFqAFkPI4b1gUv0Z7LQr9lYWsviOFSe8VzFD/IMQRKTxpYlrGMKffIikHMMEjFNEOpBZFDOriL2YQ0HckTWDekiuLO722JH0Vy0nCyJ3QYirAnuw6RgNh4xnPEQr9hShCX41uCUrvjLBc6MPAhXvWzIcCu8IOXlAS9AJUGsIMfPorMnNRmtJ4s/aZQOgrKwKEyaM5VYTPLkYCEyu5MgFATyWorXY4CVGRGkkAtSObsxF2i3iwjw2XGqcDuv1dgCoa9ksDAFca4NV4VbNgYmuDZaThYTISyHXQYWpcS6hWtJvKEFKWKdBIgFJdyISVQT/TgGObxcAIlqokeUkor+0Ml8PtrZ5fiTUHg9ZiQ8ktFPieMaqi7GbSYaJiukSLQyodFQt0GrkjoZIX8Ikd2uL2wVaqJXCMcaoTNi3qKSCcGxgoCuJhCgMVkoBAF8LM5hEvgcLFCVaBdkhRDRxv25DcQF/ZlM1lWV64MdoKbv1u3Y586dX//BWP77RdN6NKR6VsI8hciiDzMVeuX4BTp1qbT4l2+fNNii4bbzp44E8veLP9/1AKoznmYhFiVr5MKobJrmg9NF+9CbZy53n8kqZBIu6LczVmgO3d3jUq5k0U+O5cWz7eUGnyx+RFP1vYdF3dBxnMens7lnXziRUU9Nh6iCU2QJez0ZglKCTJoSDY8olLFC9KP7DeqosTO5ZSFJ7BQ1HaMmWefOj9wzMTHxJBL1sCmHbAW9VcQsVXwPBpktcS2Qd3pKojErTN/ZiiBbt4PfKlj8hquS9dG589Fyufysoiqvxbu7LQ1xSg+ppGNroodCogxBVDDFL3KoII1j7b8yLnWnkrS165aPXg2jlmVtKRaLqXiyM5eKRAogygopSP0gRubFJBME0Zg8Jk31CYNbsrt+ZjVSvT/Q7YKqZCFW3Q/LOh0NG5mopl0ADVVNhcnTsATgLMn1iYJFqzs7KKIuvtBbTqhKlm3bm+FmZxCTphRFqSu/seVpEP6GuOiolNSWYVqsgapkwc0MSZZMlA2Zhwqy2LqwCYFb+srbBLXImlRkJQ4L8/ZwDUBGllRdLDJs3ujePq5YlSxVVd+BDPL7TxYIqwfMqYM/HVg3hN0ifT2dofvl097ThtsAVcmSJOnP4XD4vpmZGblcxjamDuviPjYCVl/YoYSh0j13byN19B80/qtH6MpfD7fsCeZSoSpZqwZWXtQ0bVjT9AfzubwgoRpfrOZ2C8Jh6g79Cq1bs4a6ulbQP0+dIsXMUf6tl2js4Ndo8pWnqXT2LTC7/IJ/VbIYkiw/ZRjGp1zH2WhWbJDhYEnhkcYEseDQI8p2EdBdurerRMc+KtDzxy7Q4T/+jbRwjLZs2ULpdJri0QhV/vs2Tb78Q/r4N09Q7p2j3kO+ZYJF89XExOWN8Mk3Xn4/k3xnpBDmrQ5nO/wK8KN0G+zxmmpzvETvnZ+mP3ykeJYI65Gx3QpPvE/xS++SlB2hfC7H/w7CL76K90FJ1cm489MU3babQr1D3qCfAJr2YkgmM9WJGPbLsxMzXzlxIS+fnzYpX3bEHjGpEw0YFTKLOXrtbImOfyyTWWX5qk2dJePSCYpc/kB8QzM5OSmkUCjwuo5CKzdTZMsXyNj02ZY8vKsF1wJZ+5r4Fs1Th978biQa/9nmbuwNZZsKJZMuZcv03rhJxy85lL/65UxNyJUC6WMnKTb2NmnlKf5fGn4fnbLZLJXLJqW++lPS01v93ksD1zL/cnHfg59HteqbJY2QRV0P7x/QeoeeQdV7IH6TkOC+2ujbZJx5tSdmaPdBpU9lc9n8vU++Tp3pJcwA/DWc+aexX+w+jIOq27qGyPLBSeFGzquJDRs2rEPxjVQq9eHJkydfhLXd/KfRGJikpb5mG2200UYbbbTRRhtttNEiEP0PFf4oqCoXLZAAAAAASUVORK5CYII='

#---win1 layout
layout = [[sg.Image(data=LoginImageData)],
            [sg.Text('ENTER PASSWORD TO LOGIN')],  
            [sg.Input(key='-PASS-', password_char='*', size=(15,1))],               
            [sg.Button('Exit', size=(7,1)), sg.Button('Login', bind_return_key=True, size=(7,1))]]  

win1 = sg.Window('CASES Login Page', layout, element_justification='c', size=(350, 350))  
win2NotActive = True  
while True:  
    ev1, vals1 = win1.Read() 

    if ev1 in (sg.WIN_CLOSED, 'Exit'):  
        break  
    
    if ev1 == 'Login'  and win2NotActive:
        if vals1['-PASS-'] == '':
            sg.PopupOK('Enter a Password', title='Warning')
        elif vals1['-PASS-'] == LoginPassword:             
            try:
                #---decrypting database
                root = os.getcwd()
                targetfolder = f'{root}\\CASES_files' # o fakelos ths bashs
                db_pass = b'ar3youtalking2mep@5s'
                os.chdir(targetfolder)
                with pyzipper.AESZipFile('CASES.zip') as myzip:
                    myzip.setpassword(db_pass)
                    myzip.extractall()

                os.remove('CASES.zip')
                #---making db connection                
                conn = sqlite3.connect(f'{root}\\CASES_files\\CASES.db') 
                c = conn.cursor()
                c.execute('SELECT oid, * FROM ypotheseis')
                data = c.fetchall()
                alertRunOnce = True
                # ---ALERTS
                if alertRunOnce:
                        alertRunOnce = False # gia na treksei mono mia fora
                        c.execute('SELECT oid, * FROM ypotheseis')
                        for row in c.fetchall():
                            if row[6] != '':
                                sg.PopupOK(f'Υπενθύμιση για την {row[1]}: {row[6]}!', title='!')
                win1.Hide()
                win2NotActive = False
                 
                #---win2 layout
                menu_def = [['File', ['Export to CSV', 'Decrypt Files', 'Exit']],
                            ['Help', ['Documentation', 'About']],]           

                UpdateFrameLayout = [[sg.Input(key='-UPDPROTOCOLINPUT-', size=(52,1)), sg.Text('Πρωτόκολλο Υπόθεσης')],
                                    [sg.Input(key='-UPDADIKIMAINPUT-', size=(52,1)), sg.Text('Αδίκημα Υπόθεσης')],
                                    [sg.Input(key='-UPDKATASTASHINPUT-', size=(52,1)), sg.Text('Κατάσταση Υπόθεσης')],
                                    [sg.Input(key='-UPDPERAIWSHINPUT-', size=(52,1)), sg.Text('Χρόνος Περαίωσης')],
                                    [sg.Multiline(key='-UPDSXOLIAINPUT-', size=(50,5)), sg.Text('Σχόλια Υπόθεσης')],
                                    [sg.Input(key='-UPDALERTINPUT-', size=(52,2)), sg.Text('Μήνυμα υπενθύμισης')],                                                               
                                    [sg.Button('Update', size=(8,1)), sg.Button('Insert', size=(8,1)), sg.Button('Delete', size=(8,1))]]

                layout2 = [ [sg.Menu(menu_def, key='-MENUBAR-')],
                            [sg.Image(data=win2ImageData1), sg.Text('ΚΑΤΣΟΥΛΗΣ CASES'), sg.Image(data=win2ImageData2)],
                            [sg.Input(key='-PROTOCOLINPUT-', size=(10,1)), sg.Button('Find', size=(8,1), bind_return_key=True,), sg.Button('Refresh', size=(8,1))],
                            [sg.Table(key='-TABLE-', font=('arial',12), values= data, headings=[' ID ',' ΠΡΩΤΟΚΟΛΛΟ ', 'ΑΔΙΚΗΜΑ', 'ΚΑΤΑΣΤΑΣΗ', 'ΧΡΟΝΟΣ ΠΕΡΑΙΩΣΗΣ', 'ΣΧΟΛΙΑ', 'ΥΠΕΝΘΥΜΙΣΗ'], select_mode=sg.TABLE_SELECT_MODE_BROWSE, auto_size_columns=True, justification='center', enable_events=True, alternating_row_color= 'grey', num_rows=20, selected_row_colors=('white','#c23c53'))],     
                            [sg.Frame('Ενημέρωση-Εισαγωγή-Διαγραφή Υπόθεσης', UpdateFrameLayout, element_justification='left')],
                            [sg.Button('Έξοδος', size=(8,1)), sg.Button('Χρόνοι', size=(8,1))]]  

                win2 = sg.Window('CASES', layout2, element_justification='c')
                # ---ALERTS
                # if alertRunOnce:
                #         alertRunOnce = False # gia na treksei mono mia fora
                #         c.execute('SELECT oid, * FROM ypotheseis')
                #         for row in c.fetchall():
                #             if row[5] != '':
                #                 sg.PopupOK(f'Υπενθύμιση για την {row[1]}: {row[5]}!', title='!')  
                while True:  
                    ev2, vals2 = win2.Read() 
                    # print(ev2, vals2)

                    #---menu events
                    if ev2 == 'Export to CSV':
                        export2CSV()
                    if ev2 == 'Decrypt Files':
                        sg.PopupOK('Full paid version only!', title='!')
                    if ev2 == 'Documentation':
                        sg.PopupOK('FIND: Ψάχνει συγκεκριμένο αριθμό πρωτοκόλλου και παρουσιάζει μόνο αυτή την εγγραφή στον πίνακα\n\nREFRESH: Κάνει Refresh τον πίνακα και φέρνει όλα τα αποτελέσματα της βάσης\n\nUPDATE: Κάνει update την εγγραφή που έχει επιλεγεί στον πίνακα με τα στοιχεία που έχουν δοθεί στα πεδία \n\nINSERT: Εισάγει στη βάση τη νέα υπόθεση με τα στοιχεία που έχουν δοθεί στα πεδία\n\nDELETE: Διαγράφει από τη βάση την υπόθεση που έχει επιλεχτεί στον πίνακα \n\nΧΡΟΝΟΙ: Τσεκάρει τη βάση και ειδοποιεί για υποθέσεις με χρόνους περαίωσης που λήγουν εντός 2 εβδομάδων\n\nΜΗΝΥΜΑ ΥΠΕΝΘΥΜΙΣΗΣ: Pop Up μήνυμα με το περιεχόμενο της υπενθύμισης κάθε φορά που κάνει login ο χρήστης', title='Documentation')
                    if ev2 == 'About':
                        sg.PopupOK('CASES Ver. 1.0.1 \n\n --DKats 2020', title='-About-')

                    #---button events    
                    if ev2 == 'Χρόνοι':
                        peraiwsiAlert() 

                    if ev2 == '-TABLE-':
                        table_list = win2['-TABLE-'].Get()
                        tableRow_oid = table_list[vals2['-TABLE-'][0]][0] # oid ths eggrafhs pou pathsa sto table table_list[x][0], opou x to row tou table kai 0 to prwto stoixeio tou tuple apo to get dhladh to oid
                        c.execute('SELECT oid, * FROM ypotheseis WHERE oid == (?)', (tableRow_oid,))
                        row = c.fetchone()
                        win2['-UPDPROTOCOLINPUT-'].update(value=row[1])
                        win2['-UPDADIKIMAINPUT-'].update(value=row[2])
                        win2['-UPDKATASTASHINPUT-'].update(value=row[3]) 
                        win2['-UPDPERAIWSHINPUT-'].update(value=row[4])
                        win2['-UPDSXOLIAINPUT-'].update(value=row[5])
                        win2['-UPDALERTINPUT-'].update(value=row[6])

                    if ev2 == 'Find':
                        if vals2['-PROTOCOLINPUT-'] == '':
                            sg.PopupOK(f'Δώσε αριθμό πρωτοκόλλου για εύρεση!', title='!')
                        else:
                            c.execute('SELECT oid, * FROM ypotheseis WHERE protocol == (?)', (vals2['-PROTOCOLINPUT-'],)) 
                            if c.fetchone() != None:
                                c.execute('SELECT oid, * FROM ypotheseis WHERE protocol == (?)', (vals2['-PROTOCOLINPUT-'],))
                                data = c.fetchall()
                                win2['-TABLE-'].update(values=data)
                                win2['-UPDPROTOCOLINPUT-'].update(value='')
                                win2['-UPDADIKIMAINPUT-'].update(value='')
                                win2['-UPDKATASTASHINPUT-'].update(value='') 
                                win2['-UPDPERAIWSHINPUT-'].update(value='')
                                win2['-UPDSXOLIAINPUT-'].update(value='')
                                win2['-UPDALERTINPUT-'].update(value='')
                            else:
                                sg.PopupOK(f'Ο αριθμός πρωτοκόλλου για εύρεση δεν υπάρχει στη βάση!', title='!')
                                win2['-UPDPROTOCOLINPUT-'].update(value='')
                                win2['-UPDADIKIMAINPUT-'].update(value='')
                                win2['-UPDKATASTASHINPUT-'].update(value='') 
                                win2['-UPDPERAIWSHINPUT-'].update(value='')
                                win2['-UPDSXOLIAINPUT-'].update(value='')
                                win2['-PROTOCOLINPUT-'].update(value='')
                                win2['-UPDALERTINPUT-'].update(value='')
                    
                    if ev2 == 'Refresh':
                        refresh()

                    if ev2 == 'Update':
                        if vals2['-UPDPROTOCOLINPUT-'] == '': # pathse update xwris na exei epileksei ypothesh h' me sbhsmenh timh
                            sg.PopupOK(f'Δώσε αριθμό πρωτοκόλλου υπόθεσης για να γίνει update!', title='!')
                        else: # exei dwsei timh sto protocol otan pathse update
                            c.execute('SELECT oid, * FROM ypotheseis WHERE protocol == (?)', (vals2['-UPDPROTOCOLINPUT-'],)) 
                            if c.fetchone() != None: # ara yparxei ypothesh me ayto to protocolo sth bash
                                if vals2['-UPDPERAIWSHINPUT-'] == '': # den exei dwsei xrono peraiwshs opote proxwraw kanonika sto update
                                    updateYpothesh(vals2['-UPDPROTOCOLINPUT-'], vals2['-UPDADIKIMAINPUT-'], vals2['-UPDKATASTASHINPUT-'], vals2['-UPDPERAIWSHINPUT-'], vals2['-UPDSXOLIAINPUT-'], vals2['-UPDALERTINPUT-'])
                                else: # exei dwsei xrono peraiwshs opote prepei na kanw sanity check sto format ths hmeromhnias prin kanw to update
                                    allowed_chars = ['0','1','2','3','4','5','6','7','8','9','-']
                                    if all(chars in allowed_chars for chars in vals2['-UPDPERAIWSHINPUT-']):
                                        updateYpothesh(vals2['-UPDPROTOCOLINPUT-'], vals2['-UPDADIKIMAINPUT-'], vals2['-UPDKATASTASHINPUT-'], vals2['-UPDPERAIWSHINPUT-'], vals2['-UPDSXOLIAINPUT-'], vals2['-UPDALERTINPUT-'])
                                    else:
                                        sg.PopupOK(f'Λάθος format ημερομηνίας χρόνου περαίωσης! Αποδεκτή μορφή: π.χ. 01-01-1970', title='!') 
                            else: # den yparxei sth bash ypothesh me to protocolo pou edwse gia update 
                                sg.PopupOK(f'Δεν υπάρχει στη βάση υπόθεση με τον αριθμό πρωτοκόλλου που έδωσες για να γίνει update!', title='!') 

                    if ev2 == 'Insert':
                        if vals2['-UPDPROTOCOLINPUT-'] == '': # pathse insert xwris na exei timh sto protocolo
                            sg.PopupOK(f'Δώσε αριθμό πρωτοκόλλου υπόθεσης προς εισαγωγή!', title='!')
                        else:
                            c.execute('SELECT oid, * FROM ypotheseis WHERE protocol == (?)', (vals2['-UPDPROTOCOLINPUT-'],)) 
                            if len(c.fetchall()) != 0: # to protocolo pou edwse gia to neo insert yparxei hdh sth bash opote den ginetai na eisax8ei
                                sg.PopupOK(f'Yπόθεση με ίδιο αριθμό πρωτοκόλλου υπάρχει ήδη στη βάση! Δοκίμασε update', title='!')
                            else: # to protocolo den yparxei hdh opote mporei na eisax8ei
                                if vals2['-UPDPERAIWSHINPUT-'] == '': # den exei dwsei xrono peraiwshs opote proxwraw kanonika sto insert xwris kanena elegxo
                                    insertYpothesh(vals2['-UPDPROTOCOLINPUT-'], vals2['-UPDADIKIMAINPUT-'], vals2['-UPDKATASTASHINPUT-'], vals2['-UPDPERAIWSHINPUT-'], vals2['-UPDSXOLIAINPUT-'], vals2['-UPDALERTINPUT-'])
                                else: # exei dwsei xrono peraiwshs opote prepei na kanw sanity check sto format ths hmeromhnias prin kanw to insert
                                    allowed_chars = ['0','1','2','3','4','5','6','7','8','9','-']
                                    if all(chars in allowed_chars for chars in vals2['-UPDPERAIWSHINPUT-']):
                                        insertYpothesh(vals2['-UPDPROTOCOLINPUT-'], vals2['-UPDADIKIMAINPUT-'], vals2['-UPDKATASTASHINPUT-'], vals2['-UPDPERAIWSHINPUT-'], vals2['-UPDSXOLIAINPUT-'], vals2['-UPDALERTINPUT-'])
                                    else:
                                        sg.PopupOK(f'Λάθος format ημερομηνίας χρόνου περαίωσης! Αποδεκτή μορφή: π.χ. 01-01-1970', title='!') 

                    if ev2 == 'Delete':
                        if len(vals2['-TABLE-']) == 0:
                            sg.PopupOK('Επέλεξε υπόθεση από τον πίνακα για διαγραφή', title='Warning')
                        else:
                            table_list = win2['-TABLE-'].Get()
                            tableRow_oid = table_list[vals2['-TABLE-'][0]][0]
                            tableRow_protocol = table_list[vals2['-TABLE-'][0]][1]
                            deleteYpothesh(tableRow_oid, tableRow_protocol)
                            refresh()
                            

                    if ev2 in (sg.WIN_CLOSED, 'Exit', 'Έξοδος'): 
                        conn.close() 
                        #---encrypting db copying it to APPDATA
                        os.chdir(targetfolder)
                        with pyzipper.AESZipFile('CASES.zip', 'w', encryption=pyzipper.WZ_AES) as myzip:
                            myzip.setpassword(db_pass)
                            myzip.write('CASES.db')
                        #---copying the encrypted db to APPDATA    
                        user = getpass.getuser()                        
                        if not os.path.exists(f'C:\\Users\\{user}\\AppData\\Local\\DKatsForensics'):
                            os.mkdir(f'C:\\Users\\{user}\\AppData\\Local\\DKatsForensics')
                        shutil.copy('CASES.zip', f'C:\\Users\\{user}\\AppData\\Local\\DKatsForensics')
                        os.remove('CASES.db')                            
                        win2.Close()
                        win1.Close()  
                        break
            except Exception as e:
                # print(e)
                sg.PopupOK(f'Failure to run the app. Reason: {e}', title='!')
                win1['-PASS-'].update(value = '')
        else:
            sg.PopupOK('Wrong Password', title='Warning')
            win1['-PASS-'].update(value = '')
win1.Close()