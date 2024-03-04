import pyodbc
from datetime import datetime
import json
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from tkinter import messagebox
import os
from sql import conn

cheminexport="C:\\Cegid"
racineclient=""
comptetva=""
comptevente=""
date_premier_janvier = datetime(datetime.now().year, 1, 1).strftime('%d/%m/%Y')
date_aujourdhui = datetime.now().strftime('%d/%m/%Y')
cursor = conn.cursor()
with open('config.json', 'r') as fichier_config:
    tableau_data = json.load(fichier_config)
def main():
    start_date = date_entry.get()
    end_date = end_date_entry.get()
    checkbox_achat = purchases_var.get()
    checkbox_vente = sales_var.get()

    # Création du tableau pour stocker les résultats
    resultats = []

    #FACTURE VENTES
    if checkbox_vente == True and checkbox_achat == False:
        cursor.execute("SELECT Code, Date, HTNet0, HTNet1, HTNet2, HTNet3, HTNet4, HTNet5, HTNet6, HTNet7, HTNet8, HTNet9, TVA0, TVA1, TVA2, TVA3, TVA4, TVA5, TVA6, TVA7, TVA8, TVA9, Taux1, Taux2, Taux3, Taux4, Taux5, Taux6, Taux7, Taux8, Taux9 , Avoir, CodeClient, Nom, SPP, JOURNAL FROM Facture WHERE TransCompta='False' AND PieceEdit='True' AND JOURNAL<>'-' AND Date BETWEEN ? AND ?", (start_date, end_date))
        factures = cursor.fetchall()
        if factures:

            # Parcours des factures et construction du tableau des résultats
            for facture in factures:
                code_facture, date_facture, *htnets, TVA0, TVA1, TVA2, TVA3, TVA4, TVA5, TVA6, TVA7, TVA8, TVA9, Taux1, Taux2, Taux3, Taux4, Taux5, Taux6, Taux7, Taux8, Taux9 , avoir, Code_Client, Nom, SPP, journal = facture
                code_facture_new = code_facture[-5:]

                Taux0 = 0
                if journal == 'T':
                    journal = 'VT'
                elif journal == 'E':
                    journal = 'VE'
                tva_variable = {
                    0: TVA0,
                    1: TVA1,
                    2: TVA2,
                    3: TVA3,
                    4: TVA4,
                    5: TVA5,
                    6: TVA6,
                    7: TVA7,
                    8: TVA8,
                    9: TVA9
                }
                taux_variable = {
                    0: Taux0,
                    1: Taux1,
                    2: Taux2,
                    3: Taux3,
                    4: Taux4,
                    5: Taux5,
                    6: Taux6,
                    7: Taux7,
                    8: Taux8,
                    9: Taux9
                }
                
                for i, htnet in enumerate(htnets):
                    if htnet != 0:
                        result_tableau = []
                        date_formattee = date_facture.strftime('%d%m%Y')
                        # Récupérer le code client
                        # cursor.execute("SELECT Code, Nom, Compte FROM ClientDef WHERE Code=?", (Code_Client,))
                        # nom_client = cursor.fetchone()[1]
                        cursor.execute("SELECT Code, Compte FROM ClientDef WHERE Code=?", (Code_Client,))
                        compte_client=cursor.fetchone()[1]
                        dateecheance=""
                        cursor.execute("SELECT NumeroDoc, DateEcheance FROM FinancEcheancier WHERE NumeroDoc=?",(code_facture,))
                        dateecheance=cursor.fetchall()
                        if dateecheance:
                            print(code_facture)
                            print(dateecheance)
                            dateecheance=dateecheance[0][1].strftime('%d%m%Y')
                        # Récupérer le compte du client
                        for entree in tableau_data['tableau']:
                            if entree['journal'] == journal and entree['tauxtva'] == taux_variable.get(i):
                                if SPP=='O':
                                    racineclient=str(entree['racineclient']) + 'SPP0'
                                elif SPP=='N':
                                    racineclient=str(entree['racineclient']) + compte_client[3:7]
                                else:
                                    racineclient=""
                                comptetva=entree['comptetva']
                                comptevente=entree['comptevente']
                        libelle = str("FACTURE " + Nom)
                        totalttctva=htnet+tva_variable.get(i)
                        # Ajouter au tableau des résultats
                        if avoir == 0:
                            resultats.append(("", str(code_facture_new), str(journal), str(date_formattee), "411000", str(racineclient), str(round(totalttctva,2)), "0.00", "", "", str(libelle),dateecheance,""))
                            resultats.append(("", str(code_facture_new), str(journal), str(date_formattee), str(comptevente), "", "0.00", str(htnet), "", "", str(libelle), "",""))
                            resultats.append(("", str(code_facture_new), str(journal), str(date_formattee), str(comptetva), "", "0.00", str(tva_variable.get(i)), "", "", str(libelle), "",""))
                        elif avoir == 1:
                            resultats.append(("", str(code_facture_new), str(journal), str(date_formattee), "411000", str(racineclient), "0.00", str(round(totalttctva,2)), "", "", str(libelle),dateecheance,""))
                            resultats.append(("", str(code_facture_new), str(journal), str(date_formattee), str(comptevente), "", str(htnet), "0.00", "", "", str(libelle), "",""))
                            resultats.append(("", str(code_facture_new), str(journal), str(date_formattee), str(comptetva), "", str(tva_variable.get(i)), "0.00", "", "", str(libelle), "",""))

                        maintenant = datetime.now().strftime("%d%m%y_%H%M%S")
                        nom_fichier = "ECRITURE_" + maintenant + ".txt"
                        # Chemin complet du fichier
                        chemin_complet = os.path.join(cheminexport, nom_fichier)
                        
                        # Ouvrir le fichier en écriture
                        with open(chemin_complet, 'w') as f:
                            # Parcourir le tableau et écrire chaque ligne dans le fichier
                            for ligne in resultats:
                                ligne_formatee = ";".join(map(str, ligne))
                                f.write(";" + ligne_formatee.strip(";") + ";\n")
                        exit
            # Affichage des résultats
            # for resultat in resultats:
            #     print(resultat)
            # print(resultats)
            # Fermeture de la connexion à la base de données
            messagebox.showinfo("Terminé", "Export terminé avec succès.")
        else:
            messagebox.showinfo("Terminé", "Aucune facture à transférer.")



            #FACTURES ACHATS
    elif checkbox_achat == True and checkbox_vente == False:

        messagebox.showinfo("Erreur", "Indisponible pour le moment")
        # #----------------OUT--------------------
        # # Exécution de la requête SQL principale pour récupérer les codes de facture
        # cursor.execute("SELECT Code, Date, Avoir, CodeFou, TotalNet FROM FactureFour WHERE TransCompta='False' AND Date BETWEEN ? AND ?", (start_date, end_date))
        # factures = cursor.fetchall()
        # if factures:
        #     resultats = []
        #     # Parcourir chaque code de facture
        #     for code in factures:  
        #         code_facture = factures[0][0]
        #         date_facture = factures[0][1]
        #         avoir = factures[0][2]
        #         CodeFou=factures[0][3]
        #         TotalTTC=factures[0][4]
        #         code_facture_new = code_facture[-5:]

        #         cursor.execute("SELECT Code, Nom, Compte FROM Fournisseur WHERE Code=?", (CodeFou,))
        #         nom_client = cursor.fetchone()[1]
        #         cursor.execute("SELECT Code, Compte FROM Fournisseur WHERE Code=?", (CodeFou,))
        #         compte_client=cursor.fetchone()[1]

        #         date_formattee = date_facture.strftime('%d%m%Y')
        #         racineclient="0"+compte_client[3:7]
        #         libelle = str("FACTURE " + nom_client)
        #         if avoir == 0:
        #             resultats.append(("",code_facture_new, "AC", date_formattee, "411000", racineclient, "0.00", TotalTTC, "","",libelle,""))
        #         elif avoir == 1:
        #             resultats.append(("",code_facture_new, "AC", date_formattee, "411000", racineclient, "0.00", TotalTTC, "","",libelle,""))

        #         # Exécuter la requête pour récupérer les lignes associées à ce code
        #         cursor.execute("SELECT ffl.code, compte,SUM(ffl.qte * ffl.panet) AS total FROM facturefourligne AS ffl JOIN facturefour AS ff ON ffl.code = ? GROUP BY ffl.code, ffl.compte;", (code_facture))
        #         lignes_facture = cursor.fetchall()
        #         for lignefac in lignes_facture:
        #             compte_ligne=lignefac[1][:6]
        #             cumul=lignefac[2]
        #             resultats.append(("",code_facture_new, "AC", date_formattee, compte_ligne, "", cumul , "0.00" ,"","",libelle,""))
        #     for resultat in resultats:
        #         print(resultat)
        #     maintenant = datetime.now().strftime("%d%m%y_%H%M%S")
        #     nom_fichier = "ECRITURE_" + maintenant + ".txt"
        #     # Chemin complet du fichier
        #     chemin_complet = os.path.join(cheminexport, nom_fichier)
            
        #     # Ouvrir le fichier en écriture
        #     with open(chemin_complet, 'w') as f:
        #         # Parcourir le tableau et écrire chaque ligne dans le fichier
        #         for ligne in resultats:
        #             ligne_formatee = ";".join(map(str, ligne))
        #             f.write(";" + ligne_formatee.strip(";") + ";\n")

        #     # conn.close()

        #     # Fermeture de la connexion à la base de données
        #     messagebox.showinfo("Terminé", "Export terminé avec succès.")
        # else:
        #     messagebox.showinfo("Terminé", "Aucune facture à transférer.")
        # #----------------OUT--------------------
            


    elif checkbox_vente == False and checkbox_achat == False:
        messagebox.showinfo("Erreur", "Selectione")
    elif checkbox_achat == True and checkbox_vente == True:
        messagebox.showinfo("Erreur", "en choisir un")

def format_date(entry):
    current_text = entry.get()
    if len(current_text) == 2:
        entry.insert(tk.END, '/')
    elif len(current_text) == 5:
        entry.insert(tk.END, '/')

def validate_date(entry):
    current_text = entry.get()
    if len(current_text) == 10:
        day, month, year = current_text.split('/')
        try:
            int(day)
            int(month)
            int(year)
            entry.config(fg="black")  # Change text color to black if the format is correct
            return True
        except ValueError:
            entry.config(fg="red")  # Change text color to red if there's an invalid format
            return False
    else:
        entry.config(fg="black")  # Reset text color to black if the length is not 10
        return False

def end_format_date(entry):
    current_text = entry.get()
    if len(current_text) == 2:
        entry.insert(tk.END, '/')
    elif len(current_text) == 5:
        entry.insert(tk.END, '/')

def end_validate_date(entry):
    current_text = entry.get()
    if len(current_text) == 10:
        day, month, year = current_text.split('/')
        try:
            int(day)
            int(month)
            int(year)
            entry.config(fg="black")  # Change text color to black if the format is correct
            return True
        except ValueError:
            entry.config(fg="red")  # Change text color to red if there's an invalid format
            return False
    else:
        entry.config(fg="black")  # Reset text color to black if the length is not 10
        return False
    
def on_entry_click(event):
    if date_entry.get() == date_premier_janvier:
        date_entry.delete(0, tk.END) # Supprime le texte actuel de l'entrée lorsqu'on clique dessus
        date_entry.config(fg = 'black')
def end_entry_click(event):
    if end_date_entry.get() == date_aujourdhui:
        end_date_entry.delete(0, tk.END) # Supprime le texte actuel de l'entrée lorsqu'on clique dessus
        end_date_entry.config(fg = 'black')
def fermer_application():
    root.destroy()

# Create the main window
root = tk.Tk()
root.title('TransfoEcritures')

# Create the 'Traitement Comptable' frame
frame = tk.LabelFrame(root, text='Traitement Comptable', padx=10, pady=10)
frame.pack(padx=10, pady=10)

# 'Période' label
period_label = tk.Label(frame, text='Période')
period_label.grid(row=0, column=0, columnspan=2, sticky='W')

# 'Date de début' field
start_date_label = tk.Label(frame, text='Date de début:')
start_date_label.grid(row=1, column=0, sticky='W')

entry_var = tk.StringVar()
entry_var.trace("w", lambda name, index, mode, sv=entry_var: format_date(date_entry))

date_entry = tk.Entry(frame, textvariable=entry_var, width=10)
date_entry.grid(row=1, column=1)

date_entry.insert(0, date_premier_janvier)
date_entry.bind('<FocusIn>', on_entry_click)

validate_date_cmd = root.register(lambda: validate_date(date_entry))
date_entry.config(validate="focusout", validatecommand=(validate_date_cmd,))

# 'Date de fin' field
end_date_label = tk.Label(frame, text='Date de fin:')
end_date_label.grid(row=2, column=0, sticky='W')

end_entry_var = tk.StringVar()
end_entry_var.trace("w", lambda name, index, mode, sv=end_entry_var: end_format_date(end_date_entry))

end_date_entry = tk.Entry(frame, textvariable=end_entry_var, width=10)
end_date_entry.grid(row=2, column=1)

end_date_entry.insert(0, date_aujourdhui)
end_date_entry.bind('<FocusIn>', end_entry_click)

end_validate_date_cmd = root.register(lambda: end_validate_date(end_date_entry))
end_date_entry.config(validate="focusout", validatecommand=(end_validate_date_cmd,))

# Checkboxes
sales_var = tk.BooleanVar()
purchases_var = tk.BooleanVar()

sales_check = tk.Checkbutton(frame, text='VENTES', variable=sales_var)
sales_check.grid(row=3, column=0, sticky='W')
purchases_check = tk.Checkbutton(frame, text='ACHATS', variable=purchases_var)
purchases_check.grid(row=3, column=1, sticky='W')

# Buttons
cancel_button = tk.Button(frame, text='Fermer', command=fermer_application)
cancel_button.grid(row=4, column=0, pady=(10,0))
ok_button = tk.Button(frame, text='OK', command=main)
ok_button.grid(row=4, column=1, pady=(10,0))

root.mainloop()