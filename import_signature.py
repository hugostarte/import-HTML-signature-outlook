__author__ = "Hugo Stawiarski"
__copyright__ = "Copyright 2023"
__license__ = "CC-by-nc-nd"
__version__ = "1.0"
__email__ = "contact@hugostawiarski.fr"
__status__ = "Production" 
import os
import tkinter as tk
import shutil
import subprocess
import time

class Application(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.creer_widgets()
    
    def afficher_widgets(self):
        # Affiche le contenu de la fenetre
        self.label_chemin.pack()
        self.titre_etape1.pack()
        self.etape1.pack()
        self.label_entry.pack()
        self.saisir_nom.pack()
        self.btn_valider.pack()
    
    
    def creer_widgets(self): 
        self.session_utilisateur = os.getlogin()
        self.chemin = "C:/Users/"+self.session_utilisateur+"/AppData/Roaming/Microsoft/Signatures/"
        self.label_chemin = tk.Label(self, text="Chemin de la signature Outlook : "+self.chemin)

        self.titre_etape1 = tk.Label(self, justify=tk.LEFT, font=('Arial',9,'bold','underline'), text="Prérequis :")
        self.etape1 = tk.Label(self, justify=tk.LEFT, text="- Rendez-vous sur outlook \n -Cliquez sur Nouveau message électronique.\n- Choisissez l’onglet Une signature puis Signatures...\n- Cliquez sur Nouveau puis entrez votre nom de signature.\n- Fermez Outlook")

        self.label_entry = tk.Label(self, text="Veuillez entrer le nom de signature saisi sur Outlook ")
        self.saisir_nom = tk.Entry(self)
        self.btn_valider = tk.Button(self, text="Valider", command=self.valider)

        self.afficher_widgets()

    def valider(self):
        print('Chemin de la signature Outlook : '+self.chemin)

        subprocess.Popen(["taskkill","/IM","outlook.exe","/F"],shell=True) # Fermeture de Outlook
        print('Temps de pause 2 secondes...')
        time.sleep(2)

        chemin_script = os.path.realpath(__file__).replace("import_signature.py","")  # Récuperer le chemin du script

        signature_avant = chemin_script+'signature.htm'

        print('Suppression de la signature de base : OK')
        os.remove(self.chemin+self.saisir_nom.get()+'.htm')

        print('Copie de la signature vers Outlook : OK')
        shutil.copy(signature_avant,self.chemin)
        

        print('Renommage du fichier source : OK')
        os.rename(self.chemin+'signature.htm', self.chemin+self.saisir_nom.get()+'.htm')

        print('Temps de pause 2 secondes...')
        time.sleep(2)
        os.startfile("outlook")



if __name__ == "__main__":
    app = Application()
    app.geometry("500x200")
    app.title("Importateur signature HTML Outlook 2016")
    app.mainloop()