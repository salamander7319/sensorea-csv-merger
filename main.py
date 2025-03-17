from tkinter import *
from tkinter import filedialog
from tkinter import messagebox  # Import messagebox separately
import pandas as pd
from datetime import datetime
import os
import csv

class ExcelMerger(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.title("Excel Sheet Merger")
        self.resizable(False, False)
        self.iconbitmap("icon.ico")
        self.mode=IntVar(value=0)
        
        #déclaration de variables vide
        self.file_paths = []
        self.merged_df = None
        
        #Architecture du programme-----------------------------------------------------------------
        self.frame = Frame(self)
        self.frame.pack()
        
        self.listbox = Listbox(self.frame, height=20, width=50)
        self.listbox.pack()
        
        #scrollbar = Scrollbar(self.frame, orient="vertical", command=self.listbox.yview)
        #scrollbar.pack(side=RIGHT, fill="y")
        
        self.button_canvas = Canvas(self.frame,width=300,height=50,bg="grey")
        self.button_canvas.pack()

        #Bouton d'importation
        self.button_browse = Button(self.button_canvas,text="Sélectionner...", bd = 0,fg="black",bg="white",font=("Calibri",10), command=self.select_files)
        self.button_browse.place(relx = 0.05, rely = 0.70, anchor="sw")
        
        #Bouton d'exportation
        self.button_extract = Button(self.button_canvas,text="Extraire...", state=DISABLED, bd = 0,fg="black",bg="white",font=("Calibri",10), command=self.file_save_path)
        self.button_extract.place(relx = 0.34, rely = 0.70, anchor="sw")

        #Bouton d'exportation
        self.button_clear = Button(self.button_canvas,text="Effacer", state=DISABLED, bd = 0,fg="black",bg="white",font=("Calibri",10), command=self.clear_listbox)
        self.button_clear.place(relx = 0.55, rely = 0.70, anchor="sw")
        
    def select_files(self): #Fonction pour importer
        self.file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files (.csv)", "*.csv")])

        if self.file_paths:
            i=0 #variable d'ittération qui sera utilisée pour la suppression des mauvais fichier
            delete_list=[]
            self.listbox.delete(0, END)  # Supprimer le contenu de la list box s'il y en a
            
            for file in self.file_paths:
                substring="_AI_" #Substring qui nous permettra de quérir la mesure du fichier qui nous intéresse
                filename = os.path.basename(file)  # Acquérir le nom du fichier
                filename_without_ext = os.path.splitext(filename)[0]  # Supprimer l'extension (.csv)
                filename_value = filename_without_ext[filename_without_ext.index(substring)+len(substring):-1] #la mesure du fichier qui nous intéresse
                self.listbox.insert(END, filename_value)  # Afficher le nom dans la fenetre
                self.enable_buttons()
                
            #Verification de l'intégrité des fichiers .csv sélectionnés (s'ils contiennent la colonne TimeStamp)
            for csv_file in self.file_paths: # Ouvrir chaque fichier .csv et les lire
                with open(csv_file, mode='r', newline='', encoding='utf-8') as file:
                    reader = csv.reader(file, delimiter=";")  # Create a CSV reader object
                    
                    #Lire la première ligne pour voir s'il y a une colonne "TimeStamp"
                    first_row = next(reader)
                    #print(first_row)
                    if "TimeStamp" not in first_row:
                        messagebox.showwarning("Warning", str(csv_file)+ " ne contient pas la colonne 'TimeStamp'.")
                        delete_list.append(i) #Mettre l'index du fichier .csv incorrecte dans la liste des fichiers à supprimer
                        
                        if self.listbox.size() == 0:
                            self.clear_listbox()
                        
                    #else: #Sinon, lire chaque lignes du fichier .csv comme prévu.
                        #print(first_row)
                        #for row in reader:  # Loop through each row
                            #print(row)  # Print the row as a list
                    i=i+1
            for j in reversed(range(len(delete_list))):#Supprimer le fichier .csv qui ne contient pas de "TimeStamp" de la liste
                self.listbox.delete(delete_list[j]) 


                
        
    def file_save_path(self): #Fonction pour exporter
        #Lire les données de chaque fichiers .csv
        merged_list=[] #Liste qui va contenir toutes les lignes de chaque fichier .csv avant d'être trié
        merged_header=["TimeStamp"]
        header=[]
        i=0 #Variable d'ittération pour chaque fichier (pour avancer d'une colonne quand on passe au prochain fichier)
        for csv_file in self.file_paths: 
                with open(csv_file, mode='r', newline='', encoding='utf-8') as file: #ouvrir le fichier de l'itération en tant que "file"
                    csv_file_value=self.file_value_name(csv_file,"_AI_")
                    timestamp_col=1 #variable qui compte où se trouve la colonne "TimeStamp"
                    timestamp_trouvé=0
                    value_col=1
                    value_trouvé=0
                    reader = csv.reader(file, delimiter=";")  # Create a CSV reader object
                    header = next(reader) #Définir header comme la première ligne du fichier
                    if not timestamp_trouvé:
                        for col in header:
                            if col == "TimeStamp": #Recherche de la colonne "timestamp"
                                print("'TimeStamp' se trouve sur la "+str(timestamp_col)+"e colonne !")
                                timestamp_trouvé=1
                                break
                            else:
                                timestamp_col+=1
                    if not value_trouvé:
                        for col in header:
                            if col == "Value": #Recherche de la colonne "value"
                                print("'Value' se trouve sur la "+str(value_col)+"e colonne !")
                                value_trouvé=1
                                break
                            else:
                                value_col+=1
                    #Code pour rajouter la colonne de l'ittération dans la liste des colonnes s'il ne l'était pas avant.
                    if csv_file_value not in merged_header:
                        merged_header.append(csv_file_value)
                    
                    data = [row for row in reader] #Pour ligne par ligne 
                    for row in data:
                        #merged_list.append([row[timestamp_col-1], row[value_col-1]]) #Rajouter chaque line du fichier CSV dans la grande liste
                        #merged_list+= [[row[timestamp_col-1], row[value_col-1]]]
                        merged_list += [[row[timestamp_col-1]] + [""] * i + [row[value_col-1]]]
                    i+=1
        merged_list.sort(key=lambda row: datetime.strptime(row[0], "%Y/%m/%d %H:%M:%S.%f"), reverse=True) #Trier par le TimeStamp
        print("MERGED HEADER \n", merged_header)
        print("MERGE : \n", merged_list)
                    
        # Ouvrir le dialogue de fichier pour extraire le fichier excel
        self.save_path = filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("Excel Files", "*.csv")],title="Save Merged File")
        if self.save_path:
            with open(self.save_path, "w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(merged_header)  # Write single row (header)
                writer.writerows(merged_list)  # Write multiple rows
                messagebox.showinfo("Success", f"Fichier sauvegardé vers :\n{self.save_path}")

    def clear_listbox(self):
        self.listbox.delete(0, END)  # Effacer le contenu
        self.button_extract.config(state=DISABLED)  # Desactiver le bouton "Extract" button
        self.button_clear.config(state=DISABLED)  # Desactiver le bouton "Clear" button

    
    def enable_buttons(self):
        self.button_extract.config(state=NORMAL)  # Activer le bouton "Extract" button
        self.button_clear.config(state=NORMAL)  # Activer le bouton "Clear" button
    
    def file_value_name(self,file,substring): #fonction pour renvoyer en string la valeur mesurée du nom du fichier
        a=(os.path.splitext(os.path.basename(file))[0]) #code pour enlever le chemin du fichier et le .csv à la fin
        return a[a.index(substring)+len(substring):-1] #code pour enlever les chaînes de caractères inutiles
    


root = ExcelMerger()
root.mainloop()

