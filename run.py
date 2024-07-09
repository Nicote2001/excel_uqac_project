import excel_reader, excel_writer
import tkinter as tk
from tkinter import filedialog
import time

class FileApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Trasnformateur de l'extreme de Grand-père")

        self.file_paths = [tk.StringVar() for _ in range(4)]
        
        self.create_widgets()

    def create_widgets(self):
        for i in range(4):
            if i ==0 :
                tk.Label(self.root, text=f"SAINT-DOMINIQUE:").grid(row=i, column=0, padx=10, pady=10)
            elif i == 1:
                tk.Label(self.root, text=f"SAINTE-FAMILLE:").grid(row=i, column=0, padx=10, pady=10)
            elif i == 2:
                tk.Label(self.root, text=f"SAINT-GERARD:").grid(row=i, column=0, padx=10, pady=10)
            else :
                tk.Label(self.root, text=f"SAINTE-THERESE:").grid(row=i, column=0, padx=10, pady=10)
            tk.Entry(self.root, textvariable=self.file_paths[i], width=50).grid(row=i, column=1, padx=10, pady=10)
            tk.Button(self.root, text="Parcourir", command=lambda i=i: self.browse_file(i)).grid(row=i, column=2, padx=10, pady=10)

        tk.Button(self.root, text="Générer", command=self.generate).grid(row=4, column=1, padx=10, pady=10)

        self.message_label = tk.Label(self.root, text="")
        self.message_label.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

    def browse_file(self, index):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_paths[index].set(file_path)


    def generate(self):
        files = [var.get() for var in self.file_paths]
        print("Fichiers sélectionnés:", files)
        self.startbackend()

    def startbackend(self):
        #debut programme
        files = [var.get() for var in self.file_paths]
        self.message_label.config(text="génération en cours",fg="red")
        final_list = excel_reader.ReadAllExcel(files)
        bilan_list = excel_writer.regroupement(final_list)
        excel_writer.WriteExcel(bilan_list)
        self.root.after(2000,self.message_label.config(text="génération en terminée !", fg="green"))

if __name__ == "__main__":
    root = tk.Tk()
    app = FileApp(root)
    root.mainloop()


