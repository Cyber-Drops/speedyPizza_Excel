import os
from breezypythongui import EasyFrame
from tkinter import filedialog
from tkinter import PhotoImage
import pizza_exel_V1
import socket
#Finestra Principale
class SpeedyPizzaGui(EasyFrame):
    def __init__(self):
        EasyFrame.__init__(self, title="][---SpeedyPizzaPy---][",width=700, height=400, resizable=False)
        #self.default_path = pizza_exel_V1.percorso
        self.default_path = os.getcwd()

#Disposizione Pannelli
        dataPanel = self.addPanel(row=0, column=0, background="white")
        checkboxPanel = self.addPanel(row=1, column=0, background="black")
        buttonPanel = self.addPanel(row=2, column=0, background="black")

#Panel Uno
        dataPanel.addLabel(text= "PATH: ",row= 0, column= 0, sticky= "NSEW")
        self.inputPath = dataPanel.addTextField(text= "seleziona cartella", row= 0, column= 1, width= 70, sticky= "EW")
        self.buttonPath = dataPanel.addButton(text="Seleziona", row= 0, column= 2, command= self.selezionaPath)

#Panel Due
        self.checkInviomail = checkboxPanel.addCheckbutton(text="@-Invia Email", row=0,column=0, sticky= "NSEW")
        self.checkArchivia = checkboxPanel.addCheckbutton(text="Archivia Originale[]", row=0, column= 1, sticky= "NSEW")
        self.checkReport = checkboxPanel.addCheckbutton(text= "Report Classe", row= 0, column= 2)

#Panel tre
        self.buttonStart = buttonPanel.addButton(text= "Avvio->", row= 1, column= 1, columnspan= 2, command= self.avvio)
        self.buttonHelp = buttonPanel.addButton(text= "? Help", row= 1, column= 3, command= self.help)
        self.buttonHelp["width"] = 10
        self.buttonHelp["height"] = 2
        self.buttonStart["width"] = 10
        self.buttonStart["height"] = 2

        self.image = PhotoImage(file="pizza.gif")
        imageLabel = buttonPanel.addLabel(text="", row=1, column =0, sticky="NSEW")
        imageLabel["width"] = 200
        imageLabel["height"] = 150
        imageLabel["image"] = self.image

    def selezionaPath(self):
        self.nomeDirect = filedialog.askdirectory(parent=self,title="File Excel .xlsx")
        self.inputPath.setText(self.nomeDirect)
        self.selected_path = self.nomeDirect
        self.inputPath["state"] = "disabled"

    def avvio(self):
        try:
            new_directory, lista_excel, file_output, df_risultante,df_elaborato, data_ora = pizza_exel_V1.main(self.selected_path)
            if self.checkInviomail.isChecked():
                try: pizza_exel_V1.invia_mail(new_directory, file_output)
                except socket.gaierror as e: self.messageBox(title="Connection Error", message="Verifica lo stato della connesione internet")
            if self.checkReport.isChecked():
                df_report = pizza_exel_V1.df_risult_to_report_class(file_output)
                directory_classe = pizza_exel_V1.directory_class(percorso=self.selected_path, data_ora=data_ora)
                pizza_exel_V1.gestisci_file_classe(directory_classe,df_report,data_ora)
            if self.checkArchivia.isChecked():
                try: pizza_exel_V1.archivia(lista_excel, self.selected_path)
                except IsADirectoryError: self.messageBox(title="!!!Error!!!", message="non è possibile rimuovere cartelle")
                except FileNotFoundError: self.messageBox(title="!!!Error!!!", message="File originale non trovato in cartella")
            self.messageBox(title="Success!!!!", message="Operazione avvenuta con successo!")
        except UnboundLocalError as e:
            errore_elabora = str(e)
            if "df_elaborato" in errore_elabora or "lista_excel" in errore_elabora:
                self.messageBox(title="!!!Error!!!", message="Verifica la presenza del file .xlsx da elaborare", width=70, height=10)
        except AttributeError as e:
            e = str(e)
            self.messageBox(title="!!!Error!!!", message=f"Non hai selezionata il percorso: ({e[0:16]+e[41:60]})", width=70, height=10)

    def help(self):
        self.messageBox(title="??Help??", message= pizza_exel_V1.descrizione, width=100, height=30)
        pass

def main():
    SpeedyPizzaGui().mainloop()
if __name__ == "__main__":
    main()