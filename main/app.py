import tkinter as tk
from tkinter import ttk, filedialog

from utilities import *


if __name__ == "__main__":

    def plan4Blocs():
        essais, essaisMatricesList, gaucheIndex, droiteIndex = backendWrapper4Blocs(
            dirname
        )

        if typePlanValue.get() == "8 blocs":
            dibujante = Dessinateur4blocs(
                dirname + nomPlanValue.get(),
                essais,
                essaisMatricesList,
                droiteIndex,
                gaucheIndex,
            )

        else:
            dibujante = Dessinateur2blocs(
                dirname + nomPlanValue.get(),
                essais,
                essaisMatricesList,
            )

        dibujante.openDraw()
        dibujante.planPrincipal()
        dibujante.drawBordures()
        dibujante.matriceCompteur()
        dibujante.fisher()
        dibujante.createPavePourLesCalculs()
        dibujante.etiquettes()
        dibujante.maquillage()
        dibujante.closeDraw()
        dibujante.picasso()

    # Create a GUI app
    app = tk.Tk()
    app.resizable(False, False)

    app.update_idletasks()

    # Specify the title and dimensions to the app
    app.title("Consolidateur d'essais")
    app.geometry("600x300")

    titreLabel = tk.Label(app, text="CONSOLIDATEUR D'ESSAIS", font="Helvetica 14 bold")
    titreLabel.place(x=170, y=30)

    nomduPlanLabel = tk.Label(app, text="Nom du plan à generer", font="Helvetica 12")
    nomduPlanLabel.place(x=20, y=100)

    nomPlanValue = tk.StringVar(app)
    typePlanValue = tk.StringVar(app)

    nomduPlan = tk.Entry(app, textvariable=nomPlanValue, width=25, font="Helvetica 12")
    nomduPlan.place(x=200, y=100)

    nomduPlanLabel = tk.Label(app, text="Type de plan", font="Helvetica 12")
    nomduPlanLabel.place(x=20, y=150)

    typePlan = ttk.Combobox(app, textvariable=typePlanValue, values=["4 blocs", "8 blocs"], width=25, font="Helvetica 12")
    typePlan.place(x=200, y=150)

    contactLabel = tk.Label(app, text="Un logiciel du service agronomique d'Agrial\nEn cas de problèmes techniques, contacter : j.agudelo@agrial.com", justify="left", font="Helvetica 8")
    contactLabel.place(x=20, y=260)

    execute = ttk.Button(app, text="Construire Plan", width=15, command=plan4Blocs)
    execute.place(x=450, y=220)

    dirname = filedialog.askdirectory() + "/"

    app.mainloop()
