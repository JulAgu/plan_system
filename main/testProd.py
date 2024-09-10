import tkinter as tk
from tkinter import ttk, filedialog

from utilities import *


# if __name__ == "__main__":

#     dirname = "src" + "/"
#     nomPlanValue = "Test"

#     essais, essaisMatricesList, gaucheIndex, droiteIndex = backendWrapper4Blocs(
#         dirname
#     )
#     dibujante = Dessinateur4blocs(
#         dirname + nomPlanValue,
#         essais,
#         essaisMatricesList,
#         droiteIndex,
#         gaucheIndex,
#     )
#     dibujante.openDraw()
#     dibujante.planPrincipal()
#     dibujante.drawBordures()
#     dibujante.matriceCompteur()
#     dibujante.fisher()
#     dibujante.createPavePourLesCalculs()
#     dibujante.etiquettes()
#     dibujante.maquillage()
#     dibujante.closeDraw()
#     dibujante.picasso()


if __name__ == "__main__":

    dirname = "src" + "/"
    nomPlanValue = "Test"

    essais, essaisMatricesList, = backendWrapper2Blocs(dirname)
    dibujante = Dessinateur2blocs(
        dirname + nomPlanValue,
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