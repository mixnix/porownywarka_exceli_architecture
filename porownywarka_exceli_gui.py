from tkinter import *
from tkinter import filedialog
from porownywarka_exceli import porownaj_dwa_excele
from functools import partial
import logging
import traceback

root = Tk()


nazwy_plikow = ["", ""]


def wybierz_excela(position):
    nazwy_plikow[position] = filedialog.askopenfilename(initialdir="./", title="Wybierz excela")


firstDataCordsLabel = Label(root, text="Współrzędne tabelki z danymi w pierwszym excelu np.: \"A8:R27\"")
firstDataCordsLabel.pack()

firstDataCords = Entry(root)
firstDataCords.pack()
# label
# text input na wstawienie gdzie znajduja sie dane w pierwszym excelu, powinno byc przed przyciskiem

pickFirstExcel = Button(root, text="Wybierz pierwszego/starszego excela", command=partial(wybierz_excela, 0))
pickFirstExcel.pack()

secondDataCords = Entry(root)
secondDataCords.pack()

secondDataCordsLabel = Label(root, text="Współrzędne tabelki z danymi w drugim excelu np.: \"A8:R28\"")
secondDataCordsLabel.pack()
# label
# text input na wsatwieie gdzie znajduja sie dane w drugim excelu, powinno byc przed drugim przyciskiem

pickSecondExcel = Button(root, text="Wybierz drugiego/nowszego excela. W nim będą zaznaczone różnice", command=partial(wybierz_excela, 1))
pickSecondExcel.pack()

keyColumnLabel = Label(root, text="Kolumna służąca do identyfikacji rzędu, wartości w tej kolumnie nie mogą się "
                                  "powtarzać i powinny być takie same w obu excelach (ale nie muszą być) np.: "
                                  "\":Nazwa\"")
keyColumnLabel.pack()

keyColumn = Entry(root)
keyColumn.pack()




# label
# text input - column to be used as key

def compareOnClick():
    older_file = {"filename": nazwy_plikow[0], "data_cords": firstDataCords.get(), "key_column": keyColumn.get()}
    newer_file = {"filename": nazwy_plikow[1], "data_cords": secondDataCords.get(), "key_column": keyColumn.get()}

    #todo: uncomment and fix
    # porownaj_dwa_excele(older_file, newer_file)
    try:
        porownaj_dwa_excele(older_file, newer_file)
    except Exception as inst:
        logging.debug("wpis")
        logging.debug(type(inst))
        logging.debug(inst.args)
        logging.debug(inst)
        logging.debug(traceback.format_exc())
        traceback.print_exc()



compareExcelsButton = Button(root, text="Porównaj excele", command=compareOnClick)
compareExcelsButton.pack()

root.mainloop()
