from fast_sheet import FastSheet
from excel_functions import *


def porownaj_dwa_excele(older_excel, newer_excel):
    older_fast_sheet = FastSheet(older_excel)
    newer_fast_sheet = FastSheet(newer_excel)

    differences_list, old_rows = newer_fast_sheet.compare(older_fast_sheet=older_fast_sheet,
                                                          key_column=newer_excel['key_column'])
    new_mark_differences(newer_fast_sheet, differences_list)

    mark_old_rows(newer_fast_sheet, old_rows)


if __name__ == "__main__":
    porownaj_dwa_excele(older_excel={
        "filename": 'C:\projekty\porownywarka_exceli_architecture\excels_unique_key_same_schema_same_rows\wczesniejszy.xlsx',
        "data_cords": 'A1:C4', "key_column": 'Tytul'},
                        newer_excel={
                            "filename": 'C:\projekty\porownywarka_exceli_architecture\excels_unique_key_same_schema_same_rows\pozniejszy.xlsx',
                            "data_cords": 'A1:C4', "key_column": 'Tytul'})
