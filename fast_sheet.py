import xlwings as xw
from string import ascii_lowercase as letters
from excel_functions import *


#zalozenie: jest jedna kolumna klucz ktora identyfikuje rzad
#zalozenie: nazwy kolumn sie nie powtarzaja
class ColumnsNotUniqueException(Exception):
    pass


class DataiInRowNotUniqueException(Exception):
    pass


class FastSheet:
    def __convert_rows_to_columns(values, start_row):
        columns = {}
        for ind, e in enumerate(zip(*values[1:])):
            # key = letters[ind]
            columnd_dictionary = {}
            for ind2, k in enumerate(e):
                if k in columnd_dictionary:
                    columnd_dictionary[k].append(start_row + ind2 + 1)
                else:
                    columnd_dictionary[k] = [start_row + ind2 + 1]
            columns[values[0][ind]] = columnd_dictionary
        return columns

    # excel musi skladac sie tylko z jednego sheeta
    def __load_excel(self, data_cords, key_column):
        values = self.ws.range(data_cords).options(expand='table').value
        if len(values[0]) != len(set(values[0])):
            raise ColumnsNotUniqueException("columns are not unique")
        headers = {e: letters[ind] for ind, e in enumerate(values[0])}
        data_as_columns = FastSheet.__convert_rows_to_columns(values, int(re.search(r'([0-9])+', data_cords).group(1)))
        self.headers = headers
        self.data_as_columns = data_as_columns
        for e in values[1:]:
            self.data_as_rows[e[values[0].index(key_column)]] = dict(zip(values[0], e))

    def __init__(self, excel):
        # todo: gdy podczas ladowania z excela nazwy kolumn sie powtarzaja to rzuc wyjatek ColumnNotUniqueException
        '''example:
                {
                    'Nazwa': 'A',
                    'Ceny': 'B'
                }
        '''
        self.headers = {}
        '''example:
        {
            'Nazwa':{"ej":["A1","B2"], "mej":["A2"]},
            'Ceny':{}
         }
         '''
        self.data_as_columns = {}

        self.data_as_rows = {}
        wb = xw.Book(excel['filename'])
        self.ws = wb.sheets[0]
        self.__load_excel(data_cords=excel['data_cords'], key_column=excel['key_column'])
        # true only at initialization
        start_cords = re.search(r'([A-Z]+)([0-9]+)', excel['data_cords'].split(':')[0])
        end_cords = re.search(r'([A-Z]+)([0-9]+)', excel['data_cords'].split(':')[1])
        self.first_empty_row = start_cords.group(1) + str(int(end_cords.group(2)) + 1)

    # nazwa kolumny -> literka kolumny w excelu
    # example: 'Nazwa' -> 'S'
    # example: '' -> 'T' # zwraca pierwszą pustą kolumnę
    def get_column_letter(self, name):
        if name in self.headers:
            return self.headers[name]
        else:
            return None

    def __next_letter(char):
        return chr(ord(char) + 1)

    def get_first_empty_column_letter(self):
        return FastSheet.__next_letter(max([self.headers[e] for e in self.headers]))

    def get_empty_row_cords(self):
        return self.first_empty_row

    # wartość do wyszukania -> wspolrzedne
    # example: 'Kodowanie kalkulatora' -> 'S10'
    def get_data_row(self, value):
        for column_name in self.data_as_columns:
            if value in self.data_as_columns[column_name]:
                if len(self.data_as_columns[column_name][value]) == 1:
                    return self.data_as_columns[column_name][value][0]
                else:
                    raise DataiInRowNotUniqueException("Dane w kolumnie nie są unikalne")

    def __special_compare(val1, val2):
        if isinstance(val1, str) and isinstance(val2, str):
            val1 = val1.upper()
            val2 = val2.upper()
        # puste stringi sa rowne None
        if val1 in ['', None] and val2 in ['', None]:
            return False
        else:
            return not (is_nan(val1) and is_nan(val2)) and val1 != val2

    def __find_row_differences(old_row, new_row):
        differences = []
        # todo: kazdy row musi byc slownikiem: kolumna: wartosc, nie moge polegac na kolejnosci kolumn w rzedzie
        for i, (key, value) in enumerate(new_row.items()):
            if key not in old_row:
                differences.append(key)
                continue
            val1 = safely_try_to_cast_to_flow(old_row[key])
            val2 = safely_try_to_cast_to_flow(new_row[key])
            if FastSheet.__special_compare(val1, val2):
                differences.append(key)
        if len(differences) == 0:
            return ['IDENTYCZNE']
        else:
            return differences

    def __check_if_column_in_older_columns(new_column, old_columns):
        for key, value in old_columns.items():
            if key.endswith(new_column):
                return True
        return False

    def __check_if_column_in_newer_columns(old_column, new_columns):
        for key, value in new_columns.items():
            if old_column.endswith(key):
                return True
        return False


    def __get_column_name_with_prepended_project_name(new_column, old_columns):
        for key, value in old_columns.items():
            if key.endswith(new_column):
                return key

            # zawsze porownuj nowy ze starym a nie na odwrot!!!
    def compare(self, older_fast_sheet, key_column):
        differences_list = []

        # check if this row exist in older _sheet
        for new_column in self.data_as_columns[key_column]:
            if new_column in older_fast_sheet.data_as_columns[key_column]:
                differences_list.append((new_column, ", "
                                         .join(FastSheet.__find_row_differences(
                    old_row=older_fast_sheet.data_as_rows[new_column],
                    new_row=self.data_as_rows[new_column]))))
            else:
                differences_list.append((new_column, 'NOWY'))

        old_rows = [list(next(iter(older_fast_sheet.data_as_rows.values())).keys())]
        # todo: dopisz ten fragment gdy bedziesz testowal dla roznych miesiecy
        for old_column in older_fast_sheet.data_as_columns[key_column]:
            is_similar_column_in_new_rows = FastSheet. \
                __check_if_column_in_newer_columns(old_column, self.data_as_columns[key_column])
            is_similar_column_in_new_rows = FastSheet.\
                __check_if_column_in_newer_columns(old_column, self.data_as_columns[key_column])
            if old_column not in self.data_as_columns[key_column] and not is_similar_column_in_new_rows:
                old_rows.append(list(older_fast_sheet.data_as_rows[old_column].values()))

        return differences_list, old_rows
