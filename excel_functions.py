import math
import re


def safely_try_to_cast_to_flow(string):
    if not isinstance(string, str):
        return string
    try:
        return float(string)
    except ValueError:
        return string


def is_nan(val):
    if isinstance(val, float):
        return math.isnan(val)


def new_mark_differences(nowy_excel, differences_list):
    empty_column_letter = nowy_excel.get_first_empty_column_letter()

    for diff_row in differences_list:
        row_number = nowy_excel.get_data_row(diff_row[0])
        cell_cords = empty_column_letter + str(row_number)

        if diff_row[1] in ['IDENTYCZNE', 'NOWY']:
            nowy_excel.ws.range(cell_cords).value = diff_row[1]
            nowy_excel.ws.range(cell_cords).color = (255, 0, 0)
        else:
            nowy_excel.ws.range(cell_cords).value = diff_row[1].replace('\n', ' ')
            for column_name_with_difference in [e.strip() for e in diff_row[1].split(',')]:
                column_letter_with_difference = nowy_excel.get_column_letter(column_name_with_difference)
                nowy_excel.ws.range(column_letter_with_difference + str(row_number)).color = (148, 0, 211)


def get_next_row(cords):
    first_row_splitted = re.search(r'([A-Z]+)([0-9]+)', cords.split(':')[0])
    return first_row_splitted.group(1) + str(int(first_row_splitted.group(2)) + 1)


def mark_old_rows(nowy_excel, old_rows):
    # znajdz pierwszy pusty rzad
    correct_coords = nowy_excel.get_empty_row_cords()

    for old_row in old_rows:
        nowy_excel.ws.range(correct_coords).value = list(old_row) + ['STARY']
        correct_coords = get_next_row(correct_coords)
