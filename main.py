import ezdxf
import tkinter as tk
from tkinter import filedialog
from collections import Counter
import ctypes
import os
import csv

def choose_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory()
    return folder

def export_csv(lista, filename):
    folder = choose_folder()
    if folder:
        sciezka_pliku = folder + '/' + filename + '_tags.csv'
        with open(sciezka_pliku, 'w', newline='', encoding='utf-8') as plik_csv:
            writer = csv.writer(plik_csv)
            for item in lista:
                writer.writerow([item])

        print(f"Dane zostały wyeksportowane do pliku: {sciezka_pliku}")
        ctypes.windll.user32.MessageBoxW(0, f"Dane zostały wyeksportowane do pliku: {sciezka_pliku}", "Uwaga!", 1)

def sprawdz_powtorzenia(lista_stringow):
    liczba_powtorzen = Counter(lista_stringow)
    powtorzenia = {string: count for string, count in liczba_powtorzen.items() if count > 1}
    lista_unikalna = list(set(lista_stringow))

    warning = ""

    if powtorzenia:
        print("Oznaczenia powtarzają się:")
        for string, count in powtorzenia.items():
            print(f"Oznaczenie: {string}, Liczba powtórzeń: {count}")
            #warning.append(f"oznaczenie: {string}, Liczba powtórzeń: {count}")
            warning = warning + f"Oznaczenie: {string}, Liczba powtórzeń: {count} \n"
        ctypes.windll.user32.MessageBoxW(0, warning + ' \n Usunięto powtorzenia.' , "Uwaga!", 1)
    else:
        print("Nie znaleziono powrarzajacych sie oznaczen.")
        ctypes.windll.user32.MessageBoxW(0, "Nie znaleziono powrarzajacych sie oznaczen.", "Uwaga!", 1)
    return  lista_unikalna

def reduce_str(lista_tekstow):

    dot = '.'  # Znak do sprawdzenia
    #lista_stringow = [s for s in lista_tekstow if dot in s and len(s) <= 13]


    # skracanie stringów dwuwierszowych i pozbawianie stringów bez cyfr
    skrocone_stringi = []
    for string in lista_tekstow:
        # Sprawdzenie czy string ma dwie linie
        if '\n' in string:
            # Podział stringa na linie
            # print(string)
            linie = string.split('\n')
            # Wyłuskanie pierwszej linii
            pierwsza_linia = linie[0]
            # Dodanie skróconego stringa do listy jesli zawiera cyfry
            if any(char.isdigit() for char in string):
                skrocone_stringi.append(pierwsza_linia)
        elif ';' in string:
            # Podział stringa na linie
            # print(string)
            linie = string.split(';')
            #usuniecie ewentualnych spacjii
            linie = [s.replace(" ", "") for s in linie]
            # dodanie lini oddzielnie
            for l in linie:
                if any(char.isdigit() for char in string):
                    skrocone_stringi.append(l)

        elif '=' in string:
            # Podział stringa na linie
            # print(string)
            linie = string.split('=')
            # Wyłuskanie pierwszej linii
            pierwsza_linia = linie[0]
            # Dodanie skróconego stringa do listy jesli zawiera cyfry
            if any(char.isdigit() for char in string):
                skrocone_stringi.append(pierwsza_linia)

        else:
            # Dodanie niezmienionego stringa do listy jesli zawiera cyfry
            if any(char.isdigit() for char in string):
                skrocone_stringi.append(string)

    # Pętle sprawdzające i usuwające elementy bez kropek

    lista_stringow = [s for s in skrocone_stringi if dot in s and len(s) <= 13]

    return lista_stringow

def loadfile():
    # Inicjalizacja modułu tkinter
    root = tk.Tk()
    root.withdraw()
    # Określenie dostępnych rozszerzeń plików
    filetypes = (("Pliki DXF", "*.dxf"), ("Pliki DWG", "*.dwg"), ("Wszystkie pliki", "*.*"))
    # Wyświetlenie okna dialogowego do wyboru pliku
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    # Sprawdzenie, czy plik został wybrany
    if file_path:
        print("Wybrano plik:", file_path)
    else:
        print("Nie wybrano żadnego pliku.")

    return file_path

def extract_text_from_entity(entity):
    if entity.dxftype() == "MTEXT":
        return [entity.plain_text()]
    elif entity.dxftype() == "TEXT":
        return [entity.plain_text()]
    elif entity.dxftype() == "INSERT":
        nested_text = extract_text_from_block(entity)
        return nested_text
    else:
        return None

def extract_text_from_block(block):
    extracted_text = []
    for attrib in block.attribs:
        print()
        if attrib.dxf.text != '':
            extracted_text.append(attrib.dxf.text)
    return extracted_text

def extract_text_from_dwg(file_path):
    dwg = ezdxf.readfile(file_path)
    modelspace = dwg.modelspace()

    extracted_text = []
    for entity in modelspace:
        text = extract_text_from_entity(entity)
        if text:
            extracted_text = extracted_text + text

    return extracted_text


#===================================MAIN======================================

file_path=loadfile()

file_name = os.path.splitext(os.path.basename(file_path))[0]

lista_tekstow = extract_text_from_dwg(file_path)
#for s in lista_tekstow:
#    print(s)

reduced_list = reduce_str(lista_tekstow)

lista_unikalna = sprawdz_powtorzenia(reduced_list)

posortowana_lista = sorted(reduced_list, key=lambda s: next((int(c) for c in s if c.isdigit()), 0))

print(posortowana_lista)

export_csv(posortowana_lista, file_name)









