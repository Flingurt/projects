import os
import csv
import re
import pdfplumber
import pandas as pd
from tabulate import tabulate

from pathlib import Path
from pygments.formatters import ImageFormatter
import pygments.lexers
from PIL import Image

import time

def main():
    # Clean up files from last session
    files = os.listdir()
    for i in files:
        if ".png" in i:
            os.remove(i)
    print("cleaning folder (last sessions PNG files.)")

    # Go through main program
    get_pdf()

    # Remove junk from folder.
    print("cleaning folder")
    files_updated = os.listdir()
    for i in files_updated:
        if ".txt" in i:
            os.remove(i)
        if ".csv" in i:
            os.remove(i)
        if ".pdf" in i:
            os.remove(i)
            
    # Wait for user to print to printer, exits on enter.
    input("Press Enter to exit: ")
            

# Convert formatted file to PNG
def get_png(name):
    print("creating PNG")

    # Create white background to paste list onto (for portrait printing)
    background = Image.new('RGB', (3117, 5500), color='white')

    lexer = pygments.lexers.TextLexer()
    png = pygments.highlight(Path(f"{name}").read_text(), lexer, ImageFormatter(line_numbers = False, font_size = 50, font_name = "Consolas"))
    Path(f"{name.replace(".txt", "")}.png").write_bytes(png)

    # Paste list onto background
    page = Image.open(f"{name.replace(".txt", "")}.png")
    background.paste(page, (-180, 600))
    background.save(f"{name.replace(".txt", "")}.png", quality=95)

    # Print prompt for each iteration
    # 0.5s delay for print prompts to keep up
    print(f"printing {name.replace(".txt", "")}.png")
    os.startfile(f"{name.replace(".txt", "")}.png", "print")
    time.sleep(0.5)


# Create formatted .txt.
def get_format(table, file_name, page):
    print("making pretty lines")
    new_name = f"{file_name.replace(".csv", f"_{page}")}.txt"
    with open(new_name, 'w') as f:
        f.write(tabulate(table, headers=["_", "Product\nCatalog\nNumber", "  Customer\n  Catalog No", "   Description   ", "PoU\nLoc", "Master\nLoc", "St. Area", "Cost\nCenter", "Stock", "Pick"], tablefmt="grid", 
                         maxcolwidths=[8, 8, 15, 15, 6, 9, 10, 3, 3]))
    get_png(new_name)

# Fix invalid characters, Get tables for tabulate, new page if 25+ rows
def get_tables(current_house, file_name):
    table_1 = []
    table_2 = []
    table_3 = []

    # Fix åäö, uL.
    for n in range(len(current_house)):
        for n_2 in range(len(current_house)):
            try:
                # Replace "o"
                if "+APY-" in current_house[n][n_2]:
                    current_house[n][n_2] = current_house[n][n_2].replace("+APY-", "o")
                # Replace "O"
                if "+ANY-" in current_house[n][n_2]:
                    current_house[n][n_2] = current_house[n][n_2].replace("+ANY-", "O")
                # Replace "ä"
                if "+AOQ-" in current_house[n][n_2]:
                    current_house[n][n_2] = current_house[n][n_2].replace("+AOQ-", "a")
                # Replace "uL"
                if "+ALU-" in current_house[n][n_2]:
                    current_house[n][n_2] = current_house[n][n_2].replace("+ALU-", "u")
                   
            except IndexError:
                pass

    # Replace "Splitted" under available quantity
    # Remove newline character from staging area
    for n in range(len(current_house)):
        try:
            if "Splitted" in current_house[n][8]:
                current_house[n][8] = current_house[n][8].replace("Splitted", "")
            if "\n" in current_house[n][6]:
                current_house[n][6] = current_house[n][6].replace("\n", "")
        except IndexError:
            pass


    # Get tables for tabulate
    # Set number of lines before switching page.
    for row in current_house:
        line = row
        if int(len(table_2)) > 22:
            table_3.append(line)
        elif int(len(table_1)) > 22:
            table_2.append(line)
        else:
            table_1.append(line)

        
        
    get_format(table_1, file_name, "page_1")

    if int(len(table_2)) > 0:
        get_format(table_2, file_name, "page_2")
    if int(len(table_3)) > 0:
        get_format(table_3, file_name, "page_3")

    #print(current_house)


# Creates sorted csv for.
def update_csv(file_name, current_list):
    print("writing to file")
    with open(file_name) as f:
        r = csv.reader(f)
        lines = list(r)
    
    # Match locations to rest of list
    rows = []
    fields = lines[0]
    for i in current_list:
        for j in lines:
            for k in j:
                if i in k.replace("\n", ""):
                    if j not in rows:
                        rows.append(j)

    # Write updated list to csv.
    with open(f"sorted_{file_name}", "w") as file:
        writer = csv.writer(file)
        writer.writerow(fields)
        writer.writerows(rows)

    get_tables(rows, file_name)
    


# Filter out ma1
def sort_ma(list):
    ma1_locations = []
    for j in list:
        if filtered:= re.search(r"^([ABCDEFGHI])-(\d?\d)-(\d?\d)-(\d?\d)$", j):
            ma1_locations.append(filtered.group(0))
    
    n = len(ma1_locations)
    counter = 0
    for k in range(n):
        for l in range(0, n-k-1):
            letter, shelf, plane, pos = ma1_locations[l].split("-")
            letter_2, shelf_2, plane_2, pos_2 = ma1_locations[l + 1].split("-")
            if int(shelf) > int(shelf_2):
                temp = ma1_locations[l]
                ma1_locations[l] = ma1_locations[l+1]
                ma1_locations[l+1] = temp

    for k in range(n):
        for l in range(0, n-k-1):
            letter, shelf, plane, pos = ma1_locations[l].split("-")
            letter_2, shelf_2, plane_2, pos_2 = ma1_locations[l + 1].split("-")
            if letter > letter_2:
                temp = ma1_locations[l]
                ma1_locations[l] = ma1_locations[l+1]
                ma1_locations[l+1] = temp

    return ma1_locations


# Filter out ac057
def sort_ac057(list):
    ac057_locations = []
    for j in list:
        if filtered:= re.search(r"(AC057)-[ABCDEFGHJ]-(\d?\d)-(\d?\d)-(\d?\d)", j):
            ac057_locations.append(filtered.group(0))
    
    n = len(ac057_locations)
    counter = 0
    for k in range(n):
        for l in range(0, n-k-1):
            _, letter, shelf, plane, pos = ac057_locations[l].split("-")
            _, letter_2, shelf_2, plane_2, pos_2 = ac057_locations[l + 1].split("-")
            if int(shelf) > int(shelf_2):
                temp = ac057_locations[l]
                ac057_locations[l] = ac057_locations[l+1]
                ac057_locations[l+1] = temp

    for k in range(n):
        for l in range(0, n-k-1):
            _, letter, shelf, plane, pos = ac057_locations[l].split("-")
            _, letter_2, shelf_2, plane_2, pos_2 = ac057_locations[l + 1].split("-")
            if letter > letter_2:
                temp = ac057_locations[l]
                ac057_locations[l] = ac057_locations[l+1]
                ac057_locations[l+1] = temp

    return ac057_locations


# Filter out ac060
def sort_ac060(list):
    ac060_locations = []
    for j in list:
        if filtered:= re.search(r"(AC060)-[ABCDEFGHJ]-(\d?\d)-(\d?\d)", j):
            ac060_locations.append(filtered.group(0))
    
    n = len(ac060_locations)
    counter = 0
    for k in range(n):
        for l in range(0, n-k-1):
            _, letter, shelf, plane, pos = ac060_locations[l].split("-")
            _, letter_2, shelf_2, plane_2, pos_2 = ac060_locations[l + 1].split("-")
            if int(shelf) > int(shelf_2):
                temp = ac060_locations[l]
                ac060_locations[l] = ac060_locations[l+1]
                ac060_locations[l+1] = temp

    for k in range(n):
        for l in range(0, n-k-1):
            _, letter, shelf, plane, pos = ac060_locations[l].split("-")
            _, letter_2, shelf_2, plane_2, pos_2 = ac060_locations[l + 1].split("-")
            if letter > letter_2:
                temp = ac060_locations[l]
                ac060_locations[l] = ac060_locations[l+1]
                ac060_locations[l+1] = temp

    return ac060_locations


# Function: Sort
def get_sorted(file_name):
    print("sorting...")
    alpha_sorted_list = []
    # Get current CSV.
    current_list = []
    with open(file_name) as file:
        reader = csv.reader(file)
        lines = list(reader)
        for line in lines:
            if line not in current_list:
                current_list.append(line)

    # Get rid of newline character, append master location to unsorted list.
    unsorted_list = []
    for i in current_list:
        unsorted_list.append(i[5].replace(" ", ""))

    # Sort locations for help with formatting.
    # Keeps second and last number in location in order without the need to format manually.
    alpha_sorted_list = sorted(unsorted_list)

    # Append lists in order.
    all_sorted = []
    for i in sort_ma(alpha_sorted_list):
        all_sorted.append(i)
    print("MA1 sorted")
    for i in alpha_sorted_list:
        #================================================================== TODO: tar inte 'AC060-AC4 Ã¶st'
        if i not in sort_ma(alpha_sorted_list) and i not in sort_ac057(alpha_sorted_list) and i not in sort_ac060(alpha_sorted_list):
            if "MasterLocation" not in i:
                i = i.replace("Ã¶st", "Väst")
                all_sorted.append(i)
    print("ETC. sorted")
    for i in sort_ac057(alpha_sorted_list):
        all_sorted.append(i)
    print("AC057 sorted")
    for i in sort_ac060(alpha_sorted_list):
        all_sorted.append(i)
    print("AC060 sorted")

    update_csv(file_name, all_sorted)
    #print(all_sorted)


# Function: Get PDF
def get_pdf():
    print("converting PDF to CSV")
    # Browse folder for PDF.
    # Create list containing names of PDF files.
    files = os.listdir()
    # Repeat file creation for every file in files.
    counter = 1
    for i in files:
        if ".pdf" in i:
            # Get PDF data.
            with pdfplumber.open(i) as pdf:
                page0 = pdf.pages[0]
                table = page0.extract_table()
                table[:3]
                df = pd.DataFrame(columns = table[0])
                each_page_data = pd.DataFrame(table[1:], columns=df.columns)
                each_page_data
            try:
                for j in range(2):
                    page = pdf.pages[j]
                    table = page.extract_table()
                    each_page_data = pd.DataFrame(table[1:], columns=df.columns)
                    df = pd.concat([df, each_page_data], ignore_index=True)

                    # Remove department column and cost center column
                    newdf = df.drop("Department", axis="columns")
                    newdf = newdf.drop("PoU", axis="columns")
                    newdf = newdf.drop("Vendor", axis="columns")

                    # Create CSV file.
                    newdf.to_csv(f"{i.replace(".pdf", "")}.csv", encoding="utf_7",compression=None)
            except (IndexError):
                pass
            print(f"============== List: {counter} ==============")
            counter = counter + 1
            get_sorted(f"{i.replace(".pdf", "")}.csv")

main()