import xlrd
from random import randint
import csv

class game:
    def __init__(self, name, column_sizes, column_pos):
        self.name = name
        self.column_sizes = column_sizes
        self.column_pos = column_pos

# Number of possibilities for each category in the format:
# 1. Pokemon    (Catch)
# 2. Items      (Find)
# 3. Moves      (Learn)
# 4. Locations  (Visit)
# 5. Trainers   (Defeat)

game_names= ["platinum", "white-2", "showdown"]
game_list = [
    game("platinum", (210, 445, 467, 40, 18), (2, 5, 8, 11, 14)), # Platinum
    game("white-2", (297, 548, 559, 32, 23), (2, 5, 8, 11, 14)), # Black or white 2
    game("showdown", (583, 850, 42, 47, 583), (2, 5, 8, 11, 14)), # Showdown
]

rows = 5

# Determining which game is being played
game_input = input(
"""The supported games are listed below:
    1. Platinum
    2. Black or White 2
    3. Showdown
Enter the number of the game you are playing: """)
game_id = int(game_input) - 1
game_object = game_list[game_id]

# Reading from xls file
file = "bingo_data.xls"
sheet = game_list[game_id].name
book = xlrd.open_workbook(file)
sheet = book.sheet_by_name(sheet)

# Creates row tuples inside of a main tuple, and makes sure all values will be unique
unique = False
while unique == False:
    unique = True
    randoms = tuple(tuple(randint(0, size - 1) for size in game_object.column_sizes) for row in range(rows))
    for col in range(rows):
        for i in range(1, rows):
            for j in range(i):
                #print("({0}, {1}, {2}): {3}, {4})".format(col, i, j, randoms[i][col], randoms[j][col]))
                if randoms[i][col] == randoms[j][col]:
                    unique = False

# Finds the data pertaining to the random values
values = tuple(tuple(sheet.cell(randoms[row][i], game_object.column_pos[i]).value for i in range(rows)) for row in range(rows))

# Shifts values so that categories are skew
shifted = tuple(values[row][row:] + values[row][:row] for row in range(rows))

# Formats data to look nice
formatted = tuple(tuple(value for value in row) for row in shifted)
for row in formatted:
    print(row)

# Export to csv file format
output_file = open("output.csv", mode="w", newline="")
output = csv.writer(output_file, delimiter=",")       
for row in range(rows):
    output.writerow(formatted[row])

# Cleaning memory by closing files
book.release_resources()
del book
output_file.close()
