# Reading an excel file using Python 
import xlrd

TotalRows = 0
PlayerWins = 0
DealerWins = 0
TieGames = 0
Errors = 0

# Give the location of the file
loc = "BlackJackData.xlsx"

# To open Workbook 
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for row in range(sheet.nrows):
    tempValue = sheet.cell_value(row, 0)
    TotalRows += 1
    if tempValue == 1: DealerWins += 1
    elif tempValue == 2: PlayerWins += 1
    elif tempValue == 3: TieGames += 1
    else: Errors += 1

print("Total games: " + str(TotalRows) + " Dealer Wins: " + str(DealerWins) + " Player Wins: " + str(
    PlayerWins) + " Tie Games: " + str(TieGames) + " Errors: " + str(Errors))
