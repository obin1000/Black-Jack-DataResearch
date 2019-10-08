import xlsxwriter
import xlrd
import BlackjackDataGenerator as bdg

""" DO EXPERIMENT THINGS HERE"""
"This cannot add the charts to the existing xlsx file, so it creates it own"

DATASOURCE = bdg.WORKBOOK  # Name of the file to read from
WORKBOOK = 'charts.xlsx'  # Name of the file to write the graphs to
DIFFICULTHAND = [11, 12, 13, 14, 15, 16, 17]
WIN = 1
LOSE = 0
DRAW = 1
PASS = 0


class Counter:

    def __init__(self, handvalue):
        self.identifier = handvalue
        self.DrawTotal = 0
        self.DrawWin = 0
        self.DrawLose = 0

        self.PassTotal = 0
        self.PassWin = 0
        self.PassLose = 0

    # WinOrLose: 1 = Win 0 = Lose
    # DrawOrPass: 1 = Draw 0 = Pass
    def write_result(self, WinOrLose, DrawOrPass):
        if DrawOrPass:  # Draw
            self.DrawTotal += 1
            if WinOrLose:  # Win
                self.DrawWin += 1
            else:  # Lose
                self.DrawLose += 1
        else:  # Pass
            self.PassTotal += 1
            if WinOrLose:  # Win
                self.PassWin += 1
            else:  # Lose
                self.PassLose += 1

    def write_result_to_sheet(self):
        write_result(str(self.identifier), self.DrawTotal, self.DrawWin, self.DrawLose, self.DrawWin / self.DrawTotal,
                     self.PassTotal, self.PassWin,
                     self.PassLose, self.PassWin / self.PassTotal,
                     (self.DrawWin / self.DrawTotal) - (self.PassWin / self.PassTotal))


def write_to_counter(handvalue, winorlose, draworpass):
    if handvalue == 11:
        counter11.write_result(winorlose, draworpass)
    elif handvalue == 12:
        counter12.write_result(winorlose, draworpass)
    elif handvalue == 13:
        counter13.write_result(winorlose, draworpass)
    elif handvalue == 14:
        counter14.write_result(winorlose, draworpass)
    elif handvalue == 15:
        counter15.write_result(winorlose, draworpass)
    elif handvalue == 16:
        counter16.write_result(winorlose, draworpass)
    elif handvalue == 17:
        counter17.write_result(winorlose, draworpass)


columnpointer = 0


def write_result(*data):
    global columnpointer
    count = 0
    for dat in data:
        writesheet.write(count, columnpointer, dat)
        if isinstance(dat, str):
            if len(dat) < 5:
                writesheet.set_column(columnpointer, columnpointer, 5)
            else:
                writesheet.set_column(columnpointer, columnpointer, len(dat))
        count += 1
    columnpointer += 1


if __name__ == "__main__":
    # Reading an excel file using Python
    reader = xlrd.open_workbook(DATASOURCE)  # Create a reader to read the xlsx file
    readsheet = reader.sheet_by_index(0)  # Select the sheet with the data
    workbook = xlsxwriter.Workbook(WORKBOOK)
    writesheet = workbook.add_worksheet()  # Add new sheet for the graphs

    TotalRows = 0  # Count the number of games in the sheet
    PlayerWins = 0  # Count the number of games won by the player
    DealerWins = 0  # Count the number of games won by the dealer
    TieGames = 0  # Count the number of games tied
    Errors = 0  # Count the number of invalid rows

    counter11 = Counter(11)
    counter12 = Counter(12)
    counter13 = Counter(13)
    counter14 = Counter(14)
    counter15 = Counter(15)
    counter16 = Counter(16)
    counter17 = Counter(17)

    # Loop to refine the data
    for row in range(readsheet.nrows):
        # The first row contains column names
        if row == 0:
            continue

        tempValue = str(readsheet.cell_value(row, bdg.winnerColumn))
        TotalRows += 1
        # Sum how many wins the play and the dealer have
        if tempValue == bdg.DEALER:
            DealerWins += 1
        elif tempValue == bdg.PLAYER:
            PlayerWins += 1
        elif tempValue == bdg.TIED:
            TieGames += 1
        else:
            Errors += 1

        Hand = []
        HandValue = 0
        currentColumn = 0
        # Loop over the player hand
        for column in range(bdg.maxPlayerCards):
            currentColumn = 3 + column  # Read the cards of the player from start to end
            tempValue = readsheet.cell_value(row, currentColumn)
            Hand.append(bdg.getCardFromString(tempValue))  # Add the card to the hand
            HandValue = bdg.hand_value(Hand)  # Calculate the current value of the hand
            # Extract difficult hand
            if HandValue in DIFFICULTHAND:
                # Check if the player won
                winner = readsheet.cell_value(row, bdg.winnerColumn)
                if winner == bdg.PLAYER:
                    # If the player won check if he drew or passed
                    nextCard = bdg.getCardFromString(readsheet.cell_value(row, currentColumn + 1))
                    if nextCard[0] == '':  # if passed
                        write_to_counter(HandValue, WIN, PASS)
                    else:
                        write_to_counter(HandValue, WIN, DRAW)
                elif winner == bdg.DEALER:
                    # If the player lost, check if he drew or passed
                    nextCard = bdg.getCardFromString(readsheet.cell_value(row, currentColumn + 1))
                    if nextCard[0] == '':  # if passed
                        write_to_counter(HandValue, LOSE, PASS)
                    else:
                        write_to_counter(HandValue, LOSE, DRAW)

    # Write refined data to the sheet for the graphs
    write_result(bdg.DEALER, DealerWins)
    write_result(bdg.PLAYER, PlayerWins)
    write_result(bdg.TIED, TieGames)
    write_result('Cardvalue', 'DrawTotal', 'DrawWin', 'DrawLose', 'DrawWinChance', 'PassTotal', 'PassWin', 'PassLose',
                 'PassWinChance', 'DiffWinPass')
    counter11.write_result_to_sheet()
    counter12.write_result_to_sheet()
    counter13.write_result_to_sheet()
    counter14.write_result_to_sheet()
    counter15.write_result_to_sheet()
    counter16.write_result_to_sheet()
    counter17.write_result_to_sheet()

    chartTotalWins = workbook.add_chart({'type': 'column'})  # Create chart
    chartTotalWins.add_series({'values': ['Sheet1', 1, 0, 1, 2],  # Add configuration to chart
                               'name': 'Wins',
                               'categories': ['Sheet1', 0, 0, 0, 2]})

    chartTotalWins.set_title({
        'name': 'Total wins'
    })
    chartTotalWins.set_x_axis({
        'name': 'Game outcome'
    })
    chartTotalWins.set_y_axis({
        'name': 'Number of wins'
    })

    chartDifficultHand = workbook.add_chart({'type': 'column'})
    chartDifficultHand.add_series({'name': 'Draw card',
                                   'categories': ['Sheet1', 0, 4, 0, 9],
                                   'values': ['Sheet1', 4, 4, 4, 9]})

    chartDifficultHand.add_series({'name': 'Pass',
                                   'categories': ['Sheet1', 0, 4, 0, 9],
                                   'values': ['Sheet1', 8, 4, 8, 9]})

    chartDifficultHand.set_size({'x_scale': 1.5, 'y_scale': 2})

    chartDifficultHand.set_title({
        'name': 'Win for each card value'
    })
    chartDifficultHand.set_x_axis({
        'name': 'Card value'
    })
    chartDifficultHand.set_y_axis({
        'name': 'Win chance'
    })

    # Insert the charts into the worksheet.
    writesheet.insert_chart(12, 2, chartTotalWins)
    writesheet.insert_chart(12, 12, chartDifficultHand)

    workbook.close()
