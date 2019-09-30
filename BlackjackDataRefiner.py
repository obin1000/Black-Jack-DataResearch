import xlsxwriter
import xlrd
import BlackjackDataGenerator as bdg

""" DO EXPERIMENT THINGS HERE"""
"This cannot add the charts to the existing xlsx file, so it creates it own"

DATASOURCE = bdg.WORKBOOK  # Name of the file to read from
WORKBOOK = 'charts.xlsx'  # Name of the file to write the graphs to
DIFFICULTHAND = [14, 15, 16]

if __name__ == "__main__":
    # Reading an excel file using Python
    TotalRows = 0  # Count the number of games in the sheet
    PlayerWins = 0  # Count the number of games won by the player
    DealerWins = 0  # Count the number of games won by the dealer
    TieGames = 0  # Count the number of games tied
    Errors = 0  # Count the number of invalid rows
    reader = xlrd.open_workbook(DATASOURCE)  # Create a reader to read the xlsx file
    readsheet = reader.sheet_by_index(0)  # Select the sheet with the data
    workbook = xlsxwriter.Workbook(WORKBOOK)
    writesheet = workbook.add_worksheet()  # Add new sheet for the graphs

    difficultGames = 0  # Number of difficult cases
    difficultWins = 0  # Player wins difficult
    difficultLoses = 0  # Player wins difficult
    difficultTies = 0  # Player wins difficult
    difficultDraw = 0  # Player wins difficult
    difficultPass = 0  # Player wins difficult
    difficultWinsDraw = 0  # Player wins by drawing in difficult case
    difficultWinsPass = 0
    difficultLosesDraw = 0  # Player loses by passing in difficult case
    difficultLosesPass = 0

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

        starterHand = []
        starterHandValue = 0
        # Loop over the player hand
        for column in range(bdg.maxPlayerCards):
            tempValue = readsheet.cell_value(row, 3 + column)
            # extract the starting hand from the sheet, so the first two cards
            if column <= 1:
                starterHand.append(bdg.getCardFromString(tempValue))
                starterHandValue = bdg.hand_value(starterHand)
                # Extract difficult starterhand
                if column == 1 and starterHandValue in DIFFICULTHAND:
                    difficultGames += 1
                    # Check if the player won
                    winner = readsheet.cell_value(row, bdg.winnerColumn)
                    if winner == bdg.PLAYER:
                        difficultWins += 1
                        # If the player won check if he drew or passed
                        card3 = bdg.getCardFromString(readsheet.cell_value(row, 6))
                        if card3[0] == '':  # if passed
                            difficultPass += 1
                            difficultWinsPass += 1
                        else:
                            difficultDraw += 1
                            difficultWinsDraw += 1
                    elif winner == bdg.DEALER:
                        difficultLoses += 1
                        # If the player lost check if he drew or passed
                        card3 = bdg.getCardFromString(readsheet.cell_value(row, 6))
                        if card3[0] == '':  # if passed
                            difficultPass += 1
                            difficultLosesPass += 1
                        else:
                            difficultDraw += 1
                            difficultLosesDraw += 1
                    else:
                        difficultTies += 1

    # Write refined data to the sheet for the graphs
    columnpointer1 = 0
    writesheet.write(0, columnpointer1, bdg.DEALER)
    writesheet.write(1, columnpointer1, DealerWins)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, bdg.PLAYER)
    writesheet.write(1, columnpointer1, PlayerWins)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, bdg.TIED)
    writesheet.write(1, columnpointer1, TieGames)
    columnpointer1 += 2
    writesheet.write(0, columnpointer1, 'Difficult')
    writesheet.write(1, columnpointer1, difficultGames)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DWin')
    writesheet.write(1, columnpointer1, difficultWins)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DLose')
    writesheet.write(1, columnpointer1, difficultLoses)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DTied')
    writesheet.write(1, columnpointer1, difficultTies)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DDraw')
    writesheet.write(1, columnpointer1, difficultDraw)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DPass')
    writesheet.write(1, columnpointer1, difficultPass)
    columnpointer1 += 2
    writesheet.write(0, columnpointer1, 'DWinPass')
    writesheet.write(1, columnpointer1, difficultWinsPass)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DWinDraw')
    writesheet.write(1, columnpointer1, difficultWinsDraw)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DLosePass')
    writesheet.write(1, columnpointer1, difficultLosesPass)
    columnpointer1 += 1
    writesheet.write(0, columnpointer1, 'DLoseDraw')
    writesheet.write(1, columnpointer1, difficultLosesDraw)

    # Create a new column chart  to display the number of wins
    chart = workbook.add_chart({'type': 'column'})

    # Add the configuration to the chart
    chart.add_series({'values': ['Sheet1', 1, 0, 1, 2],
                      'name': 'Wins',
                      'categories': ['Sheet1', 0, 0, 0, 2],
                      'fill': {'color': 'blue'}})

    chart.set_title({
        'name': 'Number of wins'
    })
    chart.set_x_axis({
        'name': 'Possible outcome'
    })
    chart.set_y_axis({
        'name': 'Number of wins'
    })

    # Insert the chart into the worksheet.
    writesheet.insert_chart(4, 2, chart)

    workbook.close()
