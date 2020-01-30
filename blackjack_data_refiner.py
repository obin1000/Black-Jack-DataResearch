import xlsxwriter
import xlrd
from tqdm import tqdm
import blackjack_data_generator as bdg

""" DO EXPERIMENT THINGS HERE"""
"This cannot add the charts to the existing xlsx file, so it creates it own"

DATA_SOURCE = bdg.WORKBOOK  # Name of the file to read from
WORKBOOK = 'BlackJackDataCharts.xlsx'  # Name of the file to write the graphs to
DIFFICULT_HAND = [12, 13, 14, 15, 16, 17]
WIN = 1
LOSE = 0
DRAW = 1
PASS = 0


class Counter:

    def __init__(self, value_hand):
        self.identifier = value_hand
        self.draw_total = 0
        self.draw_win = 0
        self.draw_lose = 0

        self.pass_total = 0
        self.pass_win = 0
        self.pass_lose = 0

    # WinOrLose: 1 = Win 0 = Lose
    # DrawOrPass: 1 = Draw 0 = Pass
    def write_result(self, win_or_lose, draw_or_pass):
        if draw_or_pass:  # Draw
            self.draw_total += 1
            if win_or_lose:  # Win
                self.draw_win += 1
            else:  # Lose
                self.draw_lose += 1
        else:  # Pass
            self.pass_total += 1
            if win_or_lose:  # Win
                self.pass_win += 1
            else:  # Lose
                self.pass_lose += 1

    def write_result_to_sheet(self):
        write_result(str(self.identifier), self.draw_total, self.draw_win, self.draw_lose,
                     self.draw_win / self.draw_total,
                     self.pass_total, self.pass_win,
                     self.pass_lose, self.pass_win / self.pass_total,
                     (self.draw_win / self.draw_total) - (self.pass_win / self.pass_total))


counter12 = Counter(12)
counter13 = Counter(13)
counter14 = Counter(14)
counter15 = Counter(15)
counter16 = Counter(16)
counter17 = Counter(17)


def write_to_counter(value_hand, win_or_lose, draw_or_pass):
    if value_hand == 12:
        counter12.write_result(win_or_lose, draw_or_pass)
    elif value_hand == 13:
        counter13.write_result(win_or_lose, draw_or_pass)
    elif value_hand == 14:
        counter14.write_result(win_or_lose, draw_or_pass)
    elif value_hand == 15:
        counter15.write_result(win_or_lose, draw_or_pass)
    elif value_hand == 16:
        counter16.write_result(win_or_lose, draw_or_pass)
    elif value_hand == 17:
        counter17.write_result(win_or_lose, draw_or_pass)


column_pointer = 0


def write_result(*data):
    global column_pointer
    count = 0
    for dat in data:
        write_sheet.write(count, column_pointer, dat)
        if isinstance(dat, str):
            if len(dat) < 5:
                write_sheet.set_column(column_pointer, column_pointer, 5)
            else:
                write_sheet.set_column(column_pointer, column_pointer, len(dat))
        count += 1
    column_pointer += 1


if __name__ == "__main__":
    # Reading an excel file using Python
    reader = xlrd.open_workbook(DATA_SOURCE)  # Create a reader to read the xlsx file
    read_sheet = reader.sheet_by_index(0)  # Select the sheet with the data
    workbook = xlsxwriter.Workbook(WORKBOOK)
    write_sheet = workbook.add_worksheet()  # Add new sheet for the graphs

    total_rows = 0  # Count the number of games in the sheet
    player_wins = 0  # Count the number of games won by the player
    dealer_wins = 0  # Count the number of games won by the dealer
    tie_games = 0  # Count the number of games tied
    errors = 0  # Count the number of invalid rows

    print('Refining data')
    # Start progress bar
    progress_bar = tqdm(total=read_sheet.nrows)

    # Loop to refine the data
    for row in range(read_sheet.nrows):
        # Add 1 to progress bar
        progress_bar.update()

        # The first row contains column names
        if row == 0:
            continue

        temp_value = str(read_sheet.cell_value(row, bdg.WINNER_COLUMN))
        total_rows += 1
        # Sum how many wins the play and the dealer have
        if temp_value == bdg.DEALER:
            dealer_wins += 1
        elif temp_value == bdg.PLAYER:
            player_wins += 1
        elif temp_value == bdg.TIED:
            tie_games += 1
        else:
            errors += 1

        hand = []
        hand_value = 0
        current_column = 0
        # Loop over the player hand
        for column in range(bdg.MAX_PLAYER_CARDS):
            current_column = 3 + column  # Read the cards of the player from start to end
            temp_value = read_sheet.cell_value(row, current_column)
            hand.append(bdg.get_card_from_string(temp_value))  # Add the card to the hand
            hand_value = bdg.hand_value(hand)  # Calculate the current value of the hand
            # Extract difficult hand
            if hand_value in DIFFICULT_HAND:
                # Check if the player won
                winner = read_sheet.cell_value(row, bdg.WINNER_COLUMN)
                if winner == bdg.PLAYER:
                    # If the player won check if he drew or passed
                    nextCard = bdg.get_card_from_string(read_sheet.cell_value(row, current_column + 1))
                    if nextCard[0] == '':  # if passed
                        write_to_counter(hand_value, WIN, PASS)
                    else:
                        write_to_counter(hand_value, WIN, DRAW)
                elif winner == bdg.DEALER:
                    # If the player lost, check if he drew or passed
                    nextCard = bdg.get_card_from_string(read_sheet.cell_value(row, current_column + 1))
                    if nextCard[0] == '':  # if passed
                        write_to_counter(hand_value, LOSE, PASS)
                    else:
                        write_to_counter(hand_value, LOSE, DRAW)

    # Write refined data to the sheet for the graphs
    write_result(bdg.DEALER, dealer_wins)
    write_result(bdg.PLAYER, player_wins)
    write_result(bdg.TIED, tie_games)
    write_result('Cardvalue', 'DrawTotal', 'DrawWin', 'DrawLose', 'DrawWinChance', 'PassTotal', 'PassWin', 'PassLose',
                 'PassWinChance', 'DiffWinPass')
    counter12.write_result_to_sheet()
    counter13.write_result_to_sheet()
    counter14.write_result_to_sheet()
    counter15.write_result_to_sheet()
    counter16.write_result_to_sheet()
    counter17.write_result_to_sheet()

    chart_total_wins = workbook.add_chart({'type': 'column'})  # Create chart
    chart_total_wins.add_series({'values': ['Sheet1', 1, 0, 1, 2],  # Add configuration to chart
                                 'name': 'Wins', 'categories': ['Sheet1', 0, 0, 0, 2]})

    chart_total_wins.set_title({
        'name': 'Total wins'
    })
    chart_total_wins.set_x_axis({
        'name': 'Game outcome'
    })
    chart_total_wins.set_y_axis({
        'name': 'Number of wins'
    })

    chart_difficult_hand = workbook.add_chart({'type': 'column'})
    chart_difficult_hand.add_series({'name': 'Draw card',
                                     'categories': ['Sheet1', 0, 4, 0, 9],
                                     'values': ['Sheet1', 4, 4, 4, 9]})

    chart_difficult_hand.add_series({'name': 'Pass',
                                     'categories': ['Sheet1', 0, 4, 0, 9],
                                     'values': ['Sheet1', 8, 4, 8, 9]})

    chart_difficult_hand.set_size({'x_scale': 1.5, 'y_scale': 2})

    chart_difficult_hand.set_title({
        'name': 'Probability to win by hand value'
    })
    chart_difficult_hand.set_x_axis({
        'name': 'Hand value'
    })
    chart_difficult_hand.set_y_axis({
        'name': 'Probability to win'
    })

    # Insert the charts into the worksheet.
    write_sheet.insert_chart(12, 2, chart_total_wins)
    write_sheet.insert_chart(12, 12, chart_difficult_hand)

    progress_bar.close()
    print('Done!')

    workbook.close()
