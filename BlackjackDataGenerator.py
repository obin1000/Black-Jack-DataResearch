from random import shuffle
from random import randint
from time import sleep

import xlsxwriter
import xlrd

"""CREDITS TO: https://brilliant.org/wiki/programming-blackjack/#"""
# Enum for readability
DEALER = "Dealer"
PLAYER = "Player"
TIED = "Tied"

WORKBOOK = 'BlackJackData.xlsx'  # Name of the file to read and write from
DATATOCOLLECT = 10000  # The number of games to play
ranks = [_ for _ in range(2, 11)] + ['JACK', 'QUEEN', 'KING', 'ACE']  # Values of the cards
suits = ['SPADE', 'HEART ', 'DIAMOND', 'CLUB']  # Suits of the cards
maxPlayerCards = 8  # Max cards for the Player, used for columns
maxDealerCards = 8  # Max cards for the Dealer, used for columns
winnerColumn = 0  # The column in the sheet used to save the winner
playerValueColumn = 1  # The column in the sheet used to save the value of the hand of the player
dealerValueColumn = 2  # The column in the sheet used to save the value of the hand of the dealer
playerColumnPointer = 3  # Current column to write the player card to
dealerColumnPointer = maxPlayerCards + 3  # Current column to write the dealer card to
currentLine = 1  # Current row to write to
playing = 1  # Status of the player

workbook = xlsxwriter.Workbook(WORKBOOK)
worksheet = workbook.add_worksheet()

# Add the first row column names tot the sheet
worksheet.write(0, winnerColumn, "Winner")
worksheet.write(0, playerValueColumn, "Player value")
worksheet.write(0, dealerValueColumn, "Dealer value")
for card in range(maxPlayerCards):
    worksheet.write(0, 3 + card, "Player" + str(card))

for card in range(maxDealerCards):
    worksheet.write(0, 3 + maxPlayerCards + card, "Dealer" + str(card))


# Get a deck of cards as an array
def get_deck():
    return [[rank, suit] for rank in ranks for suit in suits]


# Get the integer value of a card
def card_value(card):
    rank = card[0]
    if rank in ranks[0:-4]:
        return int(rank)
    elif rank is 'ACE':
        return 11
    else:
        return 10


# Get the value of a hand
def hand_value(hand):
    # Naively sum up the cards in the deck.
    tmp_value = sum(card_value(_) for _ in hand)
    # Count the number of Aces in the hand.
    num_aces = len([_ for _ in hand if _[0] is 'ACE'])

    # Aces can count for 1, or 11. If it is possible to bring the value of
    # The hand under 21 by making 11 -> 1 substitutions, do so.
    while num_aces > 0:
        if tmp_value > 21 and 'ACE' in ranks:
            tmp_value -= 10
            num_aces -= 1
        else:
            break

    return tmp_value


# Run the test
for test in range(DATATOCOLLECT):
    # get a deck of cards and shuffle it
    deck = get_deck()
    shuffle(deck)

    # Draw the starting hand
    player_hand = [deck.pop(), deck.pop()]
    dealer_hand = [deck.pop()]

    # Save starting hand
    worksheet.write(currentLine, playerColumnPointer, str(player_hand[0]))
    playerColumnPointer += 1
    worksheet.write(currentLine, playerColumnPointer, str(player_hand[1]))
    playerColumnPointer += 1
    worksheet.write(currentLine, dealerColumnPointer, str(dealer_hand[0]))
    dealerColumnPointer += 1

    # This loops only ends when a winner or tie is found
    while playing:
        valuePlayer = hand_value(player_hand)
        if valuePlayer == 21:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, card_value(dealer_hand[0]))  # Dealer has not played yet
            worksheet.write(currentLine, winnerColumn, PLAYER)
            break
        elif valuePlayer >= 21:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, card_value(dealer_hand[0]))  # Dealer has not played yet
            worksheet.write(currentLine, winnerColumn, DEALER)
            break
        else:
            # There are 2 options: draw or pass, player will draw till at least a value of 12 and then draw at random
            if valuePlayer <= 11:
                player_hand.append(deck.pop())
                worksheet.write(currentLine, playerColumnPointer, str(player_hand[len(player_hand) - 1]))
                playerColumnPointer += 1
                continue

            choice = randint(0, 2)
            # Choice one is draw card
            if choice == 1:
                player_hand.append(deck.pop())
                worksheet.write(currentLine, playerColumnPointer, str(player_hand[len(player_hand) - 1]))
                playerColumnPointer += 1
                continue

            # Choice two is pass
            if choice == 0:
                # Dealer draws to at least a value of 17
                while hand_value(dealer_hand) < 17:
                    dealer_hand.append(deck.pop())
                    worksheet.write(currentLine, dealerColumnPointer, str(dealer_hand[len(dealer_hand) - 1]))
                    dealerColumnPointer += 1

                # Compare the value of both hands and determine the result
                valueDealer = hand_value(dealer_hand)
                if valueDealer >= 21:
                    worksheet.write(currentLine, playerValueColumn, valuePlayer)
                    worksheet.write(currentLine, dealerValueColumn, valueDealer)
                    worksheet.write(currentLine, winnerColumn, PLAYER)
                    break

                if valueDealer < valuePlayer:
                    worksheet.write(currentLine, playerValueColumn, valuePlayer)
                    worksheet.write(currentLine, dealerValueColumn, valueDealer)
                    worksheet.write(currentLine, winnerColumn, PLAYER)
                    break

                if valueDealer > valuePlayer:
                    worksheet.write(currentLine, playerValueColumn, valuePlayer)
                    worksheet.write(currentLine, dealerValueColumn, valueDealer)
                    worksheet.write(currentLine, winnerColumn, DEALER)
                    break

                if valueDealer == valuePlayer:
                    worksheet.write(currentLine, playerValueColumn, valuePlayer)
                    worksheet.write(currentLine, dealerValueColumn, valueDealer)
                    worksheet.write(currentLine, winnerColumn, TIED)
                    break
                break
    # Reset the variables for the next game
    currentLine += 1
    dealerColumnPointer = maxPlayerCards + 3
    playerColumnPointer = 3

""" DO EXPERIMENT THINGS HERE"""
"This cannot be done in another file, because once the excelsheet(workbook) is closed, " \
"xlsxwriter cannot edit it anymore"

# Reading an excel file using Python
TotalRows = 0  # Count the number of games in the sheet
PlayerWins = 0  # Count the number of games won by the player
DealerWins = 0  # Count the number of games won by the dealer
TieGames = 0  # Count the number of games tied
Errors = 0  # Count the number of invalid rows
reader = xlrd.open_workbook(WORKBOOK)  # Create a reader to read the xlsx file
readsheet = reader.sheet_by_index(0)  # Select the sheet with the data
chartsheet = workbook.add_worksheet()  # Add new sheet for the graphs

# Loop to refine the data
for row in range(readsheet.nrows):
    # The first row contains column names
    if row == 0:
        continue
    for column in range(maxPlayerCards):
        tempValue = str(readsheet.cell_value(row, column + 1))
        print(tempValue)
    tempValue = str(readsheet.cell_value(row, winnerColumn))
    TotalRows += 1
    # Sum how many wins the play and the dealer have
    if tempValue == DEALER:
        DealerWins += 1
    elif tempValue == PLAYER:
        PlayerWins += 1
    elif tempValue == TIED:
        TieGames += 1
    else:
        Errors += 1

print("Total games: " + str(TotalRows) + " Dealer Wins: " + str(DealerWins) + " Player Wins: " + str(
    PlayerWins) + " Tie Games: " + str(TieGames) + " Errors: " + str(Errors))

# Write refined data to the sheet for the graphs
chartsheet.write(0, 0, DEALER)
chartsheet.write(0, 1, PLAYER)
chartsheet.write(0, 2, TIED)

chartsheet.write(1, 0, DealerWins)
chartsheet.write(1, 1, PlayerWins)
chartsheet.write(1, 2, TieGames)

# Create a new column chart  to display the number of wins
chart = workbook.add_chart({'type': 'column'})

# Add the configuration to the chart
chart.add_series({'values': ['Sheet2', 1, 0, 1, 2],
                  'name': 'Wins',
                  'categories': ['Sheet2', 0, 0, 0, 2],
                  'fill': {'color': 'blue'}})

chart.set_title({
    'name': 'Number of wins',
})
chart.set_x_axis({
    'name': 'Possible outcome'
})
chart.set_y_axis({
    'name': 'Number of wins'
})

# Insert the chart into the worksheet.
chartsheet.insert_chart(2, 2, chart)

# Cleanup
workbook.close()
