from random import shuffle
from random import randint
import xlsxwriter
import xlrd

"""CREDITS TO: https://brilliant.org/wiki/programming-blackjack/#"""
# Enum for readability
DEALER = "Dealer"
PLAYER = "Player"
TIED = "Tied"

WORKBOOK = 'BlackJackData.xlsx'
DATATOCOLLECT = 10000  # The number of games to play
ranks = [_ for _ in range(2, 11)] + ['JACK', 'QUEEN', 'KING', 'ACE']  # Values of the cards
suits = ['SPADE', 'HEART ', 'DIAMOND', 'CLUB']  # Suits of the cards
maxPlayerCards = 8  # Max cards for the Player, used for columns
maxDealerCards = 8  # Max cards for the Dealer, used for columns
winnerRow = 0  # The row in the sheet used to save the winner
playerColumnPointer = 1  # Current column to write the player card to
dealerColumnPointer = maxPlayerCards + 1  # Current column to write the dealer card to
currentLine = 1  # Current row to write to
playing = 1  # Status of the player
end = maxPlayerCards + maxDealerCards + 1

workbook = xlsxwriter.Workbook(WORKBOOK)
worksheet = workbook.add_worksheet()

# Add the first row column names tot the sheet
worksheet.write(0, winnerRow, "Winner")
for card in range(maxPlayerCards):
    worksheet.write(0, 1 + card, "Player" + str(card))

for card in range(maxDealerCards):
    worksheet.write(0, 1 + maxPlayerCards + card, "Dealer" + str(card))


def get_deck():
    """Return a new deck of cards."""
    return [[rank, suit] for rank in ranks for suit in suits]


def card_value(card):
    """Returns the integer value of a single card."""
    rank = card[0]
    if rank in ranks[0:-4]:
        return int(rank)
    elif rank is 'ACE':
        return 11
    else:
        return 10


def hand_value(hand):
    """Returns the integer value of a set of cards."""
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

    # Return 99 if busted
    if tmp_value <= 21:
        return tmp_value
    else:
        return 99


# Run the test
for test in range(DATATOCOLLECT):
    # get a deck of cards, and randomly shuffle it
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
            worksheet.write(currentLine, winnerRow, PLAYER)
            break
        elif valuePlayer >= 21:
            worksheet.write(currentLine, winnerRow, DEALER)
            break
        else:
            # There are 2 options draw or pass, player will draw till at least a value of 12 and then draw at random
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
                    worksheet.write(currentLine, winnerRow, PLAYER)
                    break

                if valueDealer < valuePlayer:
                    worksheet.write(currentLine, winnerRow, PLAYER)
                    break

                if valueDealer > valuePlayer:
                    worksheet.write(currentLine, winnerRow, DEALER)
                    break

                if valueDealer == valuePlayer:
                    worksheet.write(currentLine, winnerRow, TIED)
                    break
                break

    currentLine += 1
    dealerColumnPointer = maxPlayerCards + 1
    playerColumnPointer = 1

""" DO EXPERIMENT THINGS HERE"""
"This cannot be done in another file, because once the excelsheet(workbook) is closed, " \
"xlsxwriter cannot edit it anymore"

# Reading an excel file using Python
TotalRows = 0
PlayerWins = 0
DealerWins = 0
TieGames = 0
Errors = 0
reader = xlrd.open_workbook('BlackJackData.xlsx')
readsheet = reader.sheet_by_index(0)
chartsheet = workbook.add_worksheet()

# Loop to refine the data
# The first row contains names
for row in range(readsheet.nrows):
    tempValue = str(readsheet.cell_value(row, 0))
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

chartsheet.write(0, 0, DealerWins)
chartsheet.write(0, 1, PlayerWins)
chartsheet.write(0, 2, TieGames)

# Create a new chart object.
chart = workbook.add_chart({'type': 'column'})

# Add a series to the chart.
chart.add_series({'values': ['Sheet2', 0, 0, 0, 2]})

# Insert the chart into the worksheet.
chartsheet.insert_chart(2, 2, chart)

print("Total games: " + str(TotalRows) + " Dealer Wins: " + str(DealerWins) + " Player Wins: " + str(
    PlayerWins) + " Tie Games: " + str(TieGames) + " Errors: " + str(Errors))

workbook.close()
