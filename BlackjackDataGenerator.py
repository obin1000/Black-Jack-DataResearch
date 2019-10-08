from random import shuffle
from random import randint
import xlsxwriter

"""CREDITS TO: https://brilliant.org/wiki/programming-blackjack/#"""
"Generates a xlsx file with N number of simulated blackjack games"


# Enum for readability
DEALER = "Dealer"
PLAYER = "Player"
TIED = "Tied"

WORKBOOK = 'BlackJackData.xlsx'  # Name of the file to read and write from
DATATOCOLLECT = 1000000  # The number of games to play
ranks = [_ for _ in range(2, 11)] + ['JACK', 'QUEEN', 'KING', 'ACE']  # Values of the cards
suits = ['SPADE', 'HEART ', 'DIAMOND', 'CLUB']  # Suits of the cards
maxPlayerCards = 8  # Max cards for the Player, used for columns
maxDealerCards = 8  # Max cards for the Dealer, used for columns
winnerColumn = 0  # The column in the sheet used to save the winner
playerValueColumn = 1  # The column in the sheet used to save the value of the hand of the player
dealerValueColumn = 2  # The column in the sheet used to save the value of the hand of the dealer


def getCardFromString(string):
    string = str(string).replace(']', '').replace('[', '').replace('\'', '')
    string = string.split(',')
    if str.isdigit(string[0]):
        string[0] = int(string[0])
    return string


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


if __name__ == "__main__":
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

        # Play game till a winner or tie is found
        while playing:
            valuePlayer = hand_value(player_hand)
            valueDealer = hand_value(dealer_hand)
            # Game ends if player or dealer has 21 or dies when higher than 21
            if valuePlayer >= 21:
                break
            elif valueDealer >= 21:
                break
            else:
                # Always daw till at least a value of 12
                if valuePlayer <= 11:
                    player_hand.append(deck.pop())
                    worksheet.write(currentLine, playerColumnPointer, str(player_hand[len(player_hand) - 1]))
                    playerColumnPointer += 1
                    continue
                # Always pass if value is 18 or higher
                if valuePlayer >= 18:
                    break
                # Make random choices to draw or to pass
                choice = randint(0, 2)
                # Choice one is draw card
                if choice == 1:
                    player_hand.append(deck.pop())
                    worksheet.write(currentLine, playerColumnPointer, str(player_hand[len(player_hand) - 1]))
                    playerColumnPointer += 1
                    continue

                # Choice two is pass
                if choice == 0:
                    break

        # Dealer draws to at least a value of 17
        while hand_value(dealer_hand) < 17:
            dealer_hand.append(deck.pop())
            worksheet.write(currentLine, dealerColumnPointer, str(dealer_hand[len(dealer_hand) - 1]))
            dealerColumnPointer += 1

        # Calculate value of both hands and check who won
        valuePlayer = hand_value(player_hand)
        valueDealer = hand_value(dealer_hand)
        if valuePlayer == 21:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, PLAYER)

        elif valueDealer == 21:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, DEALER)

        elif valuePlayer > 21:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, DEALER)

        elif valueDealer > 21:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, PLAYER)

        elif valueDealer < valuePlayer:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, PLAYER)

        elif valueDealer > valuePlayer:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, DEALER)

        elif valueDealer == valuePlayer:
            worksheet.write(currentLine, playerValueColumn, valuePlayer)
            worksheet.write(currentLine, dealerValueColumn, valueDealer)
            worksheet.write(currentLine, winnerColumn, TIED)

        # Reset the variables for the next game
        currentLine += 1
        dealerColumnPointer = maxPlayerCards + 3
        playerColumnPointer = 3

    # Cleanup
    workbook.close()
