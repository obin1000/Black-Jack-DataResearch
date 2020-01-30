from random import shuffle
from random import randint
from tqdm import tqdm
import xlsxwriter
import sys

"""CREDITS TO: https://brilliant.org/wiki/programming-blackjack/#"""
"Generates a xlsx file with N number of simulated blackjack games"

# Enum for readability
DEALER = "Dealer"
PLAYER = "Player"
TIED = "Tied"

WORKBOOK = 'BlackJackData.xlsx'  # Name of the file to read and write from

try:
    DATA_TO_COLLECT = int(sys.argv[1])  # The number of games to play
except IndexError:
    DATA_TO_COLLECT = 10000

RANKS = [_ for _ in range(2, 11)] + ['JACK', 'QUEEN', 'KING', 'ACE']  # Values of the cards
SUITS = ['SPADE', 'HEART ', 'DIAMOND', 'CLUB']  # Suits of the cards
MAX_PLAYER_CARDS = 8  # Max cards for the Player, used for columns
MAX_DEALER_CARDS = 8  # Max cards for the Dealer, used for columns
WINNER_COLUMN = 0  # The column in the sheet used to save the winner
PLAYER_VALUE_COLUMN = 1  # The column in the sheet used to save the value of the hand of the player
DEALER_VALUE_COLUMN = 2  # The column in the sheet used to save the value of the hand of the dealer


def get_card_from_string(string):
    string = str(string).replace(']', '').replace('[', '').replace('\'', '')
    string = string.split(',')
    if str.isdigit(string[0]):
        string[0] = int(string[0])
    return string


# Get a deck of cards as an array
def get_deck():
    return [[rank, suit] for rank in RANKS for suit in SUITS]


# Get the integer value of a card
def card_value(card_in):
    rank = card_in[0]
    if rank in RANKS[0:-4]:
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
        if tmp_value > 21 and 'ACE' in RANKS:
            tmp_value -= 10
            num_aces -= 1
        else:
            break

    return tmp_value


if __name__ == "__main__":
    print('Generating blackjack data')
    player_column_pointer = 3  # Current column to write the player card to
    dealer_column_pointer = MAX_PLAYER_CARDS + 3  # Current column to write the dealer card to
    current_line = 1  # Current row to write to
    playing = True  # Status of the player

    workbook = xlsxwriter.Workbook(WORKBOOK)
    worksheet = workbook.add_worksheet()

    # Add the first row column names tot the sheet
    worksheet.write(0, WINNER_COLUMN, "Winner")
    worksheet.write(0, PLAYER_VALUE_COLUMN, "Player value")
    worksheet.write(0, DEALER_VALUE_COLUMN, "Dealer value")
    for card in range(MAX_PLAYER_CARDS):
        worksheet.write(0, 3 + card, "Player" + str(card))

    for card in range(MAX_DEALER_CARDS):
        worksheet.write(0, 3 + MAX_PLAYER_CARDS + card, "Dealer" + str(card))

    # Start progress bar
    progress_bar = tqdm(total=DATA_TO_COLLECT)

    # Run the test
    for test in range(DATA_TO_COLLECT):
        # Add 1 to progress bar
        progress_bar.update()

        # get a deck of cards and shuffle it
        deck = get_deck()
        shuffle(deck)

        # Draw the starting hand
        player_hand = [deck.pop(), deck.pop()]
        dealer_hand = [deck.pop()]

        # Save starting hand
        worksheet.write(current_line, player_column_pointer, str(player_hand[0]))
        player_column_pointer += 1
        worksheet.write(current_line, player_column_pointer, str(player_hand[1]))
        player_column_pointer += 1
        worksheet.write(current_line, dealer_column_pointer, str(dealer_hand[0]))
        dealer_column_pointer += 1

        # Play game till a winner or tie is found
        while playing:
            value_player = hand_value(player_hand)
            value_dealer = hand_value(dealer_hand)
            # Game ends if player or dealer has 21 or dies when higher than 21
            if value_player >= 21:
                break
            elif value_dealer >= 21:
                break
            else:
                # Always daw till at least a value of 12
                if value_player <= 11:
                    player_hand.append(deck.pop())
                    worksheet.write(current_line, player_column_pointer, str(player_hand[len(player_hand) - 1]))
                    player_column_pointer += 1
                    continue
                # Always pass if value is 18 or higher
                if value_player >= 18:
                    break
                # Make random choices to draw or to pass
                choice = randint(0, 1)
                # Choice one is draw card
                if choice == 1:
                    player_hand.append(deck.pop())
                    worksheet.write(current_line, player_column_pointer, str(player_hand[len(player_hand) - 1]))
                    player_column_pointer += 1
                    continue

                # Choice two is pass
                if choice == 0:
                    break

        # Dealer draws to at least a value of 17
        while hand_value(dealer_hand) < 17:
            dealer_hand.append(deck.pop())
            worksheet.write(current_line, dealer_column_pointer, str(dealer_hand[len(dealer_hand) - 1]))
            dealer_column_pointer += 1

        # Calculate value of both hands and check who won
        value_player = hand_value(player_hand)
        value_dealer = hand_value(dealer_hand)
        if value_player == 21:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, PLAYER)

        elif value_dealer == 21:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, DEALER)

        elif value_player > 21:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, DEALER)

        elif value_dealer > 21:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, PLAYER)

        elif value_dealer < value_player:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, PLAYER)

        elif value_dealer > value_player:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, DEALER)

        elif value_dealer == value_player:
            worksheet.write(current_line, PLAYER_VALUE_COLUMN, value_player)
            worksheet.write(current_line, DEALER_VALUE_COLUMN, value_dealer)
            worksheet.write(current_line, WINNER_COLUMN, TIED)

        # Reset the variables for the next game
        current_line += 1
        dealer_column_pointer = MAX_PLAYER_CARDS + 3
        player_column_pointer = 3

    # Cleanup
    progress_bar.close()
    print('Done!')
    print('Results written to: {}'.format(WORKBOOK))
    workbook.close()
