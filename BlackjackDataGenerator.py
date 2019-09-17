from random import shuffle
from random import randint
import xlsxwriter

"""CREDITS TO: https://brilliant.org/wiki/programming-blackjack/#"""

DEALER = 1
PLAYER = 2
TIED = 3

DATATOCOLLECT = 1000000
# define the card ranks, and suits
ranks = [_ for _ in range(2, 11)] + ['JACK', 'QUEEN', 'KING', 'ACE']
suits = ['SPADE', 'HEART ', 'DIAMOND', 'CLUB']
maxPlayerCards = 8
maxDealerCards = 8
winnerRow = 0
playerColumnPointer = 1
dealerColumnPointer = maxPlayerCards + 1
currentLine = 1
playing = 1

workbook = xlsxwriter.Workbook('BlackJackData.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, winnerRow, "Winner")
for card in range(maxPlayerCards):
    worksheet.write(0, 1 + card, "Player" + str(card))

for card in range(maxDealerCards):
    worksheet.write(0, 1 + maxPlayerCards + card, "Dealer" + str(card))


# Functions
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

    # This loops only ends when a winner is found
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
                worksheet.write(currentLine, playerColumnPointer, str(player_hand[len(player_hand)-1]))
                playerColumnPointer += 1
                continue

            choice = randint(0, 2)
            # Choice one is draw card
            if choice == 1:
                player_hand.append(deck.pop())
                worksheet.write(currentLine, playerColumnPointer, str(player_hand[len(player_hand)-1]))
                playerColumnPointer += 1
                continue

            # Choice two is pass
            if choice == 0:
                # Dealer draws to at least 17 then compares
                while hand_value(dealer_hand) < 17:
                    dealer_hand.append(deck.pop())
                    worksheet.write(currentLine, dealerColumnPointer, str(dealer_hand[len(dealer_hand)-1]))
                    dealerColumnPointer += 1

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

workbook.close()
