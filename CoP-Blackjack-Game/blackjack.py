import json
import random
import logging

logger = logging.getLogger()
logger.setLevel(logging.INFO)

VALUES = ['2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K', 'A']
SUITS = ['Hearts', 'Diamonds', 'Clubs', 'Spades']

def create_deck():
    return [{'value': value, 'suit': suit} for value in VALUES for suit in SUITS]

def deal_card(deck):
    return deck.pop(random.randint(0, len(deck) - 1))

def calculate_hand_value(hand):
    value = 0
    ace_count = 0
    for card in hand:
        if card['value'] in ['J', 'Q', 'K']:
            value += 10
        elif card['value'] == 'A':
            ace_count += 1
            value += 11
        else:
            value += int(card['value'])

    while value > 21 and ace_count:
        value -= 10
        ace_count -= 1

    return value

def handler(event, context):
    logger.info("Event: " + json.dumps(event))
    try:
        action = event['action']
        player_hands = event.get('player_hands', [[]])
        dealer_hand = event.get('dealer_hand', [])
        deck = event.get('deck', create_deck())
        current_hand_index = event.get('current_hand_index', 0)

        if action == 'start':
            player_hands = [[deal_card(deck), deal_card(deck)]]
            dealer_hand = [deal_card(deck), deal_card(deck)]
        elif action == 'hit':
            player_hands[current_hand_index].append(deal_card(deck))
        elif action == 'stay':
            current_hand_index += 1
            if current_hand_index >= len(player_hands):
                while calculate_hand_value(dealer_hand) < 17:
                    dealer_hand.append(deal_card(deck))
                current_hand_index = -1  # Indicates the game is over
        elif action == 'split':
            if len(player_hands[current_hand_index]) == 2 and player_hands[current_hand_index][0]['value'] == player_hands[current_hand_index][1]['value']:
                split_card = player_hands[current_hand_index].pop()
                player_hands.append([split_card, deal_card(deck)])
                player_hands[current_hand_index].append(deal_card(deck))

        player_values = [calculate_hand_value(hand) for hand in player_hands]
        dealer_value = calculate_hand_value(dealer_hand)

        result = ""
        if current_hand_index == -1:  # Game over
            results = []
            for player_value in player_values:
                if player_value > 21:
                    results.append("Player busts! Dealer wins.")
                elif dealer_value > 21:
                    results.append("Dealer busts! Player wins.")
                elif player_value > dealer_value:
                    results.append("Player wins!")
                elif player_value < dealer_value:
                    results.append("Dealer wins!")
                else:
                    results.append("It's a tie!")
            result = " ".join(results)

        response = {
            'statusCode': 200,
            'body': json.dumps({
                'player_hands': player_hands,
                'dealer_hand': dealer_hand,
                'deck': deck,
                'player_values': player_values,
                'dealer_value': dealer_value,
                'result': result,
                'current_hand_index': current_hand_index
            }),
            'headers': {
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type, X-Amz-Date, Authorization, X-Api-Key, X-Amz-Security-Token'
            }
        }

        logger.info("Response: " + json.dumps(response))
        return response
    except Exception as e:
        logger.error("Error: " + str(e))
        return {
            'statusCode': 500,
            'body': json.dumps({'error': str(e)}),
            'headers': {
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type, X-Amz-Date, Authorization, X-Api-Key, X-Amz-Security-Token'
            }
        }
