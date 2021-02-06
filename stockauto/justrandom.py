# This is for sending an alarm/message to slacker bot

from slacker import Slacker

slack = Slacker('xoxb-1709162090453-1712267545922-zBUfyirsvPaOjIXJVqdWle4R')

# Send a message to #general channel
slack.chat.post_message('#stocks', 'Current price:'+ str(offer))