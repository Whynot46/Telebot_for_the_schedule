import telebot
from colorama import init
from colorama import Fore

init()
bot = telebot.TeleBot('5620044286:AAEluWJLvr_8FMqLjuuRKFzej2RtvXQyYKA')
print(Fore.WHITE + 'Launching the bot: ' + Fore.GREEN + 'ok\n')

@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    bot.send_message(message.chat.id, 'Janger временно недоступен, ведутся технические работы')

bot.polling()