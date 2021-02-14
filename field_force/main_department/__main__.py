

import telebot

list = ['boss', 'ff']
passwords = ["123", '456']
bot = telebot.TeleBot("1665557553:AAH6Yg7lHv-5mkyO0LLXr6DuQy3R8UXDl5U")
@bot.message_handler(content_types=['text'])
def log_in(message):
    if message.text in list:
        bot.send_message(message.from_user.id, "Введите password:")
        bot.register_next_step_handler(message, get_password)

    elif message.text == "/help":
        bot.send_message(message.from_user.id, "Напиши привет")
    else:
        bot.send_message(message.from_user.id, "Я тебя не понимаю. Напиши /help.")
def get_password(message):
    if message.text in passwords:
        bot.send_message(message.from_user.id, f"OK, введите ключ для данных, которые хотите увидеть:\nТретичные продажи - '1'")
        bot.register_next_step_handler(message, choose)

def choose(message):
    if message.text == '1':
        bot.send_message(message.from_user.id, "А ты молодец )")

bot.polling(none_stop=True, interval=0)
