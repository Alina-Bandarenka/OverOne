import telebot
from telebot import types
from openpyxl import load_workbook


bot = telebot.TeleBot("5383429021:AAG37NJUm0oOGZokV2AVnFN4hAzHDX2AIRc")


name1 = " "
phone = " "
group = " "


@bot.message_handler(commands=["start"])
def start(message):
    keyboard = types.InlineKeyboardMarkup(row_width=1)
    key_yes = types.InlineKeyboardButton("Записаться на пробное занятие", callback_data="question_1")
    keyboard.add(key_yes)
    name = f"Привет, <b>{message.from_user.first_name}!</b>\nМы ждём тебя на бесплатном пробном занятии в студии танца Амелия. Желаете записаться?"
    bot.send_message(message.chat.id, name, parse_mode="html", reply_markup=keyboard)


@bot.callback_query_handler(func=lambda call:True)
def callback(call):
    if call.message:
        if call.data == "question_1":
            bot.send_message(call.message.chat.id, "Выберете направление:\nStretching\nZumba")


@bot.message_handler(content_types=["text"])
def get_group(message):
    if message.text == "Stretching":
        bot.send_message(message.chat.id, "Выберете группу:\nПн-Ср 10:00\nВт-Чт 9:00",)
    elif message.text == "Zumba":
        bot.send_message(message.chat.id, "Выберете группу:\nПн-Ср 19:00\nВт-Чт 18:00",)
        bot.register_next_step_handler(message, reg_group)


def reg_group(message):
    global group
    group = message.text
    bot.send_message(message.from_user.id, "Ведите своё имя")
    bot.register_next_step_handler(message, reg_name)


def reg_name(message):
    global name1
    name1 = message.text
    bot.send_message(message.from_user.id, "Ведите номер телефона с кодом\nНапример: 297756457 ")
    bot.register_next_step_handler(message, reg_phone)


def reg_phone(message):
    global phone
    phone = message.text
    try:
        phone = int(message.text)
    except Exception:
        bot.send_message(message.from_user.id, "Ведите номер телефона с кодом\nНапример: 297756457 ")

    bot.send_message(message.from_user.id, f"{name1},{phone}.\nВы записаны на пробное занятие!\nЖдём вас в {group}"
                                           f" При себе иметь спортивную одежду, обувь и хорошее настроение!")


bot.polling(none_stop=True)


fn = "my_bot.xlsx"
wb = load_workbook(fn)
ws = wb["record"]
row = (name1, phone, group)
ws.append(row)
# ws.append([name1, phone, group])
wb.save(fn)
wb.close()
