import telebot
import config
import datetime
import pytz
import openpyxl


class Day:
    def __init__(self, ntext, nlink, ptext, plink):
        self.ntext = []
        self.nlink = []
        self.ptext = []
        self.plink = []


rxl = openpyxl.open("rall.xlsx", read_only=True)

rallarray = [["Непарний понеділок", "Непарний вівторок", "Непарна середа", "Непарний четвер", "Непарна п'ятниця"],
             ["Парний понеділок", "Парний вівторок", "Парна середа", "Парний четвер", "Парна п'ятниця"]]

mondey = Day([], [], [], [])
tuesday = Day([], [], [], [])
wednesday = Day([], [], [], [])
thursday = Day([], [], [], [])
friday = Day([], [], [], [])

darray = [mondey, tuesday, wednesday, thursday, friday]
for i in range(5):
    sheet = rxl.worksheets[i]
    for j in range(4):
        darray[i].ptext.append(sheet[(j + 3)][1].value)
        # print(darray[i].ptext)
        darray[i].plink.append(sheet[(j + 3)][2].value)
        darray[i].ntext.append(sheet[(j + 3)][3].value)
        darray[i].nlink.append(sheet[(j + 3)][4].value)

bot = telebot.TeleBot(config.TOKEN)


@bot.message_handler(commands=['set_avatar'])
def set_avatar(message):
    photo = open('avatar.png', 'rb')
    bot.set_my_profile_photo(photo)
    bot.send_message(message.chat.id, "Аватарка успішно встановлена!")


@bot.message_handler(commands=['start'])
def wellcum(message):
    sti = open('stickers/peremoga.webp', 'rb')
    bot.send_sticker(message.chat.id, sti)
    bot.send_message(message.chat.id, "Ярік ботяра йобаний. 2.0.2")


@bot.message_handler(commands=['tag'])
def tag_people(message):
    args = message.text.split()[1:]
    if len(args) == 0:
        bot.reply_to(message, "Довбойоб ще раз блять введи па нармальнаму '/tag ім'я'наххуй.")
        return

    found_users = []  # Список знайдених користувачів

    if args[0] == 'all':
        with open('list.txt', 'r', encoding='utf-8') as file:
            for line in file:
                parts = line.strip().split(',')
                if len(parts) >= 2:  # Перевірка наявності коми для розділення
                    name = parts[0]
                    username = parts[1]
                    found_users.append(username)
    else:
        with open('list.txt', 'r', encoding='utf-8') as file:
            for line in file:
                parts = line.strip().split(',')
                if len(parts) >= 2:  # Перевірка наявності коми для розділення
                    name = parts[0]
                    username = parts[1]
                    for arg in args:
                        if arg.lower() in name.lower():
                            found_users.append(username)

    if found_users:
        tag_limit = 3  # Кількість тегів у повідомленні
        chunks = [found_users[i:i + tag_limit] for i in range(0, len(found_users), tag_limit)]
        for chunk in chunks:
            tagged_users = ', '.join(chunk)
            bot.send_message(message.chat.id, tagged_users)
    else:
        bot.reply_to(message, "Ніхуя сука не знайшов(")


def interpreter(array, even, i, j):
    reply = ""
    if even == 0:

        if darray[i].nlink is not None:
            text_pjerdole = str(darray[i].ntext[j]).replace('.', '\.').replace('(', '\(').replace(')', '\)').replace('-', '\-')
            # print(text_pjerdole)
            link_pjerdole = str(darray[i].nlink[j]).replace('.', '\.').replace('(', '\(').replace(')', '\)').replace('-', '\-')
            # print(link_pjerdole)
            reply += f"{j + 1}\. [{text_pjerdole}]({link_pjerdole}) \n"
        else:
            reply += f"{j}. {array[i].ntext}\n"
    else:

        if darray[i].plink is not None:
            text_pjerdole = str(darray[i].ptext[j]).replace('.', '\.').replace('(', '\(').replace(')', '\)').replace('-', '\-')
            link_pjerdole = str(darray[i].plink[j]).replace('.', '\.').replace('(', '\(').replace(')', '\)').replace('-', '\-')
            reply += f"{j + 1}\. [{text_pjerdole}]({link_pjerdole}) \n"
        else:
            reply += f"{j}. {array[i].ptext}\n"
    reply += "\n"
    return reply

@bot.message_handler(commands=['rhelp'])
def rozklad_help(message):
    reply=f"/rall - весь розклад\n" \
          f"/rtweek - розклад на цей тиждень" \
          f"\n/rtomorrow - розклад на завтра" \
          f"\n/rtoday - розклад на сьогодні" \
          f"\n/rnext - наступна пара" \
          f"\n/rnow - пара що є зараз".replace('-', '\-')
    bot.reply_to(message, reply, parse_mode='MarkdownV2')

@bot.message_handler(commands=['rall'])
def rozklad_all(message):
    reply = ""
    for i in range(5):
        even = False
        reply += f"{rallarray[0][i]}\n\n"
        for j in range(4):
            reply += interpreter(rallarray, even, i, j)

        even = True
        reply += f"{rallarray[1][i]}\n\n"
        for k in range(4):
            reply += interpreter(rallarray, even, i, k)

        reply += "\n-----------------------------------------------\n".replace('-', '\-')
    bot.reply_to(message, reply, parse_mode='MarkdownV2')


@bot.message_handler(commands=['rtweek'])
def rozklad_this_week(message):
    reply = ""
    even = datetime.date.today().isocalendar()[1] % 2
    for i in range(5):
        if even == 0:

            reply += f"{rallarray[0][i]}\n\n"
            for j in range(4):
                reply += interpreter(rallarray, even, i, j)

        else:

            reply += f"{rallarray[1][i]}\n\n"
            for k in range(4):
                reply += interpreter(rallarray, even, i, k)

            reply += "\n-----------------------------------------------\n".replace('-', '\-')
    bot.reply_to(message, reply, parse_mode='MarkdownV2')


@bot.message_handler(commands=['rtoday'])
def rozklad_today(message):
    reply = ""
    even = datetime.date.today().isocalendar()[1] % 2
    day_number = datetime.date.today().weekday()
    if day_number < 5:

        if even == 0:

            reply += f"{rallarray[0][day_number]}\n\n"
            for j in range(4):
                reply += interpreter(rallarray, even, day_number, j)

        else:

            reply += f"{rallarray[1][day_number]}\n\n"
            for k in range(4):
                reply += interpreter(rallarray, even, day_number, k)

            reply += "\n-----------------------------------------------\nСлався АС-224 та її виборний! Слава Україні!\n".replace(
            '-', '\-').replace('!', '\!')
        bot.reply_to(message, reply, parse_mode='MarkdownV2')
    else:
        bot.reply_to(message, "Далбайоб сьогодні вихідний!")



@bot.message_handler(commands=['rtomorrow'])
def rozklad_tomorrow(message):
    reply = ""
    even = datetime.date.today().isocalendar()[1] % 2
    day_number = 1 + datetime.date.today().weekday()
    if day_number < 5:

        if even == 0:

            reply += f"{rallarray[0][day_number]}\n\n"
            for j in range(4):
                reply += interpreter(rallarray, even, day_number, j)

        else:

            reply += f"{rallarray[1][day_number]}\n\n"
            for k in range(4):
                reply += interpreter(rallarray, even, day_number, k)

            reply += "\n-----------------------------------------------\nСлався АС-224 та її виборний! Слава Україні!\n".replace(
            '-', '\-').replace('!', '\!')
        bot.reply_to(message, reply, parse_mode='MarkdownV2')
    else:
        if day_number == 7:
            if even == 1:

                reply += f"{rallarray[0][0]}\n\n"
                for j in range(4):
                    reply += interpreter(rallarray, even, 0, j)

            else:

                reply += f"{rallarray[1][0]}\n\n"
                for k in range(4):
                    reply += interpreter(rallarray, even, 0, k)

                reply += "\n-----------------------------------------------\nСлався АС-224 та її виборний! Слава Україні!\n".replace(
                    '-', '\-').replace('!', '\!')
            bot.reply_to(message, reply, parse_mode='MarkdownV2')
        bot.reply_to(message, "Далбайоб сьогодні вихідний!")

@bot.message_handler(commands =['rnext'])
def rozklad_next(message):
    reply =""
    number=0
    even = datetime.date.today().isocalendar()[1] % 2
    day_number = datetime.date.today().weekday()

    # Отримайте поточний час
    odessa_timezone = pytz.timezone('Europe/Kiev')

    # Отримайте поточний час у часовому поясі Одеси
    current_time = datetime.datetime.now(odessa_timezone).time()
    # Визначте інтервали
    intervals = [
        (datetime.time(8, 0), datetime.time(9, 49)),
        (datetime.time(9, 50), datetime.time(11, 39)),
        (datetime.time(11, 40), datetime.time(13, 29)),
        (datetime.time(13, 30), datetime.time(15, 5)),
    ]

    # Перевірте, в якому інтервалі знаходиться поточний час
    current_interval = None

    for i, interval in enumerate(intervals, start=1):
        start_time, end_time = interval
        if start_time <= current_time <= end_time:
            current_interval = i
            break
        elif current_time < start_time:
            current_interval=0
            break

    # Виведіть результат
    if current_interval == 4:
        rozklad_tomorrow(message)
    elif current_interval ==0:
        rozklad_today(message)

    elif current_interval is not None:
        reply += f"Некст пара\n\n"
        reply += interpreter(rallarray, even, day_number, current_interval)

        reply += "\n-----------------------------------------------\nСлався АС-224 та її виборний! Слава Україні!\n".replace(
            '-', '\-').replace('!', '\!')
        bot.reply_to(message, reply, parse_mode='MarkdownV2')
    else:
        rozklad_tomorrow(message)

@bot.message_handler(commands=['rnow'])
def rozklad_now(message):
    reply = ""
    number = 0
    even = datetime.date.today().isocalendar()[1] % 2
    day_number = datetime.date.today().weekday()

    # Отримайте поточний час
    odessa_timezone = pytz.timezone('Europe/Kiev')

    # Отримайте поточний час у часовому поясі Одеси
    current_time = datetime.datetime.now(odessa_timezone).time()

    # Визначте інтервали
    intervals = [
        (datetime.time(8, 0), datetime.time(9, 34)),
        (datetime.time(9, 35), datetime.time(9, 49)),
        (datetime.time(9, 50), datetime.time(11, 24)),
        (datetime.time(11, 25), datetime.time(11, 39)),
        (datetime.time(11, 40), datetime.time(13, 14)),
        (datetime.time(13, 15), datetime.time(13, 29)),
        (datetime.time(13, 30), datetime.time(15, 5)),
    ]

    # Перевірте, в якому інтервалі знаходиться поточний час
    current_interval = None

    for i, interval in enumerate(intervals, start=1):
        start_time, end_time = interval
        if start_time <= current_time <= end_time:
            current_interval = i
            break

    # Виведіть результат
    if current_interval is not None:
        if current_interval %2 != 0:
            if even == 0:

                reply += f"Зараз йде пара\n\n"
                reply += interpreter(rallarray, even, day_number, int(current_interval / 2))

            else:

                reply += f"Зараз йде пара\n\n"
                reply += interpreter(rallarray, even, day_number, int(current_interval / 2))

                reply += "\n-----------------------------------------------\nСлався АС-224 та її виборний! Слава Україні!\n".replace('-', '\-').replace('!', '\!')
        else:
            rozklad_next(message)
        bot.reply_to(message, reply, parse_mode='MarkdownV2')
    else:
        rozklad_tomorrow(message)


@bot.message_handler(commands=['test'])
def test(message):
    bebra = 5


bot.polling(none_stop=True)
