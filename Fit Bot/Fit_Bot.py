import config # Файл с настройками
import telebot
from telebot import types
import gspread
from gspread_formatting import *
import time
# Гугловский API
from oauth2client.service_account import ServiceAccountCredentials
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('Velogore-0b5ca1626d70.json', scope)

gc = gspread.authorize(credentials)
# Конец Гугловского API
# Поиск свободной ячейки в Учете книг
def av_cell(worksheet):
    str_list = list(filter(None, worksheet.col_values(1))) 
    return str(len(str_list)+2)
# Закрашивание ячейки у взятой книги
fmt = cellFormat(
    backgroundColor=color(0.84, 0.65, 0.74),
    horizontalAlignment='CENTER'
    )

fmt1 = cellFormat(
    backgroundColor=color(0.71, 0.84, 0.66),
    )

# Токен для бота
bot = telebot.TeleBot(config.token)
# Подключение гугл таблиц
wks = gc.open_by_key(config.sht).get_worksheet(0) # Подключение листа с пользователями
wkb = gc.open_by_key(config.sht).get_worksheet(1) # Подключение листа с книгами
wkl = gc.open_by_key(config.sht).get_worksheet(2) # Подключение листа с логом взятия книг
wkd = gc.open_by_key(config.sht).get_worksheet(3) # Подключение листа с долгами

# Переменные
bookid = 0
# Проверка пользователя при вводе /start.
@bot.message_handler(commands=['start'])
def command_start(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
        bot.send_message(m.chat.id, "Привет, я «ФИТ-бот». С моей помощью можно забронировать переговорную комнату или книгу из корпоративной библиотеки.")
        bot.send_chat_action(m.chat.id, 'typing')  # show the bot "typing" (max. 5 secs)
        time.sleep(2)
        bot.send_message(m.chat.id, "Я не нашёл информацию о тебе в базе данных. Если ты являешься сотрудником ФИТ, то свяжись с "+ config.admin + " для внесения в базу данных. В противном случае, ты не сможешь использовать бота.",)
    else:
        bot.send_message(m.chat.id, "Привет, я «ФИТ-бот». С моей помощью можно забронировать книгу из корпоративной библиотеки.")

# Команда /help.
@bot.message_handler(commands=['help'])
def command_help(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
       pass
    else:
        bot.send_message(m.chat.id, config.help)

# Команда /catalog.
@bot.message_handler(commands=['catalog'])
def command_catalog(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
       pass
    else:
        keyboard = types.InlineKeyboardMarkup()
        url_button_books = types.InlineKeyboardButton(text="Список доступных книг", url="https://docs.google.com/spreadsheets/d/1b8EiY6Ru27Oqq3mAHrF1q9vCEWsWLr5E6WkqZzcy6YY/edit?usp=sharing#gid=1655049504") # Ссылка на гугл-таблицу, лист с книгами
        keyboard.add(url_button_books)
        bot.send_message(m.chat.id, "Актуальный список книг всегда доступен в данной таблице:", reply_markup=keyboard)

# Команда /reservation.
@bot.message_handler(commands=['reservation'])
def command_reservation(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
      pass
    else:
        msg = bot.reply_to(m, "Введите номер книги, которую вы хотите взять в аренду:")
        bot.register_next_step_handler(msg, reservation)
def reservation(m):
    bookid = m.text
    if bookid.isdigit():
        if not wkb.cell(int(bookid)+1, 5).value:
            wkb.update_cell(int(bookid)+1, 5, m.from_user.first_name + ' ' + m.from_user.last_name) # Бот заносит в поле "Наличие в библиотеке" имя и фамилию пользователя
            bookid = int(bookid) + 1
            bookname = wkb.cell(int(bookid), 3).value # Название книги
            bookauthor = wkb.cell(int(bookid), 4).value # Автор
            format_cell_range(wkb, 'E'+str(bookid)+':E'+str(bookid), fmt) # Закрашивает ячейку у взятой книги
            av = av_cell(wkl)
            avd = av_cell(wkd)
            avd = int(avd) - 1
            wkl.update_cell(av, 1, bookname)
            wkl.update_cell(av, 2, bookauthor)
            wkl.update_cell(av, 3, m.from_user.first_name + ' ' + m.from_user.last_name)
            wkl.update_cell(av, 4, time.strftime("%d.%m.20%y", time.localtime()))
            wkd.update_cell(avd, 1, m.chat.id)
            wkd.update_cell(avd, 2, time.strftime("%d.%m.20%y", time.localtime()))
            bot.send_message(m.chat.id, "Книга " + bookauthor + " <<" + bookname + ">> успешно забронирована. Пожалуйста, верните её в течение 30 дней.")
        elif wkb.cell(int(bookid)+1, 5).value != None and wkb.cell(int(bookid)+1, 6).value == 'нет':
            bot.send_message(m.chat.id, "Книга уже забронирована другим пользователем. В настоящее время книга находится у пользователя " + wkb.cell(int(bookid)+1, 5).value + ".")
        elif wkb.cell(int(bookid)+1, 5).value != None and wkb.cell(int(bookid)+1, 6).value != 'нет':
            keyboard = types.InlineKeyboardMarkup()
            url_button_drive = types.InlineKeyboardButton(text="Электронный вариант", url="https://drive.google.com/drive/folders/1RaaSF_TnpLPHgNxUIcgkDMLTWYZEVCGV") # Ссылка на гугл-диск с эл. вариантами книг
            keyboard.add(url_button_drive)
            bot.send_message(m.chat.id, "Книга уже забронирована другим пользователем. В настоящее время книга находится у пользователя " + wkb.cell(int(bookid)+1, 5).value + ". Выберете другую или воспользуйте электронным вариантом.", reply_markup=keyboard)
    else:
        bot.send_message(m.chat.id, "Вы ввели неверное значение. Необходимо указать только номер книги.")

# Команда /endofreservation.
@bot.message_handler(commands=['endofreservation'])
def command_endofreservation(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
      pass
    else:
        msg = bot.reply_to(m, "Введите номер книги:")
        bot.register_next_step_handler(msg, endofreservation)
def endofreservation(m):
    bookid = m.text
    if bookid.isdigit():
        wks.update_cell(int(bookid)+1, 5, '') # Бот заносит в поле "Наличие в библиотеке" имя и фамилию пользователя
        bookid = int(bookid) + 1
        bookname = wkb.cell(int(bookid), 3).value # Название книги
        format_cell_range(wkb, 'E'+str(bookid)+':E'+str(bookid), fmt1) # Закрашивает ячейку у взятой книги
        cell = wkd.find(str(m.chat.id))
        wkd.update_cell(cell.col, cell.row, '')
        wkd.update_cell(cell.col, cell.row+1, '')
        bot.send_message(m.chat.id, "Книга успешно сдана. Спасибо, что воспользовались услугами ФИТ-бота.")
    else:
        bot.send_message(m.chat.id, "Вы ввели неверное значение. Необходимо указать только номер книги.")

# Команда /adduser.
@bot.message_handler(commands=['adduser'])
def command_adduser(m):
    if m.from_user.username in config.admin:  # Если пользователь - админ и занесен в config.py:
        msg = bot.reply_to(m, "Введите никнейм пользователя без '@':")
        bot.register_next_step_handler(msg, adduser)
def adduser(m):
    usertoadd = m.text
    avs = av_cell(wks)
    avs = int(avs) - 1
    wks.update_cell(avs, 1, usertoadd) # Бот заносит в поле "Наличие в библиотеке" имя и фамилию пользователя
    bot.send_message(m.chat.id, "Пользователь успешно добавлен.")

# Команда /deleteuser.
@bot.message_handler(commands=['deleteuser'])
def command_deleteuser(m):
    if m.from_user.username in config.admin:  # Если пользователь - админ и занесен в config.py:
        msg = bot.reply_to(m, "Введите никнейм пользователя без '@':")
        bot.register_next_step_handler(msg, deleteuser)
def deleteuser(m):
    usertodelete = m.text
    celltodelete = wks.find(str(usertodelete))
    wks.update_cell(celltodelete.row, celltodelete.col, '') # Бот заносит в поле "Наличие в библиотеке" имя и фамилию пользователя
    bot.send_message(m.chat.id, "Пользователь успешно удален.")

# Команды для переговорных

#Команда /infoperegovornye
@bot.message_handler(commands=['infoperegovornye'])
def command_infoperegovornye(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
       pass
    else:
        bot.send_message(m.chat.id, config.infoperegovornye)

#Команда /reservationleft
@bot.message_handler(commands=['reservationleft'])
def command_reservationleft(m):
    cell_list = wks.findall(m.from_user.username) # Поиск имени пользователя в гугл-таблице
    if not cell_list:  # Если пользователя нет в базе данных:
       pass
    else:
        msg = bot.reply_to(m, "Когда вы хотите забронировать переговорную? Напишите в формате: ДД/ММ/ГГГГ ЧЧ:ММ")
        bot.register_next_step_handler(msg, reservationleft2)
def reservationleft2(m):


if __name__ == '__main__':
    print('Запускаю проверку должников.')
    if time.strftime("%A", time.localtime()) == 'Monday':
        chats_list = wkd.col_values(1) # Получаем список ID диалогов из листа с должниками
        chats_list_date = wkd.col_values(3) # Получаем список дат, когда необходимо было сдать книгу
        for i in range(0, len(chats_list)):
            date_end = time.strptime(chats_list_date[i], "%d.%m.%Y") # Преобразуем дату, когда необходимо сдать книгу в формат даты Python
            date_now = time.strptime(time.strftime("%d.%m.%Y", time.localtime()), "%d.%m.%Y") # Преобразуем текущую дату в такой же формат
            if date_end < date_now: # Сравниваем даты и если время сдачи просрочено - присылаем уведомление
                bot.send_message(chats_list[i], "Вы не сдали книгу. Пожалуйста, верните её в библиотеку ФИТ и затем снимите бронь с помощью команды /endofreservation .")
    print('Проверка должников окончена.')
    print('Запускаю работу бота. Бип-пуп, бип.')
    bot.polling(none_stop=True)