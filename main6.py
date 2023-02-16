# coding=utf-8
import config
import telebot
import requests
import logging
import pandas as pd
import json
from telegram import ParseMode
import datetime
import time
import xlwt
import numpy as np
from telebot import types
from telegram_bot_calendar import DetailedTelegramCalendar, LSTEP
from bs4 import BeautifulSoup as BS
import sqlite3
from sqlite3 import Error
from time import sleep, ctime
from json import JSONDecodeError
requestime = datetime.datetime.now()
bot = telebot.TeleBot(config.token)

flag = ''
dateflag = ''
dateresult = ''
today = datetime.datetime.today()

markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
markup.row('Ввести ключ API 🔑')
markup.row('Мне нужна помощь ℹ️')

markup2 = types.ReplyKeyboardMarkup(resize_keyboard=True)
markup2.row('Узнать продажи 💸')
markup2.row('Выбрать дату 📅')
markup2.row('Посмотреть склад 📦')
markup2.row('Ввести ключ API 🔑')
markup2.row('Помочь проекту 💵️')
markup2.row('Мне нужна помощь ℹ️')


#Подключение sqlite3
def post_sql_query(sql_query):
    with sqlite3.connect('my.db') as connection:
        cursor = connection.cursor()
        try:
            cursor.execute(sql_query)
        except Error:
            pass
        result = cursor.fetchall()
        return result

def create_tables():
    users_query = '''CREATE TABLE IF NOT EXISTS USERS 
                        (user_id INTEGER PRIMARY KEY NOT NULL,
                        username TEXT,
                        first_name TEXT,
                        last_name TEXT,
                        reg_date TEXT,
                        api_key TEXT);'''
    post_sql_query(users_query)

def register_user(user, username, first_name, last_name, api_key):
    user_check_query = f'SELECT * FROM USERS WHERE user_id = {user};'
    user_check_data = post_sql_query(user_check_query)
    if not user_check_data:
        insert_to_db_query = f'INSERT INTO USERS (user_id, username, first_name,  last_name, reg_date, api_key) VALUES ({user}, "{username}", "{first_name}", "{last_name}", "{ctime()}", "{api_key}");'
        post_sql_query(insert_to_db_query )
create_tables()

def take_key(user_id1):
    select = 'SELECT api_key FROM "USERS" where user_id=' + str(user_id1)
    con = sqlite3.connect('my.db')
    cur = con.cursor()
    cur.execute(select)
    user_key = cur.fetchone()
    user_key = user_key[0]
    con.close()
    return (user_key)

#Часть календаря
@bot.callback_query_handler(func=DetailedTelegramCalendar.func())
def cal(c):
    global dateflag
    global dateresult
    result, key, step = DetailedTelegramCalendar().process(c.data)
    if not result and key:
        bot.edit_message_text(f"Select {LSTEP[step]}",
                              c.message.chat.id,
                              c.message.message_id,
                              reply_markup=key)
    elif result:
        bot.edit_message_text(f"Выбранная дата: {result}",
                              c.message.chat.id,
                              c.message.message_id)
        dateresult = result
        dateflag = 1

@bot.message_handler(commands=['start', 'help'])
def main(message):
    global flag
    flag = 0
    bot.send_message(message.chat.id,"🙋‍♂️", reply_markup=markup)
    bot.send_message(message.chat.id,
                     "Привет, я помогу тебе следить за продажами на Wildberries.ru и помогу их улучшить\n🛠 Для работы с ботом используйте меню👇\nДанные полученные через API иногда опаздывают от данных в приложении.")
    bot.send_photo(message.chat.id, open('menu_icon.jpg', 'rb'))
    bot.send_message(message.chat.id,
                     "Для начала работы нажмите на кнопку [Ввести ключ API 🔑]")
    ## Подключение аналитики
    data_analitics = {
        "api_key": "51ad1f80d51b48fb66cad0f36362b5f4",
        "events": {
            "user_id": str(message.chat.id),
            "event_type": "[Amplitude] Start Session"
        }

    }
    headers = {
        'Content-Type': 'application/json',
        'Accept': '*/*'
    }

    r = requests.post('https://api.amplitude.com/2/httpapi', data=json.dumps(data_analitics), headers=headers)
@bot.message_handler(content_types=['text'])
def process_date(m):
    if 'Выбрать дату 📅' == str(m.text):
            calendar, step = DetailedTelegramCalendar().build()
            bot.send_message(m.chat.id,
                             f"Select {LSTEP[step]}",
                             reply_markup=calendar)
    if 'Ввести ключ API 🔑' == str(m.text):
        bot.send_message(m.chat.id,
                         'Отправьте мне свой ключ API Wildberries\n\n*Как сформировать API-ключ?*\nДля формирования API-ключа, а также если ваш API-ключ был заблокирован или у вас есть подозрения, что он был скомпрометирован, необходимо создать инцидент в разделе Service Desk на категорию "Поддержка"', parse_mode=ParseMode.MARKDOWN )
        global flag
        flag = 1
    if 'Узнать продажи 💸' == str(m.text):
            global file
            global order_data
            global requestime
            global today
            global dateresult
            try:
                if 1 == dateflag:
                    datesales = dateresult
                else:
                    datesales = today.strftime("%Y-%m-%d")
                if (datetime.datetime.now() - requestime).seconds < 60:
                    msg1 = bot.send_sticker(m.chat.id, 'CAACAgIAAxkBAAJDvF-dcYBM4MHZHt6J2cf56YiITMCIAALcAAP3AsgPUNi8Bnu98HwbBA')
                    msg2 = bot.send_message(m.chat.id,
                                     'Идет загрузка данных⏱')
                    time.sleep(60 - (datetime.datetime.now() - requestime).seconds)
                    bot.delete_message(m.chat.id, msg1.message_id)
                    bot.delete_message(m.chat.id, msg2.message_id)
                rSclad = requests.get(
                    "https://suppliers-stats.wildberries.ru/api/v1/supplier/stocks?dateFrom=2017-03-25T21%3A00%3A00.000Z&key=" + str(take_key(m.from_user.id)))
                filename2 = 'sclad.json'
                file = open(filename2, 'w')
                file.write(json.dumps(rSclad.text))
                file.close()
                file = open(filename2, 'r')
                sclad_data = json.loads(rSclad.text)
                # Сборка артикулов со склада
                nmIdlistSclad = []
                nmIdquantitySclad = []
                nmIdtotalSclad = len(sclad_data)

                i = 0
                # Сборка в 2 массива Артикул, количество со склада
                while i < nmIdtotalSclad:
                    try:
                        q = nmIdlistSclad.index(sclad_data[i]["nmId"])
                        nmIdquantitySclad[q] += sclad_data[i]["quantity"]
                    except ValueError:
                        nmIdlistSclad.append(sclad_data[i]["nmId"])
                        nmIdquantitySclad.append(sclad_data[i]["quantity"])
                    i += 1

                r = requests.get(
                    "https://suppliers-stats.wildberries.ru/api/v1/supplier/orders?dateFrom=" + str(
                        datesales) + "T21%3A00%3A00.000Z&flag=1&key=" + str(take_key(m.from_user.id)))
                requestime = datetime.datetime.now()
                filename = 'orderstoday.json'
                file = open(filename, 'w')
                file.write(json.dumps(r.text))
                file.close()
                file = open(filename, 'r')
                order_data = json.loads(r.text)
                name = ''
                # Чтение Файла
                file = open(filename, 'r')
                order_data = json.loads(r.text)
                # Сборка артикулов со всех заказов
                nmIdlist = []
                nmIdquantity = []
                nmIdprice = []
                nmIdtotal = len(order_data)

                i = 0
                # Сборка в 3 массива Артикул, количество, цена за вычетом скидки
                while i < nmIdtotal:
                    if order_data[i]["nmId"] in nmIdlist:
                        artIndex = nmIdlist.index(order_data[i]["nmId"])
                        nmIdquantity[artIndex] += order_data[i]["quantity"]
                        nmIdprice[artIndex] += (
                                order_data[i]["totalPrice"] / 100 * (100 - int(order_data[i]["discountPercent"])))
                        i += 1
                    else:
                        nmIdlist.append(order_data[i]["nmId"])
                        nmIdquantity.append(order_data[i]["quantity"])
                        nmIdprice.append(
                            order_data[i]["totalPrice"] / 100 * (100 - int(order_data[i]["discountPercent"])))
                        i += 1

                clen = len(nmIdlist) - 1
                i = 0;
                while i <= clen:
                    try:
                        quantityScladStr = str(nmIdquantitySclad[nmIdlistSclad.index(nmIdlist[i])])
                    except ValueError:
                        quantityScladStr = '0'
                    zapros = requests.get(
                        "https://www.wildberries.ru/catalog/" + str(nmIdlist[i]) + "/detail.aspx?targetUrl=mp")
                    html = BS(zapros.content, 'html.parser')
                    name += '<b>' + html.find("span", {
                        'class': 'name'}).getText() + '</b>' + '<code>\n</code>' + 'артикул WB:' + '<a href="https://www.wildberries.ru/catalog/' + str(
                        nmIdlist[i]) + '/detail.aspx?targetUrl=MI"' + '>' + str(
                        nmIdlist[i]) + '</a>' + '<code>\n</code>' + 'количество:' + str(
                        nmIdquantity[i]) + '<code>\n</code>' + 'средняя цена:' + str(
                        nmIdprice[i] // nmIdquantity[
                            i]) + '₽' + '<code>\n</code>' + 'остаток на складе:' + quantityScladStr + '<code>\n</code>' + '<code>\n</code>'
                    i += 1
                file.close()
                totalprice = round(np.sum(nmIdprice))
                totalquantity = np.sum(nmIdquantity)
                totalstring = 'Всего продано:' + str(totalprice) + '₽' + '\n' + 'количество:' + str(
                    totalquantity) + ' шт' + '\n' + 'Данные на ' + str(datesales)
                # Проверка длины сообщения и разделения его на несколько
                if len(name) > 4000:
                    try:
                        # создание и запись в файл
                        my_file = open("long.txt", "w")
                        my_file.write(name)
                        my_file.close()

                        # Отправка файла ботом
                        uis_long = open('long.txt', 'rb')
                        bot.send_document(m.chat.id, uis_long)
                        uis_long.close()
                    except ValueError:
                        bot.send_message(m.chat.id,
                                         "Подождите минутку и повторите отправку команды, Wildberries не разрешает нам делать столько запросов")
                    except NameError:
                        bot.send_message(m.chat.id,
                                         'API ключ не задан, нажмите на кнопку [Ввести ключ API 🔑]')
                    except Exception:
                        bot.send_message(m.chat.id,
                                         'Данных пока нет или они не доступны через API')
                else:
                    bot.send_message(m.chat.id, name, parse_mode=ParseMode.HTML)
                bot.send_message(m.chat.id, totalstring, parse_mode=ParseMode.HTML)
                # Подключение аналитики на получение данных
                data_analitics = {
                    "api_key": "51ad1f80d51b48fb66cad0f36362b5f4",
                    "events": {
                        "user_id": str(m.chat.id),
                        "event_type": "[Amplitude] Revenue (Verified)"
                    }

                }
                headers = {
                    'Content-Type': 'application/json',
                    'Accept': '*/*'
                }
                r = requests.post('https://api.amplitude.com/2/httpapi', data=json.dumps(data_analitics),
                                  headers=headers)
            except ValueError:
                bot.send_message(m.chat.id,
                                 "Подождите минутку и повторите отправку команды, Wildberries не разрешает нам делать столько запросов")
            except NameError:
                bot.send_message(m.chat.id,
                                 'API ключ не задан, нажмите на кнопку [Ввести ключ API 🔑]')
            except Exception:
                bot.send_message(m.chat.id,
                                 'Продаж пока нет или они не доступны через API')
    if len(m.text) == 48:
            if 1 == flag:
                user_key = str(m.text)
                if (datetime.datetime.now() - requestime).seconds < 60:
                    msg1 = bot.send_sticker(m.chat.id,
                                            'CAACAgIAAxkBAAJDvF-dcYBM4MHZHt6J2cf56YiITMCIAALcAAP3AsgPUNi8Bnu98HwbBA')
                    msg2 = bot.send_message(m.chat.id,
                                            'Идет загрузка данных⏱')
                    time.sleep(60 - (datetime.datetime.now() - requestime).seconds)
                    bot.delete_message(m.chat.id, msg1.message_id)
                    bot.delete_message(m.chat.id, msg2.message_id)
                r = requests.get(
                    "https://suppliers-stats.wildberries.ru/api/v1/supplier/orders?dateFrom=" + today.strftime(
                        "%Y-%m-%d") + "T21%3A00%3A00.000Z&flag=1&key=" + user_key)

                requestime = datetime.datetime.now()
                filename = 'orderstoday.json'
                file = open(filename, 'w')
                file.write(json.dumps(r.text))
                file.close()
                file = open(filename, 'r')
                order_data = json.loads(r.text)
                try:
                    if order_data["errors"] == ["can't decode supplier key"]:
                        bot.send_message(m.chat.id, 'Неверный ключ API')
                except TypeError:
                    bot.send_message(m.chat.id, 'Ключ API принят', reply_markup=markup2)
                    flag = 0;
                    register_user(m.from_user.id, m.from_user.username,
                                  m.from_user.first_name, m.from_user.last_name, user_key)

                file.close()
    if 'Помочь проекту 💵️' == str(m.text):
        helpmessage = 'Если вам нравится наш бот и вы хотите поддержать нас пройдите по ссылке: <a href="https://capu.st/marketplaceorder_bot">Поддержать</a> <code>\n</code>Нам очень приятно!😊<code>\n</code> Спасибо!'
        bot.send_message(m.chat.id,
                         helpmessage, parse_mode=ParseMode.HTML)
    if 'Посмотреть склад 📦' == str(m.text):
        try:
            if (datetime.datetime.now() - requestime).seconds < 60:
                msg1 = bot.send_sticker(m.chat.id,
                                        'CAACAgIAAxkBAAJDvF-dcYBM4MHZHt6J2cf56YiITMCIAALcAAP3AsgPUNi8Bnu98HwbBA')
                msg2 = bot.send_message(m.chat.id,
                                        'Идет загрузка данных⏱')
                time.sleep(60 - (datetime.datetime.now() - requestime).seconds)
                bot.delete_message(m.chat.id, msg1.message_id)
                bot.delete_message(m.chat.id, msg2.message_id)
            rSclad = requests.get(
                "https://suppliers-stats.wildberries.ru/api/v1/supplier/stocks?dateFrom=2017-03-25T21%3A00%3A00.000Z&key=" + str(take_key(m.from_user.id)))
            filename2 = 'sclad.json'
            file = open(filename2, 'w')
            file.write(json.dumps(rSclad.text))
            file.close()
            file = open(filename2, 'r')
            sclad_data = json.loads(rSclad.text)
            # Сборка артикулов со склада
            nmIdlistSclad = []
            nmIdquantitySclad = []
            nmIdtotalSclad = len(sclad_data)

            i = 0
            # Сборка в 2 массива Артикул, количество со склада
            while i < nmIdtotalSclad:
                try:
                    q = nmIdlistSclad.index(sclad_data[i]["nmId"])
                    nmIdquantitySclad[q] += sclad_data[i]["quantity"]
                except ValueError:
                    nmIdlistSclad.append(sclad_data[i]["nmId"])
                    nmIdquantitySclad.append(sclad_data[i]["quantity"])
                i += 1
            # создание и запись в файл excel
            i = 0
            result = pd.DataFrame()
            def sclad_table(i):
                global maxnamelen
                try:
                    zapros = requests.get(
                        "https://www.wildberries.ru/catalog/" + str(i) + "/detail.aspx?targetUrl=mp")
                    html = BS(zapros.content, 'html.parser')
                    name = html.find('span', class_= 'name').text
                except AttributeError:
                    name = '-'
                res = pd.DataFrame()
                res = res.append(pd.DataFrame([[i, name, nmIdquantitySclad[k]]],
                                                  columns=['Articul', 'Name', 'quantity']), ignore_index=True)
                return res

            k = 0
            for i in nmIdlistSclad:
                i = sclad_table(i)
                k += 1
                result = result.append(i, ignore_index=True)
            # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter("Sclad.xlsx", engine='xlsxwriter')
            # Convert the dataframe to an XlsxWriter Excel object.
            result.to_excel(writer, sheet_name='Sheet1')
            # Get the xlsxwriter workbook and worksheet objects.
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            # Set the column width and format.
            worksheet.set_column('C:C', 100)
            cell_format = workbook.add_format()
            cell_format.set_pattern(1)  # This is optional when using a solid fill.
            cell_format.set_bg_color('red')
            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
            #Отправка файла ботом
            uis_excel = open('Sclad.xlsx', 'rb')
            bot.send_document(m.chat.id, uis_excel)
            uis_excel.close()
        except ValueError:
            bot.send_message(m.chat.id,
                             "Подождите минутку и повторите отправку команды, Wildberries не разрешает нам делать столько запросов")
        except NameError:
            bot.send_message(m.chat.id,
                             'API ключ не задан, нажмите на кнопку [Ввести ключ API 🔑]')
        except Exception:
            bot.send_message(m.chat.id,
                             'Данных пока нет или они не доступны через API')
    if 'Мне нужна помощь ℹ️' == str(m.text):
        bot.send_message(m.chat.id,
                         'Если вам нужна помощь или что то сломалось напишите: @DmitryZhilin')
    #else:
     #   bot.send_message(m.chat.id,
     #                    'Если вам нужна помощь или что то сломалось напишите: @DmitryZhilin')





if __name__ == '__main__':
    bot.polling(none_stop=True, timeout=300)
