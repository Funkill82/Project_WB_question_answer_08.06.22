import requests
import xlsxwriter
import json
from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor


def unloading_excel(user, country, question, answer,
                    name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks, imtId):
    """ Выгружаю данные в файл excel на двух листах"""
    try:
        file_name = 'data' + str(imtId) + '.xlsx'
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet('Вопрос-ответ')
        worksheet2 = workbook.add_worksheet('Отзыв-ответ')

        cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})


        worksheet.write(0, 0, 'Имя пользователя', cell_format)
        worksheet.write(0, 1, 'Страна', cell_format)
        worksheet.write(0, 2, 'Вопрос пользователя', cell_format)
        worksheet.write(0, 3, 'Ответ', cell_format)

        worksheet2.write(0, 0, 'Имя пользователя', cell_format)
        worksheet2.write(0, 1, 'Страна', cell_format)
        worksheet2.write(0, 2, 'Вопрос пользователя', cell_format)
        worksheet2.write(0, 3, 'Ответ', cell_format)

        worksheet.set_column(0, 1, 15)
        worksheet.set_column(2, 3, 110)

        worksheet2.set_column(0, 1, 15)
        worksheet2.set_column(2, 3, 110)

        cell_format = workbook.add_format({'bold': False, 'font_color': 'black', 'align': 'left', 'valign': 'top'})
        cell_format.set_text_wrap()


        for i in range(1, len(user)):
            worksheet.write(i, 0, user[i - 1], cell_format)
            worksheet.write(i, 1, country[i - 1], cell_format)
            worksheet.write(i, 2, question[i - 1], cell_format)
            worksheet.write(i, 3, answer[i - 1], cell_format)

        for i in range(1, len(name_user_feedbacks)):
            worksheet2.write(i, 0, name_user_feedbacks[i - 1], cell_format)
            worksheet2.write(i, 1, user_country_feedbacks[i - 1], cell_format)
            worksheet2.write(i, 2, question_text_feedbacks[i - 1], cell_format)
            worksheet2.write(i, 3, answer_text_feedbacks[i - 1], cell_format)

    except Exception as e:
        print('Ошибка создания файла')

    finally:
        workbook.close()
    return file_name

def parse_questions(data, name_user, user_country, question_text, answer_text):
    """ Вытаскиваю данные по вопросам """

    for item in data['questions']:
        name_user.append(item.get('wbUserDetails', 'Данные отсутствуют').get('name', 'Данные отсутствуют'))
        user_country.append(item.get('wbUserDetails', 'Данные отсутствуют').get('country', 'Данные отсутствуют'))
        question_text.append(item.get('text', 'Данные отсутствуют').replace('\n', ' '))
        answer_text.append(item.get('answer', 'Данные отсутствуют').get('text', 'Данные отсутствуют').replace('\n', ''))


def parse_feedbacks(data, name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks):
    """ Вытаскиваю данные по отзывам """
    for item in data['feedbacks']:
        name_user_feedbacks.append(item.get('wbUserDetails', 'Данные отсутствуют').get('name', 'Данные отсутствуют'))
        user_country_feedbacks.append(
            item.get('wbUserDetails', 'Данные отсутствуют').get('country', 'Данные отсутствуют'))
        question_text_feedbacks.append(item.get('text', 'Данные отсутствуют').replace('\n', ' '))
        if item['answer']:
            answer_text_feedbacks.append(
                item.get('answer', 'Данные отсутствуют').get('text', 'Данные отсутствуют').replace('\n', ''))
        else:
            answer_text_feedbacks.append(item.get('answer', 'Данные отсутствуют'))


def questions(imtId, name_user, user_country, question_text, answer_text):
    """ По API вопросов достаю json и вызываю парсинг этих данных """
    step = 0
    data = {}
    while True:
        URL = "https://questions.wildberries.ru/api/v1/questions?imtId=" + str(imtId) + "&skip=" + str(
            step) + "&take=20"
        r = requests.get(url=URL)
        data = r.json()
        if not data['questions']:
            break
        parse_questions(data, name_user, user_country, question_text, answer_text)
        step += 20


def feedbacks(imtId, name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks):
    """ По API отзывов достаю json и вызываю парсинг этих данных """

    step = 0
    URL = "https://public-feedbacks.wildberries.ru/api/v1/summary/full"

    while True:
        params = {"imtId": imtId, "skip": step, "take": 20}
        r = requests.post(url=URL, data=json.dumps(params))
        data = r.json()

        if not data["feedbacks"]:
            break
        parse_feedbacks(data, name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks)
        step += 20


def get_imt_id(nm_id:int):
    """ По nm_id вытаскиваю imtId """
    URL = "https://wbx-content-v2.wbstatic.net/ru/" + str(nm_id) + ".json"
    r = requests.get(url=URL)
    data = r.json()
    return data.get("imt_id")




bot = Bot(token='5301069444:AAFRT7o9Uue5J_BOP8-d6gYac2Cv0TdQKB0')
dp = Dispatcher(bot)


@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):
    await message.reply("Привет!\nНапиши мне nm_id цифрами")


@dp.message_handler(commands=['help'])
async def process_help_command(message: types.Message):
    await message.reply("Я могу формировать excel отчет на основании отправленного мне nm_id."
                        " В отчете будет выборка по отзывам-ответам и вопросам-ответам. ")


@dp.message_handler()
async def echo_message(msg: types.Message):

    try:

        nm_id = int(msg.text)
        name_user, user_country, question_text, answer_text = [], [], [], []
        name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks = [], [], [], []

        # По nm_id достаю imtId
        imtId = get_imt_id(nm_id)

        # Достаю вопросы
        questions(imtId, name_user, user_country, question_text, answer_text)

        # Достаю отзывы
        feedbacks(imtId, name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks)

        # Выгружаю в excel
        file_name = unloading_excel(name_user, user_country, question_text, answer_text,
                                    name_user_feedbacks, user_country_feedbacks, question_text_feedbacks,
                                    answer_text_feedbacks, imtId)

        docs = open(file_name, 'rb')
        await bot.send_document(msg.from_user.id, document=docs)
        await bot.send_message(msg.from_user.id, 'Вот твой файл :)')

    except:
        await bot.send_message(msg.from_user.id, 'nm_id Должен состоять только из чисел!')



if __name__ == '__main__':
    executor.start_polling(dp)