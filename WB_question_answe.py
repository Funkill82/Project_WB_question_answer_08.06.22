import requests
import xlsxwriter
import json


def unloading_excel(user, country, question, answer,
                    name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks, imtId):
    """ Выгружаю данные в файл excel на двух листах"""
    try:

        workbook = xlsxwriter.Workbook('data' + str(imtId) + '.xlsx')
        worksheet = workbook.add_worksheet('Вопрос-ответ')
        worksheet2 = workbook.add_worksheet('Отзыв-ответ')

        bold = workbook.add_format({'bold': True})

        worksheet.write(0, 0, 'Имя пользователя', bold)
        worksheet.write(0, 1, 'Страна', bold)
        worksheet.write(0, 2, 'Вопрос пользователя', bold)
        worksheet.write(0, 3, 'Ответ', bold)

        worksheet2.write(0, 0, 'Имя пользователя', bold)
        worksheet2.write(0, 1, 'Страна', bold)
        worksheet2.write(0, 2, 'Вопрос пользователя', bold)
        worksheet2.write(0, 3, 'Ответ', bold)

        worksheet.set_column(0, 1, 15)
        worksheet.set_column(2, 3, 120)

        worksheet2.set_column(0, 1, 15)
        worksheet2.set_column(2, 3, 120)


        for i in range(1, len(user)):
            worksheet.write(i, 0, user[i - 1])
            worksheet.write(i, 1, country[i - 1])
            worksheet.write(i, 2, question[i - 1])
            worksheet.write(i, 3, answer[i - 1])

        for i in range(1, len(name_user_feedbacks)):
            worksheet2.write(i, 0, name_user_feedbacks[i - 1])
            worksheet2.write(i, 1, user_country_feedbacks[i - 1])
            worksheet2.write(i, 2, question_text_feedbacks[i - 1])
            worksheet2.write(i, 3, answer_text_feedbacks[i - 1])

    except Exception as e:
        print('Ошибка создания файла')

    finally:
        workbook.close()


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


imtId = 12663157

name_user, user_country, question_text, answer_text = [], [], [], []
name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks = [], [], [], []

questions(imtId, name_user, user_country, question_text, answer_text)

feedbacks(imtId, name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks)

unloading_excel(name_user, user_country, question_text, answer_text,
                name_user_feedbacks, user_country_feedbacks, question_text_feedbacks, answer_text_feedbacks, imtId)
