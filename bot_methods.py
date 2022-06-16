import requests
import xlsxwriter
import json
from io import BytesIO


def unloading_excel(questions_list, feedbacks_list):
    """ Выгружаю данные в файл excel на двух листах"""
    try:
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
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

        for i in range(1, len(questions_list[0].get("name_user"))):

            worksheet.write(i, 0, questions_list[0].get("name_user")[i - 1], cell_format)
            worksheet.write(i, 1, questions_list[1].get("user_country")[i - 1], cell_format)
            worksheet.write(i, 2, questions_list[2].get("question_text")[i - 1], cell_format)
            worksheet.write(i, 3, questions_list[3].get("answer_text")[i - 1], cell_format)

        for i in range(1, len(feedbacks_list[0].get("name_user_feedbacks"))):
            worksheet2.write(i, 0, feedbacks_list[0].get("name_user_feedbacks")[i - 1], cell_format)
            worksheet2.write(i, 1, feedbacks_list[1].get("user_country_feedbacks")[i - 1], cell_format)
            worksheet2.write(i, 2, feedbacks_list[2].get("question_text_feedbacks")[i - 1], cell_format)
            worksheet2.write(i, 3, feedbacks_list[3].get("answer_text_feedbacks")[i - 1], cell_format)

    except:
        return False
    finally:
        workbook.close()
    return output


def parse_questions(data, questions_list):
    """ Вытаскиваю данные по вопросам """

    for item in data['questions']:
        questions_list[0].get("name_user").append(
            item.get('wbUserDetails', 'Данные отсутствуют').get('name', 'Данные отсутствуют'))
        questions_list[1].get("user_country").append(
            item.get('wbUserDetails', 'Данные отсутствуют').get('country', 'Данные отсутствуют'))
        questions_list[2].get("question_text").append(item.get('text', 'Данные отсутствуют').replace('\n', ' '))
        questions_list[3].get("answer_text").append(
            item.get('answer', 'Данные отсутствуют').get('text', 'Данные отсутствуют').replace('\n', ''))


def parse_feedbacks(data, feedbacks_list):
    """ Вытаскиваю данные по отзывам """
    for item in data['feedbacks']:

        feedbacks_list[0].get("name_user_feedbacks").append(item.get('wbUserDetails', 'Данные отсутствуют')
                                                            .get('name', 'Данные отсутствуют'))
        feedbacks_list[1].get("user_country_feedbacks").append(item.get('wbUserDetails', 'Данные отсутствуют')
                                                               .get('country', 'Данные отсутствуют'))
        feedbacks_list[2].get("question_text_feedbacks").append(item.
                                                                get('text', 'Данные отсутствуют').replace('\n', ' '))
        if item["answer"]:
            feedbacks_list[3].get("answer_text_feedbacks").append(item.get('answer')
                                                                  .get('text', 'Данные отсутствуют').replace('\n', ''))
        else:
            feedbacks_list[3].get("answer_text_feedbacks").append(item.get('answer', 'Данные отсутствуют'))


def questions(imtId: int, questions_list: list) -> bool:
    """ По API вопросов достаю json и вызываю парсинг этих данных """
    step = 0
    data = {}
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36'
                             ' (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
    while True:
        URL = f"https://questions.wildberries.ru/api/v1/questions?imtId={imtId}&skip={step}&take=20"
        try:
            r = requests.get(url=URL, headers=headers)
            data = r.json()
        except:
            return False

        if not data['questions']:
            break
        parse_questions(data, questions_list)
        step += 20
    return True


def feedbacks(imtId: int, feedbacks_list: list) -> bool:
    """ По API отзывов достаю json и вызываю парсинг этих данных """

    step = 0
    URL = "https://public-feedbacks.wildberries.ru/api/v1/summary/full"

    while True:
        try:
            params = {"imtId": imtId, "skip": step, "take": 20}
            r = requests.post(url=URL, data=json.dumps(params))
            data = r.json()
            step += 20
        except:
            return False

        if not data["feedbacks"]:
            break
        parse_feedbacks(data, feedbacks_list)
    return True


def get_imt_id(nm_id: int):
    """ По nm_id вытаскиваю imtId """

    URL = f"https://wbx-content-v2.wbstatic.net/ru/{nm_id}.json"
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36'
                             ' (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
    try:
        r = requests.get(url=URL, headers=headers)
        data = r.json()
    except:
        return False
    return data.get("imt_id")
