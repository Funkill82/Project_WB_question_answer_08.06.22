import requests
import json
from unloading import unloading_excel
from app.bot.bot_methods import send_message, send_file
import os




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


def get_imt_id(nm_id: int):
    """ По nm_id вытаскиваю imtId """
    URL = "https://wbx-content-v2.wbstatic.net/ru/" + str(nm_id) + ".json"
    r = requests.get(url=URL)
    data = r.json()
    return data.get("imt_id")

def start_app(msg: int):
    username = msg.from_user.username
    try:

        nm_id = int(msg)
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

        file = open(file_name, 'rb')
        send_file(username, file_name, file)

    except:
        send_message(username, 'nm_id Должен состоять только из чисел!')


    finally:
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), file_name)
        os.remove(path)












