import requests
import xlsxwriter


def unloading_excel(user, country, question, answer):
    """ Выгружаю данные в файл excel """
    try:
        workbook = xlsxwriter.Workbook('data.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 'Имя пользователя')
        worksheet.write(0, 1, 'Страна')
        worksheet.write(0, 2, 'Вопрос пользователя')
        worksheet.write(0, 3, 'Ответ')
        for i in range(1, len(user)):
            worksheet.write(i, 0, user[i])
            worksheet.write(i, 1, country[i])
            worksheet.write(i, 2, question[i])
            worksheet.write(i, 3, answer[i])
    except Exception as e:
        print('Ошибка создания файла')

    finally:
        workbook.close()

def parse(data):
    """ Вытаскиваю данные из json """
    name_user, user_country, question_text, answer_text = [], [], [], []

    for item in data['questions']:
        name_user.append(item.get('wbUserDetails', 'Данные отсутствуют').get('name', 'Данные отсутствуют'))
        user_country.append(item.get('wbUserDetails', 'Данные отсутствуют').get('country', 'Данные отсутствуют'))
        question_text.append(item.get('text', 'Данные отсутствуют').replace('\n', ' '))
        answer_text.append(item.get('answer', 'Данные отсутствуют').get('text', 'Данные отсутствуют').replace('\n', ''))

    unloading_excel(name_user, user_country, question_text, answer_text)


URL = "https://questions.wildberries.ru/api/v1/questions?imtId=30499672&skip=0&take=20"

r = requests.get(url=URL)

data = r.json()

parse(data)

