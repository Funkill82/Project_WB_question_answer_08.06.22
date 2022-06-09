import requests
import xlsxwriter


def unloading_excel(user, country, question, answer, imtId):
    """ Выгружаю данные в файл excel """
    try:
        workbook = xlsxwriter.Workbook('data' + str(imtId) + '.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 'Имя пользователя')
        worksheet.write(0, 1, 'Страна')
        worksheet.write(0, 2, 'Вопрос пользователя')
        worksheet.write(0, 3, 'Ответ')
        for i in range(1, len(user)):
            worksheet.write(i, 0, user[i-1])
            worksheet.write(i, 1, country[i-1])
            worksheet.write(i, 2, question[i-1])
            worksheet.write(i, 3, answer[i-1])
    except Exception as e:
        print('Ошибка создания файла')

    finally:
        workbook.close()

def parse(data, name_user, user_country, question_text, answer_text):
    """ Вытаскиваю данные  """
    # name_user_temp, user_country_temp, question_text_temp, answer_text_temp = [], [], [], []

    for item in data['questions']:
        name_user.append(item.get('wbUserDetails', 'Данные отсутствуют').get('name', 'Данные отсутствуют'))
        user_country.append(item.get('wbUserDetails', 'Данные отсутствуют').get('country', 'Данные отсутствуют'))
        question_text.append(item.get('text', 'Данные отсутствуют').replace('\n', ' '))
        answer_text.append(item.get('answer', 'Данные отсутствуют').get('text', 'Данные отсутствуют').replace('\n', ''))




imtId = 12663157
step = 0
data = {}
name_user, user_country, question_text, answer_text = [], [], [], []


while True:
    URL = "https://questions.wildberries.ru/api/v1/questions?imtId=" + str(imtId) + "&skip=" + str(step) + "&take=20"
    r = requests.get(url=URL)
    data = r.json()
    if not data['questions']:
        break
    parse(data, name_user, user_country, question_text, answer_text)
    step += 20


unloading_excel(name_user, user_country, question_text, answer_text, imtId)