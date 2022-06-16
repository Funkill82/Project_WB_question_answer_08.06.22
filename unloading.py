import xlsxwriter



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