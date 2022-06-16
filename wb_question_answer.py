from aiogram import types
from bot_methods import get_imt_id, questions, feedbacks, unloading_excel
from app.bot.bot_methods import send_message, send_file

def question_app(msg: types.Message):
    username = msg.from_user.username
    try:
        nm_id = int(msg.text)
    except:
        send_message(username, 'nm_id Должен состоять только из чисел!')


    questions_list = [{"name_user": []}, {"user_country": []},
                      {"question_text": []}, {"answer_text": []}]
    feedbacks_list = [{"name_user_feedbacks": []}, {"user_country_feedbacks": []},
                      {"question_text_feedbacks": []}, {"answer_text_feedbacks": []}]

    # По nm_id достаю imtId
    imtId = get_imt_id(nm_id)
    # if imtId:
    #     # Достаю вопросы
    #     if questions(imtId, questions_list):
    #         # Достаю отзывы
    #         if feedbacks(imtId, feedbacks_list):
    #             # Выгружаю в excel
    #             output = unloading_excel(questions_list, feedbacks_list)
    #             if output:
    #                 file = output.getvalue()
    #                 file_name = f"data'{imtId}.xlsx"
    #                 send_file(username, file_name, file)
    #             else:
    #                 send_message(username, 'Ошибка создания файла')
    #         else:
    #             send_message(username, 'Сайт недоступен')
    #     else:
    #         send_message(username, 'Сайт недоступен')
    # else:
    #     send_message(username, 'Неправильный nm_id или сайт недоступен')
    if not imtId:
        send_message(username, 'Неправильный nm_id или сайт недоступен')
    else:
        if not questions(imtId, questions_list):
            send_message(username, 'Сайт недоступен')
        else:
            if not feedbacks(imtId, feedbacks_list):
                send_message(username, 'Сайт недоступен')
            else:
                output = unloading_excel(questions_list, feedbacks_list)
                if not output:
                    send_message(username, 'Ошибка создания файла')
                else:
                    file = output.getvalue()
                    file_name = f"data'{imtId}.xlsx"
                    send_file(username, file_name, file)


