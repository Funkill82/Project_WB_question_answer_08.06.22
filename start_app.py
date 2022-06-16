from aiogram import types
from aiogram.dispatcher import Dispatcher
from wb_question_answer_temp import question_app


async def bot_message(msg: types.Message):
    question_app(msg)


def register_load_feedbacks(dp: Dispatcher):
    dp.register_message_handler(bot_message, commands='question')