from aiogram import types
from aiogram.dispatcher import Dispatcher
from wb_question_answe import start_app




async def echo_message(msg: types.Message):
    start_app(msg)


def register_wb_question_answe(dp: Dispatcher):
    dp.register_message_handler(echo_message, commands='question')