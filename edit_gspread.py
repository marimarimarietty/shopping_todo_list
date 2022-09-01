from flask import Flask, request, abort

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, ImageMessage, TextSendMessage, FollowEvent, SourceGroup, SourceRoom
)
from re_con_gspread import RemoteControlGoogleSpreadSheet

import os
from dotenv import load_dotenv
load_dotenv()


worksheets = {}#initialize dictionary
separator = "==================="


app = Flask(__name__)

line_bot_api = LineBotApi(os.environ['LINE_TOKEN'])
handler = WebhookHandler(os.environ['CHANNEL_SECRET'])

@app.route("/callback", methods=['POST'])
def callback():
  signature = request.headers['X-Line-Signature']

  body = request.get_data(as_text=True)
  app.logger.info("Request body: " + body)

  try:
      handler.handle(body, signature)
  except InvalidSignatureError:
    print("Invalid signature. Please check your channel access token/channel secret.")
    abort(400)


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
  text = event.message.text
  profile = line_bot_api.get_profile(event.source.user_id)

  try:
    #store worksheet object of the user when user send text message.
    worksheet = worksheets[profile.user_id]

  except KeyError:
    #if it doesn't be registered in dictionary, register it again.
    worksheets[profile.user_id] = RemoteControlGoogleSpreadSheet(profile.display_name)
    worksheet = worksheets[profile.user_id]


  #ユーザーから送信されたテキストメッセージの最初の２文字が"追加"だったら、その後の文字列をTodo列に書き込む
  if text[:2] == "追加" :
    worksheet.write_to_Todo(text)

    line_bot_api.reply_message(
      event.reply_token,
      TextSendMessage(text = text[2:] + "\nをリストに追加しました"))


  #when user send message include the text '完了', move the list to raw Done.
  #if it doesn't exist, send message to user.
  if text[:2] == "完了" :
    is_there_cell = worksheet.from_Todo_to_Done(text)

    if is_there_cell == True:
      line_bot_api.reply_message(
      event.reply_token,
      TextSendMessage(text = text[2:] + "\nをリストから削除しました"))
    else :
      line_bot_api.reply_message(
      event.reply_token,
      TextSendMessage(text = text[2:] + "\nはリストに存在しません"))


  #if user send '買うもの', return all lists
  if text == "買うもの" :
    respons = worksheet.get_Todo()

    line_bot_api.reply_message(
    event.reply_token,
    TextSendMessage(text = "買うもの\n・" + respons))


  #if user send 'クリア', reset worksheet.
  if text == "クリア" :
    worksheet.clear_worksheet()
    line_bot_api.reply_message(
    event.reply_token,
    TextSendMessage(text = "リストをクリアしました"))


  #if user send '削除', remove worksheet
  if text == "削除" :
    worksheet.delete_worksheet()
    del worksheets[profile.user_id]
    line_bot_api.reply_message(
    event.reply_token,
    TextSendMessage(text = "リストを削除しました"))
