import os
import re
import json

import openai
import gspread

from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from oauth2client.service_account import ServiceAccountCredentials

# ============================
# Slack の認証情報 (環境変数)
# ============================
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ["SLACK_SIGNING_SECRET"]

# ============================
# Google認証 (Secret Filesを利用)
# ============================
SERVICE_ACCOUNT_FILE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # Secret Filesパス
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")  # スプレッドシートID

# ============================
# OpenAI APIキー
# ============================
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
openai.api_key = OPENAI_API_KEY  # ここでセット

# Slack Bolt アプリを初期化
app_bolt = App(
    token=SLACK_BOT_TOKEN,
    signing_secret=SLACK_SIGNING_SECRET
)

# Flask アプリ生成（Boltのイベントを受け取る用）
flask_app = Flask(__name__)
handler = SlackRequestHandler(app=app_bolt)

# -----------------------
# Google Sheets クライアントの初期化
# -----------------------
def get_gspread_client():
    """
    Secret Filesによるサービスアカウントファイルを読み込み、
    Googleスプレッドシートにアクセス可能なクライアントを返す
    """
    if not SERVICE_ACCOUNT_FILE:
        raise ValueError("環境変数 GCP_SERVICE_ACCOUNT_JSON が設定されていません。")

    with open(SERVICE_ACCOUNT_FILE, "r", encoding="utf-8") as f:
        service_account_dict = json.load(f)

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_dict, scope)
    gc = gspread.authorize(credentials)
    return gc

# -----------------------
# 「【○○○様】」から医院名を抜き出す
# -----------------------
def extract_hospital_name(text: str) -> str:
    """
    例: 【センター北あだち歯科様】 → "センター北あだち歯科"
    """
    pattern = r"【([^】]+)】"  # 「【...】」の中身
    match = re.search(pattern, text)
    if not match:
        return ""

    raw_name = match.group(1).strip()
    # 末尾が「様」であれば削除
    if raw_name.endswith("様"):
        raw_name = raw_name[:-1]
    return raw_name

# -----------------------
# 「○○○よりXXXXの応募がございました。」から媒体名を抜き出す
# -----------------------
def extract_media_name(text: str) -> str:
    """
    例: "ジョブメドレーより歯科医師の応募がございました。"
        → "ジョブメドレー"
    """
    pattern = r"(.+?)より(.+?)の応募がございました。"
    match = re.search(pattern, text)
    if not match:
        return ""
    media_name = match.group(1).strip()
    return media_name

# -----------------------
# OpenAIを用いてプロフィール情報を抽出
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    ChatCompletionで以下をJSON形式で抽出:
      name, member_id, age, job, experience, address, status, cert, education
    """
    if not OPENAI_API_KEY:
        return {}

    system_prompt = (
        "あなたはテキストから以下の情報を抽出するアシスタントです。\n"
        "抽出すべき項目: name(氏名), member_id(会員番号), age(年齢), "
        "job(職種), experience(経験), address(お住まい), status(就業状況), "
        "cert(資格), education(最終学歴)\n"
        "出力は必ず JSON 形式のみで、キー名は上記の英語でお願いします。\n"
        "値が不明の場合は空文字にしてください。"
    )

    user_prompt = (
        f"以下のテキストから必要項目を抜き出して、JSON形式で返してください。\n\n"
        f"{text}\n"
    )

    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
        )
        content = response["choices"][0]["message"]["content"].strip()
        extracted_data = json.loads(content)

        return {
            "name":       extracted_data.get("name", ""),
            "member_id":  extracted_data.get("member_id", ""),
            "age":        extracted_data.get("age", ""),
            "job":        extracted_data.get("job", ""),
            "experience": extracted_data.get("experience", ""),
            "address":    extracted_data.get("address", ""),
            "status":     extracted_data.get("status", ""),
            "cert":       extracted_data.get("cert", ""),
            "education":  extracted_data.get("education", ""),
        }
    except Exception as e:
        print(f"OpenAI API error: {e}")
        return {}

# -----------------------
# スプレッドシートへ書き込み
# -----------------------
def write_to_spreadsheet(data: dict):
    """
    ディクショナリ data の内容をスプレッドシートに1行追加
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    worksheet = sh.worksheet("Sheet1")  # シート名はSheet1とする（存在しない場合は作成要）

    # 列を拡張（病院名・媒体名を先頭に加えた例: hospital_name, media_name, name, ..., education）
    new_row = [
        data.get("hospital_name", ""),
        data.get("media_name", ""),
        data.get("name", ""),
        data.get("member_id", ""),
        data.get("age", ""),
        data.get("job", ""),
        data.get("experience", ""),
        data.get("address", ""),
        data.get("status", ""),
        data.get("cert", ""),
        data.get("education", ""),
    ]
    worksheet.append_row(new_row, value_input_option="USER_ENTERED")

# -----------------------
# Slack Bolt: メッセージイベントのハンドラ
# -----------------------
@app_bolt.event("message")
def handle_message_events(body, say, logger):
    """
    チャンネルへの新規メッセージイベントを受け取る
    """
    event = body.get("event", {})
    text = event.get("text", "")
    thread_ts = event.get("ts")

    # 「応募がございました。」を含むかをチェック
    if "応募がございました。" in text:
        # 1) まず病院名(医院名)を抽出
        hospital_name = extract_hospital_name(text)
        # 2) 媒体名を抽出
        media_name = extract_media_name(text)

        # 3) OpenAIでプロフィール情報抽出
        parsed_profile = parse_profile_info(text)

        # 4) 全情報をまとめたdictを作成
        merged_data = {
            "hospital_name": hospital_name,
            "media_name": media_name,
            **parsed_profile,  # 既存の name, age, job etc を展開
        }

        # 名前 or 会員番号など、何かしら得られている場合のみ書き込みを行う
        # (必須条件がなければ省略しても良い)
        if merged_data["name"] or merged_data["member_id"]:
            try:
                write_to_spreadsheet(merged_data)
                logger.info("スプレッドシートへの書き込みに成功しました。")
                # 書き込み完了をSlackに返信
                say(
                    text="スプレッドシート書き込みが完了しました。",
                    thread_ts=thread_ts
                )
            except Exception as e:
                import traceback
                logger.error(f"スプレッドシートへの書き込みでエラーが発生: {e}")
                traceback.print_exc()
                logger.exception("スプレッドシートへの書き込みでエラーの詳細スタックトレース")
                say(
                    text=f"スプレッドシートへの書き込みでエラー: {e}",
                    thread_ts=thread_ts
                )

# -----------------------
# Flaskルート設定
# -----------------------
@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

@flask_app.route("/", methods=["GET"])
def healthcheck():
    return "OK", 200

# -----------------------
# アプリ起動 (RenderでのGunicorn運用を想定)
# -----------------------
if __name__ == "__main__":
    flask_app.run(host="0.0.0.0", port=5000)
