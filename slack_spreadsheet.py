import os
import json
import openai
import gspread
from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import WorksheetNotFound

# ============================
# Slack の認証情報 (環境変数)
# ============================
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ["SLACK_SIGNING_SECRET"]

# ============================
# Google認証 (Secret Filesを利用)
# ============================
SERVICE_ACCOUNT_FILE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # Secret Filesパス
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")  # スプレッドシートのID

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
        raise ValueError("環境変数 GCP_SERVICE_ACCOUNT_JSON (Secret Filesパス) が設定されていません。")

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
# チャンネル名のワークシートを取得 or 作成
# -----------------------
def get_or_create_worksheet(sh, worksheet_title: str):
    """
    スプレッドシートオブジェクト sh から、
    worksheet_title に一致するワークシートを探す。
    なければ新規作成して返す。
    """
    try:
        worksheet = sh.worksheet(worksheet_title)
    except WorksheetNotFound:
        # 新規にワークシートを作成
        sh.add_worksheet(title=worksheet_title, rows=100, cols=20)
        worksheet = sh.worksheet(worksheet_title)
    return worksheet

# -----------------------
# OpenAI を用いて情報抽出
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    OpenAI API (ChatCompletion) を使って、応募メッセージから各種情報を抽出する。
    抽出すべき項目: name, member_id, age, job, experience, address, status, cert, education
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
def write_to_spreadsheet(profile_data: dict, channel_name: str):
    """
    スプレッドシートに1行追加する
    - channel_name というタイトルのワークシートを取得 or 作成して書き込み
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)

    # チャンネル名のワークシートを取得 or 作成
    worksheet = get_or_create_worksheet(sh, channel_name)

    new_row = [
        profile_data["name"],
        profile_data["member_id"],
        profile_data["age"],
        profile_data["job"],
        profile_data["experience"],
        profile_data["address"],
        profile_data["status"],
        profile_data["cert"],
        profile_data["education"]
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
    channel_id = event.get("channel")

    # 「ジョブメドレーより〇〇の応募がございました。」を含むかチェック
    if "ジョブメドレーより" in text and "の応募がございました" in text:
        # 1) チャンネル名を取得
        # conversations_info を呼ぶことで、チャンネルIDからチャンネル名を取得
        # ※ Botに channels:read または groups:read スコープが必要
        try:
            channel_info = app_bolt.client.conversations_info(channel=channel_id)
            # 公開チャンネルなら channel["name"], プライベートなら channel["name"] が取得できる
            slack_channel_name = channel_info["channel"]["name"]
        except Exception as e:
            logger.error(f"チャンネル名の取得に失敗しました: {e}")
            slack_channel_name = f"UnknownChannel_{channel_id}"

        # 2) OpenAI で情報を抽出
        parsed_data = parse_profile_info(text)

        # 3) スプレッドシートに書き込み
        if parsed_data["name"] or parsed_data["member_id"]:
            try:
                write_to_spreadsheet(parsed_data, slack_channel_name)
                logger.info("スプレッドシートへの書き込みに成功しました。")

                # 書き込み完了メッセージ
                say(
                    text=f"スプレッドシートへの書き込みが完了しました。（ワークシート名: {slack_channel_name}）",
                    thread_ts=thread_ts
                )
            except Exception as e:
                import traceback

                logger.error(f"スプレッドシートへの書き込みでエラーが発生: {e}")
                traceback.print_exc()
                logger.exception("スプレッドシートへの書き込みでエラーの詳細スタックトレース")

                say(
                    text=f"スプレッドシートへの書き込みでエラーが発生しました: {e}",
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
