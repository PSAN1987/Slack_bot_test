import os
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

    # ファイルを読み込んでJSONパース
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
# OpenAIを用いて情報抽出する処理
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    OpenAI API (ChatCompletion) を使って、応募メッセージから各種情報を抽出する。
    抽出すべき項目: name, member_id, age, job, experience, address, status, cert, education
    """
    # APIキーが無い場合は空dictを返す
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
            model="gpt-3.5-turbo",  # 必要に応じて gpt-4 等
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
        )
        content = response["choices"][0]["message"]["content"].strip()
        extracted_data = json.loads(content)

        # 必要なキーをセット（無い場合は空文字）
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
def write_to_spreadsheet(profile_data: dict):
    """
    スプレッドシートに1行追加する
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    worksheet = sh.worksheet("Sheet1")

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
    # メッセージのタイムスタンプ (返信のthread_tsに使用)
    thread_ts = event.get("ts")

    # 「ジョブメドレーより〇〇の応募がございました。」を含むかチェック
    if "ジョブメドレーより" in text and "の応募がございました" in text:
        # --- 1) OpenAI で情報を抽出 ---
        parsed_data = parse_profile_info(text)

        # --- 2) スプレッドシートに書き込み ---
        if parsed_data["name"] or parsed_data["member_id"]:
            try:
                write_to_spreadsheet(parsed_data)
                logger.info("スプレッドシートへの書き込みに成功しました。")

                # 書き込み完了したらスレッドに返事
                say(
                    text="スプレッドシート書き込みが完了しました。",
                    thread_ts=thread_ts
                )

            except Exception as e:
                import traceback

                # ① 例外メッセージを含むログを出す
                logger.error(f"スプレッドシートへの書き込みでエラーが発生: {e}")
                # ② スタックトレースをコンソールに出力（Renderのログに表示される）
                traceback.print_exc()
                # ③ logger.exception() でもスタックトレースを出す（任意）
                logger.exception("スプレッドシートへの書き込みでエラーの詳細スタックトレース")
                # Slackへの通知 (任意)
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
