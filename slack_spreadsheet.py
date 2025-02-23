import os
import json
import re
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
SERVICE_ACCOUNT_FILE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # Secret Filesのパス
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")               # スプレッドシートID

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
# チャンネルごとのワークシート取得 or 作成
# -----------------------
def get_or_create_worksheet(sh, worksheet_title: str):
    """
    指定ワークシートを取得。なければ新規作成。
    1件もデータがない場合は英語のヘッダ行（A列から）を書き込む。
    """
    try:
        worksheet = sh.worksheet(worksheet_title)
    except WorksheetNotFound:
        sh.add_worksheet(title=worksheet_title, rows=100, cols=20)
        worksheet = sh.worksheet(worksheet_title)

    all_values = worksheet.get_all_values()
    if len(all_values) == 0:
        # データが無い=空ワークシートならヘッダを書き込む（英語の変数名）
        header = [
            "hospital_name",
            "media_name",
            "name",
            "member_id",
            "age",
            "job",
            "experience",
            "address",
            "status",
            "cert",
            "education"
        ]
        worksheet.append_row(header, value_input_option="USER_ENTERED")

    return worksheet

# -----------------------
# OpenAI を用いて情報抽出
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    OpenAIを使い、hospital_name, media_name, name, member_id, age, job, experience,
    address, status, cert, education を抽出。
    - 会員番号 (member_id) はtel形式の場合でも数字のみ抽出
    - ()の中身も省略せず出力
    - 年齢は「歳」を除去して数字のみに
    """
    if not OPENAI_API_KEY:
        return {}

    system_prompt = (
        "You are an assistant that extracts information from the text. "
        "Please extract the following fields in JSON format:\n"
        "  - hospital_name (医院名)\n"
        "  - media_name (媒体名)\n"
        "  - name (氏名, with all parentheses included)\n"
        "  - member_id (会員番号: if it has tel-like format, remove all non-digit chars)\n"
        "  - age (年齢: remove any trailing '歳' and keep only digits)\n"
        "  - job (職種)\n"
        "  - experience (経験)\n"
        "  - address (お住まい)\n"
        "  - status (就業状況)\n"
        "  - cert (資格)\n"
        "  - education (最終学歴)\n\n"
        "If a field is unknown or empty, return an empty string.\n"
        "Do not omit parentheses or content inside them. Output must be valid JSON with the above keys.\n"
        "For member_id, please remove all non-digit characters.\n"
        "For age, please remove the suffix '歳' and only keep digits.\n"
        "Output must be strictly JSON."
    )

    user_prompt = (
        f"Below is the text. Please parse out the fields. \n\n"
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
        data = json.loads(content)

        # post-processing: ensure member_id is digits only, age is digits only
        member_id_raw = data.get("member_id", "")
        # 数字のみ抽出
        member_id_digits = re.sub(r"\D", "", member_id_raw)

        age_raw = data.get("age", "")
        age_digits = re.sub(r"\D", "", age_raw)

        return {
            "hospital_name": data.get("hospital_name", ""),
            "media_name":    data.get("media_name", ""),
            "name":          data.get("name", ""),
            "member_id":     member_id_digits,
            "age":           age_digits,
            "job":           data.get("job", ""),
            "experience":    data.get("experience", ""),
            "address":       data.get("address", ""),
            "status":        data.get("status", ""),
            "cert":          data.get("cert", ""),
            "education":     data.get("education", ""),
        }
    except Exception as e:
        print(f"OpenAI API error: {e}")
        return {}

# -----------------------
# スプレッドシートへ書き込み
# -----------------------
def write_to_spreadsheet(profile_data: dict, channel_name: str):
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    worksheet = get_or_create_worksheet(sh, channel_name)

    new_row = [
        profile_data.get("hospital_name", ""),
        profile_data.get("media_name", ""),
        profile_data.get("name", ""),
        profile_data.get("member_id", ""),
        profile_data.get("age", ""),
        profile_data.get("job", ""),
        profile_data.get("experience", ""),
        profile_data.get("address", ""),
        profile_data.get("status", ""),
        profile_data.get("cert", ""),
        profile_data.get("education", ""),
    ]
    worksheet.append_row(new_row, value_input_option="USER_ENTERED")

# -----------------------
# Slack Bolt: メッセージイベントのハンドラ
# -----------------------
@app_bolt.event("message")
def handle_message_events(body, say, logger):
    event = body.get("event", {})
    text = event.get("text", "")
    thread_ts = event.get("ts")
    channel_id = event.get("channel")

    # 応募がございました。を含むメッセージかチェック
    if "応募がございました。" in text:
        # チャンネル名を取得。取れなければIDを使う
        try:
            channel_info = app_bolt.client.conversations_info(channel=channel_id)
            slack_channel_name = channel_info["channel"].get("name")
            if not slack_channel_name:
                slack_channel_name = channel_id
        except Exception as e:
            logger.error(f"チャンネル名の取得に失敗: {e}")
            slack_channel_name = channel_id

        # --- 1) OpenAIで情報抽出 ---
        parsed_profile = parse_profile_info(text)

        # --- 2) 書き込み ---
        #   name or member_id がある程度は必須と仮定
        if parsed_profile.get("name") or parsed_profile.get("member_id"):
            try:
                write_to_spreadsheet(parsed_profile, slack_channel_name)
                logger.info("スプレッドシートへの書き込みに成功しました。")
                say(
                    text="スプレッドシート書き込みが完了しました。",
                    thread_ts=thread_ts
                )
            except Exception as e:
                import traceback
                logger.error(f"スプレッドシートへの書き込みでエラー: {e}")
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
