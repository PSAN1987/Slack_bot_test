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
SERVICE_ACCOUNT_FILE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # Secret Filesパス
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")               # スプレッドシートのID

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
# チャンネルごとのワークシート取得 or 作成
# -----------------------
def get_or_create_worksheet(sh, worksheet_title: str):
    """
    スプレッドシート sh の中で、worksheet_title に対応するワークシートを取得。
    なければ新規作成して返す。
    さらに、1行目が空なら日本語の見出し行を書き込む。
    """
    try:
        worksheet = sh.worksheet(worksheet_title)
    except WorksheetNotFound:
        sh.add_worksheet(title=worksheet_title, rows=100, cols=20)
        worksheet = sh.worksheet(worksheet_title)

    # 1行目に見出しが書かれていなければ書き込む（日本語の列名）
    existing_rows = len(worksheet.get_all_values())
    if existing_rows < 1:
        # 変数名を1行目A列から書き出し（日本語）
        header = [
            "医院名",         # hospital_name
            "媒体名",         # media_name
            "氏名",           # name
            "会員番号",       # member_id
            "年齢",           # age
            "職種",           # job
            "経験",           # experience
            "お住まい",       # address
            "就業状況",       # status
            "資格",           # cert
            "最終学歴"        # education
        ]
        worksheet.append_row(header, value_input_option="USER_ENTERED")

    return worksheet


# -----------------------
# OpenAI を用いて情報抽出
#   ※ ()の中身を省略せず抽出
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    OpenAI API (ChatCompletion) を使って、応募メッセージから各種情報を抽出する。
    抽出すべき項目: name, member_id, age, job, experience, address, status, cert, education
    ()の中身も省略せず全て出力してもらうようプロンプトを補足。
    """
    if not OPENAI_API_KEY:
        return {}

    system_prompt = (
        "あなたはテキストから以下の情報を抽出するアシスタントです。\n"
        "抽出すべき項目: name(氏名), member_id(会員番号), age(年齢), "
        "job(職種), experience(経験), address(お住まい), status(就業状況), "
        "cert(資格), education(最終学歴)\n"
        "カッコ（）の中身も省略せずにすべて出力してください。ただし年齢の場合は歳を除いて数字にみにしてください\n"
        "出力は必ず JSON 形式のみで、キー名は上記の英語でお願いします。\n"
        "値が不明の場合は空文字にしてください。"
    )

    user_prompt = (
        f"以下のテキストから必要項目を抜き出して、JSON形式で返してください。\n"
        f"丸カッコや波カッコの中身も省略しないでください。\n\n"
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
    - channel_name というタイトルのワークシートを取得 or 作成して書き込み
    - append_rowでA列から書き出し、既存データは消さずに追加
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    worksheet = get_or_create_worksheet(sh, channel_name)

    # A列から順にデータを並べる
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
#   ※ チャンネル名が取得できない場合は channel_id を使う
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
                # チャンネル名が取得できない場合
                slack_channel_name = channel_id
        except Exception as e:
            logger.error(f"チャンネル名の取得に失敗: {e}")
            slack_channel_name = channel_id

        # 必要なら病院名や媒体名などをここで抽出（例: extract_hospital_name, extract_media_name）
        # ここでは例としてダミーのキーを使う
        hospital_name = re.search(r"【([^】]+)】", text)
        if hospital_name:
            raw_name = hospital_name.group(1)
            # 末尾が「様」なら削る
            if raw_name.endswith("様"):
                raw_name = raw_name[:-1]
        else:
            raw_name = ""

        media_match = re.search(r"(.+?)より(.+?)の応募がございました。", text)
        if media_match:
            media_name = media_match.group(1).strip()
        else:
            media_name = ""

        # OpenAI でプロフィール情報抽出
        parsed_profile = parse_profile_info(text)

        # マージ
        merged_data = {
            "hospital_name": raw_name,
            "media_name": media_name,
            **parsed_profile
        }

        # 書き込み
        if merged_data["name"] or merged_data["member_id"]:
            try:
                write_to_spreadsheet(merged_data, slack_channel_name)
                logger.info("スプレッドシートへの書き込みに成功しました。")

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
