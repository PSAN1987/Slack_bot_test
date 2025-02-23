import os
import re
import json
import openai
import gspread

from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from oauth2client.service_account import ServiceAccountCredentials

# 新規追加: ワークシートが無いときに発生する例外を扱うため
from gspread.exceptions import WorksheetNotFound

import datetime  # 日本時間への変換等に使うため

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
# 「【○○○様】」から医院名を抜き出す (現行コードを同一)
# -----------------------
def extract_hospital_name(text: str) -> str:
    pattern = r"【([^】]+)】"  # 「【...】」の中身を取得
    match = re.search(pattern, text)
    if not match:
        return ""
    raw_name = match.group(1).strip()
    # 末尾が「様」であれば削除
    if raw_name.endswith("様"):
        raw_name = raw_name[:-1]
    return raw_name

# -----------------------
# 「○○○よりXXXXの応募がございました。」から媒体名を抜き出す (現行コードを同一)
# -----------------------
def extract_media_name(text: str) -> str:
    pattern = r"(.+?)より(.+?)応募がございました。"
    match = re.search(pattern, text)
    if not match:
        return ""
    media_name = match.group(1).strip()
    return media_name

# -----------------------
# OpenAIを用いてプロフィール情報を抽出 (現行コードを同一)
# -----------------------
def parse_profile_info(text: str) -> dict:
    if not OPENAI_API_KEY:
        return {}

    system_prompt = (
        "あなたはテキストから以下の情報を抽出するアシスタントです。\n"
        "抽出すべき項目: name(氏名), member_id(会員番号), age(年齢), "
        "job(職種), experience(経験), address(お住まい), status(就業状況), "
        "cert(資格), education(最終学歴), spot_dates(スポット希望日)\n"
        "職種(job)には括弧内の情報(例: (正社員))も含めてください。\n"
        "経験(experience)には職歴情報をすべて文字列としてまとめてください。\n"
        "spot_dates(スポット希望日)があれば、複数日でも1つの文字列にまとめてください。\n"
        "年齢(age)は「歳」を除いて数字のみ出力してください。\n"
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

        # 年齢は数字のみ残す
        age_str = extracted_data.get("age", "")
        extracted_data["age"] = re.sub(r"\D", "", age_str)

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
            "spot_dates": extracted_data.get("spot_dates", ""),
        }
    except Exception as e:
        print(f"OpenAI API error: {e}")
        return {}

# ============================
# 追加: ヘッダを常に確認・補正
# ============================
def ensure_header(worksheet):
    """
    ワークシートの1行目には書き込まない。
    A1にSubtotal関数、A2:M2にヘッダを書き込む。
    """
    expected_header = [
        "hospital_name",
        "slack_timestamp",
        "media_name",
        "name",
        "member_id",
        "age",
        "job",
        "experience",
        "address",
        "status",
        "cert",
        "education",
        "spot_dates",
    ]
    
    # A1: Subtotal 関数
    worksheet.update_acell('A1', '=SUBTOTAL(103,A3:A1000)')

    # A2:M2 にヘッダを書き込む (常に上書き)
    worksheet.update('A2:M2', [expected_header])

# ============================
# 追加: チャンネル用ワークシートを取得 or 作成する関数
# ============================
def get_or_create_worksheet(sh, sheet_title: str):
    """
    sheet_title に一致するワークシートを探す。
    なければ新規作成するが、1行目には書かない。
    """
    try:
        worksheet = sh.worksheet(sheet_title)
    except WorksheetNotFound:
        # 新規作成 (行数や列数は必要に応じて拡張)
        worksheet = sh.add_worksheet(title=sheet_title, rows=100, cols=35)

    # ここでヘッダとSUBTOTAL関数をセット
    ensure_header(worksheet)
    return worksheet

# -----------------------
# スプレッドシートへ書き込み (現行のwrite_to_spreadsheetを置き換え)
# -----------------------
def write_to_spreadsheet(data: dict):
    """
    ディクショナリ data の内容をスプレッドシートに1行追加する。
    A2にヘッダ、A3以降にデータ。
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)

    channel_name = data.get("channel_name", "UnknownChannel")
    worksheet = get_or_create_worksheet(sh, channel_name)

    new_row = [
        data.get("hospital_name", ""),
        data.get("slack_timestamp", ""),
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
        data.get("spot_dates", ""),
    ]
    worksheet.append_row(new_row, value_input_option="USER_ENTERED")

# -----------------------
# Slack Bolt: メッセージイベントのハンドラ (現行コード同一)
# -----------------------
@app_bolt.event("message")
def handle_message_events(body, say, logger):
    event = body.get("event", {})
    text = event.get("text", "")
    thread_ts = event.get("ts")
    channel_id = event.get("channel")

    if "応募がございました。" in text:
        hospital_name = extract_hospital_name(text)
        media_name = extract_media_name(text)
        parsed_profile = parse_profile_info(text)

        slack_timestamp_str = ""
        if thread_ts:
            try:
                dt = datetime.datetime.fromtimestamp(float(thread_ts))
                dt_jst = dt + datetime.timedelta(hours=9)
                slack_timestamp_str = dt_jst.strftime("%Y-%m-%d")
            except:
                pass

        merged_data = {
            "hospital_name": hospital_name,
            "slack_timestamp": slack_timestamp_str,
            "media_name": media_name,
            **parsed_profile
        }

        try:
            channel_info = app_bolt.client.conversations_info(channel=channel_id)
            slack_channel_name = channel_info["channel"]["name"]
        except Exception as e:
            logger.error(f"チャンネル名の取得に失敗: {e}")
            slack_channel_name = f"UnknownChannel_{channel_id}"

        merged_data["channel_name"] = slack_channel_name

        if merged_data["name"] or merged_data["member_id"]:
            try:
                write_to_spreadsheet(merged_data)
                logger.info("スプレッドシートへの書き込みに成功しました。")
                say(text="スプレッドシート書き込みが完了しました。", thread_ts=thread_ts)
            except Exception as e:
                import traceback
                logger.error(f"スプレッドシートへの書き込みでエラーが発生: {e}")
                traceback.print_exc()
                logger.exception("スプレッドシートへの書き込みでエラーの詳細スタックトレース")
                say(text=f"スプレッドシートへの書き込みでエラー: {e}", thread_ts=thread_ts)

# -----------------------
# Flaskルート設定 (現行コード同一)
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
