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
import time      # リトライ時のスリープに使用（必要に応じて）

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
    if raw_name.endswith("様"):
        raw_name = raw_name[:-1]
    return raw_name

# -----------------------
# 「○○○よりXXXXの応募がございました。」から媒体名を抜き出す (現行コードを同一)
# -----------------------
def extract_media_name(text: str) -> str:
    """
    OOOより ◯◯◯(応募|見学希望)がございました。
    の場合に OOO をメディア名として取り出す
    """
    # 「より」以前をグループ1として取得
    # その後の任意の文字列の中に 「応募」 または 「見学希望」 があり、
    # 最後に「がございました。」で終わるパターンにマッチ
    pattern = r"(.+?)より.+?(応募|見学希望)がございました。"
    match = re.search(pattern, text)
    if not match:
        return ""
    # group(1) が "より" より前、つまりメディア名
    media_name = match.group(1).strip()
    return media_name

# ============================
# 必須キーを空文字で用意した安全なdictを返す関数
# ============================
def empty_profile_dict() -> dict:
    """
    必要なキーをすべて持つ空文字入りディクショナリを返す。
    """
    return {
        "name": "",
        "member_id": "",
        "age": "",
        "job": "",
        "experience": "",
        "address": "",
        "status": "",
        "cert": "",
        "education": "",
        "spot_dates": "",
    }

# -----------------------
# OpenAIを用いてプロフィール情報を抽出 (堅牢版)
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    GPT呼び出しを最大2回リトライし、それでも失敗時は必須キーを含む空dictを返す。
    """
    # 1. OpenAIキーが無い場合は空辞書を返す
    if not OPENAI_API_KEY:
        return empty_profile_dict()

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
    user_prompt = f"以下のテキストから必要項目を抜き出して、JSON形式で返してください。\n\n{text}\n"

    # 2. リトライロジック
    for attempt in range(2):  # 最大2回
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

            # 年齢は数字のみ
            age_str = extracted_data.get("age", "")
            extracted_data["age"] = re.sub(r"\D", "", age_str)

            # 必須キー以外は無視。必須キーがあれば取り出し、無ければ空文字
            safe_data = empty_profile_dict()
            for k in safe_data.keys():
                safe_data[k] = extracted_data.get(k, "")

            return safe_data

        except Exception as e:
            print(f"OpenAI API error (attempt {attempt+1}): {e}")
            # 一時的なエラーかもしれないので短いスリープ
            time.sleep(1.0)

    # 3. 2回とも失敗なら、必須キーだけ空で返す
    return empty_profile_dict()

# ============================
# 追加: ヘッダを常に確認・補正
# ============================
def ensure_header(worksheet):
    """
    1行目は =SUBTOTAL(103,A3:A1000)、
    2行目(A2:M2) はヘッダを書き込む
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

    worksheet.update_acell('A1', '=SUBTOTAL(103,A3:A1000)')
    worksheet.update('A2:M2', [expected_header])

# ============================
# 追加: チャンネル用ワークシートを取得 or 作成する関数
# ============================
def get_or_create_worksheet(sh, sheet_title: str):
    """
    sheet_title に一致するワークシートを探し。
    なければ新規作成し、A1=SUBTOTAL, A2=ヘッダをセット。
    """
    try:
        worksheet = sh.worksheet(sheet_title)
    except WorksheetNotFound:
        worksheet = sh.add_worksheet(title=sheet_title, rows=100, cols=35)

    ensure_header(worksheet)
    return worksheet

# -----------------------
# スプレッドシートへ書き込み (現行のwrite_to_spreadsheetを置き換え)
# -----------------------
def write_to_spreadsheet(data: dict):
    """
    ディクショナリ data の内容をスプレッドシートに1行追加する。
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
import datetime
def handle_message_events(body, say, logger):
    event = body.get("event", {})
    text = event.get("text", "")
    thread_ts = event.get("ts")
    channel_id = event.get("channel")

    # ★「応募...がございました。」または「見学希望がございました。」なら処理
    if ("がございました。" in text) and ("応募" in text or "見学希望" in text):
        hospital_name = extract_hospital_name(text)
        media_name = extract_media_name(text)

        # 堅牢版 parse_profile_info (必ず必須キーを含むdictが返る)
        parsed_profile = parse_profile_info(text)

        # Slackのtsを日時文字列(YYYY-MM-DD)に変換
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

        # チャンネル名取得 (権限不足の場合は例外発生)
        try:
            channel_info = app_bolt.client.conversations_info(channel=channel_id)
            slack_channel_name = channel_info["channel"]["name"]
        except Exception as e:
            logger.error(f"チャンネル名の取得に失敗: {e}")
            slack_channel_name = f"UnknownChannel_{channel_id}"

        merged_data["channel_name"] = slack_channel_name

        if merged_data.get("name") or merged_data.get("member_id"):
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
