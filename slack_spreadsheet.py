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

# ★追加：日本時間への変換に使うため
import datetime

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

    # ★追加：スポット希望日(spot_dates)も含めて抽出させるよう指示
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

        # 年齢は数字のみ残す（万が一 GPT が「32歳」のように返した時の対策）
        age_str = extracted_data.get("age", "")
        extracted_data["age"] = re.sub(r"\D", "", age_str)  # 数字以外を除去

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
            # ★追加: spot_dates (スポット希望日)
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
    ワークシートの1行目を確認し、期待するヘッダと異なる場合は上書きする。
    """
    # ★修正: slack_timestamp を hospital_name と media_name の間に追加
    expected_header = [
        "hospital_name",
        "slack_timestamp",  # ★追加列
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
        "spot_dates",  # 新規追加列
    ]
    
    current_header = worksheet.row_values(1)
    
    if current_header != expected_header:
        worksheet.update('A1:M1', [expected_header])

# ============================
# 追加: チャンネル用ワークシートを取得 or 作成する関数
# ============================
def get_or_create_worksheet(sh, sheet_title: str):
    """
    sheet_title に一致するワークシートを探す。
    なければ新規作成し、1行目(A列から)にヘッダ（変数名）を書き込む。
    """
    try:
        worksheet = sh.worksheet(sheet_title)
        newly_created = False
    except WorksheetNotFound:
        # 新規作成 (行数や列数は必要に応じて拡張)
        worksheet = sh.add_worksheet(title=sheet_title, rows=100, cols=35)
        newly_created = True

    # 新規作成された場合、1行目に変数名ヘッダーを追加 (A列から)
    if newly_created:
        header = [
            "hospital_name",
            "slack_timestamp",  # ★追加列
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
            "spot_dates",  # 新規追加列
        ]
        worksheet.append_row(header, value_input_option="USER_ENTERED")

    # ヘッダを確認・補正
    ensure_header(worksheet)

    return worksheet

# -----------------------
# スプレッドシートへ書き込み (現行のwrite_to_spreadsheetを置き換え)
# -----------------------
def write_to_spreadsheet(data: dict):
    """
    ディクショナリ data の内容をスプレッドシートに1行追加する。
    チャンネル名を読み出し、そのチャンネル名のワークシートを取得 or 作成して書き込む。
    1行目(A列)にはヘッダー、既存データは消さずに追加。
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)

    # dataの中に channel_name が含まれている想定
    channel_name = data.get("channel_name", "UnknownChannel")

    # ワークシート取得 or 新規作成
    worksheet = get_or_create_worksheet(sh, channel_name)

    # ★修正: slack_timestamp 列を hospital_name と media_name の間へ
    new_row = [
        data.get("hospital_name", ""),
        data.get("slack_timestamp", ""),  # ★追加
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

        # ★追加: Slackのtsを日時文字列に変換
        slack_timestamp_str = ""
        if thread_ts:
            try:
                # UTCでの日時を取得
                dt = datetime.datetime.fromtimestamp(float(thread_ts))
                # JSTへ変換
                dt_jst = dt + datetime.timedelta(hours=9)
                # ★ここで「年-月-日」のみの文字列
                slack_timestamp_str = dt_jst.strftime("%Y-%m-%d")
            except:
                pass

        # まとめたdictに、病院名・媒体名を含める
        merged_data = {
            "hospital_name": hospital_name,
            "slack_timestamp": slack_timestamp_str,  # ★追加
            "media_name": media_name,
            **parsed_profile
        }

        # 追加: チャンネル名を取得して merged_data に格納
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
