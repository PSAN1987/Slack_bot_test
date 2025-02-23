import os
import re
import json

from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ====== New: openai ライブラリ ======
import openai

# ============================
# Slack の認証情報 (環境変数)
# ============================
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ["SLACK_SIGNING_SECRET"]

# ============================
# Google認証 (Secret Filesを利用)
# ============================
SERVICE_ACCOUNT_FILE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # Secret Filesパス
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
# テキスト解析用の正規表現
# -----------------------
re_name         = re.compile(r"・氏名：([^（\n]+)")  # （）が含まれない部分を取得
re_member_id    = re.compile(r"・会員番号：(\S+)")
re_age          = re.compile(r"・年齢：(\d+)歳")
re_job          = re.compile(r"・職種：(.+)")
re_experience   = re.compile(r"・経験：(.+)")
re_address      = re.compile(r"・お住まい：(.+)")
re_status       = re.compile(r"・就業状況：(.+)")
re_cert         = re.compile(r"・資格：(.+)")
re_education    = re.compile(r"・最終学歴：(.+)")

def parse_profile_info_by_regex(text: str) -> dict:
    """
    既存の正規表現を使った抽出
    """
    data = {
        "name":        "",
        "member_id":   "",
        "age":         "",
        "job":         "",
        "experience":  "",
        "address":     "",
        "status":      "",
        "cert":        "",
        "education":   "",
    }

    m_name       = re_name.search(text)
    m_member_id  = re_member_id.search(text)
    m_age        = re_age.search(text)
    m_job        = re_job.search(text)
    m_experience = re_experience.search(text)
    m_address    = re_address.search(text)
    m_status     = re_status.search(text)
    m_cert       = re_cert.search(text)
    m_education  = re_education.search(text)

    if m_name:
        data["name"] = m_name.group(1).strip()
    if m_member_id:
        data["member_id"] = m_member_id.group(1).strip()
    if m_age:
        data["age"] = m_age.group(1).strip()
    if m_job:
        data["job"] = m_job.group(1).strip()
    if m_experience:
        data["experience"] = m_experience.group(1).strip()
    if m_address:
        data["address"] = m_address.group(1).strip()
    if m_status:
        data["status"] = m_status.group(1).strip()
    if m_cert:
        data["cert"] = m_cert.group(1).strip()
    if m_education:
        data["education"] = m_education.group(1).strip()

    return data

# -----------------------
# OpenAIを用いて情報抽出する処理
# -----------------------
def parse_profile_info_by_openai(text: str) -> dict:
    """
    OpenAI API (ChatCompletion) を使って、応募メッセージから各種情報を抽出する。
    サンプルでは JSON 形式で出力している前提。
    """
    if not OPENAI_API_KEY:
        # APIキーがない場合はスキップ
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

        result = {
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
        return result

    except Exception as e:
        print(f"OpenAI API error: {e}")
        return {}

# -----------------------
# まとめて情報抽出
# -----------------------
def parse_profile_info(text: str) -> dict:
    """
    Slackメッセージ本文からプロフィール情報を抽出する。
    - OpenAIで解析した結果を得る
    - 既存の正規表現で解析した結果を得る
    - 双方をマージし、最終的な情報を返す
    """
    data_ai = parse_profile_info_by_openai(text)
    data_re = parse_profile_info_by_regex(text)

    final_data = {}
    for key in ["name", "member_id", "age", "job", "experience", "address", "status", "cert", "education"]:
        if data_ai.get(key):
            final_data[key] = data_ai[key]
        else:
            final_data[key] = data_re[key]

    return final_data

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
    # メッセージのタイムスタンプ (返信のthread_tsに使う)
    thread_ts = event.get("ts")

    # 「ジョブメドレーより〇〇の応募がございました。」を含むかチェック
    if "ジョブメドレーより" in text and "の応募がございました" in text:
        # プロフィール情報を抽出
        parsed_data = parse_profile_info(text)

        # 1つでも値が取れた場合のみスプレッドシート書き込み
        if parsed_data["name"] or parsed_data["member_id"]:
            try:
                write_to_spreadsheet(parsed_data)
                logger.info("スプレッドシートへの書き込みに成功しました。")

                # 書き込みが完了したら返信(同じメッセージのスレッド上に通知)
                say(
                    text="スプレッドシート書き込みが完了しました。",
                    thread_ts=thread_ts
                )
            except Exception as e:
                logger.error(f"スプレッドシートへの書き込みでエラー: {e}")
                # 失敗時にも返信する場合
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
