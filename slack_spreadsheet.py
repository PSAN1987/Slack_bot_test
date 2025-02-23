
import os
import re
import json

from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ============================
# Slack の認証情報 (環境変数)
# ============================
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ["SLACK_SIGNING_SECRET"]

# ============================
# Google認証
# ============================
# Render.comなどにデプロイするときは、
# JSONを直接環境変数に埋め込むか、Base64やファイルパスにするなど工夫してください
# ここではシンプルにファイルとして読み込むパターンは省略
SERVICE_ACCOUNT_INFO = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # JSON文字列
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")  # スプレッドシートのID

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
    環境変数からサービスアカウントJSONを読み込み、
    Googleスプレッドシートにアクセス可能なクライアントを返す
    """
    if SERVICE_ACCOUNT_INFO is None:
        raise ValueError("環境変数 GCP_SERVICE_ACCOUNT_JSON が設定されていません")

    service_account_dict = json.loads(SERVICE_ACCOUNT_INFO)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_dict, scope)
    gc = gspread.authorize(credentials)
    return gc


# -----------------------
# テキスト解析用の正規表現
# -----------------------
# 例:
# 「・氏名：遊道 俊雄（ゆうどう としお）先生」から「遊道 俊雄」を抜き出す 等
# （氏名の括弧や先生の部分をどう取り扱うかは運用に合わせて調整してください）
re_name         = re.compile(r"・氏名：([^（\n]+)")     # （）が含まれない部分を取得
re_member_id    = re.compile(r"・会員番号：(\S+)")
re_age          = re.compile(r"・年齢：(\d+)歳")
re_job          = re.compile(r"・職種：(.+)")
re_experience   = re.compile(r"・経験：(.+)")
re_address      = re.compile(r"・お住まい：(.+)")
re_status       = re.compile(r"・就業状況：(.+)")
re_cert         = re.compile(r"・資格：(.+)")
re_education    = re.compile(r"・最終学歴：(.+)")

# -----------------------
# メッセージの解析処理
# -----------------------
def parse_profile_info(text: str):
    """
    Slackメッセージ本文からプロフィール情報を抽出する
    戻り値: dict (キー: 'name', 'member_id', 'age', 'job', 'experience', 'address', 'status', 'cert', 'education')
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

    # 正規表現で抽出
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
        data["age"] = m_age.group(1).strip()  # "30" といった年齢数字
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
# スプレッドシートへ書き込み
# -----------------------
def write_to_spreadsheet(profile_data: dict):
    """
    スプレッドシートに1行追加する
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    # シート名は運用次第で変更してください
    worksheet = sh.worksheet("Sheet1")

    # ここでは単純に append する例
    # カラム順は必要に応じて調整してください
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

    # 「ジョブメドレーより〇〇の応募がございました。」を含むかチェック
    if "ジョブメドレーより" in text and "の応募がございました" in text:
        # プロフィール情報を抽出
        parsed_data = parse_profile_info(text)
        # 氏名が取れた等、何かしらプロフィール項目がある場合のみスプレッドシート書き込み
        if parsed_data["name"] or parsed_data["member_id"]:
            try:
                write_to_spreadsheet(parsed_data)
                logger.info("スプレッドシートへの書き込みに成功しました。")
            except Exception as e:
                logger.error(f"スプレッドシートへの書き込みでエラー: {e}")

    # ここではsay等で返信はしないが、必要に応じて応答メッセージを送信しても良い


# -----------------------
# Flaskルート設定
# -----------------------
@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

# ヘルスチェック用のエンドポイントなど
@flask_app.route("/", methods=["GET"])
def healthcheck():
    return "OK", 200

# -----------------------
# アプリ起動 (RenderでのGunicorn運用を想定)
# -----------------------
# Renderなどでは gunicorn コマンドで起動する想定であり、
# python main.py で直接起動する場合には以下の if __name__ == "__main__": が必要。
if __name__ == "__main__":
    flask_app.run(host="0.0.0.0", port=5000)
