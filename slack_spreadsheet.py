
import os
import re
import json

from flask import Flask, request
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ============================
# Slack �̔F�؏�� (���ϐ�)
# ============================
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ["SLACK_SIGNING_SECRET"]

# ============================
# Google�F��
# ============================
# Render.com�ȂǂɃf�v���C����Ƃ��́A
# JSON�𒼐ڊ��ϐ��ɖ��ߍ��ނ��ABase64��t�@�C���p�X�ɂ���ȂǍH�v���Ă�������
# �����ł̓V���v���Ƀt�@�C���Ƃ��ēǂݍ��ރp�^�[���͏ȗ�
SERVICE_ACCOUNT_INFO = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")  # JSON������
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")  # �X�v���b�h�V�[�g��ID

# Slack Bolt �A�v����������
app_bolt = App(
    token=SLACK_BOT_TOKEN,
    signing_secret=SLACK_SIGNING_SECRET
)

# Flask �A�v�������iBolt�̃C�x���g���󂯎��p�j
flask_app = Flask(__name__)
handler = SlackRequestHandler(app=app_bolt)

# -----------------------
# Google Sheets �N���C�A���g�̏�����
# -----------------------
def get_gspread_client():
    """
    ���ϐ�����T�[�r�X�A�J�E���gJSON��ǂݍ��݁A
    Google�X�v���b�h�V�[�g�ɃA�N�Z�X�\�ȃN���C�A���g��Ԃ�
    """
    if SERVICE_ACCOUNT_INFO is None:
        raise ValueError("���ϐ� GCP_SERVICE_ACCOUNT_JSON ���ݒ肳��Ă��܂���")

    service_account_dict = json.loads(SERVICE_ACCOUNT_INFO)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_dict, scope)
    gc = gspread.authorize(credentials)
    return gc


# -----------------------
# �e�L�X�g��͗p�̐��K�\��
# -----------------------
# ��:
# �u�E�����F�V�� �r�Y�i�䂤�ǂ� �Ƃ����j�搶�v����u�V�� �r�Y�v�𔲂��o�� ��
# �i�����̊��ʂ�搶�̕������ǂ���舵�����͉^�p�ɍ��킹�Ē������Ă��������j
re_name         = re.compile(r"�E�����F([^�i\n]+)")     # �i�j���܂܂�Ȃ��������擾
re_member_id    = re.compile(r"�E����ԍ��F(\S+)")
re_age          = re.compile(r"�E�N��F(\d+)��")
re_job          = re.compile(r"�E�E��F(.+)")
re_experience   = re.compile(r"�E�o���F(.+)")
re_address      = re.compile(r"�E���Z�܂��F(.+)")
re_status       = re.compile(r"�E�A�Ə󋵁F(.+)")
re_cert         = re.compile(r"�E���i�F(.+)")
re_education    = re.compile(r"�E�ŏI�w���F(.+)")

# -----------------------
# ���b�Z�[�W�̉�͏���
# -----------------------
def parse_profile_info(text: str):
    """
    Slack���b�Z�[�W�{������v���t�B�[�����𒊏o����
    �߂�l: dict (�L�[: 'name', 'member_id', 'age', 'job', 'experience', 'address', 'status', 'cert', 'education')
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

    # ���K�\���Œ��o
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
        data["age"] = m_age.group(1).strip()  # "30" �Ƃ������N���
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
# �X�v���b�h�V�[�g�֏�������
# -----------------------
def write_to_spreadsheet(profile_data: dict):
    """
    �X�v���b�h�V�[�g��1�s�ǉ�����
    """
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    # �V�[�g���͉^�p����ŕύX���Ă�������
    worksheet = sh.worksheet("Sheet1")

    # �����ł͒P���� append �����
    # �J�������͕K�v�ɉ����Ē������Ă�������
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
# Slack Bolt: ���b�Z�[�W�C�x���g�̃n���h��
# -----------------------
@app_bolt.event("message")
def handle_message_events(body, say, logger):
    """
    �`�����l���ւ̐V�K���b�Z�[�W�C�x���g���󂯎��
    """
    event = body.get("event", {})
    text = event.get("text", "")

    # �u�W���u���h���[���Z�Z�̉��傪�������܂����B�v���܂ނ��`�F�b�N
    if "�W���u���h���[���" in text and "�̉��傪�������܂���" in text:
        # �v���t�B�[�����𒊏o
        parsed_data = parse_profile_info(text)
        # ��������ꂽ���A��������v���t�B�[�����ڂ�����ꍇ�̂݃X�v���b�h�V�[�g��������
        if parsed_data["name"] or parsed_data["member_id"]:
            try:
                write_to_spreadsheet(parsed_data)
                logger.info("�X�v���b�h�V�[�g�ւ̏������݂ɐ������܂����B")
            except Exception as e:
                logger.error(f"�X�v���b�h�V�[�g�ւ̏������݂ŃG���[: {e}")

    # �����ł�say���ŕԐM�͂��Ȃ����A�K�v�ɉ����ĉ������b�Z�[�W�𑗐M���Ă��ǂ�


# -----------------------
# Flask���[�g�ݒ�
# -----------------------
@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

# �w���X�`�F�b�N�p�̃G���h�|�C���g�Ȃ�
@flask_app.route("/", methods=["GET"])
def healthcheck():
    return "OK", 200

# -----------------------
# �A�v���N�� (Render�ł�Gunicorn�^�p��z��)
# -----------------------
# Render�Ȃǂł� gunicorn �R�}���h�ŋN������z��ł���A
# python main.py �Œ��ڋN������ꍇ�ɂ͈ȉ��� if __name__ == "__main__": ���K�v�B
if __name__ == "__main__":
    flask_app.run(host="0.0.0.0", port=5000)
