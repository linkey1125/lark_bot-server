import os
import re
import json
import requests
import openpyxl
from flask import Flask, request, Response
from dotenv import load_dotenv

load_dotenv()
APP_ID = os.environ["LARK_APP_ID"]
APP_SECRET = os.environ["LARK_APP_SECRET"]

app = Flask(__name__)

def get_tenant_access_token():
    url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    payload = {"app_id": APP_ID, "app_secret": APP_SECRET}
    headers = {"Content-Type": "application/json"}
    res = requests.post(url, headers=headers, json=payload)
    return res.json().get("tenant_access_token", "")

def send_file_to_lark(chat_id, file_path="output.xlsx"):
    token = get_tenant_access_token()
    upload_url = "https://open.larksuite.com/open-apis/im/v1/files"
    headers = {"Authorization": f"Bearer {token}"}
    files = {"file": open(file_path, "rb")}
    data = {"file_name": os.path.basename(file_path), "file_type": "stream"}
    res = requests.post(upload_url, headers=headers, data=data, files=files)

    res_data = res.json()
    if "data" not in res_data:
        print("ファイルアップロード失敗:", res_data)
        return

    file_key = res_data["data"]["file_key"]

    send_url = "https://open.larksuite.com/open-apis/im/v1/messages?receive_id_type=chat_id"
    payload = {
        "receive_id": chat_id,
        "msg_type": "file",
        "content": json.dumps({"file_key": file_key})
    }
    headers.update({"Content-Type": "application/json"})
    requests.post(send_url, headers=headers, json=payload)

def split_email_into_blocks(text):
    blocks = re.split(r'\n(?=(?:【)?案件名|案件概要|募集案件|プロジェクト名|案件情報)', text.strip(), flags=re.IGNORECASE)
    return [block.strip() for block in blocks if len(block.strip()) > 10]


def parse_email_block(text):
    def search(patterns, text, default='未記入'):
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match and match.group(1):
                return match.group(1).strip()
        return default

    return {
        '案件名': search([
            r'(?:【)?案件名(?:】)?[:：]?\s*(.*)',
            r'(?:【)?案件概要(?:】)?[:：]?\s*(.*)',
            r'(?:【)?案件情報(?:】)?[:：]?\s*(.*)',
            r'(?:【)?案件タイトル(?:】)?[:：]?\s*(.*)'
        ], text),
        '作業内容': search([
            r'(?:【)?作業内容(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)',
            r'(?:【)?業務内容(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)',
            r'(?:【)?工程(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)'
        ], text),
        '募集要件': search([
            r'(?:【)?必須スキル(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)',
            r'(?:【)?スキル(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)',
            r'(?:【)?募集要件(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)'
        ], text),
        '募集人数': search([
            r'(?:【)?募集人数(?:】)?[:：]?\s*(.*)',
            r'(?:【)?人数(?:】)?[:：]?\s*(.*)',
            r'(?:【)?募集枠(?:】)?[:：]?\s*(.*)'
        ], text),
        '期間': search([
            r'(?:【)?期間(?:】)?[:：]?\s*(.*)',
            r'(?:【)?作業期間(?:】)?[:：]?\s*(.*)',
            r'(?:【)?時期(?:】)?[:：]?\s*(.*)'
        ], text),
        '勤務場所': search([
            r'(?:【)?勤務地(?:】)?[:：]?\s*(.*)',
            r'(?:【)?場所(?:】)?[:：]?\s*(.*)'
        ], text),
        'その他': search([
            r'(?:【)?勤務時間(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)',
            r'(?:【)?備考(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)',
            r'(?:【)?その他(?:】)?[:：]?\s*([\s\S]*?)(?=\n\S|$)'
        ], text)
    }

def export_all_to_excel(data_list, filename="output.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['案件名', '作業内容', '募集要件', '募集人数', '期間', '勤務場所', 'その他'])

    filtered_blocks = [block for block in data_list if any(value != "未記入" for value in block.values())]

    for data in filtered_blocks:
        ws.append([
            data['案件名'],
            data['作業内容'],
            data['募集要件'],
            data['募集人数'],
            data['期間'],
            data['勤務場所'],
            data['その他']
        ])

    wb.save(filename)

@app.route('/webhook', methods=['POST'])
def lark_webhook():
    data = request.json
    print("Webhook受信データ:", json.dumps(data, indent=4, ensure_ascii=False))

    if 'challenge' in data:
        return Response(json.dumps({'challenge': data['challenge']}), status=200, mimetype='application/json')

    try:
        msg_type = data.get('event', {}).get('message', {}).get('message_type')
        if msg_type != 'text':
            print(f"無視されたメッセージタイプ: {msg_type}")
            return Response(json.dumps({'status': 'ignored'}), status=200)

        message = data['event']['message']
        content_str = message.get('content', '{}')
        content_dict = json.loads(content_str)
        email_text = content_dict.get('text', '')

    except Exception as e:
        print(f"メッセージ抽出エラー: {str(e)}")
        return Response(json.dumps({'error': str(e)}), status=400)

    blocks = split_email_into_blocks(email_text)
    parsed_blocks = [parse_email_block(block) for block in blocks]

    for i, block in enumerate(parsed_blocks):
        print(f"\n--- Block {i+1} ---")
        print(json.dumps(block, ensure_ascii=False, indent=2))

    filtered_blocks = [block for block in parsed_blocks if any(value != "未記入" for value in block.values())]
    export_all_to_excel(filtered_blocks)

    chat_id = message.get("chat_id") or message.get("conversation_id")
    send_file_to_lark(chat_id)

    return Response(json.dumps({'status': 'success'}), status=200)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
