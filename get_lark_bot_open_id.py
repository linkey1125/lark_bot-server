import requests
import os
from dotenv import load_dotenv

# ✅ 環境変数の読み込み
load_dotenv()
APP_ID = os.getenv("LARK_APP_ID")
APP_SECRET = os.getenv("LARK_APP_SECRET")

# ✅ Larkのアクセストークン取得
def get_tenant_access_token():
    url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    payload = {"app_id": APP_ID, "app_secret": APP_SECRET}
    headers = {"Content-Type": "application/json"}
    res = requests.post(url, headers=headers, json=payload)
    return res.json().get("tenant_access_token", "")

# ✅ Botの情報を取得
def get_lark_bot_open_id():
    access_token = get_tenant_access_token()
    url = "https://open.larksuite.com/open-apis/bot/v3/info"
    headers = {"Authorization": f"Bearer {access_token}"}

    res = requests.get(url, headers=headers)
    bot_info = res.json()
    
    # ✅ `open_id` を取得する処理を追加
    return bot_info.get("bot", {}).get("open_id")



if __name__ == "__main__":
    bot_open_id = get_lark_bot_open_id()
    print(f"Lark Botの OPEN_ID: {bot_open_id}")
