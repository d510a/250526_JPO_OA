import requests
import json
import os
import time
import logging
import tkinter as tk
from tkinter import simpledialog
import traceback

# ---------- 設定ファイル ----------
documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
config_file = os.path.join(documents_folder, "proxy_jpo_config.json")

def ask_credentials():
    """
    プロキシ／JPO 資格情報と Token URL を GUI で取得し JSON に保存
    """
    logging.info("===== 資格情報入力開始 =====")
    root = tk.Tk()
    root.withdraw()

    saved = {}
    if os.path.exists(config_file):
        with open(config_file, "r", encoding="utf-8") as f:
            saved = json.load(f)

    try:
        proxy_user = simpledialog.askstring(
            "プロキシユーザ名", "プロキシのユーザ名を入力してください",
            initialvalue=saved.get("proxy_user", "")
        ) or ""
        proxy_pass = simpledialog.askstring(
            "プロキシパスワード", "プロキシのパスワードを入力してください",
            show="*", initialvalue=saved.get("proxy_pass", "")
        ) or ""
        jpo_user = simpledialog.askstring(
            "JPO APIユーザID", "特許庁APIのユーザIDを入力してください",
            initialvalue=saved.get("jpo_user", "")
        ) or ""
        jpo_pass = simpledialog.askstring(
            "JPO APIパスワード", "特許庁APIのパスワードを入力してください",
            show="*", initialvalue=saved.get("jpo_pass", "")
        ) or ""
        token_url = simpledialog.askstring(
            "トークン取得URL",
            "登録メール記載の Token Generation URL をフルで入力してください\n"
            "(例: https://ip-data.jpo.go.jp/auth/token)",
            initialvalue=saved.get("token_url", "https://ip-data.jpo.go.jp/auth/token")
        ) or ""

        if not all([proxy_user, proxy_pass, jpo_user, jpo_pass, token_url]):
            raise ValueError("必要項目が入力されませんでした。")

        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "proxy_user": proxy_user,
                    "proxy_pass": proxy_pass,
                    "jpo_user": jpo_user,
                    "jpo_pass": jpo_pass,
                    "token_url": token_url
                },
                f, ensure_ascii=False, indent=2
            )

        return proxy_user, proxy_pass, jpo_user, jpo_pass, token_url

    except Exception as e:
        logging.error("資格情報取得エラー\n%s", traceback.format_exc())
        print(e)
        raise SystemExit(1)

def build_proxies(user: str, pwd: str) -> dict:
    """requests 用 proxies 辞書を生成"""
    url = f"http://{user}:{pwd}@proxy01.hm.jp.honda.com:8080"
    return {"http": url, "https": url}

# ---------- JPO API ----------
APP_NUMBER = "2024147578"  # 例
API_URL_FMT = (
    "https://ip-data.jpo.go.jp/api/patent/v1/app_doc_cont_refusal_reason/{}"
)
HEADERS = {"Content-Type": "application/x-www-form-urlencoded"}

def get_tokens(url: str, user: str, pwd: str, proxies: dict):
    """ID/パスワードでアクセストークン取得"""
    payload = {
        "grant_type": "password",
        "username": user,
        "password": pwd
    }
    r = requests.post(url, data=payload, headers=HEADERS,
                      proxies=proxies, timeout=30)
    logging.info("token POST %s status=%s", url, r.status_code)
    logging.debug("response: %s", r.text[:300])
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]

def refresh_access_token(url: str, refresh_token: str, proxies: dict):
    """リフレッシュトークンでアクセストークン更新"""
    payload = {"grant_type": "refresh_token", "refresh_token": refresh_token}
    r = requests.post(url, data=payload, headers=HEADERS,
                      proxies=proxies, timeout=30)
    logging.info("refresh POST status=%s", r.status_code)
    logging.debug("response: %s", r.text[:300])
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]

def download_refusal_reason(token: str, app_no: str, proxies: dict):
    """拒絶理由通知書 ZIP をダウンロード"""
    headers = {"Authorization": f"Bearer {token}"}
    url = API_URL_FMT.format(app_no)
    r = requests.get(url, headers=headers, proxies=proxies, timeout=60)
    logging.info("GET %s status=%s", url, r.status_code)
    if r.status_code == 200 and r.headers.get("Content-Type", "").startswith("application/zip"):
        fname = f"refusal_reason_{app_no}.zip"
        with open(fname, "wb") as f:
            f.write(r.content)
        print("ダウンロード成功:", fname)
    else:
        print("ダウンロード失敗:", r.status_code)
        print(r.text[:300])

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s"
    )

    proxy_user, proxy_pass, jpo_user, jpo_pass, token_url = ask_credentials()
    proxies = build_proxies(proxy_user, proxy_pass)

    try:
        access_token, refresh_token = get_tokens(
            token_url, jpo_user, jpo_pass, proxies
        )
    except requests.HTTPError as e:
        print("トークン取得失敗:", e, "\nレスポンス本文:", e.response.text[:300])
        return

    token_expiry = time.time() + 3600
    download_refusal_reason(access_token, APP_NUMBER, proxies)

    # 残 1 分で自動更新
    if time.time() > token_expiry - 60:
        access_token, refresh_token = refresh_access_token(
            token_url, refresh_token, proxies
        )
        download_refusal_reason(access_token, APP_NUMBER, proxies)

if __name__ == "__main__":
    main()
