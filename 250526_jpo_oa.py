import requests
import json
import os
import time
import logging
import tkinter as tk
from tkinter import simpledialog, filedialog
import traceback
import pandas as pd  # Excel 読み込み用

# ---------- 固定値 ----------
TOKEN_URL = "https://ip-data.jpo.go.jp/auth/token"

# ---------- 設定ファイル ----------
documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
config_file = os.path.join(documents_folder, "proxy_jpo_config.json")


def ask_credentials():
    """
    プロキシ／JPO 資格情報を GUI で取得し JSON に保存
    （token_url は固定で入力不要）
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
            "プロキシユーザ名",
            "プロキシのユーザ名を入力してください",
            initialvalue=saved.get("proxy_user", "")
        ) or ""
        proxy_pass = simpledialog.askstring(
            "プロキシパスワード",
            "プロキシのパスワードを入力してください",
            show="*",
            initialvalue=saved.get("proxy_pass", "")
        ) or ""
        jpo_user = simpledialog.askstring(
            "JPO APIユーザID",
            "特許庁APIのユーザIDを入力してください",
            initialvalue=saved.get("jpo_user", "")
        ) or ""
        jpo_pass = simpledialog.askstring(
            "JPO APIパスワード",
            "特許庁APIのパスワードを入力してください",
            show="*",
            initialvalue=saved.get("jpo_pass", "")
        ) or ""

        if not all([proxy_user, proxy_pass, jpo_user, jpo_pass]):
            raise ValueError("必要項目が入力されませんでした。")

        # 保存（token_url は固定なので保存しない）
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "proxy_user": proxy_user,
                    "proxy_pass": proxy_pass,
                    "jpo_user": jpo_user,
                    "jpo_pass": jpo_pass
                },
                f,
                ensure_ascii=False,
                indent=2,
            )

        return proxy_user, proxy_pass, jpo_user, jpo_pass

    except Exception as e:
        logging.error("資格情報取得エラー\n%s", traceback.format_exc())
        print(e)
        raise SystemExit(1)


def choose_excel_file():
    """
    Excel ファイルを選択し、出願番号リストとそのフォルダパスを返す
    """
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="出願番号一覧の Excel ファイルを選択",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )
    if not file_path:
        print("Excel ファイルが選択されませんでした。処理を終了します。")
        raise SystemExit(1)

    app_numbers = load_app_numbers(file_path)
    output_dir = os.path.dirname(file_path)
    return app_numbers, output_dir


def load_app_numbers(excel_path):
    """
    Excel B列 (列 index=1) の B2 から下方向に連続する値を出願番号として読み込む
    空セルに到達した時点で停止
    """
    df = pd.read_excel(excel_path, sheet_name=0, header=None)

    values = []
    for v in df.iloc[1:, 1]:  # B2 (index 1,1) から下へ
        if pd.isna(v) or str(v).strip() == "":
            break
        s = str(v)
        if s.endswith(".0"):
            s = s[:-2]
        values.append(s)

    if not values:
        raise ValueError("B2 以降に出願番号が見つかりませんでした。")

    logging.info("Excel から %d 件の出願番号を取得", len(values))
    return values


def build_proxies(user: str, pwd: str) -> dict:
    """requests 用 proxies 辞書を生成"""
    url = f"http://{user}:{pwd}@proxy01.hm.jp.honda.com:8080"
    return {"http": url, "https": url}


# ---------- JPO API ----------
API_URL_FMT = (
    "https://ip-data.jpo.go.jp/api/patent/v1/app_doc_cont_refusal_reason/{}"
)
HEADERS = {"Content-Type": "application/x-www-form-urlencoded"}


def get_tokens(user: str, pwd: str, proxies: dict):
    """ID/パスワードでアクセストークン取得"""
    payload = {"grant_type": "password", "username": user, "password": pwd}
    r = requests.post(TOKEN_URL, data=payload, headers=HEADERS, proxies=proxies, timeout=30)
    logging.info("token POST status=%s", r.status_code)
    logging.debug("response: %s", r.text[:300])
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]


def refresh_access_token(refresh_token: str, proxies: dict):
    """リフレッシュトークンでアクセストークン更新"""
    payload = {"grant_type": "refresh_token", "refresh_token": refresh_token}
    r = requests.post(TOKEN_URL, data=payload, headers=HEADERS, proxies=proxies, timeout=30)
    logging.info("refresh POST status=%s", r.status_code)
    logging.debug("response: %s", r.text[:300])
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]


def download_refusal_reason(token: str, app_no: str, proxies: dict, out_dir: str):
    """拒絶理由通知書 ZIP をダウンロード"""
    headers = {"Authorization": f"Bearer {token}"}
    url = API_URL_FMT.format(app_no)
    r = requests.get(url, headers=headers, proxies=proxies, timeout=60)
    logging.info("GET %s status=%s", url, r.status_code)
    if r.status_code == 200 and r.headers.get("Content-Type", "").startswith("application/zip"):
        fname = os.path.join(out_dir, f"refusal_reason_{app_no}.zip")
        with open(fname, "wb") as f:
            f.write(r.content)
        print(f"ダウンロード成功: {fname}")
    else:
        print(f"[{app_no}] ダウンロード失敗:", r.status_code)
        print(r.text[:300])


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    # 1) 資格情報取得
    proxy_user, proxy_pass, jpo_user, jpo_pass = ask_credentials()
    proxies = build_proxies(proxy_user, proxy_pass)

    # 2) Excel 選択 → 出願番号リスト & 出力フォルダ
    try:
        app_numbers, output_dir = choose_excel_file()
    except Exception as e:
        print("Excel 読み込み失敗:", e)
        return

    # 3) トークン取得
    try:
        access_token, refresh_token = get_tokens(jpo_user, jpo_pass, proxies)
    except requests.HTTPError as e:
        print("トークン取得失敗:", e, "\nレスポンス本文:", e.response.text[:300])
        return

    token_expiry = time.time() + 3600

    # 4) 各出願番号について ZIP ダウンロード
    for app_no in app_numbers:
        # トークン残り 1 分で更新
        if time.time() > token_expiry - 60:
            access_token, refresh_token = refresh_access_token(refresh_token, proxies)
            token_expiry = time.time() + 3600

        try:
            download_refusal_reason(access_token, app_no, proxies, output_dir)
        except Exception as e:
            print(f"[{app_no}] ダウンロード処理で例外:", e)


if __name__ == "__main__":
    main()
