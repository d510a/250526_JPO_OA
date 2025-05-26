import requests
import json
import os
import time
import logging
import tkinter as tk
from tkinter import simpledialog, filedialog
import traceback
import pandas as pd      # Excel 読み込み
import zipfile           # ZIP 展開
import shutil            # ファイル操作

# ---------- 固定値 ----------
TOKEN_URL = "https://ip-data.jpo.go.jp/auth/token"

# ---------- 設定ファイル ----------
documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
config_file = os.path.join(documents_folder, "proxy_jpo_config.json")


def ask_credentials():
    """プロキシ／JPO 資格情報を GUI で取得し JSON に保存"""
    logging.info("===== 資格情報入力開始 =====")
    root = tk.Tk(); root.withdraw()

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

        if not all([proxy_user, proxy_pass, jpo_user, jpo_pass]):
            raise ValueError("必要項目が入力されませんでした。")

        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(
                {"proxy_user": proxy_user, "proxy_pass": proxy_pass,
                 "jpo_user": jpo_user, "jpo_pass": jpo_pass},
                f, ensure_ascii=False, indent=2
            )
        return proxy_user, proxy_pass, jpo_user, jpo_pass

    except Exception as e:
        logging.error("資格情報取得エラー\n%s", traceback.format_exc())
        print(e)
        raise SystemExit(1)


def choose_excel_file():
    """Excel を選択し、[(A値, 出願番号)] と Excel のあるフォルダを返す"""
    root = tk.Tk(); root.withdraw()
    path = filedialog.askopenfilename(
        title="出願番号一覧の Excel ファイルを選択",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )
    if not path:
        print("Excel ファイルが選択されませんでした。処理を終了します。")
        raise SystemExit(1)

    pairs = load_entries(path)
    base_dir = os.path.dirname(path)
    return pairs, base_dir


def load_entries(excel_path):
    """
    A2 から下へ A列が空になる直前まで読み込み
    (A列値, B列=出願番号) のタプルをリストで返す
    """
    df = pd.read_excel(excel_path, sheet_name=0, header=None)
    entries = []

    for a_val, b_val in zip(df.iloc[1:, 0], df.iloc[1:, 1]):
        if pd.isna(a_val) or str(a_val).strip() == "":
            break
        if pd.isna(b_val) or str(b_val).strip() == "":
            raise ValueError(f"A列 {a_val} に対し B列が空です。")

        a_str, b_str = str(a_val), str(b_val)
        if a_str.endswith(".0"): a_str = a_str[:-2]
        if b_str.endswith(".0"): b_str = b_str[:-2]
        entries.append((a_str, b_str))

    if not entries:
        raise ValueError("A2 以降にデータが見つかりませんでした。")

    logging.info("Excel から %d 件の組を取得", len(entries))
    return entries


def build_proxies(user: str, pwd: str) -> dict:
    proxy = f"http://{user}:{pwd}@proxy01.hm.jp.honda.com:8080"
    return {"http": proxy, "https": proxy}


# ---------- JPO API ----------
API_URL_FMT = "https://ip-data.jpo.go.jp/api/patent/v1/app_doc_cont_refusal_reason/{}"
HEADERS = {"Content-Type": "application/x-www-form-urlencoded"}


def get_tokens(user: str, pwd: str, proxies: dict):
    payload = {"grant_type": "password", "username": user, "password": pwd}
    r = requests.post(TOKEN_URL, data=payload, headers=HEADERS,
                      proxies=proxies, timeout=30)
    logging.info("token POST status=%s", r.status_code)
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]


def refresh_access_token(refresh_token: str, proxies: dict):
    payload = {"grant_type": "refresh_token", "refresh_token": refresh_token}
    r = requests.post(TOKEN_URL, data=payload, headers=HEADERS,
                      proxies=proxies, timeout=30)
    logging.info("refresh POST status=%s", r.status_code)
    r.raise_for_status()
    data = r.json()
    return data["access_token"], data["refresh_token"]


def _extract_xml(zip_path: str, dest_folder: str):
    """
    ZIP 内の .xml ファイルを dest_folder に展開し、ZIP を削除
    """
    extracted = False
    with zipfile.ZipFile(zip_path) as zf:
        for member in zf.namelist():
            if member.lower().endswith(".xml"):
                # 同名ファイル衝突回避のため basename で保存
                out_path = os.path.join(dest_folder, os.path.basename(member))
                with zf.open(member) as src, open(out_path, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                extracted = True
    if extracted:
        os.remove(zip_path)
        logging.info("ZIP %s を削除し XML を展開しました", zip_path)
    else:
        logging.warning("ZIP %s に XML が見つかりませんでした", zip_path)


def download_refusal_reason(token: str, app_no: str,
                            proxies: dict, folder: str):
    """
    ZIP をダウンロード → XML を展開 → ZIP 削除
    """
    os.makedirs(folder, exist_ok=True)
    headers = {"Authorization": f"Bearer {token}"}
    url = API_URL_FMT.format(app_no)
    r = requests.get(url, headers=headers, proxies=proxies, timeout=60)
    logging.info("GET %s status=%s", url, r.status_code)

    if r.status_code == 200 and r.headers.get("Content-Type", "").startswith("application/zip"):
        zip_path = os.path.join(folder, f"refusal_reason_{app_no}.zip")
        with open(zip_path, "wb") as f:
            f.write(r.content)
        print(f"ダウンロード成功: {zip_path}")

        # ---- ZIP から XML 抽出 ----
        _extract_xml(zip_path, folder)
    else:
        print(f"[{app_no}] ダウンロード失敗:", r.status_code)
        print(r.text[:300])


def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")

    # 1) 認証情報
    proxy_user, proxy_pass, jpo_user, jpo_pass = ask_credentials()
    proxies = build_proxies(proxy_user, proxy_pass)

    # 2) Excel 読み込み
    try:
        entries, base_dir = choose_excel_file()
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

    # 4) 行ごとに処理
    for a_val, app_no in entries:
        # トークン残り 60 秒で更新
        if time.time() > token_expiry - 60:
            access_token, refresh_token = refresh_access_token(refresh_token, proxies)
            token_expiry = time.time() + 3600

        folder_name = f"{a_val}_{app_no}"
        folder_path = os.path.join(base_dir, folder_name)

        try:
            download_refusal_reason(access_token, app_no, proxies, folder_path)
        except Exception as e:
            print(f"[{app_no}] 処理中に例外:", e)


if __name__ == "__main__":
    main()
