#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
JPO API から拒絶理由通知 ZIP を取得・展開し、Excel (A=任意ID, B=出願番号, C=Publication Number) を読み取り
  • <A>_<出願番号> フォルダを作成
  • ZIP → XML 展開
  • JP<Publication Number>A.txt を 1 行目に作成
さらに
  • XML 内「＜引用文献等一覧＞」〜「＜先行技術文献調査結果の記録＞」に挟まれた引用文献を抽出
  • ChatGPT-API（OpenAI API）経由で Patsnap 形式 (ISO 国コード+西暦+番号+末尾 A/B/U) に正規化
  • 2 行目以降に追記（重複排除）

プロキシ／JPO API／OpenAI API 資格情報は GUI から取得し JSON 保存。
'''

from __future__ import annotations

import json
import os
import re
import shutil
import time
import logging
import zipfile
import tkinter as tk
from tkinter import filedialog, simpledialog
from typing import List, Tuple

import pandas as pd
import requests

# ------------------------------ 定数 ------------------------------
TOKEN_URL = 'https://ip-data.jpo.go.jp/auth/token'
OPENAI_ENDPOINT = 'https://api.openai.com/v1/chat/completions'
OPENAI_MODEL = 'gpt-4.1'  # 契約状況に合わせて変更可

DOCS_DIR = os.path.join(os.path.expanduser('~'), 'Documents')
CONFIG_PATH = os.path.join(DOCS_DIR, 'proxy_jpo_config.json')

# ----------------------- 資格情報ダイアログ -----------------------

def ask_credentials() -> Tuple[str, str, str, str, str]:
    """プロキシ・JPO API・OpenAI API キーを GUI で取得し保存"""
    logging.info('=== 資格情報入力開始 ===')
    root = tk.Tk(); root.withdraw()

    saved = {}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            saved = json.load(f)

    try:
        proxy_user = simpledialog.askstring('プロキシユーザ名', 'プロキシのユーザ名', initialvalue=saved.get('proxy_user', '')) or ''
        proxy_pass = simpledialog.askstring('プロキシパスワード', 'プロキシのパスワード', show='*', initialvalue=saved.get('proxy_pass', '')) or ''
        jpo_user   = simpledialog.askstring('JPO APIユーザID', 'JPO API ユーザID', initialvalue=saved.get('jpo_user', '')) or ''
        jpo_pass   = simpledialog.askstring('JPO APIパスワード', 'JPO API パスワード', show='*', initialvalue=saved.get('jpo_pass', '')) or ''
        openai_key = simpledialog.askstring('OpenAI APIキー', 'OpenAI API キー', show='*', initialvalue=saved.get('openai_key', '')) or ''

        if not all([proxy_user, proxy_pass, jpo_user, jpo_pass, openai_key]):
            raise ValueError('全ての項目を入力してください')

        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump({
                'proxy_user': proxy_user, 'proxy_pass': proxy_pass,
                'jpo_user': jpo_user, 'jpo_pass': jpo_pass,
                'openai_key': openai_key
            }, f, ensure_ascii=False, indent=2)
        return proxy_user, proxy_pass, jpo_user, jpo_pass, openai_key

    except Exception:
        logging.error('資格情報取得エラー', exc_info=True)
        raise SystemExit(1)

# ------------------------ プロキシ設定 ------------------------

def build_proxies(user: str, pwd: str) -> dict:
    proxy = f'http://{user}:{pwd}@proxy01.hm.jp.honda.com:8080'
    return {'http': proxy, 'https': proxy}

# ------------------------- Excel 読み込み -------------------------

def choose_excel_file():
    root = tk.Tk(); root.withdraw()
    path = filedialog.askopenfilename(title='Excel ファイルを選択', filetypes=[('Excel', '*.xlsx *.xlsm *.xls')])
    if not path:
        print('Excel 未選択. 終了します'); raise SystemExit(1)
    entries = load_entries(path)
    return entries, os.path.dirname(path)

def load_entries(xls: str) -> List[Tuple[str, str, str]]:
    df = pd.read_excel(xls, header=None)
    out = []
    for a, b, c in zip(df.iloc[1:,0], df.iloc[1:,1], df.iloc[1:,2]):
        if pd.isna(a) or str(a).strip()=='' : break
        if pd.isna(b) or str(b).strip()=='' : raise ValueError(f'A={a} の B列が空')
        if pd.isna(c) or str(c).strip()=='' : raise ValueError(f'A={a} の C列が空')
        def norm(x:str)->str:
            x=str(x); return x[:-2] if x.endswith('.0') else x
        out.append((norm(a), norm(b), norm(c)))
    if not out: raise ValueError('データ無し')
    logging.info('Excel から %d 件取得', len(out))
    return out

# ------------------------- JPO 認証 -------------------------
HEADERS_FORM = {'Content-Type': 'application/x-www-form-urlencoded'}
API_URL_FMT = 'https://ip-data.jpo.go.jp/api/patent/v1/app_doc_cont_refusal_reason/{}'

def get_tokens(uid,pwd,proxies):
    r=requests.post(TOKEN_URL,data={'grant_type':'password','username':uid,'password':pwd},headers=HEADERS_FORM,proxies=proxies,timeout=30)
    r.raise_for_status(); d=r.json(); return d['access_token'], d['refresh_token']

def refresh_token(refresh, proxies):
    r=requests.post(TOKEN_URL,data={'grant_type':'refresh_token','refresh_token':refresh},headers=HEADERS_FORM,proxies=proxies,timeout=30)
    r.raise_for_status(); d=r.json(); return d['access_token'], d['refresh_token']

# ------------------------- ファイル処理 -------------------------

def extract_zip(zip_path:str, dst:str):
    ext=False
    with zipfile.ZipFile(zip_path) as zf:
        for m in zf.namelist():
            if m.lower().endswith('.xml'):
                with zf.open(m) as src, open(os.path.join(dst, os.path.basename(m)), 'wb') as out:
                    shutil.copyfileobj(src,out); ext=True
    if ext: os.remove(zip_path)

# ---------------------- テキストファイル ----------------------

def pub_txt_path(folder:str, pub:str)->str:
    return os.path.join(folder, f'JP{pub}A.txt')

def ensure_pub_txt(folder:str, pub:str):
    os.makedirs(folder, exist_ok=True)
    path=pub_txt_path(folder,pub)
    if not os.path.exists(path):
        with open(path,'w',encoding='utf-8') as f: f.write(f'JP{pub}A\n')
    return path

# -------------------- XML → 引用抽出 --------------------
SYS_PROMPT=(
    """ #役割 あなたは特許庁のベテラン審査官です。

 #指示
 引用文献の公報番号のみを抽出してください。 
 引用文献は＜引用文献等一覧＞から＜先行技術文献調査結果の記録＞の間に記載されています。
 抽出した引用文献は、以下の変換ルールでpatsnap（特許検索ツール）で利用できる形式に変換してください。引用文献が何もない場合には何も出力しないで下さい。
 
 #変換のルール
 変換後は、頭にISO 3166-1 alpha-2（アルファツー）国名コードとなります。その後、西暦が続きます。その後、６～7桁の数字の番号が続きます。
 公開公報の場合は末尾に「Ａ」登録公報の場合は、末尾に「B」が付きます。実用新案は、末尾に「Ｕ」を付けてください。末尾につくアルファベットは１つのみです。実用新案は、優先的に「Ｕ」を付けてください。日本の元号はすべて西暦に変換してください。
 例えば
 「２．特開平１－４７８８０号公報（特に、特許請求の範囲,第2頁右上欄第7-12行、参照）」と記載のある場合は、JP198947880Aと変換ください。
 「１．米国特許出願公開第２０１５／０１８３０８３号明細書」と記載のあるものは、US20150183083A1と変換ください。
 「５.国際公開第２０１７／０１８０１６号（周知技術を示す文献）」は、WO2017018016と変換ください。
 「１．特開平５－１７８２６８号公報」は、JP1993178268Aと変換ください。
 「２．実願昭５８－１６７９９５号（実開昭６０－７５１９８号）のマイクロフィルム」は、JP1985075198Uと変換ください。
 「２．特開平１１－２０７７５号公報」は、JP199920775Aと変換ください。
 「１．特開平７－１６８９９４号公報」は、JP1995168994Aと変換ください。
 「１．特開２０１４－１７８９２８号公報」は、JP2014178928Aと変換ください。
 「２．国際公開第２０１１／０９２８１３号」は、WO2011092813と変換ください。

 #出力形式（公報番号のみ出力ください） 
  JP19991789822A
  jp2011167995A
""")

def detect_encoding(raw:bytes)->str:
    m=re.search(br'encoding="([^"]+)"', raw[:300])
    if m:
        enc=m.group(1).decode(errors='replace')
        return enc.lower().replace('shift_jis','cp932')
    return 'utf-8'

def read_xml_text(path:str)->str:
    with open(path,'rb') as f: raw=f.read()
    enc=detect_encoding(raw)
    return raw.decode(enc, errors='replace')


def gather_citation_section(folder:str)->str|None:
    buf=[]
    for f in os.listdir(folder):
        if not f.lower().endswith('.xml'): continue
        txt=read_xml_text(os.path.join(folder,f))
        m=re.search('引用文献等一覧', txt)
        if not m: continue
        seg=txt[m.end():]
        m2=re.search('先行技術文献調査結果の記録', seg)
        if m2: seg=seg[:m2.start()]
        buf.append(seg)
    return '\n'.join(buf) if buf else None


def gpt_normalize(raw:str, key:str, proxies:dict)->List[str]:
    if not raw.strip(): return []
    payload={
        'model': OPENAI_MODEL,
        'messages': [
            {'role':'system','content':SYS_PROMPT},
            {'role':'user','content': raw[:12000]}  # トークン節約
        ],
        'temperature':0}
    headers={'Authorization': f'Bearer {key}', 'Content-Type':'application/json'}
    r=requests.post(OPENAI_ENDPOINT,json=payload,headers=headers,proxies=proxies,timeout=120)
    r.raise_for_status(); txt=r.json()['choices'][0]['message']['content']
    return [ln.strip() for ln in txt.splitlines() if ln.strip()]


def append_citations(folder:str, pub:str, key:str, proxies:dict):
    raw=gather_citation_section(folder)
    if raw is None:
        logging.info('引用文献セクション無し: %s', folder); return

    try: ids=gpt_normalize(raw,key,proxies)
    except Exception as e:
        logging.error('ChatGPT 失敗 (%s): %s', folder, e); return
    if not ids:
        logging.info('正規化結果無し: %s', folder); return

    path=ensure_pub_txt(folder,pub)
    with open(path,encoding='utf-8') as f: existing={ln.strip() for ln in f if ln.strip()}
    new=[i for i in ids if i not in existing]
    if not new:
        logging.info('新規引用無し: %s', folder); return
    with open(path,'a',encoding='utf-8') as f:
        for i in new: f.write(i+'\n')
    logging.info('%d 件追記: %s', len(new), path)

# -------------------- JPO ダウンロード --------------------

def download_xml(token:str, app:str, proxies:dict, folder:str):
    os.makedirs(folder, exist_ok=True)
    r=requests.get(API_URL_FMT.format(app), headers={'Authorization': f'Bearer {token}'}, proxies=proxies, timeout=60)
    if r.status_code==200 and r.headers.get('Content-Type','').startswith('application/zip'):
        zp=os.path.join(folder, f'refusal_{app}.zip')
        with open(zp,'wb') as f: f.write(r.content)
        extract_zip(zp, folder)
    else:
        logging.warning('[%s] ダウンロード失敗: %s', app, r.status_code)

# ------------------------------ main ------------------------------

def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

    proxy_user, proxy_pass, jpo_user, jpo_pass, openai_key = ask_credentials()
    proxies = build_proxies(proxy_user, proxy_pass)

    entries, base_dir = choose_excel_file()

    try:
        access, refresh = get_tokens(jpo_user, jpo_pass, proxies)
    except Exception as e:
        print('JPO 認証失敗:', e); return
    expiry=time.time()+3600

    for aid, app, pub in entries:
        if time.time()>expiry-60:
            access, refresh = refresh_token(refresh, proxies); expiry=time.time()+3600
        folder=os.path.join(base_dir, f'{aid}_{app}')
        download_xml(access, app, proxies, folder)
        ensure_pub_txt(folder, pub)
        append_citations(folder, pub, openai_key, proxies)

if __name__=='__main__':
    main()
