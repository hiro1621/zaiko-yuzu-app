# -*- coding: utf-8 -*-
"""
在庫融通アプリの「裏の保管庫」＝Googleスプレッドシート読み書き（gsheet_store.py）

【何をするスクリプトか】
店舗セルフアップロード型アプリ（streamlit_app.py）の“貯金箱”です。
各店がバラバラにアップロードした在庫（前処理・スリム化済み）を1つのGoogleシートへ貯め、
全店ぶんがそろった時点でマッチング結果を書き戻します。店舗はこのシートを直接は触らず、
必ずアプリ経由で読み書きします（サービスアカウントにだけ編集者共有）。

融通は「他店の出庫実績」がそろって初めて計算できるため、各店のアップを貯める
共有保管庫が必須です。だからアプリは、現状データから毎回すべてを再計算し、
「現在 N/15店 アップ済み」を表示します。

■ このシートのタブ構成
  _index                … 店名／対象年月(YYYYMM)／アップ日時／行数／様式OK・NG／ファイル名。○/15の判定元。
  raw_<店名>            … その店のスリム在庫（約35列・前処理済）。同月再提出は上書き（最新採用）。
  融通提案／不足品目一覧／品目×店舗マトリクス／店舗別サマリ … 結果（毎回再計算して書き戻し）。
  前月_YYYYMM           … 月替わり時に、直前の結果（融通提案）を退避したスナップショット。

■ 月替わりの考え方
  ・各rawと_indexに対象年月(YYYYMM)を持たせる。
  ・マッチングは「当月（＝_index内で最新のYYYYMM）」のデータだけで計算する。
  ・新しい年月が最初に来たら、直前の結果を 前月_<直前YYYYMM> へ退避してから当月を始める。
  ・○/15は「当月データを持つ店数」で数える。

★重要：認証情報が無い環境でも import・構文が通るよう、gspread と google-auth は
  関数の中で遅延importしています（ここを import しただけでは何も接続しません）。
  実際のGoogleシート接続は、Streamlit Cloud の Secrets（サービスアカウント鍵）が
  設定されて初めて成立します。ローカルの構文チェックはこのままで通ります。

【必要なライブラリのインストール（Googleシート連携を使うときだけ）】
    pip install gspread google-auth

※ venv不要。Windows専用パス。コメント・メッセージはすべて日本語です。
"""

import datetime

# yuzu_core（同じフォルダ）から、保管庫に入れる列の定義と道具を借りる
from yuzu_core import KEEP_COLS, g


# ============================================================================
# タブ名（Googleスプレッドシート側のシート名）
# ============================================================================
INDEX_TAB = '_index'
RAW_PREFIX = 'raw_'
TAB_PROPOSAL = '融通提案'
TAB_SHORTAGE = '不足品目一覧'
TAB_MATRIX = '品目×店舗マトリクス'
TAB_SUMMARY = '店舗別サマリ'
PREV_PREFIX = '前月_'   # 前月_YYYYMM

# _index の見出し（列順）
INDEX_HEADERS = ['店名', '対象年月', 'アップ日時', '行数', '様式', 'ファイル名']


# ============================================================================
# 接続
# ============================================================================
def open_spreadsheet(sa_info, spreadsheet_id):
    """
    サービスアカウント情報（辞書）とスプレッドシートIDから、シートを開いて返す。
      sa_info        … Secrets（TOML）に貼ったサービスアカウント鍵の中身（辞書）。
      spreadsheet_id … 対象スプレッドシートURL /d/【ここ】/edit のID。
    ※ ここで初めて gspread / google-auth を import する（未設定環境では呼ばれない）。
    """
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(dict(sa_info), scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(spreadsheet_id)


def _get_or_create_ws(sh, title, rows=2000, cols=40):
    """ 指定名のタブを取得。無ければ作る。 """
    try:
        return sh.worksheet(title)
    except Exception:
        return sh.add_worksheet(title=title, rows=str(rows), cols=str(cols))


def _update(ws, values, range_name='A1'):
    """ ws.update をバージョン差（引数順）に強い形（キーワード指定）で呼ぶ。
        gspread 5系（range_name, values）／6系（values, range_name）どちらでも通る。 """
    ws.update(range_name=range_name, values=values, value_input_option='RAW')


# ============================================================================
# _index の読み書き
# ============================================================================
def read_index(sh):
    """
    _index タブを読んで {店名: {'ym','uploaded_at','rows','format','filename'}} を返す。
    タブが無ければ空の辞書。
    """
    try:
        ws = sh.worksheet(INDEX_TAB)
    except Exception:
        return {}
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return {}
    header = values[0]
    idx = {h: i for i, h in enumerate(header)}
    out = {}
    for row in values[1:]:
        if not row or not (row[idx.get('店名', 0)] if idx.get('店名', 0) < len(row) else ''):
            continue

        def cell(name):
            i = idx.get(name)
            return row[i] if (i is not None and i < len(row)) else ''

        name = cell('店名').strip()
        if not name:
            continue
        out[name] = {
            'ym': cell('対象年月').strip(),
            'uploaded_at': cell('アップ日時').strip(),
            'rows': cell('行数').strip(),
            'format': cell('様式').strip(),
            'filename': cell('ファイル名').strip(),
        }
    return out


def _write_index(sh, index):
    """ index 辞書を _index タブへ丸ごと書き出す（店名の五十音でなく登録順を保つため店名でソート）。 """
    ws = _get_or_create_ws(sh, INDEX_TAB)
    ws.clear()
    body = [list(INDEX_HEADERS)]
    for name in sorted(index.keys()):
        e = index[name]
        body.append([name, e.get('ym', ''), e.get('uploaded_at', ''),
                     e.get('rows', ''), e.get('format', ''), e.get('filename', '')])
    _update(ws, body)


def latest_ym(index):
    """ _index の中で最新の対象年月(YYYYMM)を返す。1件も無ければ None。 """
    yms = [e.get('ym', '') for e in index.values() if e.get('ym', '')]
    return max(yms) if yms else None


# ============================================================================
# raw_<店名> の読み書き
# ============================================================================
def _raw_tab_name(store_name):
    return RAW_PREFIX + store_name


def write_raw(sh, store_name, slim_rows):
    """ その店のスリム在庫（約35列・前処理済）を raw_<店名> へ書き込む（毎回上書き）。
        slim_rows は {列名:文字列} の辞書のリスト（yuzu_core.slim_rows の出力）。 """
    ws = _get_or_create_ws(sh, _raw_tab_name(store_name),
                           rows=max(2000, len(slim_rows) + 10), cols=len(KEEP_COLS) + 2)
    ws.clear()
    body = [list(KEEP_COLS)]
    for r in slim_rows:
        body.append([r.get(c, '') for c in KEEP_COLS])
    _update(ws, body)


def read_raw(sh, store_name):
    """ raw_<店名> を読み、{列名:文字列} の辞書のリストで返す。タブが無ければ空リスト。 """
    try:
        ws = sh.worksheet(_raw_tab_name(store_name))
    except Exception:
        return []
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return []
    header = values[0]
    rows = []
    for raw in values[1:]:
        d = {}
        for ci, h in enumerate(header):
            d[h] = raw[ci] if ci < len(raw) else ''
        # KEEP_COLS のうちタブに無い列は空文字で補う（列欠けに強くする）
        for c in KEEP_COLS:
            d.setdefault(c, '')
        rows.append(d)
    return rows


# ============================================================================
# 結果4タブの書き戻し
# ============================================================================
def write_results(sh, payload):
    """
    compute_matching の結果（を payload 化したもの）を、結果4タブへ書き戻す。
    payload の作り方は build_results_payload() を参照。
    """
    # 融通提案
    ws = _get_or_create_ws(sh, TAB_PROPOSAL)
    ws.clear()
    note = ('基準年月：%s ／ CSV基準日：%s ／ ※月初スナップショットです（現在庫と異なる場合があります）'
            % (payload.get('base_ym', ''), payload.get('csv_base', '')))
    body = [[note], list(payload['proposal_headers'])]
    body += [list(r) for r in payload['proposal_rows']]
    _update(ws, body)

    # 不足品目一覧
    ws2 = _get_or_create_ws(sh, TAB_SHORTAGE)
    ws2.clear()
    _update(ws2, [list(payload['shortage_headers'])] + [list(r) for r in payload['shortage_rows']])

    # 品目×店舗マトリクス
    ws3 = _get_or_create_ws(sh, TAB_MATRIX)
    ws3.clear()
    _update(ws3, [list(payload['matrix_headers'])] + [list(r) for r in payload['matrix_rows']])

    # 店舗別サマリ
    ws4 = _get_or_create_ws(sh, TAB_SUMMARY)
    ws4.clear()
    _update(ws4, [list(payload['summary_headers'])] + [list(r) for r in payload['summary_rows']])


def snapshot_results_to_prev(sh, prev_ym):
    """
    月替わり時、直前の結果（融通提案タブ）を 前月_<prev_ym> へ退避する。
    ※ 融通提案タブがまだ無い初回は何もしない。
    """
    if not prev_ym:
        return
    try:
        old = sh.worksheet(TAB_PROPOSAL)
    except Exception:
        return
    values = old.get_all_values()
    if not values:
        return
    prev = _get_or_create_ws(sh, PREV_PREFIX + prev_ym)
    prev.clear()
    _update(prev, values)


# ============================================================================
# アップロード1件を保管する（月替わり退避つき）
# ============================================================================
def save_store_upload(sh, store_name, ym, slim_rows, filename, format_ok):
    """
    店のアップロード1件を保管庫へ入れる。
      1) _index を読む。既存の最新年月 old_latest を得る。
      2) 今回の ym が old_latest より新しい（＝新しい月が始まる）なら、
         直前の結果を 前月_<old_latest> へ退避してから当月を始める。
      3) raw_<店名> を上書きし、_index の当該店を更新する。
    戻り値：更新後の _index（辞書）。
    """
    index = read_index(sh)
    old_latest = latest_ym(index)

    # 2) 月替わり退避（今回のymが既存最新より新しいときだけ）
    if old_latest is not None and ym > old_latest:
        snapshot_results_to_prev(sh, old_latest)

    # 3) raw と _index を更新
    write_raw(sh, store_name, slim_rows)
    index[store_name] = {
        'ym': ym,
        'uploaded_at': datetime.datetime.now().strftime('%Y/%m/%d %H:%M'),
        'rows': str(len(slim_rows)),
        'format': 'OK' if format_ok else 'NG(別様式)',
        'filename': filename,
    }
    _write_index(sh, index)
    return index


# ============================================================================
# 当月データを持つ全店を読み込む（マッチング入力を組み立てる）
# ============================================================================
def load_current_month_stores(sh):
    """
    _index を見て「当月（最新年月）のデータを持つ店」だけを raw から読み込み、
    compute_matching に渡せる stores のリストを組み立てて返す。
      戻り値：(stores, latest, index)
        stores … [{'name','ym','base_date','rows'}, ...]（当月のみ）
        latest … 当月の年月(YYYYMM) または None
        index  … _index の全内容（画面の「未アップ店」表示等に使う）
    """
    index = read_index(sh)
    latest = latest_ym(index)
    stores = []
    if latest is None:
        return stores, latest, index
    for name, e in index.items():
        if e.get('ym', '') != latest:
            continue
        if e.get('format', '').startswith('NG'):
            # 別様式でNG判定の店は、当月マッチングには入れない（画面には「様式NG」で出す）
            continue
        rows = read_raw(sh, name)
        y, m = int(latest[:4]), int(latest[4:6])
        base_date = datetime.date(y, m, 1)
        stores.append({'name': name, 'ym': latest, 'base_date': base_date, 'rows': rows})
    return stores, latest, index


# ============================================================================
# compute_matching の戻り値を、write_results 用 payload に変換
# ============================================================================
def build_results_payload(result, base_ym_disp, csv_base_disp):
    """
    yuzu_core.compute_matching(...) の戻り値 result を、結果4タブ書き戻し用の
    「文字列2次元配列」に整える。Excel出力（write_excel）と同じ列・同じ並び。
    """
    proposal_headers = ['出し手店', '薬品名', '単位', 'メーカ名', '過剰数', '過剰数金額',
                        '過剰在庫区分', '不動区分', '期限切迫区分', '有効期限', 'ロットNO',
                        '最終出庫日', '要記録警告', '引取候補店', '参考:過剰だが使用中の店',
                        '6ヶ月出庫回数', '医薬品CD']
    proposal_rows = [[r['出し手店'], r['薬品名'], r['単位'], r['メーカ名'], r['過剰数'],
                      r['過剰数金額'], r['過剰在庫区分'], r['不動区分'], r['期限切迫区分'],
                      r['有効期限'], r['ロットNO'], r['最終出庫日'], r['要記録警告'],
                      r['引取候補店'], r['参考:過剰だが使用中の店'],
                      r['6ヶ月出庫回数'], r['医薬品CD']] for r in result['proposal_rows']]

    shortage_headers = ['店', '薬品名', '在庫数', '安全在庫数', '不足数', '医薬品CD', '過剰に持つ他店']
    shortage_rows = [[r['店'], r['薬品名'], r['在庫数'], r['安全在庫数'], r['不足数'],
                      r['医薬品CD'], r['過剰に持つ他店']] for r in result['shortage_rows']]

    matrix_headers = ['薬品名', '医薬品CD'] + result['store_names']
    matrix_rows = result['matrix_rows']

    summary_headers = ['店', '過剰品目数', '過剰金額計', '期限切迫品目数', '期限切迫金額計', '不足品目数']
    summary_rows = [[r['店'], r['過剰品目数'], r['過剰金額計'], r['期限切迫品目数'],
                     r['期限切迫金額計'], r['不足品目数']] for r in result['summary_rows']]

    return {
        'base_ym': base_ym_disp,
        'csv_base': csv_base_disp,
        'proposal_headers': proposal_headers, 'proposal_rows': proposal_rows,
        'shortage_headers': shortage_headers, 'shortage_rows': shortage_rows,
        'matrix_headers': matrix_headers, 'matrix_rows': matrix_rows,
        'summary_headers': summary_headers, 'summary_rows': summary_rows,
    }
