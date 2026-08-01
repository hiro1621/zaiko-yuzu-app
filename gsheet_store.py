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
import time

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
EXCLUDE_TAB = '_除外'   # 店が「この品は融通に出さない」と外した品目
RESERVE_TAB = '_予約'   # 受け手の店が「この品はうちが引き取ります」と押さえた品目

# _index の見出し（列順）
INDEX_HEADERS = ['店名', '対象年月', 'アップ日時', '行数', '様式', 'ファイル名']

# _除外 の見出し（列順）
EXCLUDE_HEADERS = ['店名', '除外キー', '薬品名', '除外日時']

# _予約 の見出し（列順）
#   1行＝1予約。(出し手店, 予約キー) で1件を特定する＝同じ品を2店が予約することはない。
#   ・対象年月 … 「予約を入れた月」の記録（2026-08-01に意味を変更。旧：この月だけ有効）。
#   ・受取予定月 … 「いつ引き取るか（YYYYMM）」。2026-08-01に追加した列。
#       有効判定を「対象年月＝当月か」→「当月 ≤ 受取予定月か」に変えたことで、予約を
#       翌月以降へ持ち越せるようにした（最大3ヶ月）。★列を消さず末尾でなく所定位置に
#       増やすだけなので、受取予定月が空の古い行は「今すぐ＝対象年月と同じ」とみなして
#       そのまま読める（後方互換。read_reservations は列名で拾うので位置に依存しない）。
RESERVE_HEADERS = ['予約した店', '出し手店', '対象年月', '受取予定月', '予約キー', '薬品名', '予約日時']


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


# ============================================================================
# Google APIの呼び出し回数を減らす／上限に当たっても自動でやり直す（2026-07-29）
#   ★★背景：Googleスプレッドシートは「1分あたり60回」まで（読み取り・書き込みで別枠）。
#     この上限はサービスアカウント1つあたり＝全店で共有なので、店数が増えるほど当たりやすい。
#
#   問題だったのは、gspread の sh.worksheet('タブ名') が呼ぶたびにブックの目次を
#   取りに行く（＝1回ぶん消費する）こと。タブを触る＝「目次を取る」＋「中身を読む」の
#   2回だったため、起動1回で14店なら約40回に達していた。
#   → 目次は最初の1回だけ取って使い回す（_ws_map）。読み取りはほぼ半分になる。
#      3店なら18回→7回、14店なら約40回→約18回。
# ============================================================================
_WS_CACHE = {}   # {ブックのID: {タブ名: ワークシート}}

# 上限（429）に当たったときに待つ秒数。3回までやり直す。
_RETRY_WAITS = [2, 5, 10]


def _is_rate_limited(e):
    """ 例外が「1分あたりの上限に当たった（429）」ものかどうか。 """
    status = getattr(getattr(e, 'response', None), 'status_code', None)
    if status == 429:
        return True
    s = str(e)
    return ('429' in s) or ('Quota exceeded' in s) or ('RATE_LIMIT_EXCEEDED' in s)


def _call(fn, *args, **kwargs):
    """
    Google APIを1回呼ぶ。1分あたりの上限に当たったら少し待って自動でやり直す。
      ・待ち時間は 2秒 → 5秒 → 10秒（最大3回）。たいていは1回目のやり直しで通る。
      ・上限以外のエラーは、そのまま呼び出し元へ返す（握りつぶさない）。
        ただしタブの取り違え（外でタブを消した等）に備え、目次のキャッシュは捨てておく。
    """
    last = None
    for wait in [0] + _RETRY_WAITS:
        if wait:
            time.sleep(wait)
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if not _is_rate_limited(e):
                _WS_CACHE.clear()
                raise
            last = e
    raise last


def _book_key(sh):
    return getattr(sh, 'id', None) or id(sh)


def _ws_map(sh, refresh=False):
    """ タブ名→ワークシートの対応表。初回だけ目次を取りに行き、あとは使い回す。 """
    key = _book_key(sh)
    if refresh or key not in _WS_CACHE:
        _WS_CACHE[key] = {ws.title: ws for ws in _call(sh.worksheets)}
    return _WS_CACHE[key]


def reset_ws_cache():
    """ タブの目次キャッシュを捨てる（タブを新しく作った直後などに呼ぶ）。 """
    _WS_CACHE.clear()


def _find_ws(sh, title):
    """ 指定名のタブを返す。無ければ None。
        キャッシュに無いときだけ目次を1回取り直す（新しく作られたタブを拾うため）。 """
    m = _ws_map(sh)
    if title in m:
        return m[title]
    m = _ws_map(sh, refresh=True)
    return m.get(title)


def _get_or_create_ws(sh, title, rows=2000, cols=40):
    """ 指定名のタブを取得。無ければ作る。 """
    ws = _find_ws(sh, title)
    if ws is not None:
        return ws
    ws = _call(sh.add_worksheet, title=title, rows=str(rows), cols=str(cols))
    _ws_map(sh)[title] = ws      # 作ったタブを目次キャッシュにも足す（取り直さない）
    return ws


def _values(ws):
    """ タブの中身を全部読む（上限に当たったら自動でやり直す）。 """
    return _call(ws.get_all_values)


def _clear(ws):
    """ タブを空にする（上限に当たったら自動でやり直す）。 """
    return _call(ws.clear)


def _update(ws, values, range_name='A1'):
    """ ws.update をバージョン差（引数順）に強い形（キーワード指定）で呼ぶ。
        gspread 5系（range_name, values）／6系（values, range_name）どちらでも通る。 """
    _call(ws.update, range_name=range_name, values=values, value_input_option='RAW')


# ============================================================================
# _index の読み書き
# ============================================================================
def read_index(sh):
    """
    _index タブを読んで {店名: {'ym','uploaded_at','rows','format','filename'}} を返す。
    タブが無ければ空の辞書。
    """
    ws = _find_ws(sh, INDEX_TAB)
    if ws is None:
        return {}
    values = _values(ws)
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
    _clear(ws)
    body = [list(INDEX_HEADERS)]
    for name in sorted(index.keys()):
        e = index[name]
        body.append([name, e.get('ym', ''), e.get('uploaded_at', ''),
                     e.get('rows', ''), e.get('format', ''), e.get('filename', '')])
    _update(ws, body)


# ============================================================================
# _除外（店が「融通に出さない」と外した品目）の読み書き
#   ・1行＝1品目。店名＋除外キー（＝個別医薬品CD、無い品は「名:薬品名」）で特定する。
#   ・件数が知れているので、読むときは全部読み、書くときは全部書き直す（丸ごと上書き）。
# ============================================================================
def read_exclusions(sh):
    """ _除外 タブを読んで [{'店名','除外キー','薬品名','除外日時'}, ...] を返す。
        タブがまだ無ければ空リスト。 """
    ws = _find_ws(sh, EXCLUDE_TAB)
    if ws is None:
        return []
    values = _values(ws)
    if not values or len(values) < 2:
        return []
    header = values[0]
    idx = {h: i for i, h in enumerate(header)}
    out = []
    for row in values[1:]:

        def cell(name):
            i = idx.get(name)
            return (row[i] if (i is not None and i < len(row)) else '').strip()

        if not cell('店名') or not cell('除外キー'):
            continue
        out.append({'店名': cell('店名'), '除外キー': cell('除外キー'),
                    '薬品名': cell('薬品名'), '除外日時': cell('除外日時')})
    return out


def write_exclusions(sh, rows):
    """ 除外リストを丸ごと書き直す（rows は read_exclusions と同じ形の辞書リスト）。 """
    ws = _get_or_create_ws(sh, EXCLUDE_TAB, rows=max(2000, len(rows) + 10), cols=8)
    _clear(ws)
    body = [list(EXCLUDE_HEADERS)]
    for r in rows:
        body.append([r.get('店名', ''), r.get('除外キー', ''),
                     r.get('薬品名', ''), r.get('除外日時', '')])
    _update(ws, body)


# ============================================================================
# _予約（受け手の店が「この品はうちが引き取ります」と押さえた品目）の読み書き
#   ・_除外 とまったく同じ作り。1行＝1予約。全部読んで全部書き直す（丸ごと上書き）。
#   ・(出し手店, 予約キー) で1件。同じ品を2店が予約することはない（保存前に画面側で突合する）。
# ============================================================================
def read_reservations(sh):
    """ _予約 タブを読んで
        [{'予約した店','出し手店','対象年月','受取予定月','予約キー','薬品名','予約日時'}, ...] を返す。
        タブがまだ無ければ空リスト。
        ★『受取予定月』列がまだ無い古いシート・古い行では空文字で返る（列名で拾うので位置に依存しない）。
          空文字は呼び出し側（app_logic）で「今すぐ＝対象年月と同じ」とみなす＝後方互換。 """
    ws = _find_ws(sh, RESERVE_TAB)
    if ws is None:
        return []
    values = _values(ws)
    if not values or len(values) < 2:
        return []
    header = values[0]
    idx = {h: i for i, h in enumerate(header)}
    out = []
    for row in values[1:]:

        def cell(name):
            i = idx.get(name)
            return (row[i] if (i is not None and i < len(row)) else '').strip()

        # 予約した店・出し手店・予約キーの3つがそろっていない行は読み飛ばす
        if not cell('予約した店') or not cell('出し手店') or not cell('予約キー'):
            continue
        out.append({'予約した店': cell('予約した店'), '出し手店': cell('出し手店'),
                    '対象年月': cell('対象年月'), '受取予定月': cell('受取予定月'),
                    '予約キー': cell('予約キー'),
                    '薬品名': cell('薬品名'), '予約日時': cell('予約日時')})
    return out


def write_reservations(sh, rows):
    """ 予約リストを丸ごと書き直す（rows は read_reservations と同じ形の辞書リスト）。
        ★受取予定月が入っていない古い行（辞書に無い）は空文字で書き出す＝後方互換。 """
    ws = _get_or_create_ws(sh, RESERVE_TAB, rows=max(2000, len(rows) + 10), cols=8)
    _clear(ws)
    body = [list(RESERVE_HEADERS)]
    for r in rows:
        body.append([r.get('予約した店', ''), r.get('出し手店', ''), r.get('対象年月', ''),
                     r.get('受取予定月', ''), r.get('予約キー', ''),
                     r.get('薬品名', ''), r.get('予約日時', '')])
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
    _clear(ws)
    body = [list(KEEP_COLS)]
    for r in slim_rows:
        body.append([r.get(c, '') for c in KEEP_COLS])
    _update(ws, body)


def read_raw(sh, store_name):
    """ raw_<店名> を読み、{列名:文字列} の辞書のリストで返す。タブが無ければ空リスト。 """
    ws = _find_ws(sh, _raw_tab_name(store_name))
    if ws is None:
        return []
    values = _values(ws)
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
    _clear(ws)
    note = ('基準年月：%s ／ CSV基準日：%s ／ ※月初スナップショットです（現在庫と異なる場合があります）'
            % (payload.get('base_ym', ''), payload.get('csv_base', '')))
    body = [[note], list(payload['proposal_headers'])]
    body += [list(r) for r in payload['proposal_rows']]
    _update(ws, body)

    # 不足品目一覧
    ws2 = _get_or_create_ws(sh, TAB_SHORTAGE)
    _clear(ws2)
    _update(ws2, [list(payload['shortage_headers'])] + [list(r) for r in payload['shortage_rows']])

    # 品目×店舗マトリクス
    ws3 = _get_or_create_ws(sh, TAB_MATRIX)
    _clear(ws3)
    _update(ws3, [list(payload['matrix_headers'])] + [list(r) for r in payload['matrix_rows']])

    # 店舗別サマリ
    ws4 = _get_or_create_ws(sh, TAB_SUMMARY)
    _clear(ws4)
    _update(ws4, [list(payload['summary_headers'])] + [list(r) for r in payload['summary_rows']])


def snapshot_results_to_prev(sh, prev_ym):
    """
    月替わり時、直前の結果（融通提案タブ）を 前月_<prev_ym> へ退避する。
    ※ 融通提案タブがまだ無い初回は何もしない。
    """
    if not prev_ym:
        return
    old = _find_ws(sh, TAB_PROPOSAL)
    if old is None:
        return
    values = _values(old)
    if not values:
        return
    prev = _get_or_create_ws(sh, PREV_PREFIX + prev_ym)
    _clear(prev)
    _update(prev, values)


# ============================================================================
# 前月の結果を読む（滞留＝何ヶ月つづけて載っているか の判定に使う）
# ----------------------------------------------------------------------------
# 2026-07-30 追加。前月_YYYYMM タブはこれまで「書くだけで誰も読んでいない」状態だった。
#
# 【なぜ前月の1枚だけ読むのか（過去12ヶ月ぶんを読まない理由）】
#   滞留月数を前月タブに書き込んでおき、翌月は「前月の月数＋1」で数える方式にしている。
#   こうすると読むタブは常に1枚で済み、それでいて何ヶ月でもさかのぼって数えられる。
#   過去タブを何枚も読む方式だと、Googleの読み取り上限（このファイル冒頭の対策の理由）に
#   近づくうえ、さかのぼれる月数がタブの枚数で頭打ちになる。
# ============================================================================
def _prev_tab_yms(sh):
    """ 保管庫にある 前月_YYYYMM タブの年月を、新しい順のリストで返す。 """
    yms = []
    for title in _ws_map(sh).keys():
        if not title.startswith(PREV_PREFIX):
            continue
        ym = title[len(PREV_PREFIX):].strip()
        if len(ym) == 6 and ym.isdigit():
            yms.append(ym)
    return sorted(yms, reverse=True)


def _prev_row_key(cd, name):
    """ 前月タブの1行から品目キーを作る。
        ★yuzu_core.exclusion_key と必ず同じ作り方にすること
          （個別医薬品CDが取れればそれを使い、取れない品だけ『名:薬品名』で代用）。
          ここがズレると、同じ品なのに前月と別物と見なされ、滞留がいつまでも1ヶ月目になる。 """
    cd = str(cd or '').strip()
    if cd:
        return cd
    return '名:' + str(name or '').strip()


def read_prev_proposal(sh, current_ym=None):
    """
    直前の月の 前月_YYYYMM タブを読み、
      {(出し手店, 品目キー): {'滞留月数': int, '引取候補店': str, '予約': str}}
    を返す。タブが1枚も無ければ空の辞書（＝全部「今月から」になる）。

      current_ym … 当月の年月(YYYYMM)。これより前で最も新しいタブを読む。
                   None なら単純に最も新しいタブを読む。

    ※タブの中身は「融通提案」をそのまま退避したもの＝1行目が注記、2行目が見出し。
      見出しの位置は決め打ちせず『出し手店』を含む行を探して特定する
      （注記の有無が将来変わっても壊れないようにするため）。
    """
    yms = _prev_tab_yms(sh)
    if current_ym:
        yms = [y for y in yms if y < str(current_ym)]
    if not yms:
        return {}

    ws = _find_ws(sh, PREV_PREFIX + yms[0])
    if ws is None:
        return {}
    values = _values(ws)
    if not values:
        return {}

    # 見出し行を探す（'出し手店' が入っている最初の行）
    hi = None
    for i, rowv in enumerate(values[:5]):
        if any(str(c).strip() == '出し手店' for c in rowv):
            hi = i
            break
    if hi is None:
        return {}
    header = [str(c).strip() for c in values[hi]]

    def col(name):
        return header.index(name) if name in header else None

    i_store = col('出し手店')
    i_cd = col('医薬品CD')
    i_name = col('薬品名')
    i_cand = col('引取候補店')
    i_book = col('予約')
    i_stag = col('滞留')          # 滞留列を足す前の月のタブには存在しない
    if i_store is None or i_name is None:
        return {}

    out = {}
    for rowv in values[hi + 1:]:
        def cell(i):
            return str(rowv[i]).strip() if (i is not None and i < len(rowv)) else ''
        store = cell(i_store)
        if not store:
            continue
        key = _prev_row_key(cell(i_cd), cell(i_name))

        # 前月の滞留月数を数字として取り出す（'3ヶ月目' → 3 ／ '⚠ 4ヶ月目・先月予約済' → 4）。
        #   ★滞留列がまだ無い月のタブでは 1 とみなす。
        #     こうすると「前月に載っていた」という事実だけは拾えるので、
        #     この機能を入れた最初の月から、ちゃんと『2ヶ月目』が出せる。
        months = 1
        s = cell(i_stag)
        if s:
            digits = ''
            for ch in s:
                if ch.isdigit():
                    digits += ch
                elif digits:
                    break
            if digits:
                months = int(digits)

        out[(store, key)] = {
            '滞留月数': months,
            '引取候補店': cell(i_cand),
            '予約': cell(i_book),
        }
    return out


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
def load_current_month_stores(sh, index=None):
    """
    _index を見て「当月（最新年月）のデータを持つ店」だけを raw から読み込み、
    compute_matching に渡せる stores のリストを組み立てて返す。
      戻り値：(stores, latest, index)
        stores … [{'name','ym','base_date','rows'}, ...]（当月のみ）
        latest … 当月の年月(YYYYMM) または None
        index  … _index の全内容（画面の「未アップ店」表示等に使う）
    ★index を渡すと _index を読み直さない（2026-07-29）。
      呼び出し側（load_stores_cached）はキャッシュ判定のために _index を先に読んでおり、
      ここでもう一度読むと同じタブを2回読むことになるため。
    """
    if index is None:
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
    # ※末尾の『予約』は2026-07-28、『滞留』は2026-07-30に追加した列（write_excel と同じ並び）。
    #   ★『滞留』はここに書き出すことが翌月の判定材料そのものになる（read_prev_proposal が読む）。
    #     列名・書式を変えるときは read_prev_proposal も必ず合わせること。
    proposal_headers = ['出し手店', '種別', '薬品名', '単位', 'メーカ名', '在庫数', '在庫金額',
                        '過剰在庫区分', '不動区分', '期限切迫区分', '有効期限', 'ロットNO',
                        '最終出庫日', '区分', '引取候補店', '参考:過剰だが使用中の店',
                        '6ヶ月出庫回数', '医薬品CD', '予約', '滞留']
    proposal_rows = [[r['出し手店'], r['種別'], r['薬品名'], r['単位'], r['メーカ名'], r['在庫数'],
                      r['在庫金額'], r['過剰在庫区分'], r['不動区分'], r['期限切迫区分'],
                      r['有効期限'], r['ロットNO'], r['最終出庫日'], r['区分'],
                      r['引取候補店'], r['参考:過剰だが使用中の店'],
                      r['6ヶ月出庫回数'], r['医薬品CD'],
                      r.get('予約', ''), r.get('滞留', '')] for r in result['proposal_rows']]

    shortage_headers = ['店', '薬品名', '在庫数', '安全在庫数', '不足数', '医薬品CD', 'デッド/期限切迫で持つ他店']
    shortage_rows = [[r['店'], r['薬品名'], r['在庫数'], r['安全在庫数'], r['不足数'],
                      r['医薬品CD'], r['デッド/期限切迫で持つ他店']] for r in result['shortage_rows']]

    matrix_headers = ['薬品名', '医薬品CD'] + result['store_names']
    matrix_rows = result['matrix_rows']

    summary_headers = ['店', 'デッド品目数', 'デッド金額計', '期限切迫品目数', '期限切迫金額計', '不足品目数']
    summary_rows = [[r['店'], r['デッド品目数'], r['デッド金額計'], r['期限切迫品目数'],
                     r['期限切迫金額計'], r['不足品目数']] for r in result['summary_rows']]

    return {
        'base_ym': base_ym_disp,
        'csv_base': csv_base_disp,
        'proposal_headers': proposal_headers, 'proposal_rows': proposal_rows,
        'shortage_headers': shortage_headers, 'shortage_rows': shortage_rows,
        'matrix_headers': matrix_headers, 'matrix_rows': matrix_rows,
        'summary_headers': summary_headers, 'summary_rows': summary_rows,
    }
