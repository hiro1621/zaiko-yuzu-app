# -*- coding: utf-8 -*-
"""
デッドストックリスト（streamlit_app.py）＝店舗セルフアップロード型の画面本体
  ※画面の名前は2026-07-28に「店舗間 在庫融通」から「デッドストックリスト」へ変更（本間部長指示）。
    フォルダ名・リポジトリ名・関数名は在庫融通のままなので、探すときは両方の呼び名で当たること。

【何をするアプリか】
各店がブラウザから薬VANの在庫ファイル（.xls / .csv / .xlsx）をアップロードすると、
共通エンジン（yuzu_core）が全店ぶんを毎回まとめて再計算し、
  ・自店（②）のデッド品／（③）の期限切迫品（それぞれ欲しがる店つき）
  ・自店（④）が引き取れる薬（他店のデッド・期限切迫のうち、引取候補店＝自店の品）＝見るだけの一覧
    （2026-07-27 追加。ボタン・申込みは無し。連絡は従来どおり電話・デスクネッツ）
  ・（参考）自店が不足している薬を、デッド／期限切迫で持っている店
  ・全店の融通提案一覧（デッド＋期限切迫・在庫金額の大きい順／「種別」列つき）
  ・「現在 N/14店 アップ済み」（店数は stores_config.py の STORE_COUNT）
を返します。裏側の保管庫はGoogleスプレッドシート（サービスアカウント経由のみ）。
店舗はGシートを直接触らず、必ずこのアプリ経由で読み書きします。

★出し手（融通候補）の定義（A案・本間部長承認 2026-07-27）★
  純粋な過剰（デッドでも期限切迫でもない、ただ多いだけの品）は提案から外し、
  「デッド（不動）」と「期限切迫」だけを出し手にします。

【使い方（店舗）】
  1. 共有パスワードを入れる（1回だけ）。
  2. 自分の店をドロップダウンで選ぶ。
     （URLに ?store=店名 が付くので、そのURLをブックマークすれば次回から選ばずに済む）
  3. その月ぶんをまだ出していなければ、薬VANの在庫ファイルをアップロードする。
     ★アップロードは月に1回でよい。データは保管庫（Gシート）に残るので、2回目以降は
       店を選ぶだけで表が見られる。当月ぶんが済んでいればアップロード欄は折りたたまれ、
       「アップ済みです／このまま下の表を見られます」と表示される（2026-07-28）。
  4. 結果が表示される。②③④は切替ボタンで1つずつ表示する（2026-07-28。3つを縦に
     並べると④まで延々スクロールが要るため。Excelボタンは切替の外＝常に一番下にある）。

★数量・金額の定義（本間部長判断 2026-07-27）★
  「在庫数」＝その店がいま持っている全量、「在庫金額」＝在庫数×薬価（薬VANの薬価金額列）。
  旧仕様の「過剰数／過剰数金額」（＝安全在庫を超えた分だけ）は使いません。デッド品は
  在庫まるごとが動かす対象で、過剰数だと0になってしまう品が実在するためです。

★載せない品（本間部長判断 2026-07-27）★
  1) 在庫金額が 1,500円未満の少額品（yuzu_core.CONFIG['min_supply_amount'] で変更可）
  2) 店が「除外」にチェックを入れた品（保管庫の _除外 タブに保存。いつでも戻せる）
  どちらも自店の表だけでなく、全店一覧・他店の参考ビュー・Excel・Gシートから同時に消えます。

★誤アップロード防止（2026-07-27 本間部長確定・追加）★
  アップロード前に、ファイル名の店名部分（例「さと和光_202607」の「さと和光」）と、画面で選んだ
  店名を「完全一致」で突き合わせ、別店のファイルや店名不明のファイルはアップロードボタンを止めます
  （認証が無く全店が同じURLを使うため、店を取り違えると別店のデータを丸ごと上書きしてしまう対策）。

【必要なライブラリのインストール（コマンドプロンプトで実行）】
    pip install streamlit gspread google-auth openpyxl xlrd==2.0.1
    ローカルで起動：  streamlit run streamlit_app.py

【秘密（Streamlit Cloud の Secrets ＝ TOML に置く。リポジトリには入れない）】
    app_password = "共有パスワード"
    spreadsheet_id = "GスプレッドシートのID"
    [gcp_service_account]
    ... サービスアカウント鍵(JSON)の中身をTOML表として貼る ...
  ※ Secrets が未設定のときは、この画面だけで動く「ローカル保管庫モード」で起動します
    （＝Gシートに接続せず、その場のブラウザのメモリに貯めます。動作確認用）。

※ venv不要。Windows専用パス。コメント・メッセージはすべて日本語です。
"""

import datetime
import hashlib

import streamlit as st
import pandas as pd   # ※ streamlit に同梱されるので requirements への追記は不要

import yuzu_core
import app_logic
import gsheet_store
import mailer
from stores_config import STORE_NAMES, STORE_COUNT, COMPANY_OF

# ⑤やり取りの相手店の並び順を『店番順（STORE_NAMES の並び）』に固定するための順位表。
#   ★並びを固定にするのが要件：未読件数や投稿で順番が変わると、st.dataframe の行選択が
#     「行の位置」で効くため、選んだ行がずれて別の店を開いてしまう事故になる（本間部長指示 2026-08-10）。
#   STORE_NAMES に無い店名（将来の追加漏れ・表記ゆれ）は末尾へまとめ、その中は五十音順にする
#   （黙って消さない）。
_STORE_ORDER_RANK = {name: i for i, name in enumerate(STORE_NAMES)}


def _thread_sort_key(thread):
    """ ⑤の相手店スレッドを『店番順（STORE_NAMES）』で並べるためのキー。
        STORE_NAMES にある店は (0, 店番順位, '')、無い店は (1, 0, 相手店名) を返す
        ＝登録店を先に店番順、未登録店は末尾で五十音順（黙って消さない）。 """
    name = thread.get('相手店名', '')
    if name in _STORE_ORDER_RANK:
        return (0, _STORE_ORDER_RANK[name], '')
    return (1, 0, name)

# ドロップダウンの先頭に置く「未選択」の選択肢
SENTINEL_STORE = '選択してください'


# ============================================================================
# 画面の基本設定
# ============================================================================
st.set_page_config(page_title='デッドストックリスト', page_icon='💊', layout='wide')


# ============================================================================
# Secrets（秘密）を安全に読む
# ============================================================================
def _get_secret(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default


def gsheet_configured():
    """ Gシート接続に必要な秘密が全部そろっているか。 """
    return bool(_get_secret('spreadsheet_id')) and bool(_get_secret('gcp_service_account'))


# ============================================================================
# Googleシートへのアクセス回数を減らす（2026-07-28）
#   ★★背景：Streamlitは操作のたびにスクリプト全体を作り直す。表の行を1つ選ぶだけでも
#     （on_select='rerun'）1回まるごと走るため、素直に書くと1クリックごとに
#       _index 1回 ＋ raw_<店> 店数ぶん ＋ _除外 ＋ _予約 ＋ 結果4タブの書き戻し8回
#     ＝14店なら30回近いAPI呼び出しになる。GoogleのAPIは1分60回が上限で、
#     そこに当たると gspread が待ち続け、画面が固まったように見える（実際に発生）。
#   そこで次の2つを入れて、ふつうの操作（行を選ぶ・画面を切り替える）では
#   Googleシートをほとんど触らないようにする。
# ============================================================================
def _signature(obj):
    """ 中身が変わったかどうかを見るための短い指紋。 """
    return hashlib.md5(repr(obj).encode('utf-8', 'replace')).hexdigest()


def load_stores_cached(backend):
    """
    当月の全店データを読む。ただし **_index が前回と同じなら読み直さない**。

    ★これは「古いデータを見せる」ことにはならない。店がアップロードすると必ず
      _index の行（アップ日時・行数・ファイル名）が書き換わるので、_index が
      1文字も変わっていなければ raw_ の中身も変わっていない、と言い切れるため。
      念のため、アップロードを保存した直後にはこのキャッシュを明示的に捨てている
      （同じ分・同じ行数・同じファイル名で上げ直した場合の取りこぼしを防ぐ）。
    読むのは _index の1タブだけ（1回）で済むので、行を選ぶたびの読み込みが
    「店数ぶん」から「1回」に減る。
    """
    if not hasattr(backend, 'load_index'):
        return backend.load_current_month_stores()      # 古い保管庫でもそのまま動く
    index = backend.load_index()
    sig = _signature(sorted(index.items()))
    cache = st.session_state.get('stores_cache')
    if cache and cache.get('sig') == sig:
        return cache['stores'], cache['latest'], index
    # ★読み直すときは、いま読んだ _index をそのまま渡す（同じタブを2回読まない・2026-07-29）
    stores, latest, index = backend.load_current_month_stores(index)
    st.session_state['stores_cache'] = {'sig': sig, 'stores': stores, 'latest': latest}
    return stores, latest, index


def load_prev_cached(backend, latest):
    """
    前月の結果（滞留の判定材料）を読む。**同じ月のあいだは1回しか読まない**。

    ★キャッシュしてよい理由：読む相手は 前月_YYYYMM タブで、これは月替わりのときにしか
      作られない＝当月のあいだは中身が変わらないため。行を選ぶたびに読み直すと、
      Googleの1分あたりの上限に近づくだけで得るものが無い。
    読めなかったときは空の辞書を返す（＝全部「今月から」扱い。画面は普通に出る）。
    """
    if not hasattr(backend, 'load_prev_proposal'):
        return {}
    cache = st.session_state.get('prev_cache')
    if cache and cache.get('ym') == latest:
        return cache['map']
    try:
        m = backend.load_prev_proposal(latest)
    except Exception as e:
        show_gsheet_error(e, '前月の記録を読めませんでした（滞留の表示は省きます）', 'warning')
        m = {}
    st.session_state['prev_cache'] = {'ym': latest, 'map': m}
    return m


def show_gsheet_error(e, what, level='error'):
    """
    Googleシートのエラーを画面に出す。1分あたりの上限（429）に当たったときは、
    英語のAPIエラーをそのまま見せず、店が取るべき行動だけを日本語で伝える。
      ★gsheet_store 側で2秒→5秒→10秒の自動やり直しを入れてあるので、
        ここまで来るのは「かなり混み合っている」ときだけ。
      level='warning' … 画面表示は続けられる（読めなくても致命的でない）場面用。
    """
    s = str(e)
    out = st.warning if level == 'warning' else st.error
    if ('429' in s) or ('Quota exceeded' in s) or ('RATE_LIMIT_EXCEEDED' in s):
        out('ただいま混み合っています（Googleスプレッドシートの1分あたりの上限）。'
            '**1分ほど待ってから、ページを再読み込み（F5）してください。**'
            'データは無事です。何度も再読み込みすると、かえって開けなくなります。')
    else:
        out('%s：%s' % (what, e))


def clear_stores_cache():
    """ アップロードを保存した直後に呼ぶ（次の描画で必ず読み直させる）。 """
    st.session_state.pop('stores_cache', None)
    st.session_state.pop('results_sig', None)
    # 月替わりのアップロードでは 前月_YYYYMM タブが新しく作られるので、前月の記録も捨てる。
    #   （当月の年月が変わればキャッシュのキーも変わるが、念のためここでも落としておく）
    st.session_state.pop('prev_cache', None)
    # 案1・案2（2026-08-12）：アップロード後は結果キャッシュと表示用リストキャッシュも捨てて、
    #   新しい在庫ですぐ計算し直させる（署名にも index が入っているので二重の保険）。
    st.session_state.pop('results_cache', None)
    st.session_state.pop('lists_cache', None)


# ============================================================================
# やり取り（掲示板）のセッションキャッシュ（2026-08-10 第2弾）
#   ★★背景：Google Sheets API はサービスアカウント合計で1分60回が上限。ふつうの操作
#     （行を選ぶ・画面を切り替える）のたびに _やり取り と _やり取り既読 を読むと上限に近づき、
#     「予約するとフリーズする」類の不具合が再発する。
#   → この2タブは 60秒のセッションキャッシュ越しに読む（load_prev_cached と同じ型）。
#     ふだんの再描画では読み直さず（＝+0回）、投稿・既読更新の直後だけキャッシュを捨てて読み直す。
#     待てないときのために ⑤の上に「最新に更新」ボタンを置き、手で読み直せるようにする。
# ============================================================================
# 投稿日時・既読日時の形式。ゼロ詰め固定幅なので文字列比較でそのまま時系列になる。
#   秒まで持つのは、同じ分に届いた投稿を「既読にした瞬間より後」と正しく判定できるようにするため。
MSG_TS_FMT = '%Y/%m/%d %H:%M:%S'
_MSG_CACHE_TTL = 60   # 秒


def load_messages_cached(backend, force=False):
    """ _やり取り・_やり取り既読 を 60秒のセッションキャッシュ越しに読む。
        force=True でキャッシュを無視して読み直す（「最新に更新」ボタン・投稿/既読更新の直後）。
        戻り値：(messages, msg_reads)。読めないタブは（あれば）前回値、無ければ空リストで返す。 """
    import time as _time
    if not hasattr(backend, 'load_messages'):
        return [], []      # 古い保管庫でも落ちない
    now = _time.time()
    cache = st.session_state.get('messages_cache')
    if (not force) and cache and (now - cache.get('t', 0) < _MSG_CACHE_TTL):
        return cache['messages'], cache['reads']
    prev = cache or {}
    try:
        messages = backend.load_messages()
    except Exception as e:
        show_gsheet_error(e, 'やり取りを読めませんでした', 'warning')
        messages = prev.get('messages', [])
    try:
        reads = backend.load_msg_reads()
    except Exception as e:
        show_gsheet_error(e, 'やり取りの既読情報を読めませんでした', 'warning')
        reads = prev.get('reads', [])
    st.session_state['messages_cache'] = {'t': now, 'messages': messages, 'reads': reads}
    return messages, reads


def clear_messages_cache():
    """ 投稿・既読更新の直後に呼ぶ（次の描画でやり取りを必ず読み直させる）。 """
    st.session_state.pop('messages_cache', None)


# ============================================================================
# 表示用の3リスト（除外・予約・出庫可能数）を60秒キャッシュ越しに読む（2026-08-12・案2）
#   ★★背景：チェックを1つ入れるたびの再描画で、除外・予約・出庫可能数を毎回読むと
#     Google Sheets API の1分60回上限に近づき、画面が固まる一因になる（品質管理部の実測で
#     チェック1つ＝往復4回＝_index＋除外＋予約＋出庫可能数）。
#   → この3タブは60秒のセッションキャッシュ越しに“表示のためだけ”読む（load_messages_cached と同じ型）。
#     ふだんの再描画では読み直さず（＝+0回）、保存の直後だけ clear_lists_cache で捨てて読み直す。
#
#   ★★★絶対に守る一線（二重予約の防止）：
#     「保存の直前に、いまの本当の値を突き合わせるための読み直し」——_save_reservations の突合、
#     _save_supply_qty_ui・_clear_all_supply_qty_ui、_save_exclusions_ui、予約取消——は
#     いまどおり backend.load_*() を直接呼ぶこと（このキャッシュを絶対に通さない）。
#     ここをキャッシュ越しにすると、2つの店が同じ品を同時に予約できてしまう。
#     このキャッシュは「画面に出すための読み」だけに使う。
# ============================================================================
_LISTS_CACHE_TTL = 60   # 秒


def load_lists_cached(backend, force=False):
    """ 除外・予約・出庫可能数の3リストを60秒のセッションキャッシュ越しに読む（表示専用）。
        戻り値：(exclusions, reservations, supply_rows)。
        読めないリストは（あれば）前回値、無ければ空リストで返す（画面は普通に出る）。 """
    import time as _time
    now = _time.time()
    cache = st.session_state.get('lists_cache')
    if (not force) and cache and (now - cache.get('t', 0) < _LISTS_CACHE_TTL):
        return cache['exclusions'], cache['reservations'], cache['supply']
    prev = cache or {}
    try:
        exclusions = backend.load_exclusions()
    except Exception as e:
        show_gsheet_error(e, '除外リストを読めませんでした（除外なしで表示します）', 'warning')
        exclusions = prev.get('exclusions', [])
    try:
        reservations = backend.load_reservations()
    except Exception as e:
        show_gsheet_error(e, '予約リストを読めませんでした（予約なしで表示します）', 'warning')
        reservations = prev.get('reservations', [])
    try:
        supply = backend.load_supply_qty()
    except Exception as e:
        show_gsheet_error(e, '提供数量を読めませんでした（全量で表示します）', 'warning')
        supply = prev.get('supply', [])
    st.session_state['lists_cache'] = {'t': now, 'exclusions': exclusions,
                                       'reservations': reservations, 'supply': supply}
    return exclusions, reservations, supply


def clear_lists_cache():
    """ 除外・予約・出庫可能数を保存した“直後”に呼ぶ（次の描画で必ず読み直させる）。
        あわせて結果キャッシュ（案1）も捨てる＝材料が変わったので照合とExcelを作り直させる。
        ★これを呼ばないと、保存した本人が自分の変更を最大60秒見られない。 """
    st.session_state.pop('lists_cache', None)
    st.session_state.pop('results_cache', None)


def _fmt_msg_time(s):
    """ 投稿日時（YYYY/MM/DD HH:MM:SS 等）を画面用の『MM/DD HH:MM』にする。
        形式が違う・短い値はそのまま返す（黙って落とさない）。 """
    s = str(s or '')
    if len(s) >= 16 and s[4:5] == '/':
        return s[5:16]      # 'MM/DD HH:MM'
    return s


# ============================================================================
# 保管庫バックエンド（Gシート or ローカル）を用意する
#   ・両者は同じメソッド名（save_store_upload / load_current_month_stores）を持つ。
# ============================================================================
class _GSheetAdapter:
    """ gsheet_store（モジュール関数）を、ローカル保管庫と同じ形で呼べるようにする薄いラッパ。 """

    def __init__(self, sh):
        self.sh = sh

    def save_store_upload(self, store_name, ym, slim_rows, filename, format_ok):
        return gsheet_store.save_store_upload(self.sh, store_name, ym, slim_rows, filename, format_ok)

    def load_current_month_stores(self, index=None):
        return gsheet_store.load_current_month_stores(self.sh, index)

    def write_results(self, payload):
        gsheet_store.write_results(self.sh, payload)

    def load_exclusions(self):
        return gsheet_store.read_exclusions(self.sh)

    def save_exclusions(self, rows):
        gsheet_store.write_exclusions(self.sh, rows)

    def load_reservations(self):
        return gsheet_store.read_reservations(self.sh)

    def save_reservations(self, rows):
        gsheet_store.write_reservations(self.sh, rows)

    def load_supply_qty(self):
        return gsheet_store.read_supply_qty(self.sh)

    def save_supply_qty(self, rows):
        gsheet_store.write_supply_qty(self.sh, rows)

    def load_messages(self):
        return gsheet_store.read_messages(self.sh)

    def append_message(self, row):
        gsheet_store.append_message(self.sh, row)

    def load_msg_reads(self):
        return gsheet_store.read_msg_reads(self.sh)

    def save_msg_reads(self, rows):
        gsheet_store.write_msg_reads(self.sh, rows)

    def load_index(self):
        """ _index だけを読む（1タブ）。在庫本体（raw_）を読み直すかどうかの判定に使う。 """
        return gsheet_store.read_index(self.sh)

    def load_prev_proposal(self, current_ym=None):
        """ 前月の結果（滞留＝何ヶ月つづけて載っているか の判定材料）を読む。 """
        return gsheet_store.read_prev_proposal(self.sh, current_ym)


@st.cache_resource(show_spinner='Googleシートに接続しています…')
def _open_sheet():
    """ Gシートを開く（重い接続なのでキャッシュ）。失敗したら例外。 """
    sa_info = _get_secret('gcp_service_account')
    ssid = _get_secret('spreadsheet_id')
    return gsheet_store.open_spreadsheet(sa_info, ssid)


def get_backend():
    """ Secrets があればGシート、無ければローカル保管庫を返す。 """
    if gsheet_configured():
        sh = _open_sheet()
        return _GSheetAdapter(sh), True
    # ローカル保管庫（このブラウザのセッションに貯める）
    if 'local_state' not in st.session_state:
        st.session_state['local_state'] = {}
    return app_logic.LocalBackend(st.session_state['local_state']), False


# ============================================================================
# パスワードゲート
# ============================================================================
def password_gate():
    """ 正しい共有パスワードを入れるまで先へ進ませない。 """
    if st.session_state.get('authed'):
        return True

    st.title('💊 デッドストックリスト')
    st.caption('社内限定ツール。共有パスワードを入力してください。')

    expected = _get_secret('app_password')
    with st.form('gate'):
        pw = st.text_input('共有パスワード', type='password')
        ok = st.form_submit_button('入る')
    if ok:
        if not expected:
            # Secrets未設定（ローカル検証・開発モード）では、そのまま入れる
            st.session_state['authed'] = True
            st.rerun()
        elif pw == str(expected):
            st.session_state['authed'] = True
            st.rerun()
        else:
            st.error('パスワードが違います。')
    if not expected:
        st.info('（開発モード：共有パスワード未設定のため、空欄のまま「入る」で進めます）')
    return False


# ============================================================================
# 表示ヘルパ
# ============================================================================
def _df(rows, columns=None):
    """ 辞書リスト → 列順を保った DataFrame。空でも見出しだけは出す。 """
    if not rows:
        return pd.DataFrame(columns=columns or [])
    df = pd.DataFrame(rows)
    if columns:
        df = df[[c for c in columns if c in df.columns]]
    return df


# 表の数値列の体裁。★2026-08-10（第2弾）で「数量系」と「金額系」を分けた（本間部長指示）。
#   ・数量系（_QTY_COLS）… 末尾の .00 を落として表示する（77.00→77／12.50→12.5）。カンマ区切りは維持。
#       ほとんどが整数なので「.00」が並ぶと見づらい、という指摘への対応。フォーマッタは _fmt_qty。
#   ・金額系（_AMOUNT_COLS）… 従来どおり小数第2位のまま（例 20,759.20）。金額は小数2桁が自然なため据え置き。
#   ※どちらも【画面表示だけ】の体裁。Excel の number_format（#,##0.00）は一切変更していない。
_QTY_COLS = ['在庫数', '出庫可能数', '安全在庫数', '不足数', '出し手の出庫可能数', '自店の在庫数']
_AMOUNT_COLS = ['在庫金額']


def _fmt_qty(v):
    """ 数量の表示。末尾の .00 は落とす（77.00→77／12.50→12.5／12.25→12.25）。カンマ区切りは維持する。
        ・整数（77.0 など）は小数点以下を付けずカンマ区切りで返す（1234.0→1,234）。
        ・端数がある値は小数2桁にしてから末尾の 0 だけ落とす（12.50→12.5）。
          整数は上の分岐で処理済みなので、ここに「12.00」が来ることはなく、rstrip('0') だけで足りる
          （末尾の『.』が残る心配がない）。
        ・数値に直せない値（空文字など）はそのまま返す。 """
    try:
        f = float(v)
    except (TypeError, ValueError):
        return v
    if f != f:                              # NaN はそのまま返す（int(NaN) で落ちるのを防ぐ）
        return v
    if f == int(f):
        return '{:,}'.format(int(f))
    return '{:,.2f}'.format(f).rstrip('0')


def _fmt_amount(v):
    """ 金額の表示。カンマ区切りの小数第2位（例 20,759.20）。空欄はそのまま空欄で返す。

        ★以前はフォーマット文字列 '{:,.2f}' を Styler に直接渡していたが、それだと画面が落ちた。
          再現手順：予約された品を、出し手の店が一覧から外す
                  → 受け手の④「予約中の品」で在庫金額が空文字になる
                  → '{:,.2f}'.format('') が ValueError で落ちて、画面ごと真っ白になる。
          このエラーは _style_expiry の try/except では拾えない。Styler.format は
          その場では計算せず、表を描くときに初めて動くので、例外が try の外で起きるため。
          だからフォーマッタ側で空欄を受け止める（2026-08-10 品質管理部が再現・裏取り済み）。
          ※この不具合は掲示板の改修より前から存在していた（改修前のファイルでも再現）。 """
    if v is None:
        return ''
    try:
        f = float(v)
    except (TypeError, ValueError):
        return v                            # 空文字や文字列はそのまま返す
    if f != f:                              # NaN はそのまま返す
        return v
    return '{:,.2f}'.format(f)


def _style_expiry(df, paint=True, stag_levels=None, expiry_flags=None):
    """ 表を見やすく整えた pandas Styler を返す。
          ・数量列（出庫可能数など）は末尾の .00 を落として表示（例 77／12.5・カンマ区切り）。
            金額列（在庫金額）は従来どおり小数第2位のまま（例 49,630.20・カンマ区切り）
          ・paint=True のとき、「有効期限まで5ヶ月以内」の行を【薄赤の背景＋黒文字】で行ごと目立たせる
            （②デッド一覧の中で"期限が近いデッド"が一目で分かるように）。
            どの行が5ヶ月以内かは expiry_flags（行ごとの True/False の配列。滞留色の stag_levels と
            同じ「行ごとの配列を渡す」方式）で受け取る＝判定は yuzu_core で1度だけ行い、画面は
            その結果を塗るだけにして定義をズラさない（以前は期限切迫区分の非空で塗っていた）。
          ・paint=False のときは背景を塗らない（③期限切迫品は全行が期限切迫で、
            塗ると表全体が赤くなるだけで情報量がゼロのため。数値フォーマットは②③とも効かせる）。
          ・stag_levels（行ごとの滞留区分のリスト）を渡すと、「滞留」列のセルだけを
            滞留区分の色で塗る。★行の色（＝期限のこと）とは別のことを表しているので、
            行全体ではなくセル1つだけを塗って意味を混ぜないようにしている。
        ★文字色を黒（#000000）で明示するのは、ダークテーマだと白文字×薄赤で読めなくなるため。
          ライトテーマでは元から黒文字なので見た目は変わらない。
        Styler が使えない環境では素の DataFrame をそのまま返す。 """
    if df is None or df.empty:
        return df

    # 薄赤背景＋黒文字。行全体を塗るので有効期限も一緒に目立つ。
    # 色は #FFE3E6（旧 #FFC7CE より薄いピンク＝現行色と白のちょうど中間）。
    HIT_STYLE = 'background-color: #FFE3E6; color: #000000'

    def _paint(row):
        # 行の位置 i（既定インデックスなので row.name＝0始まりの位置）で expiry_flags を引く。
        #   ★「期限切迫区分が非空か」ではなく「その行が有効期限まで5ヶ月以内か」で塗る（A案）。
        i = row.name
        hit = bool(expiry_flags[i]) if (expiry_flags and i < len(expiry_flags)) else False
        return [HIT_STYLE if hit else '' for _ in row]

    def _paint_stag(col):
        """ 「滞留」列のセルを、行ごとの滞留区分の色で塗る。 """
        out = []
        for i in range(len(col)):
            lv = stag_levels[i] if i < len(stag_levels) else 'new'
            bg, fg, _x, _d = yuzu_core.STAGNATION_STYLES.get(lv, ('', '', None, ''))
            out.append(('background-color: %s; color: %s' % (bg, fg)) if bg else '')
        return out

    try:
        sty = df.style
        # 数量系は _fmt_qty（.00 を落とす）／金額系は _fmt_amount（小数第2位のまま）で分けて整える。
        # ★どちらも「数値に直せない値はそのまま返す」フォーマッタにしてある。
        #   表の一部が空欄でも画面を落とさないため（_fmt_amount の説明を参照）。
        fmt = {}
        for c in df.columns:
            if c in _QTY_COLS:
                fmt[c] = _fmt_qty
            elif c in _AMOUNT_COLS:
                fmt[c] = _fmt_amount
        if fmt:
            sty = sty.format(fmt)
        if paint and expiry_flags:
            sty = sty.apply(_paint, axis=1)
        if stag_levels and '滞留' in df.columns:
            sty = sty.apply(_paint_stag, subset=['滞留'])
        return sty
    except Exception:
        # 万一 Styler が使えない環境でも、表示自体は素のDataFrameで続行する
        return df


def stagnation_legend(rows, chips=True):
    """
    「滞留」列の意味を画面に出す。
      chips=True （既定）… 色チップつき（色を見ただけで意味が分かる）。★色つきの st.dataframe を
                          使う④・予約中・除外中の表で使う（表の色とチップの色が対応する）。
      chips=False        … 色チップをやめ、文字だけ（区分名＋件数）で出す。★②③は st.data_editor に
                          切り替えて表に色が無いので、色チップだけ残すと表と対応が取れず誤解を招く。
                          そのため文字だけの説明にする（2026-08-10・本間部長指示）。
      rows … いま表示している行リスト。実際に出ている区分だけを並べるので、
             関係のない説明で画面がうるさくならない。
    1件も滞留が無い月は何も出さない。
    """
    counts = app_logic.stagnation_summary(rows)
    if not counts:
        return
    if not chips:
        # ②③用：色を使わず、区分名＋件数の文字だけで出す。
        parts = []
        for lv in yuzu_core.STAGNATION_LEGEND_ORDER:
            n = counts.get(lv)
            if not n:
                continue
            _bg, _fg, _x, desc = yuzu_core.STAGNATION_STYLES[lv]
            parts.append('%s：%d件' % (desc, n))
        st.caption('「滞留」の意味　' + '／'.join(parts))
        return
    chips_html = []
    for lv in yuzu_core.STAGNATION_LEGEND_ORDER:
        n = counts.get(lv)
        if not n:
            continue
        bg, fg, _x, desc = yuzu_core.STAGNATION_STYLES[lv]
        chips_html.append(
            '<span style="background-color:%s;color:%s;padding:2px 8px;'
            'border-radius:4px;border:1px solid #BBB;white-space:nowrap;">'
            '%s <b>%d件</b></span>' % (bg, fg, desc, n))
    st.markdown(
        '<div style="font-size:0.86rem;line-height:2.1;">'
        '<b>「滞留」列の色の見方</b>　' + '　'.join(chips_html) + '</div>',
        unsafe_allow_html=True)


# ============================================================================
# 除外（この品は融通に出さない）の表示と保存
# ============================================================================
def _exclusion_set(exclusions):
    """ 保存されている除外リスト → {(店名, 除外キー), ...} の集合（エンジンに渡す形）。 """
    return {(r['店名'], r['除外キー']) for r in exclusions}


def _selected_rows(event):
    """ st.dataframe(on_select='rerun') の戻りから、選択された行の位置（0始まりの整数リスト）を安全に取り出す。 """
    try:
        return list(event.selection.rows)
    except Exception:
        try:
            return list(event['selection']['rows'])
        except Exception:
            return []


def _to_float(v, default=0.0):
    """ 文字列・数値を float にする（桁区切りカンマ・前後空白は無視）。変換できなければ default。 """
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return default


def _save_exclusions_ui(checked, my_store, backend, exclusions):
    """ 選ばれた品を除外（デッドストックから外す）に保存する。従来の「除外を保存」の中身。
        ★予約が入っている品は除外しない（引き取るつもりの店の予約が理由も分からず消えるため）。 """
    now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')
    # ★★保存の直前に、除外リストを保管庫から“直接”読み直してから足す（2026-08-12）。
    #   表示は60秒キャッシュ（load_lists_cached）越しだが、書き込みの土台は必ず最新にする。
    #   これをしないと、直前60秒の間にほかの店が入れた除外を、丸ごと上書きで消してしまう。
    try:
        exclusions = backend.load_exclusions()
    except Exception:
        pass   # 読み直せなければ、渡された表示用の一覧を土台にする（従来どおり動く）
    keep = list(exclusions)
    have = {(r['店名'], r['除外キー']) for r in keep}
    # ★予約が入っている品は除外しない。品名と相手店を出して、話を付けてもらう。
    booked = [d for d in checked if d.get('_予約店')]
    if booked:
        st.warning('  \n'.join(
            ['次の品はすでに引取先が決まっているため、除外しませんでした。'
             '取りやめる場合は相手店に連絡してください：']
            + ['%s → %s が引取予定' % (d['薬品名'], d['_予約店']) for d in booked]))
    added = 0
    for d in checked:
        if d.get('_予約店'):
            continue
        pair = (my_store, d['_key'])
        if pair in have:
            continue
        keep.append({'店名': my_store, '除外キー': d['_key'],
                     '薬品名': d['薬品名'], '除外日時': now})
        have.add(pair)
        added += 1
    backend.save_exclusions(keep)
    clear_lists_cache()   # 保存直後：表示用キャッシュと結果キャッシュを捨て、次の描画ですぐ反映する
    st.success('%d件をデッドストックから外しました。' % added)
    st.rerun()


def _flash(level, text):
    """ 保存直後の st.rerun をまたいでメッセージを見せるため、いったん session_state にためる。
        Streamlit は st.rerun でその回の描画を捨てるので、rerun 前に出したメッセージは消える。
        ためておき、次の描画の先頭で _render_flash が1回だけ出す。 """
    st.session_state.setdefault('_supqty_flash', []).append((level, text))


def _render_flash():
    """ ためておいたメッセージ（_flash）を画面の先頭で1回だけ出して消す。 """
    for level, text in st.session_state.pop('_supqty_flash', []):
        getattr(st, level, st.info)(text)


# ============================================================================
# メール通知（mailer.py）との橋渡し … 2026-08-10（③の改修）
#   ・掲示板・予約の保存が成功した“直後”に1回だけ呼ぶ（下の各セクション参照）。
#     st.rerun をまたぐと画面メッセージが消えるので、_flash にためて次の描画で出す
#     （＝メール結果の案内・警告が確実に見える）。
#   ・メールはおまけ。送れなくても投稿・予約は必ず残す（例外は外へ出さない）。
# ============================================================================
def _mail_secrets():
    """ Secrets を安全に取り出す（無い環境でも落ちない）。中身の判定は mailer 側が行う。 """
    try:
        return st.secrets
    except Exception:
        return {}


def _mail_flush(res):
    """ mailer の戻り（{'messages': [(レベル, 文言), ...]}）を _flash に流す。
        [smtp] 未設定なら 'メール通知は未設定です。' が1行入っている。 """
    for level, text in ((res or {}).get('messages') or []):
        _flash(level, text)


def _bump_editor(table_key):
    """ 保存・取消のあと data_editor を作り直すため、版番号を1つ上げる（＝キーが変わる）。
        ★これをしないと、除外で行が減った後も古い編集差分（行の位置で持つ）が残り、別の行へ
          誤って適用される（st.data_editor の既知の落とし穴）。 """
    k = 'supqty_ver_%s' % table_key
    st.session_state[k] = st.session_state.get(k, 0) + 1


def _validate_supply_desired(rows, desired, now=''):
    """
    セル直接編集された『出庫可能数』（desired: {_key: 編集後の数値}）を検証する純関数
    （st に依存しない＝単体テスト可能）。NumberColumn の max_value は列に1つしか指定できず、
    行ごとの在庫上限にできないので、その担保をここで行う。

      戻り値 …
        {'picked':  plan_supply_qty に渡す品のリスト（値が変わった品だけ）。各要素
                    {'_key','薬品名','出せる数','在庫数','_予約店','_now'},
         'clamped': [(薬品名, 在庫全量), ...]  在庫（全量）を超えたので全量まで下げた品,
         'zero':    [薬品名, ...]              0以下で保存しなかった品}

    ルール（本間部長指示 2026-08-10）：
      ・値が変わっていない品は picked に入れない（無駄な保存・警告・rerun を避ける）。
      ・予約が入っている品（_予約店 が非空）は、値が変わっていれば picked に入れる
        （数量は plan_supply_qty が blocked にして変えない。0以下・在庫超過の判定はしない）。
      ・0以下は保存しない＝zero に入れ、呼び出し側が「除外を使って」と案内する。
      ・在庫（全量）を超える値は在庫（全量）まで自動で下げ、clamped に入れて知らせる
        （全量と同値になるので plan_supply_qty 側で指定が消える＝全量に戻る）。
    """
    picked, clamped, zero = [], [], []
    for r in (rows or []):
        key = (r.get('_key', '') or '').strip()
        if not key:
            continue
        stock = _to_float(r.get('在庫数', 0))               # 在庫数（全量）
        cur = _to_float(r.get('出庫可能数', stock), stock)   # いまの実効数量（未指定なら全量）
        q = _to_float(desired.get(key, cur), cur)
        if q != q:            # NaN（セルを空にした等）は現在値に戻す
            q = cur
        reserved = (r.get('_予約店', '') or '').strip()
        name = r.get('薬品名', '')
        if reserved:
            # 予約品：値が変わっていれば plan に渡して blocked にしてもらう（数量は変わらない）
            if round(q, 2) != round(cur, 2):
                picked.append({'_key': key, '薬品名': name, '出せる数': q,
                               '在庫数': stock, '_予約店': reserved, '_now': now})
            continue
        if q <= 0:
            zero.append(name)
            continue
        if stock > 0 and round(q, 2) > round(stock, 2):
            clamped.append((name, stock))
            q = stock
        if round(q, 2) == round(cur, 2):
            continue          # 変更なし
        picked.append({'_key': key, '薬品名': name, '出せる数': q,
                       '在庫数': stock, '_予約店': '', '_now': now})
    return {'picked': picked, 'clamped': clamped, 'zero': zero}


def _save_supply_qty_ui(rows, my_store, backend, desired):
    """ セル直接編集された『出庫可能数』を保存する（純関数 app_logic.plan_supply_qty を呼ぶ薄い皮）。
        ★2026-08-10（第3弾）：行選択＋number_input の2段方式から st.data_editor へ切り替えたのに伴い、
          「編集後の数量マップ desired を検証してから保存する」形に作り替えた。
        ・在庫超過の自動クランプ・0以下の拒否は _validate_supply_desired で行う（黙って直さない）。
        ・予約が入っている品は plan_supply_qty の blocked の仕組みでそのまま守る（数量を変えない）。
        ・メッセージは _flash にためてから（呼び出し側が）st.rerun する（rerun で消えないように）。
        ・保存の直前に _提供数量 を読み直してから書く（同時書き込み対策・従来どおり）。 """
    now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')
    v = _validate_supply_desired(rows, desired, now=now)
    # 在庫（全量）を超えた品は全量まで下げたことを知らせる（黙って直さない）
    for name, stock in v['clamped']:
        _flash('info', '「%s」は在庫%sを超えていたため%sにしました。'
               % (name, _fmt_qty(stock), _fmt_qty(stock)))
    # 0以下は保存しない案内
    if v['zero']:
        _flash('warning', '  \n'.join(
            ['次の品は0以下のため保存しませんでした。'
             '一切出さない品は「除外」にチェックを入れてください：']
            + ['・%s' % n for n in v['zero']]))
    try:
        latest_rows = backend.load_supply_qty()   # ★書く直前に読み直す（同時書き込み対策）
    except Exception:
        latest_rows = []
    plan = app_logic.plan_supply_qty(latest_rows, my_store, v['picked'])
    if plan['blocked']:
        _flash('warning', '  \n'.join(
            ['次の品はすでに引取先が決まっているため、出庫可能数を変えませんでした。'
             '取りやめる場合は相手店に連絡してください：']
            + ['%s → %s が引取予定' % (b['薬品名'], b['予約店']) for b in plan['blocked']]))
    changed = plan['added'] + plan['updated'] + plan['removed']
    if changed:
        backend.save_supply_qty(plan['keep'])
        clear_lists_cache()   # 保存直後：表示用キャッシュと結果キャッシュを捨て、次の描画ですぐ反映する
        parts = []
        if plan['added']:
            parts.append('新規%d件' % plan['added'])
        if plan['updated']:
            parts.append('変更%d件' % plan['updated'])
        if plan['removed']:
            parts.append('全量に戻す%d件' % plan['removed'])
        _flash('success', '出庫可能数を保存しました（%s）。' % '／'.join(parts))
    elif not (v['clamped'] or v['zero'] or plan['blocked']):
        _flash('info', '出庫可能数の変更はありませんでした。')


def _clear_all_supply_qty_ui(my_store, backend):
    """ この店の『出庫可能数』の指定を一括で消す（全品を在庫まるごと＝全量に戻す）。
        ★行選択が無くなったため、旧「選んだ行を全量に戻す」の代わりに新設（2026-08-10）。
          全量に戻すのは供給量を増やす向きなので、予約が入っている品があっても相手の不利にならない。 """
    try:
        latest_rows = backend.load_supply_qty()   # ★書く直前に読み直す（同時書き込み対策）
    except Exception:
        latest_rows = []
    keep = [r for r in latest_rows if (r.get('店名', '') or '').strip() != my_store]
    removed = len(latest_rows) - len(keep)
    if removed:
        backend.save_supply_qty(keep)
        clear_lists_cache()   # 保存直後：表示用キャッシュと結果キャッシュを捨て、次の描画ですぐ反映する
        _flash('success', 'この店の出庫可能数の指定を%d件すべて取り消し、全品を在庫まるごとに戻しました。'
               % removed)
    else:
        _flash('info', '取り消す出庫可能数の指定はありませんでした。')


def _supply_reduced_caption(rows):
    """ 出庫可能数を在庫（全量）より減らしている品を、表の下にキャプションで出す。
        ★表に『在庫（全量）』列を足さない代わりの手当て（本間部長：2列並べると見づらい）。
          指定が0件なら何も出さない。数値は _fmt_qty で末尾の .00 を落とす。 """
    parts = []
    for r in (rows or []):
        stock = _to_float(r.get('在庫数', 0))
        eff = _to_float(r.get('出庫可能数', stock), stock)
        if stock > 0 and round(eff, 2) < round(stock, 2):
            parts.append('%s（在庫 %s → %s）'
                         % (r.get('薬品名', ''), _fmt_qty(stock), _fmt_qty(eff)))
    if parts:
        st.caption('出庫可能数を減らしている品：' + '、'.join(parts))


def supply_editor(rows, my_store, backend, exclusions, table_key,
                  paint_expiry=True, n_supply_specified=0, msg_by_store=None):
    """
    ②デッド品・③期限切迫品の表を「セル直接編集方式（st.data_editor）」で描く。
      rows              … app_logic.build_view_a / build_view_expiry の戻り
      table_key         … 画面部品を区別する名前（'dead' / 'expiry'）
      paint_expiry      … ②かどうか（True＝②）。②のときだけ、有効期限まで5ヶ月以内の品の
                          薬品名の先頭に ⚠ を付ける（③は全行が期限切迫なので付けない）。
      n_supply_specified… この店で現在『出庫可能数』の指定が入っている件数
                          （一括取消ボタンの表示・活性に使う）。
      msg_by_store      … {相手店名: 'N件 ●'} のやり取り早見表（2026-08-10 第2弾）。
                          予約が入っている品は、予約した店とのやり取り件数を『やり取り』列に出す。
                          ★この列は読み取り専用（disabled）＝表示だけで、編集は「除外」「出庫可能数」だけ。

    ★2026-08-10（第3弾・本間部長指示）★
      「セルをクリックして数を直接書き換えたい」という指示に応え、行選択＋number_input の2段方式を
      やめ、st.data_editor（表内で直接編集）に切り替えた。
      ・st.data_editor は pandas Styler の色（期限切迫の薄赤・滞留色）を表示できない。
        本間部長は「表を直接編集できるようにする（色は消える）」を選択済み（トレードオフ承知）。
      ・色が消えるぶんの情報の代替：
        - 期限切迫（②のみ）… 薬品名の先頭に ⚠（判定は yuzu_core の _expiry_flag をそのまま使う）
        - 滞留             … 『滞留』列の文字（2ヶ月目など）はそのまま残る
        - 滞留の色の意味   … stagnation_legend(..., chips=False) の文字だけの説明に切替
      ・列は：除外／薬品名／滞留／単位／出庫可能数／在庫金額／有効期限／期限切迫区分／区分／引取候補店。
        編集できるのは「除外」（チェック）と「出庫可能数」だけ。ほかは disabled で読み取り専用。
      ・NumberColumn は printf 系ではカンマ区切りにできない（2026-07-27確認）ので、読み取り専用の
        金額（在庫金額）はあらかじめ '20,759.20' の文字列にして TextColumn で出す。出庫可能数は
        編集させたいので NumberColumn。
        ★書式は format='%.10g'（77／12.5 のように末尾の .00 が出ない）。
          いったん format='localized' にしたが、localized は step（=0.01）から小数桁を決めるため
          「35.00」のように必ず2桁付いてしまい、本間部長の「.00は表示しない」指示に反した
          （2026-08-10 実機で確認して差し戻し）。桁区切りのカンマは出なくなるが、
          数量は3桁以下がほとんどで実害が小さいため、.00 を消す方を優先する。
      ・NumberColumn の max_value は列に1つしか指定できず行ごとの在庫上限にできないので指定せず、
        保存時に _validate_supply_desired で在庫（全量）まで自動クランプする。
      ・★返り値からは _key などの内部キーを取らず、必ず「行の位置（0始まり）」で元の rows から引く
        （column_config=None の隠し列が返るかは仕様依存で危険。2026-07-27 の記録）。
    """
    # 直前の保存・取消でためたメッセージを、表の上で1回だけ出す（st.rerun で消えないように）
    _render_flash()

    if not rows:
        st.info('該当する品目はありません。')
        return

    # ---- 表示用 DataFrame（行の順番＝rows の順番。突き合わせは行の位置で行う）----
    msg_by_store = msg_by_store or {}
    records = []
    for r in rows:
        name = r.get('薬品名', '')
        # ②で期限切迫（有効期限まで5ヶ月以内）の品は薬品名の先頭に ⚠（行の薄赤の代替）
        if paint_expiry and bool(r.get('_expiry_flag')):
            name = '⚠ ' + name
        # 予約が入っている品は、その予約店とのやり取り件数を出す（無ければ空）
        talk = msg_by_store.get((r.get('_予約店', '') or '').strip(), '')
        records.append({
            '除外': False,                                            # チェックで除外（編集可）
            '薬品名': name,                                           # 読み取り専用
            '滞留': r.get('滞留', ''),                                # 読み取り専用（色は無いが文字は残る）
            '単位': r.get('単位', ''),                                # 読み取り専用
            '出庫可能数': _to_float(r.get('出庫可能数', 0)),          # ★編集可（これが本命）
            '在庫金額': '{:,.2f}'.format(_to_float(r.get('在庫金額', 0))),  # カンマ付き文字列（読み取り専用）
            '有効期限': r.get('有効期限', ''),                        # 読み取り専用
            '期限切迫区分': r.get('期限切迫区分', ''),                # 読み取り専用
            '区分': r.get('区分', ''),                                # 読み取り専用
            'やり取り': talk,                                         # 読み取り専用（予約店とのやり取り件数）
            '引取候補店': r.get('引取候補店', ''),                    # 読み取り専用
        })
    df = pd.DataFrame(records, columns=[
        '除外', '薬品名', '滞留', '単位', '出庫可能数', '在庫金額',
        '有効期限', '期限切迫区分', '区分', 'やり取り', '引取候補店'])

    col_cfg = {
        '除外': st.column_config.CheckboxColumn(
            '除外',
            help='デッドストックリストから外す品にチェックを入れて「除外を保存」を押してください。'),
        '出庫可能数': st.column_config.NumberColumn(
            '出庫可能数',
            help='この店から出す数量です。数字をクリックして直接書き換えられます。'
                 '在庫（全量）と同じ数にすると「全量を出す」に戻ります'
                 '（在庫を超える数を入れると自動で全量まで下げます）。',
            min_value=0.01, step=0.01, format='%.10g'),
        '在庫金額': st.column_config.TextColumn('在庫金額'),
        'やり取り': st.column_config.TextColumn(
            'やり取り',
            help='予約が入っている品は、その相手店とのやり取り件数を出します（● は自店の未読あり）。'
                 '会話は⑤「やり取り」で見られます。'),
    }
    # 除外・出庫可能数だけ編集可。ほかは読み取り専用にする（やり取り列も必ず読み取り専用）。
    disabled = ['薬品名', '滞留', '単位', '在庫金額', '有効期限', '期限切迫区分', '区分',
                'やり取り', '引取候補店']

    # 保存・取消のたびに版番号でキーを変え、data_editor を作り直す（古い編集差分の誤適用を防ぐ）
    ver = st.session_state.get('supqty_ver_%s' % table_key, 0)
    editor_key = 'editor_%s_%d' % (table_key, ver)
    edited = st.data_editor(
        df, column_config=col_cfg, disabled=disabled,
        hide_index=True, width='stretch', num_rows='fixed', key=editor_key)

    # 滞留の色の意味は、色チップをやめて文字だけで出す（表に色が無いのにチップだけ色付きは誤解のもと）
    stagnation_legend(rows, chips=False)

    # ---- 編集結果を「行の位置」で元の rows と突き合わせる（隠し列の値に頼らない）----
    checked = []       # 除外にチェックが入った品（元の rows の辞書）
    desired = {}       # {_key: 編集後の出庫可能数}
    for i, r in enumerate(rows):
        try:
            row_e = edited.iloc[i]
        except Exception:
            continue
        if bool(row_e.get('除外', False)):
            checked.append(r)
        key = (r.get('_key', '') or '').strip()
        if key:
            # セルを空にすると NaN が返るので、その品は現在値のまま（変更なし扱い）にする
            fallback = _to_float(r.get('出庫可能数', 0))
            val = _to_float(row_e.get('出庫可能数'), fallback)
            desired[key] = fallback if val != val else val
    n_checked = len(checked)

    # 全量から減らしている品を表の下にキャプションで出す（在庫全量列を足さない代わりの手当て）
    _supply_reduced_caption(rows)

    # ---- ボタン3つ：除外を保存／出庫可能数を保存／指定をすべて取り消す ----
    b1, b2, b3 = st.columns(3)
    with b1:
        do_excl = st.button('除外を保存（%d件）' % n_checked,
                            key='btn_%s' % table_key, disabled=(n_checked == 0))
    with b2:
        do_qty = st.button('出庫可能数を保存', key='btnqty_%s' % table_key)
    with b3:
        do_clear = st.button('出庫可能数の指定をすべて取り消す（%d件）' % n_supply_specified,
                            key='btnrst_%s' % table_key, disabled=(n_supply_specified == 0))
        st.caption('この店で指定している出庫可能数をすべて取り消し、全品を在庫まるごとに戻します。')

    if do_excl:
        _bump_editor(table_key)   # 除外で行が減るので、古い編集差分を残さないよう作り直す
        _save_exclusions_ui(checked, my_store, backend, exclusions)
    elif do_qty:
        _save_supply_qty_ui(rows, my_store, backend, desired)
        _bump_editor(table_key)
        st.rerun()
    elif do_clear:
        _clear_all_supply_qty_ui(my_store, backend)
        _bump_editor(table_key)
        st.rerun()


# ============================================================================
# ②③④の切替ボタン（2026-07-28 本間部長指示）
#   3つの表を縦に並べると画面がとても長くなり、④を見るのに延々スクロールが要る。
#   → ボタンで「選んだ1つだけを描く」＝ページを分けたのと同じ使い勝手にする。
#     （描く表が1つになるので表示も軽くなる。）
# ============================================================================
VIEW_DEAD = 'dead'        # ②自店のデッド品
VIEW_EXPIRY = 'expiry'    # ③自店の期限切迫品
VIEW_RECEIVE = 'receive'  # ④自店が引き取れる薬
VIEW_MESSAGE = 'message'  # ⑤店舗間のやり取り（掲示板）
VIEW_ORDER = [VIEW_DEAD, VIEW_EXPIRY, VIEW_RECEIVE, VIEW_MESSAGE]

# ④の予約で選ぶ「受取時期」（今すぐ／1〜3ヶ月後の4択・最大3ヶ月・本間部長確定）。
#   実際に予約する月は、品ごとの有効期限キャップ（app_logic.pickup_cap）で頭打ちにする。
PICKUP_OFFSETS = [0, 1, 2, 3]
PICKUP_LABELS = {0: '今すぐ', 1: '1ヶ月後', 2: '2ヶ月後', 3: '3ヶ月後'}


def _pickup_offset_selector(key):
    """ 受取時期（今すぐ／1〜3ヶ月後）を選ぶドロップダウンを描き、選ばれたオフセット（0〜3）を返す。 """
    return st.selectbox(
        '受取時期（いつ引き取るか）', PICKUP_OFFSETS,
        format_func=lambda o: PICKUP_LABELS[o], key=key,
        help='いま在庫があって使い切ってから引き取りたいときは、先の月を選べます（最大3ヶ月）。'
             '有効期限が近い品は、期限の月より先は選べません（自動で早めます）。')


def view_switcher(n_dead, n_expiry, n_receive, n_unread=0):
    """
    ②③④⑤を切り替えるボタンを描き、選ばれた画面の記号（VIEW_*）を返す。
      ・件数をボタンに入れて、開く前に中身があるかどうか分かるようにする。
        ⑤は投稿数ではなく『新着（未読）件数』を出す（0件のときは「（新着…）」を付けない）。
      ・★選択肢そのものは VIEW_* の記号にして、見た目の文字は format_func で作る。
        ボタンの文字（件数入り）を選択肢にしてしまうと、除外を保存して件数が変わった瞬間に
        「保存されている選択」が選択肢の中から消えて、③④を見ていても②に戻されてしまう。
        ⑤の新着件数も未読が0になると変わるので、記号を選択肢にするこの作りが必須（設計の肝）。
      ・★segmented_control は選択中のボタンをもう一度押すと「選択なし（None）」を返すので、
        そのときは②に戻す（画面が空になるのを防ぐ）。
      ・segmented_control が無い古いStreamlitでは st.radio（横並び）に自動で切り替える。
    """
    labels = {
        VIEW_DEAD:    '②  デッド品（%d件）' % n_dead,
        VIEW_EXPIRY:  '③  期限切迫品（%d件）' % n_expiry,
        VIEW_RECEIVE: '④  引き取れる薬（%d件）' % n_receive,
        VIEW_MESSAGE: ('⑤  やり取り（新着%d件）' % n_unread) if n_unread else '⑤  やり取り',
    }
    st.caption('見たい表のボタンを押してください（選んだものだけを表示します）。')
    picker = getattr(st, 'segmented_control', None)
    if picker is None:
        chosen = st.radio('表示する内容', VIEW_ORDER, horizontal=True,
                          format_func=lambda k: labels[k],
                          label_visibility='collapsed', key='view_switch')
    else:
        chosen = picker('表示する内容', VIEW_ORDER, default=VIEW_DEAD, required=True,
                        format_func=lambda k: labels[k],
                        label_visibility='collapsed', key='view_switch')
    return chosen or VIEW_DEAD


def _receive_view_with_talk(r, hidden, msg_by_store):
    """ ④受け手側の行 r を画面表示用の辞書に整える（隠し列を除き、末尾に『やり取り』列を足す）。
        『やり取り』＝その品の出し手店とのやり取り件数（'N件 ●' 形式・無ければ空）。会話は⑤で見る。
        ★④本体・別枠・予約中の3つの表で同じ整え方をする（同じ列を2通りに作らないため）。 """
    d = {k: v for k, v in r.items() if k not in hidden}
    d['やり取り'] = (msg_by_store or {}).get((r.get('出し手店', '') or '').strip(), '')
    return d


def receive_section(view_receive, my_store, backend, reservations, ym, msg_by_store=None):
    """
    ④受け手ビュー：他店がデッド・期限切迫で持っていて、自店が引き取れば活かせる品の一覧。
      ・2026-07-28に「予約」を付けた。左端の□で選んで「予約する」を押すと、その品は
        ほかの店の④から消え、出し手の②③には「◯◯が引取予定」と出る。
      ・予約は品目まるごと。数量の相談は従来どおり電話・デスクネッツ。
      ・★保存の直前に予約表を読み直して重複を止める（下の _save_reservations 参照）。
      ・2026-08-10 第2弾：出し手店ごとのやり取り件数を『やり取り』列で出す（会話は⑤で見る）。
    """
    st.subheader('④（%s）が引き取れる薬（他店のデッド・期限切迫）' % my_store)
    st.caption('他店がデッド・期限切迫で持っていて、自店が引き取れば活かせる品です。'
               '引き取りたい品は左端の□にチェックを入れて「予約する」を押してください'
               '（予約するとほかの店の一覧からは消え、出し手の画面に自店名が出ます）。'
               '『滞留』の欄は、その品が何ヶ月つづけて動いていないかです。'
               '長く残っている品ほど上に並べています（色の意味は表の下に出ます）。'
               '『区分』に表示がある薬（向精神薬・毒薬・劇薬）は、受け取る側でも記録が必要です。'
               '数量の相談・連絡は従来どおり電話・デスクネッツでお願いします。')
    if not view_receive:
        st.info('いま他店から引き取れる品はありません。')
        return

    # 内部用の列（_出し手店・_key・滞留区分・有効期限キャップ用など）は画面に出さない
    hidden = ('_出し手店', '_key', '_予約店', '_滞留区分', '_滞留月数', '_有効期限')
    df = pd.DataFrame([_receive_view_with_talk(r, hidden, msg_by_store) for r in view_receive])
    stag_levels = [r.get('_滞留区分', 'new') for r in view_receive]
    event = st.dataframe(
        _style_expiry(df, paint=False, stag_levels=stag_levels),
        hide_index=True,
        width='stretch',
        on_select='rerun',
        selection_mode='multi-row',
        key='table_receive')
    stagnation_legend(view_receive)

    # 選ばれた行を「行の位置」で拾い、位置で元の view_receive から引く（隠し列の値に頼らない）
    picked = _selected_rows(event)
    checked = [view_receive[i] for i in picked if 0 <= i < len(view_receive)]
    n = len(checked)
    offset = _pickup_offset_selector('pickup_offset_receive')
    if st.button('予約する（%d件）' % n, type='primary', key='btn_reserve', disabled=(n == 0)):
        _save_reservations(backend, my_store, ym, checked, reservations, offset)


def _save_reservations(backend, my_store, ym, checked, reservations, offset=0):
    """
    選ばれた品を予約として保存する。

    ★★保存の直前に予約表を読み直す。
      Googleシートは同時書き込みを止められないので、画面を開いてからボタンを押すまでの間に
      別の店が同じ品を予約している可能性がある。読み直さずに書くと、あとから押した店が
      先の予約を上書きしてしまい、どちらも「自分が予約できた」と思い込む事故になる。
      すでに他店が押さえていた品は保存せず、赤でその品名と店名を知らせる。
      残り（重複していない品）はそのまま保存する＝全部やり直しにはしない。

    ★2026-08-01：受取時期（offset＝今から何ヶ月後に引き取るか・0〜3）を受け取る。
      品ごとに有効期限キャップ（app_logic.pickup_cap）で頭打ちにしてから受取予定月を決める
      ＝期限の月より先の受け取りは選ばせない（自動で早める）。早めた品は下で知らせる。
    """
    try:
        latest_rows = backend.load_reservations()
    except Exception:
        latest_rows = list(reservations)   # 読み直せなければ画面の内容で進む

    now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')
    picked = []
    clamped = []   # 有効期限が近く、受取時期を早めた品（薬品名, 実際の受取ラベル）
    for d in checked:
        cap = app_logic.pickup_cap(d.get('_有効期限', d.get('有効期限', '')), ym)
        eff_off = max(0, min(int(offset), cap))
        if eff_off < int(offset):
            clamped.append((d.get('薬品名', ''), yuzu_core.pickup_label(eff_off)))
        picked.append(dict(d, _now=now, _受取予定月=yuzu_core.ym_add(ym, eff_off)))
    plan = app_logic.plan_reservations(latest_rows, my_store, ym, picked)

    if clamped:
        st.info('  \n'.join(
            ['次の品は有効期限が近いため、受取時期を早めて予約しました：']
            + ['%s → 受取：%s' % (name, lab) for name, lab in clamped]))

    if plan['added']:
        backend.save_reservations(plan['keep'])
        clear_lists_cache()   # 保存直後：表示用キャッシュと結果キャッシュを捨て、次の描画ですぐ反映する
        # 新しく予約できた行だけを出し手店へメール通知（保存後・rerun 前に1回だけ＝二重送信しない）。
        #   plan['keep'] のうち、読み直した latest_rows に無かった (出し手店, 予約キー) が新規。
        before_keys = {(r.get('出し手店', ''), r.get('予約キー', '')) for r in latest_rows}
        added_rows = [r for r in plan['keep']
                      if (r.get('出し手店', ''), r.get('予約キー', '')) not in before_keys]
        _mail_flush(mailer.notify_reservation(_mail_secrets(), my_store, added_rows))
    if plan['conflicts']:
        st.error('  \n'.join(
            ['次の品は、ひと足先にほかの店が予約していました（予約できていません）：']
            + ['%s（%s）→ %s が予約済み' % (c['薬品名'], c['出し手店'], c['予約した店'])
               for c in plan['conflicts']]))
    if plan['added']:
        st.success('%d件を予約しました。出し手の画面に「%s が引取予定」と表示されます。'
                   % (plan['added'], my_store))
    if plan.get('broken'):
        # 通常は起こらない。起きたら黙って落とさず知らせる（気づけない不具合を防ぐ）。
        st.error('%d件は内部キーが取れず予約できませんでした。'
                 'お手数ですが管理本部に連絡してください。' % plan['broken'])
    if not plan['added'] and not plan['conflicts'] and not plan.get('broken'):
        st.info('新しく予約する品はありませんでした（すでに予約済みです）。')
    if plan['added'] or plan['conflicts']:
        st.rerun()


def receive_ref_section(view_ref, my_store, backend, reservations, ym, msg_by_store=None):
    """
    ④の別枠『いまは在庫があるが、先になら引き取れる薬』（③の改修・2026-08-01）。
      ・自店もその薬を使っているが、いま在庫を余らせている品（tier③参考）。今すぐ引き取ると
        移した先で新しいデッドを作りかねないので、本来の④からは外している。
      ・ここから「1〜3ヶ月後」の受取予定月つきで予約できる（本命の狙い）。
      ・④本体と同じ方式（st.dataframe + 行選択）・同じ2列構成（出し手の出庫可能数＋自店の在庫数）。
        ★本来の④に出る品目は1件も変えない（増えるのはこの別枠だけ）。
    """
    with st.expander('いまは在庫があるが、先になら引き取れる薬（%d件）' % len(view_ref)):
        st.caption('自店もこの薬を使っていますが、いまは在庫を余らせているため、今すぐ引き取ると'
                   'かえって自店で新しいデッドを作りかねない品です。'
                   'いまの在庫を使い切る先の時期（1〜3ヶ月後）を選んで予約できます。'
                   '数量の相談・連絡は従来どおり電話・デスクネッツでお願いします。')
        if not view_ref:
            st.info('該当する品はありません。')
            return
        # ④本体と同じ隠し列
        hidden = ('_出し手店', '_key', '_予約店', '_滞留区分', '_滞留月数', '_有効期限')
        df = pd.DataFrame([_receive_view_with_talk(r, hidden, msg_by_store) for r in view_ref])
        stag_levels = [r.get('_滞留区分', 'new') for r in view_ref]
        event = st.dataframe(
            _style_expiry(df, paint=False, stag_levels=stag_levels),
            hide_index=True, width='stretch', on_select='rerun',
            selection_mode='multi-row', key='table_receive_ref')
        stagnation_legend(view_ref)

        picked = _selected_rows(event)
        checked = [view_ref[i] for i in picked if 0 <= i < len(view_ref)]
        n = len(checked)
        offset = _pickup_offset_selector('pickup_offset_ref')
        if st.button('予約する（%d件）' % n, type='primary', key='btn_reserve_ref', disabled=(n == 0)):
            _save_reservations(backend, my_store, ym, checked, reservations, offset)


def reserved_section(view_reserved, my_store, backend, reservations, result, latest,
                     msg_by_store=None):
    """ 「予約中の品」を折りたたみで出し、選んで取り消せるようにする（②③の除外と同じ操作感）。
        ★2026-08-01：一番下に「引取依頼書（出し手店ごとのExcel）」を作るボタンを足した（④の改修）。
        ★2026-08-10 第2弾：出し手店ごとのやり取り件数を『やり取り』列で出す（会話は⑤で見る）。 """
    with st.expander('予約中の品（%d件）＝自店が引き取ると押さえている薬' % len(view_reserved)):
        if not view_reserved:
            st.write('いまは1件もありません。')
            return
        st.caption('予約をやめる品を左端の□で選び、下のボタンを押してください。'
                   '取り消すと、その品はまたほかの店の一覧にも出るようになります。')
        hidden = ('_出し手店', '_key', '_受取予定月')
        df = pd.DataFrame([_receive_view_with_talk(r, hidden, msg_by_store)
                           for r in view_reserved])
        event = st.dataframe(
            _style_expiry(df, paint=False), hide_index=True, width='stretch',
            on_select='rerun', selection_mode='multi-row', key='table_reserved')
        picked = _selected_rows(event)
        checked = [view_reserved[i] for i in picked if 0 <= i < len(view_reserved)]
        n = len(checked)
        if st.button('予約を取り消す（%d件）' % n, key='btn_unreserve', disabled=(n == 0)):
            # ★★保存の直前に予約リストを“直接”読み直してから取り消す（ほかの店の予約を消さない・2026-08-12）
            try:
                reservations = backend.load_reservations()
            except Exception:
                pass
            keep = app_logic.cancel_reservations(reservations, my_store, checked, latest)
            backend.save_reservations(keep)
            clear_lists_cache()   # 取消直後：表示用キャッシュと結果キャッシュを捨て、次の描画ですぐ反映する
            st.session_state.pop('pickup_xls', None)   # 予約が変わったので作りかけの帳票は捨てる
            # 取り消した品を出し手店へメール通知（保存後・rerun 前に1回だけ＝二重送信しない）。
            _mail_flush(mailer.notify_cancellation(_mail_secrets(), my_store, checked))
            st.success('%d件の予約を取り消しました。' % (len(reservations) - len(keep)))
            st.rerun()

        # ---- 引取依頼書（出し手店ごとのExcel）＝FAX・デスクネッツ用 ----
        #   ★st.download_button はページを描くたびに中身を作るので、
        #     「作成する」ボタンを押したときだけ生成し、そのあとダウンロードボタンを出す2段にする。
        #   ★画面最下部の「この結果をExcel（4シート）でダウンロード」（全店分析用）とは
        #     場所も名前も混ぜない（こちらは自店の発注用の紙）。
        st.divider()
        st.caption('予約した品を、もらう先（出し手店）ごとにまとめた「引取依頼書」をExcelで作れます'
                   '（A4横・出し手店ごとにシート・FAX/デスクネッツ用）。'
                   '数量は在庫まるごとを初期値にしているので、必要に応じて紙の上で書き換えてください。')
        if st.button('引取依頼書を作成する', key='btn_make_pickup'):
            try:
                st.session_state['pickup_xls'] = app_logic.pickup_request_bytes(
                    result, my_store, reservations, latest)
            except Exception as e:
                st.session_state.pop('pickup_xls', None)
                st.warning('引取依頼書の作成に失敗しました：%s' % e)
        if st.session_state.get('pickup_xls'):
            if latest:
                pk_name = '引取依頼書_%s_%s-%s.xlsx' % (my_store, latest[:4], latest[4:6])
            else:
                pk_name = '引取依頼書_%s.xlsx' % my_store
            st.download_button(
                '⬇ 出し手店ごとの引取依頼書（Excel）を作成',
                data=st.session_state['pickup_xls'], file_name=pk_name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key='dl_pickup')


# ============================================================================
# ⑤ 店舗間のやり取り（掲示板）＝第2弾・本間部長確定 2026-08-10
#   ・スレッドは「相手店ごとに1本」。同じ相手との会話は必ず1つにまとまる（予約1件ごとではない）。
#   ・上段で相手店を選び、下段でその相手との会話を時系列で見て投稿する。
#   ・スレッドを開いたら既読日時を「いま」に更新する（未読が消える）。
#   ・API節約：やり取りは60秒のセッションキャッシュ越しに読む。投稿・既読更新の直後だけ捨てる。
# ============================================================================
def message_section(my_store, backend, threads, msg_reads):
    """
    ⑤ 店舗間のやり取り。threads は app_logic.build_threads の戻り（自店が関わるスレッド一覧）。
      ・上段：相手店の一覧（新着＝未読件数／予約中の品数つき）を「1行＝1店＋[開く]ボタン」で出す。
        ★一覧の並びは店番順（STORE_NAMES の並び）で固定する（投稿・未読で順番が変わらない）。
          STORE_NAMES に無い店名は末尾へ五十音順でまとめる（黙って消さない）。
          未読のあるスレッドは相手店名の左に ● を付けて目立たせる（並べ替えでは動かさない）。
      ・下段：選んだ相手との会話を時系列で表示し、任意で『どの薬の話』を添えて投稿できる。
        ★★相手店の[開く]を押すまでは下段そのものを出さない（誤送信を防ぐため）。

    ★★なぜ st.dataframe の行選択をやめて[開く]ボタン方式にしたか（将来また表に戻そうとする人へ）★★
      st.dataframe(on_select='rerun') は「表の中身（データ）が変わると frontend（glide-data-grid）が
      行選択を捨てる」という Streamlit の未解決の不具合を抱えている（issue #10701）。
      ⑤やり取りでは、これが次の形で牙を剥いていた：
        ・投稿すると一覧の『やり取り』欄が 空→「1件」に変わる（送信者側）
        ・未読スレッドを開くと、下で自動的に既読化して ● と『新着』欄が消える（受信者側）
      どちらも“アプリが自分で起こした rerun の直後”に表の中身が変わるため、frontend が行選択を
      捨て → 会話の窓が勝手に閉じて「投稿が消えた」ように見えていた。以前は msg_keep_sel という
      「1回だけ選択を復元する」応急処置で打ち消していたが、Streamlit の不具合を相殺する繊細な作りで
      壊れやすく、投稿のたびに□を選び直さないと連続で会話できなかった。
      → そこで **st.dataframe の行選択そのものを使うのをやめ**、いま開いている相手店を
        session_state['msg_open'] に明示的に持つ方式へ変えた。ボタンの押下・rerun は
        session_state を消さないので、投稿や既読化で表の中身が変わっても会話は開いたままになる。
        ＝表（st.dataframe）に戻すと #10701 の不具合が再発するので、戻さないこと。
    """
    st.subheader('⑤（%s）店舗間のやり取り' % my_store)

    # 「最新に更新」：60秒キャッシュを捨てて読み直す（他店の新着をすぐ見たいとき用）
    c1, c2 = st.columns([1, 3])
    with c1:
        if st.button('最新に更新', key='btn_msg_refresh'):
            clear_messages_cache()
            # ★タブの目次キャッシュも捨てる。まだ `_やり取り` タブが無い状態では
            #   _find_ws の空振り抑制（30秒）が効いていて、他店が今まさに作ったタブを
            #   見落とすことがあるため。手で押したときだけは必ず取り直す。
            gsheet_store.reset_ws_cache()
            st.rerun()
    with c2:
        st.caption('相手ごとに会話は1本にまとまります（予約が入ると、その相手とやり取りできます）。'
                   '他店の新着は最大60秒ほど遅れて出ます。すぐ見たいときは「最新に更新」を押してください。')

    if not threads:
        st.info('やり取りできる相手店がまだありません。'
                '他店のデッド品を予約するか、他店から自店の品を予約されると、その相手との会話が始まります。')
        return

    # --- 並びは店番順（STORE_NAMES の並び）で固定。投稿・未読で順番が変わらない。
    #     STORE_NAMES に無い店名は末尾へまとめ、その中は五十音順にする（黙って消さない）。---
    threads = sorted(threads, key=_thread_sort_key)

    # いま開いている相手店（[開く]で入れ、[◀ 一覧に戻る]や②③④への切替で消す）。
    #   ★選択を st.dataframe に持たせず session_state に持つのが今回の要点（#10701 回避）。
    open_name = st.session_state.get('msg_open')

    # --- msg_open が空＝一覧だけを出す。会話も投稿欄もいっさい出さない（誤送信防止）---
    if not open_name:
        st.divider()
        st.caption('やり取りする相手店の右の［開く］を押してください。')
        # 見出し行（相手店／新着／予約中／やり取り）。会話を開くボタンは各行の右端に置く。
        h1, h2, h3, h4, h5 = st.columns([4, 2, 2, 2, 2])
        h1.markdown('**相手店**')
        h2.markdown('**新着**')
        h3.markdown('**予約中**')
        h4.markdown('**やり取り**')
        h5.markdown('')
        for t in threads:
            ur = app_logic.unread_count(my_store, t, msg_reads)
            r1, r2, r3, r4, r5 = st.columns([4, 2, 2, 2, 2])
            # 未読があれば相手店名の左に ●（表のときと同じ見せ方）
            r1.write(('● ' if ur else '') + t['相手店名'])
            r2.write(('%d件' % ur) if ur else '')
            r3.write(('%d品' % len(t['予約中の品'])) if t['予約中の品'] else '')
            r4.write(('%d件' % t['件数']) if t['件数'] else '')
            # ★ボタンの key は「行の位置」ではなく「相手店名」から作る。
            #   位置で作ると、並び順が変わったとき別の店を開いてしまうため。
            if r5.button('開く', key='btn_open_%s' % t['相手店名']):
                st.session_state['msg_open'] = t['相手店名']
                st.rerun()
        return

    # --- msg_open に相手店名が入っている＝一覧は出さず会話だけを出す ---
    #   一覧と会話を同時に出さないのは、誰と話しているかを1画面で取り違えないため。
    sel = next((t for t in threads if t['相手店名'] == open_name), None)
    if sel is None:
        # 開いていた相手が消えた（予約が全部取り消された等）。例外で落とさず一覧へ戻す。
        st.session_state.pop('msg_open', None)
        st.info('開いていた相手店（%s）とのやり取りが見つかりませんでした。一覧に戻ります。' % open_name)
        return

    st.divider()
    # 見出しと「◀ 一覧に戻る」。戻るを押したら msg_open を捨てて一覧へ（会話を閉じる）。
    hcol, bcol = st.columns([4, 1])
    hcol.markdown('#### %s とのやり取り' % sel['相手店名'])
    if bcol.button('◀ 一覧に戻る', key='btn_msg_back'):
        st.session_state.pop('msg_open', None)
        st.rerun()
    if sel['予約中の品']:
        st.caption('予約中：' + '　'.join(sel['予約中の品']))

    if not sel['messages']:
        st.write('まだ投稿はありません。')
    else:
        for m in sel['messages']:
            st.markdown('**%s　%s**'
                        % (_fmt_msg_time(m.get('投稿日時', '')), m.get('投稿店', '')))
            st.write(m.get('本文', ''))
            drug = str(m.get('薬品名', '') or '').strip()
            if drug:
                st.caption('（%s の話）' % drug)

    # --- 投稿フォーム（送信後は入力欄を空に戻すため、版番号でキーを作り直す）---
    pair_key = '%s__%s' % (sel['店A'], sel['店B'])
    ver = st.session_state.get('msgver_%s' % pair_key, 0)
    # ★『どの薬の話』のタグには“素の薬品名”（数量を付けない）を使う。
    #   予約中の品（数量つき）を使うと『20錠』などの数量が投稿ログに残り、
    #   翌月に数量が変わっても古い数量が残って紛らわしいため（②の改修・2026-08-10）。
    drug_opts = [''] + list(sel.get('予約中の品名', sel['予約中の品']))
    drug = st.selectbox(
        'どの薬の話（任意）', drug_opts,
        format_func=lambda x: x if x else '（薬を特定しない）',
        key='msgdrug_%s_%d' % (pair_key, ver))
    body = st.text_area('本文', key='msgbody_%s_%d' % (pair_key, ver),
                        placeholder='例）4件まとめてお願いできますか。火曜に取りに伺います。')
    if st.button('送信', type='primary', key='msgsend_%s' % pair_key):
        text = (body or '').strip()
        if not text:
            st.warning('本文が空です。ひとこと書いてから送信してください。')
        else:
            now = datetime.datetime.now().strftime(MSG_TS_FMT)
            row = {'投稿日時': now, '店A': sel['店A'], '店B': sel['店B'],
                   '薬品名': drug, '投稿店': my_store, '本文': text}
            try:
                backend.append_message(row)
            except Exception as e:
                show_gsheet_error(e, '投稿の保存に失敗しました', 'error')
            else:
                clear_messages_cache()                       # 投稿直後だけキャッシュを捨てる
                st.session_state['msgver_%s' % pair_key] = ver + 1   # 入力欄を空に戻す
                # ★msg_open は触らない＝rerun 後もこの相手との会話が開いたまま（続けて投稿できる）。
                #   投稿で『やり取り』欄が 空→1件 に変わっても、選択を session_state で持っているので
                #   会話は閉じない（旧 msg_keep_sel の応急処置は不要になった）。
                # メール通知（相手店へ）。保存後・rerun 前に1回だけ＝二重送信しない。
                #   失敗しても投稿は保存済み＝止めない（結果の案内は _flash で rerun 後に出す）。
                _mail_flush(mailer.notify_new_message(
                    _mail_secrets(), my_store, sel['相手店名'], drug, text))
                st.rerun()

    # --- スレッドを開いた＝既読にする。未読があるときだけ書く（ムダな書き込み・API消費を避ける）---
    if app_logic.unread_count(my_store, sel, msg_reads) > 0:
        now = datetime.datetime.now().strftime(MSG_TS_FMT)
        new_reads = app_logic.mark_thread_read(
            msg_reads, my_store, sel['店A'], sel['店B'], now)
        try:
            backend.save_msg_reads(new_reads)
        except Exception as e:
            show_gsheet_error(e, '既読の更新に失敗しました', 'warning')
        else:
            clear_messages_cache()   # 既読を書いた直後だけキャッシュを捨てて読み直す
            # ★msg_open は触らない＝既読化の rerun で ● と『新着』欄が消えても会話は開いたまま。
            #   選択を session_state で持っているので、表のときのように会話が閉じない。
            st.rerun()


def excluded_section(my_store, backend, exclusions):
    """ 「除外中の品目」を折りたたみで出し、選んで元に戻せるようにする（②③と同じ行選択方式）。
        操作方法が画面内で2種類あると迷うため、②③のチェックと同じ「左端の□で選ぶ」に統一している。 """
    mine = [r for r in exclusions if r['店名'] == my_store]
    with st.expander('除外中の品目（%d件）＝デッドストックから外している薬' % len(mine)):
        if not mine:
            st.write('いまは1件もありません。')
            return
        st.caption('デッドストックに戻したい品を左端の□で選び、下のボタンを押してください。')
        df = pd.DataFrame([{'薬品名': r['薬品名'], '除外日時': r['除外日時']} for r in mine])
        event = st.dataframe(
            df, hide_index=True, width='stretch',
            on_select='rerun', selection_mode='multi-row', key='table_undo')
        # 選んだ行を「行の位置」で拾い、位置で元の mine から引く
        picked = _selected_rows(event)
        checked = [mine[i] for i in picked if 0 <= i < len(mine)]
        n = len(checked)
        if st.button('選んだ品目を戻す（%d件）' % n, key='btn_undo', disabled=(n == 0)):
            back = {r['除外キー'] for r in checked}
            # ★★保存の直前に除外リストを“直接”読み直してから引く（ほかの店の除外を消さない・2026-08-12）
            try:
                exclusions = backend.load_exclusions()
            except Exception:
                pass
            keep = [r for r in exclusions
                    if not (r['店名'] == my_store and r['除外キー'] in back)]
            backend.save_exclusions(keep)
            clear_lists_cache()   # 戻した直後：表示用キャッシュと結果キャッシュを捨て、次の描画ですぐ反映する
            st.success('%d件をデッドストックに戻しました。' % n)
            st.rerun()


def show_upload_status(status, latest):
    """ 画面上部に『現在 N/15店 アップ済み』と未アップ店を出す。 """
    n = status['n']
    ym_disp = ('%s年%s月' % (latest[:4], latest[4:6])) if latest else '（まだデータがありません）'
    c1, c2 = st.columns([1, 3])
    with c1:
        st.metric('アップ済み', '%d / %d 店' % (n, STORE_COUNT))
    with c2:
        st.write('**対象月：** %s' % ym_disp)
        if status['missing']:
            st.write('**未アップの店：** ' + '、'.join(status['missing']))
        if status['ng']:
            st.warning('様式が他店と違う（別様式）ため計算に入れていない店：' + '、'.join(status['ng']))
    if n < STORE_COUNT:
        st.info('全店（%d店）そろうと、他店の使用実績まで見えてマッチングの精度が上がります。' % STORE_COUNT)


# ============================================================================
# 誤アップロード防止（ファイル名の店名 × 選んだ店名の突き合わせ）
#   全店が同じURL・同じ共有パスワードを使い、認証が無いため、A店の担当者が誤ってB店を
#   選んで上げると、保管庫が ws.clear() → 書き込みの順で B店のデータを丸ごと上書きしてしまう
#   （誰も気づかない）。その最小限の防波堤として、ここでファイル名と選択店名を突き合わせる。
# ============================================================================
def _normalize_store(name):
    """ 店名を突き合わせ用にそろえる。前後の空白を除去し、全角スペースを取り除くだけ。
        ※表記ゆれの吸収まではしない（完全一致で弾くのが目的なので、余計な変換はしない）。 """
    if name is None:
        return ''
    # 　 ＝ 全角スペース。前後の空白を落としてから全角スペースを除去し、もう一度前後を落とす。
    return str(name).strip().replace('　', '').strip()


def evaluate_upload_guard(filename, selected_store, store_names):
    """ ファイル名の店名部分と、いま選んでいる店名を突き合わせる純関数（画面部品に依存しない＝テスト可能）。

        戻り値は辞書：
          status  … 'match'    … ファイル名の店名が選択店と完全一致（そのまま進めてよい）
                    'mismatch' … ファイル名の店名は別の登録店（＝別店のデータを上書きしかねない）
                    'unknown'  … ファイル名から登録店を読み取れない（確認チェックを求める）
          file_store     … 表示用。'match'/'mismatch' は一致した登録店名、
                            'unknown' はファイル名から拾った文字列
          selected_store … いま選んでいる店名（そのまま）

        ★突き合わせは必ず完全一致（==）で行う。部分一致（in／startswith）は絶対に使わない。
          理由：「和光」（ソユーズ和光）は「さと和光」（内観堂さと和光・別会社）に文字列として
          含まれるため、部分一致にすると別会社のファイルが通ってしまい、このガードの意味が消える。 """
    # ファイル名から店名部分を取り出す（末尾が _YYYYMM でなければ全体が店名扱いで返る）
    raw_store, _ym = yuzu_core.parse_filename(filename)
    file_norm = _normalize_store(raw_store)

    # 登録店（STORE_NAMES）を正規化した名前をキーにした辞書にしておく。
    #   下の「file_norm in known」は辞書キーの“完全一致”判定であり、部分一致（含む）ではない。
    known = {}
    for s in store_names:
        known[_normalize_store(s)] = s

    if file_norm in known:            # ← 辞書キーの完全一致（部分一致ではない）
        matched = known[file_norm]    # 実際の登録表記（例：'さと和光'）
        if matched == selected_store:
            return {'status': 'match', 'file_store': matched, 'selected_store': selected_store}
        return {'status': 'mismatch', 'file_store': matched, 'selected_store': selected_store}
    # どの登録店とも完全一致しない＝店名を読み取れない
    return {'status': 'unknown', 'file_store': raw_store, 'selected_store': selected_store}


# ============================================================================
# 選んだ店をURLに覚えさせる（2026-07-28）
#   店の選択はブラウザのセッションにしか残らないため、タブを閉じたり、アプリが眠って
#   起き直したりすると毎回選び直しになる。URLに ?store=店名 を残しておけば、店が
#   そのURLをブックマークするだけで次回から自分の店で開ける。
#   ※データそのものは保管庫（Gシート）にあるので、URLに何が入っていても中身は変わらない。
#     アップロード時の取り違えは従来どおり「ファイル名 × 選択店名」の突合が止める。
# ============================================================================
def _raw_store_in_url():
    """ URLの ?store= の生の値を1つだけ取り出す。
        ★同じキーが複数ある場合にリストで返る実装があるため、必ず1つ目を取り出して文字列にそろえる。
          （そろえないと下の「同じ値なら書かない」判定が毎回不一致になり、URLを書き続けてしまう。） """
    try:
        raw = st.query_params.get('store')
    except Exception:
        return None      # query_params が無い古いStreamlit＝URL記憶なしで従来どおり動く
    if isinstance(raw, (list, tuple)):
        raw = raw[0] if raw else None
    return raw


def _store_from_query():
    """ URLの ?store=店名 を読む。★登録店との完全一致のときだけ採用する（部分一致は使わない）。 """
    raw = _raw_store_in_url()
    if not raw:
        return None
    norm = _normalize_store(raw)
    for s in STORE_NAMES:
        if _normalize_store(s) == norm:
            return s
    return None


def _remember_store_in_url(store):
    """ 選んだ店をURLに残す（未選択なら消す）。失敗しても画面は止めない。
        ★「いまURLに入っている値と違うときだけ書く」＝同じ値を書き続けないようにする。 """
    try:
        current = _raw_store_in_url()
        if store:
            if current != store:
                st.query_params['store'] = store
        elif current is not None:
            del st.query_params['store']
    except Exception:
        pass


# ============================================================================
# ①店の選択 ＋ アップロード欄
#   ★2026-07-28：アップロードは「毎回やること」ではない。当月ぶんが保管庫に入っていれば、
#     店を選ぶだけで②③④の表が見られる（データはGシートに残っていて、セッションとは無関係）。
#     以前は店の選択とアップロードが1つの塊だったため「毎回アップしないと見られない」ように
#     見えていた。そこで、アップロード欄は当月ぶんが済んでいれば折りたたんでおき、
#     「アップ済みです。このまま下の表を見られます」とはっきり出すようにした。
# ============================================================================
def upload_section(backend, index, latest):
    st.subheader('① 店舗を選択')

    # 店舗名の選択（先頭は「選択してください」。会社名（ソユーズ/内観堂）の付記は無し）
    options = [SENTINEL_STORE] + STORE_NAMES
    # 思い出す順番：1) このセッションで選んだ店 → 2) URLの ?store=
    prev = st.session_state.get('my_store')
    if prev not in STORE_NAMES:
        prev = _store_from_query()
    default_index = options.index(prev) if prev in STORE_NAMES else 0

    # ラベルはこの位置に後から入れる（選択状態に応じて色を変えるため）
    label_ph = st.empty()
    choice = st.selectbox(
        '店舗名（必須）',
        options,
        index=default_index,
        format_func=lambda n: n,
        label_visibility='collapsed')

    if choice in STORE_NAMES:
        # 選択済み → ラベルは通常色
        label_ph.markdown('店舗名（必須）')
        my_store = choice
        st.session_state['my_store'] = my_store
    else:
        # 未選択 → ラベルを赤の太字で強調
        label_ph.markdown(':red[**店舗名（必須）**]')
        my_store = None
        st.session_state['my_store'] = None
        st.caption('※ まず店舗名を選んでください。')
    _remember_store_in_url(my_store)

    # ------------------------------------------------------------------
    # 選んだ店の「当月ぶんアップ済みか」を出し、アップロード欄の開き方を決める
    #   済んでいる … 折りたたんでおく（ふだんは触らない）
    #   まだ      … 開いておく（今すぐアップしてほしい）
    # ------------------------------------------------------------------
    ym_disp = ('%s年%s月' % (latest[:4], latest[4:6])) if latest else '（まだデータがありません）'
    entry = index.get(my_store) if my_store else None
    done = bool(entry) and entry.get('ym') == latest \
        and not str(entry.get('format', '')).startswith('NG')

    if my_store and done:
        st.success('**%s**の%s分はアップ済みです（%s／%s行／%s）。  \n'
                   '**このまま下の表を見られます。毎回アップロードし直す必要はありません。**'
                   % (my_store, ym_disp, entry.get('uploaded_at', '不明'),
                      entry.get('rows', '?'), entry.get('filename', '')))
        exp_label = '新しいデータに差し替える（アップロードし直す）'
        exp_open = False
    elif my_store:
        st.warning('**%s**の在庫ファイルは、まだアップされていません%s。'
                   '下の欄からアップロードしてください。'
                   % (my_store, ('' if latest is None else '（今の対象月は%s）' % ym_disp)))
        exp_label = '在庫ファイルをアップロードする'
        exp_open = True
    else:
        exp_label = '在庫ファイルをアップロードする'
        exp_open = True

    with st.expander(exp_label, expanded=exp_open):
        _upload_form(backend, my_store)


def _upload_form(backend, my_store):
    """ 在庫ファイルを選んでアップロードする欄。upload_section から折りたたみの中に置いて呼ぶ。 """
    if my_store:
        st.caption('→ **%s** で提出します。' % my_store)

    up = st.file_uploader(
        '薬VANの在庫ファイル（.xls / .csv / .xlsx）',
        type=['xls', 'csv', 'xlsx'], accept_multiple_files=False)

    # 対象年月：ファイル名に _YYYYMM があればそれを初期値に、無ければ当月
    default_ym = datetime.date.today().strftime('%Y%m')
    if up is not None:
        _, ym_from_name = yuzu_core.parse_filename(up.name)
        if ym_from_name:
            default_ym = ym_from_name
    ym = st.text_input('対象年月（YYYYMM の6桁）', value=default_ym,
                       help='薬VANを出力した月。ファイル名が「店名_202607」ならその6桁が入ります。')

    # ------------------------------------------------------------------
    # 誤アップロード防止：ファイル名の店名と、選んでいる店名を突き合わせる
    #   ・ファイルと店の両方がそろって初めて判定できる（どちらか未選択なら素通り）
    #   ・完全一致（==）のみ。「和光」と「さと和光」を取り違えないため部分一致は使わない
    # ------------------------------------------------------------------
    guard_ok = True  # 突き合わせに問題が無ければアップロードを許可する
    if up is not None and my_store is not None:
        guard = evaluate_upload_guard(up.name, my_store, STORE_NAMES)
        if guard['status'] == 'match':
            # ファイル名の店名が、選んでいる店とぴったり一致した
            st.caption('ファイル名の店名（%s）と一致しました。' % guard['file_store'])
        elif guard['status'] == 'mismatch':
            # ファイル名は別の登録店のもの＝このまま上げると別店のデータを上書きしてしまう
            guard_ok = False
            st.error('  \n'.join([
                'このファイルは『%s』のファイルに見えます（ファイル名：%s）。' % (guard['file_store'], up.name),
                'いま選んでいる店は『%s』です。' % my_store,
                'このまま進めると%sのデータが上書きされてしまいます。' % guard['file_store'],
                '店舗名を選び直すか、ファイルを確認してください。',
            ]))
        else:
            # 'unknown'：ファイル名から店名を読み取れない＝確認チェックを入れてもらってから許可する
            st.warning('  \n'.join([
                'ファイル名から店名を確認できませんでした（ファイル名：%s）。' % up.name,
                '『店名_%s』の形にしておくと、店の取り違えを自動で止められます。'
                % (ym if (len(ym) == 6 and ym.isdigit()) else '202607'),
            ]))
            # チェックが入るまでアップロードボタンを無効にする。
            #   キーにファイル名を含めて、別のファイルに差し替えたら確認をやり直させる。
            guard_ok = st.checkbox(
                'このファイルは確かに『%s』のものです' % my_store,
                key='confirm_unknown_store_%s' % up.name)

    can_upload = (up is not None) and (my_store is not None) and guard_ok
    if st.button('この内容でアップロードする', type='primary', disabled=not can_upload):
        if my_store is None:
            st.warning('先に店舗名を選んでください。')
            return
        if up is None:
            st.warning('先にファイルを選んでください。')
            return
        if not (len(ym) == 6 and ym.isdigit()):
            st.error('対象年月は「YYYYMM」の6桁の数字で入れてください（例：202607）。')
            return
        try:
            data = up.getvalue()
            res = app_logic.process_upload_bytes(up.name, data)
        except Exception as e:
            st.error('ファイルを読めませんでした：%s' % e)
            return

        if not res['format_ok']:
            st.error('このファイルには、突合に必要な列が足りません（別様式の可能性）。'
                     '不足列：' + '、'.join(res['missing']))
            st.stop()

        backend.save_store_upload(my_store, ym, res['slim'], up.name, res['format_ok'])
        clear_stores_cache()   # 次の描画で必ず読み直す（アップした中身をすぐ反映する）
        st.success('%s の %s年%s月分を受け付けました（%d行・前処理後）。'
                   % (my_store, ym[:4], ym[4:6], res['n_kept']))
        st.rerun()


# ============================================================================
# 結果表示（②自店のデッド品／③自店の期限切迫品／④自店が引き取れる薬＋Excel）
#   ※2026-07-27：画面から「参考ビューB（自店の不足を持つ店）」と「④全店一覧」を外した。
#     どちらも Excel（4シート）・Gシートには従来どおり出力している（画面が長すぎる対策）。
#   ※2026-07-28：②③④を切替ボタン（view_switcher）で1つずつ表示するようにした。
#     3つを縦に並べると④まで延々スクロールが要るため。Excelボタンは切替の外＝常に一番下にある。
# ============================================================================
def results_section(backend, stores, latest, index):
    # ※保管庫の読み込み（load_current_month_stores）は main() で1回だけ行い、
    #   アップロード欄と結果表示で使い回す（Gシートへの往復を増やさないため）。
    #   ★ためておいた通知（メール結果・出庫可能数の保存結果など）を画面の先頭で1回だけ出す。
    #     投稿・予約・取消は保存後に st.rerun するため、その回のメッセージは消える。ここで拾って見せる。
    _render_flash()
    status = app_logic.uploaded_status(index, latest, STORE_NAMES)

    st.subheader('現在の状況')
    show_upload_status(status, latest)

    if not stores:
        st.stop()

    # 旧形式で保管された店の検出
    #   在庫金額は「薬価金額」列から出すが、この列を保管庫に残すようにしたのは2026-07-27から。
    #   それ以前にアップされた店のデータには入っていないため、在庫金額が0円になってしまう。
    #   黙って0円を出すと誤解のもとなので、その店だけ再アップロードを促す。
    old_format = [s['name'] for s in stores
                  if not any(str(r.get('薬価金額', '') or '').strip()
                             or str(r.get('薬価', '') or '').strip() for r in s['rows'])]
    if old_format:
        st.warning('次の店のデータは旧形式で保管されているため、在庫金額が0円で表示されます。'
                   'お手数ですが、もう一度アップロードしてください：' + '、'.join(old_format))

    # ---- 表示用の3リスト（除外・予約・出庫可能数）を60秒キャッシュ越しにまとめて読む（案2）----
    #   ★ふだんの再描画では読み直さず、往復を「_index の1回」だけに減らす。
    #   ★保存の直前の読み直し（二重予約の防止）は各保存関数が backend.load_*() を直接呼ぶ。ここは表示用。
    #   supply_rows_raw … apply_supply_cap で頭打ちする前の“生の指定リスト”。案1の署名に使う。
    exclusions, reservations, supply_rows_raw = load_lists_cached(backend)
    reserved = app_logic.reservation_map(reservations, latest)

    # 提供数量（出し手が「この品はN錠だけ出す」と決めた数）を、当月在庫まで頭打ちにしてから計算へ。
    #   当月の (店名, 品目キー) → 在庫数（全量・ロット合算）を作る。この下ごしらえは両方の道で使う
    #   （n_supply_mine・表の下のキャプションは頭打ち後の supply_rows を見るため）ので毎回作る（軽い）。
    stock_by_key = {}
    for s in stores:
        for row in s['rows']:
            k = (s['name'], yuzu_core.exclusion_key(row))
            stock_by_key[k] = stock_by_key.get(k, 0.0) + yuzu_core.stock_qty(row)
    supply_rows = app_logic.apply_supply_cap(supply_rows_raw, stock_by_key)

    base_ym_disp = ('%s年%s月' % (latest[:4], latest[4:6])) if latest else '不明'
    csv_base_disp = datetime.date.today().strftime('%Y/%m/%d')

    # ============================================================================
    # 案1（2026-08-12）：材料が前回と同じなら「照合もExcelもまるごと作り直さない」
    #   入力の署名（各店のアップ記録＝index／対象年月／除外／予約／出庫可能数の生リスト／本日）を作り、
    #   前回と同じなら compute_matching・apply_stagnation・build_results_payload・excel_bytes を
    #   すべてスキップして、前回の結果（apply_stagnation 済み）と前回のExcelバイト列を使い回す。
    #   ・index と latest が署名に入っているので、アップロード・月替わりで自動的に作り直される。
    #   ・除外/予約/出庫可能数は保存直後に clear_lists_cache で結果キャッシュごと捨てるので、すぐ作り直す。
    #   ・result は apply_stagnation でその場書き換え済み。ビュー作成は result を読むだけ（壊さない）ので、
    #     同じ result を使い回しても値はずれない。
    # ============================================================================
    sig_input = app_logic.results_signature(
        index, latest, exclusions, reservations, supply_rows_raw, csv_base_disp)
    rc = st.session_state.get('results_cache')
    if rc and rc.get('sig') == sig_input:
        # 材料が前回と同じ → 照合もExcelも作り直さない（前回の成果をそのまま使う）
        result = rc['result']      # apply_stagnation 済み
        xls = rc['xls']            # 前回のExcelバイト列（同じ材料なら同じExcelになる）
    else:
        # 材料が変わった（または初回）→ ふつうに計算する
        #   ※出せる数（supply）が空なら、compute_matching は改修前とまったく同じ計算になる（既定 None 相当）。
        supply = app_logic.supply_qty_map(supply_rows)
        result = yuzu_core.compute_matching(stores, excluded=_exclusion_set(exclusions),
                                            reserved=reserved, supply_qty=supply)
        # 滞留（同じ品が何ヶ月つづけて載っているか）を書き込む（result をその場で書き換える）。
        #   前月の記録がまだ無い月は全部「今月から」になる（画面は従来どおりの見た目）。
        prev_map = load_prev_cached(backend, latest)
        app_logic.apply_stagnation(result, prev_map)

        # 結果を保管庫にも書き戻す（次の月替わりで前月退避できるように）。
        #   ★中身が前回と同じなら書かない（results_sig）。結果4タブの書き戻しは clear＋update で
        #     8回のAPI呼び出しになるため、変わっていないなら書き直さない。
        try:
            payload = gsheet_store.build_results_payload(result, base_ym_disp, csv_base_disp)
            sig = _signature(payload)
            if st.session_state.get('results_sig') != sig:
                if hasattr(backend, 'write_results'):
                    backend.write_results(payload)
                elif hasattr(backend, 'save_results'):
                    backend.save_results(payload)
                st.session_state['results_sig'] = sig   # 書けたときだけ覚える
        except Exception as e:
            show_gsheet_error(e, '結果の保管庫への書き戻しに失敗しました（画面表示は続行します）', 'warning')

        # Excel（ダウンロードボタン用）をここで1回だけ作る。同じ材料なら次回は作り直さない。
        try:
            xls = app_logic.excel_bytes(result, base_ym_disp, csv_base_disp)
        except Exception as e:
            xls = None
            st.warning('Excel生成に失敗しました：%s' % e)

        # 今回の材料と成果を覚えておく（次の再描画で材料が同じなら丸ごと使い回す）
        st.session_state['results_cache'] = {'sig': sig_input, 'result': result, 'xls': xls}

    my_store = st.session_state.get('my_store')

    st.divider()
    if not my_store:
        st.subheader('②（自店）のデッド品')
        st.info('上で自店の店舗名を選ぶと、自店のデッド品（②）・期限切迫品（③）と、'
                'それぞれを欲しがっている店が表示されます。')
    else:
        my_uploaded = (my_store in status['uploaded'])
        if not my_uploaded:
            st.subheader('②（%s）のデッド品' % my_store)
            st.warning('まず自店（%s）の在庫ファイルをアップロードしてください。'
                       '自店のデッド・期限切迫が入って初めて、自店視点の提案が出ます。' % my_store)
        else:
            view_a = app_logic.build_view_a(result, my_store)       # 種別＝デッド
            view_expiry = app_logic.build_view_expiry(result, my_store)  # 種別＝期限切迫
            view_receive = app_logic.build_view_receive(result, my_store)  # ④受け手ビュー
            # ④の別枠『いまは在庫があるが、先になら引き取れる薬』（③の改修・use_ref）
            view_receive_ref = app_logic.build_view_receive_ref(result, my_store)
            # 自店が押さえている品（予約中）。④の下に折りたたみで出し、ここから取り消せる。
            view_reserved = app_logic.build_view_reserved(result, my_store, reservations, latest)

            # この店で現在『出庫可能数』の指定が入っている件数（②③の一括取消ボタン用）。
            #   supply_rows は apply_supply_cap 後でも行数は変わらない＝取り消す件数と一致する。
            n_supply_mine = sum(1 for r in supply_rows
                                if (r.get('店名', '') or '').strip() == my_store)

            # ---- ⑤やり取り（掲示板）の下ごしらえ（第2弾）----
            #   ★60秒キャッシュ越しに読む＝ふだんの再描画ではAPIを増やさない（load_messages_cached）。
            #   スレッドは「相手店ごとに1本」。各表の『やり取り』列と⑤の新着バッジに使う早見表を作る。
            messages, msg_reads = load_messages_cached(backend)
            # ⑤『予約中：』に数量を添えるための早見表（②の改修・2026-08-10）。
            #   当月の提案行から (出し手店, 予約キー) で引く。数量＝提案行の『在庫数』
            #   （＝実効値＝画面の出庫可能数）、単位＝『単位』。表示は _fmt_qty で末尾の .00 を落とす。
            #   ★引けない品（出し手が除外した／当月の提案から消えた）は辞書に入れない
            #     ＝薬品名だけになる（0錠などと誤解させない）。予約キー＝提案行の _ex_key。
            qty_by_key = {}
            for pr in result.get('proposal_rows', []):
                q = _fmt_qty(pr.get('在庫数', ''))
                if str(q).strip() == '':
                    continue
                unit = pr.get('単位', '') or ''
                qty_by_key[(pr.get('出し手店', ''), pr.get('_ex_key', ''))] = '%s%s' % (q, unit)
            threads = app_logic.build_threads(my_store, messages, reservations, qty_by_key)
            #   相手店 → 'N件 ●'（自店の未読があれば ●）。②③の予約店・④の出し手店で引く。
            msg_by_store = {}
            total_unread = 0
            for t in threads:
                ur = app_logic.unread_count(my_store, t, msg_reads)
                total_unread += ur
                if t['件数'] > 0:
                    msg_by_store[t['相手店名']] = '%d件%s' % (t['件数'], ' ●' if ur else '')

            # ---- ②③④⑤の切替ボタン。選ばれた1つだけを下に描く ----
            #   ④のボタンには「引き取れる薬の件数」、⑤には「新着（未読）件数」を出す。
            chosen = view_switcher(len(view_a), len(view_expiry), len(view_receive),
                                   n_unread=total_unread)

            # ⑤以外へ切り替えたら、いま開いている会話（msg_open）を閉じる。
            #   ＝⑤へ戻ってきたときは必ず一覧から始まり、前に開いていた相手が勝手に開かない
            #     （本間部長指示「⑤に切り替えただけで相手が選ばれている状態にしない」・誤送信防止）。
            if chosen != VIEW_MESSAGE:
                st.session_state.pop('msg_open', None)

            if chosen == VIEW_DEAD:
                # ---- ②（自店）のデッド品 ----
                st.subheader('②（%s）のデッド品' % my_store)
                # ②の注記：少額カット（在庫金額しきい値未満で非表示）の件数・金額は「自店分」で出す。
                #   全店合計ではなく small_by_store の自店分を使う。0件なら少額の注記は出さない。
                #   しきい値（1,500円）は result['min_supply_amount'] から取る（ベタ書きしない）。
                min_amt = result.get('min_supply_amount', 0)
                sbs = result.get('small_by_store', {}).get(my_store, {'count': 0, 'amt': 0.0})
                cap = ('デッド＝直近6ヶ月以上、出庫（払い出し）がない在庫です。'
                       '薬品名に ⚠ が付いた品は有効期限まで5ヶ月以内です（期限切迫を兼ねています）。'
                       '『出庫可能数』のセルをクリックすると、その場で数量を書き換えられます。'
                       '『滞留』の欄は、その品が何ヶ月つづけてこの表に載っているかです'
                       '（意味は表の下に出ます）。'
                       'デッドストックリストに載せない医薬品は左端の『除外』にチェックを入れて'
                       '「除外を保存」を押してください。'
                       '／在庫金額%s円以上のものだけを載せています。'
                       % '{:,}'.format(min_amt))
                if sbs.get('count'):
                    cap += '（少額のため非表示：%d件・計%s円）' % (
                        sbs['count'], '{:,.0f}'.format(sbs.get('amt', 0.0)))
                st.caption(cap)
                supply_editor(view_a, my_store, backend, exclusions, table_key='dead',
                              n_supply_specified=n_supply_mine, msg_by_store=msg_by_store)
                excluded_section(my_store, backend, exclusions)

            elif chosen == VIEW_EXPIRY:
                # ---- ③（自店）の期限切迫品 ----
                st.subheader('③（%s）の期限切迫品' % my_store)
                st.caption('自店の期限切迫在庫（有効期限まで5ヶ月以内・デッドではないもの）と、その引取候補店。'
                           '期限が近い在庫なので、使ってくれる店へ早めに動かすのが有効です。'
                           '『出庫可能数』のセルをクリックすると、その場で数量を書き換えられます。'
                           '『滞留』の欄は、その品が何ヶ月つづけてこの表に載っているかです'
                           '（意味は表の下に出ます）。'
                           '（③は全行が期限切迫のため、薬品名の ⚠ は付けていません）')
                supply_editor(view_expiry, my_store, backend, exclusions, table_key='expiry',
                              paint_expiry=False, n_supply_specified=n_supply_mine,
                              msg_by_store=msg_by_store)
                excluded_section(my_store, backend, exclusions)

            elif chosen == VIEW_RECEIVE:
                # ---- ④（自店）が引き取れる薬（他店のデッド・期限切迫）＝受け手ビュー ----
                receive_section(view_receive, my_store, backend, reservations, latest,
                                msg_by_store=msg_by_store)
                # ④の下に別枠『いまは在庫があるが、先になら引き取れる薬』（③の改修）
                receive_ref_section(view_receive_ref, my_store, backend, reservations, latest,
                                    msg_by_store=msg_by_store)
                # さらに下に「予約中の品」＋引取依頼書ボタン（④の改修）
                reserved_section(view_reserved, my_store, backend, reservations, result, latest,
                                 msg_by_store=msg_by_store)

            else:
                # ---- ⑤ 店舗間のやり取り（掲示板）＝第2弾 ----
                message_section(my_store, backend, threads, msg_reads)

    # ---- Excelダウンロード（全店一覧・不足一覧は従来どおりこのExcelに入っている）----
    #   ★Excelバイト列（xls）は上のブロックで作る（材料が同じなら前回のものを使い回す・案1）。
    #     ボタンを押していなくても毎回作り直していたのが最大の山だったため、ここでは作らない。
    st.divider()
    if xls is not None:
        # ダウンロードされるファイル名も画面の名前にそろえる（2026-07-28の改称に追随）
        if latest:
            xls_name = 'デッドストックリスト_%s-%s.xlsx' % (latest[:4], latest[4:6])
        else:
            xls_name = 'デッドストックリスト.xlsx'
        st.download_button(
            'この結果をExcel（4シート）でダウンロード',
            data=xls,
            file_name=xls_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ============================================================================
# メイン
# ============================================================================
def main():
    if not password_gate():
        return

    st.title('💊 デッドストックリスト')
    if not gsheet_configured():
        st.warning('（開発モード）Googleシート未接続のため、このブラウザのセッションにだけ保存します。'
                   '本番では Streamlit Cloud の Secrets を設定してください。')

    try:
        backend, is_cloud = get_backend()
    except Exception as e:
        st.error('保管庫（Googleシート）に接続できませんでした。Secrets の設定をご確認ください：%s' % e)
        return

    # 保管庫の読み込みは1回だけ。アップロード欄（当月アップ済みかの判定）と
    #   結果表示の両方で同じ結果を使い回す（Gシートへの往復を2倍にしないため）。
    try:
        stores, latest, index = load_stores_cached(backend)
    except Exception as e:
        show_gsheet_error(e, '保管庫からデータを読めませんでした')
        return

    upload_section(backend, index, latest)
    st.divider()
    results_section(backend, stores, latest, index)


if __name__ == '__main__':
    # streamlit run で起動したときは、この入口スクリプトが __main__ として実行されるため
    # ここが走ってアプリが表示される。テスト目的で import しただけのときは main() は動かない。
    main()
