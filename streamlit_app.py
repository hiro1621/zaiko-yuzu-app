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
from stores_config import STORE_NAMES, STORE_COUNT, COMPANY_OF

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


# 小数第2位まで表示する数値列（在庫数・在庫金額など）
#   ★2026-08-01：④の列を「出し手の在庫数／自店の在庫数」に分けたので、その2列も同じ体裁にする。
#     これらの列名は④（と別枠）にしか出ないので、②③の見た目には影響しない。
_NUM2_COLS = ['在庫数', '在庫金額', '安全在庫数', '不足数', '出し手の在庫数', '自店の在庫数']


def _style_expiry(df, paint=True, stag_levels=None, expiry_flags=None):
    """ 表を見やすく整えた pandas Styler を返す。
          ・数値列（在庫数・在庫金額など）は小数第2位まで（例 7.00 / 49,630.00・カンマ区切り）
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
        fmt = {c: '{:,.2f}' for c in _NUM2_COLS if c in df.columns}
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


def stagnation_legend(rows):
    """
    「滞留」列の色の意味を、色チップつきで画面に出す（＝色を見ただけで意味が分かるように）。
      rows … いま表示している行リスト。実際に出ている区分だけを並べるので、
             関係のない色の説明で画面がうるさくならない。
    1件も滞留が無い月は何も出さない。
    """
    counts = app_logic.stagnation_summary(rows)
    if not counts:
        return
    chips = []
    for lv in yuzu_core.STAGNATION_LEGEND_ORDER:
        n = counts.get(lv)
        if not n:
            continue
        bg, fg, _x, desc = yuzu_core.STAGNATION_STYLES[lv]
        chips.append(
            '<span style="background-color:%s;color:%s;padding:2px 8px;'
            'border-radius:4px;border:1px solid #BBB;white-space:nowrap;">'
            '%s <b>%d件</b></span>' % (bg, fg, desc, n))
    st.markdown(
        '<div style="font-size:0.86rem;line-height:2.1;">'
        '<b>「滞留」列の色の見方</b>　' + '　'.join(chips) + '</div>',
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


def supply_editor(rows, my_store, backend, exclusions, table_key, paint_expiry=True):
    """
    ②デッド品・③期限切迫品の表を「行選択方式」で描く。
      rows          … app_logic.build_view_a / build_view_expiry の戻り
      table_key     … 画面部品を区別するための名前（'dead' / 'expiry'）
      paint_expiry  … 有効期限まで5ヶ月以内（＝期限切迫）の行を薄赤に塗るか（②=True／③=False）。
                      ③は全行が期限切迫なので塗ると真っ赤になるだけ＝Falseで白のままにする。
    表示は8列（薬品名／単位／在庫数／在庫金額／有効期限／期限切迫区分／区分／引取候補店）。
    ・②では有効期限まで5ヶ月以内の行を pandas Styler で薄赤（背景 #FFE3E6・黒文字）に塗る。
      チェック欄つきの旧方式では行の色を付けられず ⚠ 記号で代替していたが、
      行選択方式なら Styler の背景色と左端の選択チェックが両立するので色を戻した。
    ・数値（在庫数・在庫金額）はカンマ区切り＋小数第2位で表示する（②③とも共通）。
    ・左端の□で行を選び「除外を保存」を押すと、その品目は融通提案から完全に消える
      （自店の表・全店一覧・他店の参考ビュー・Excel・Gシートのすべてから）。
    """
    if not rows:
        st.info('該当する品目はありません。')
        return

    # 「滞留」は薬品名のすぐ隣に置く。★いちばん右に足すと横スクロールの先に隠れて
    #   気づかれないため、品名と並べて必ず目に入る位置にしている。
    disp_cols = ['薬品名', '滞留', '単位', '在庫数', '在庫金額',
                 '有効期限', '期限切迫区分', '区分', '引取候補店']
    df = pd.DataFrame([{c: r.get(c, '') for c in disp_cols} for r in rows])
    stag_levels = [r.get('_滞留区分', 'new') for r in rows]
    # 行ごとの「有効期限まで5ヶ月以内か」フラグ（②の薄赤ハイライト用）。
    #   yuzu_core が計算した _expiry_flag をそのまま並べるだけ＝判定の出どころを1つに保つ。
    expiry_flags = [bool(r.get('_expiry_flag')) for r in rows]

    event = st.dataframe(
        _style_expiry(df, paint=paint_expiry, stag_levels=stag_levels,
                      expiry_flags=expiry_flags),
        hide_index=True,
        width='stretch',
        on_select='rerun',
        selection_mode='multi-row',
        key='table_%s' % table_key)
    stagnation_legend(rows)

    # 選択された行を「行の位置」で拾い、位置で元の rows から引く（隠し列の値に頼らない確実な方法）
    picked = _selected_rows(event)
    checked = [rows[i] for i in picked if 0 <= i < len(rows)]
    n = len(checked)
    if st.button('除外を保存（%d件）' % n, key='btn_%s' % table_key, disabled=(n == 0)):
        now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')
        keep = list(exclusions)
        have = {(r['店名'], r['除外キー']) for r in keep}
        added = 0
        # ★予約が入っている品は除外しない。黙って外すと、引き取るつもりでいる店の
        #   予約が理由も分からず消えるため。品名と相手店を出して、話を付けてもらう。
        booked = [d for d in checked if d.get('_予約店')]
        if booked:
            st.warning('  \n'.join(
                ['次の品はすでに引取先が決まっているため、除外しませんでした。'
                 '取りやめる場合は相手店に連絡してください：']
                + ['%s → %s が引取予定' % (d['薬品名'], d['_予約店']) for d in booked]))
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
        st.success('%d件をデッドストックから外しました。' % added)
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
VIEW_ORDER = [VIEW_DEAD, VIEW_EXPIRY, VIEW_RECEIVE]

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


def view_switcher(n_dead, n_expiry, n_receive):
    """
    ②③④を切り替えるボタンを描き、選ばれた画面の記号（VIEW_*）を返す。
      ・件数をボタンに入れて、開く前に中身があるかどうか分かるようにする。
      ・★選択肢そのものは VIEW_* の記号にして、見た目の文字は format_func で作る。
        ボタンの文字（件数入り）を選択肢にしてしまうと、除外を保存して件数が変わった瞬間に
        「保存されている選択」が選択肢の中から消えて、③④を見ていても②に戻されてしまう。
      ・★segmented_control は選択中のボタンをもう一度押すと「選択なし（None）」を返すので、
        そのときは②に戻す（画面が空になるのを防ぐ）。
      ・segmented_control が無い古いStreamlitでは st.radio（横並び）に自動で切り替える。
    """
    labels = {
        VIEW_DEAD:    '②  デッド品（%d件）' % n_dead,
        VIEW_EXPIRY:  '③  期限切迫品（%d件）' % n_expiry,
        VIEW_RECEIVE: '④  引き取れる薬（%d件）' % n_receive,
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


def receive_section(view_receive, my_store, backend, reservations, ym):
    """
    ④受け手ビュー：他店がデッド・期限切迫で持っていて、自店が引き取れば活かせる品の一覧。
      ・2026-07-28に「予約」を付けた。左端の□で選んで「予約する」を押すと、その品は
        ほかの店の④から消え、出し手の②③には「◯◯が引取予定」と出る。
      ・予約は品目まるごと。数量の相談は従来どおり電話・デスクネッツ。
      ・★保存の直前に予約表を読み直して重複を止める（下の _save_reservations 参照）。
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
    df = pd.DataFrame([{k: v for k, v in r.items() if k not in hidden} for r in view_receive])
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


def receive_ref_section(view_ref, my_store, backend, reservations, ym):
    """
    ④の別枠『いまは在庫があるが、先になら引き取れる薬』（③の改修・2026-08-01）。
      ・自店もその薬を使っているが、いま在庫を余らせている品（tier③参考）。今すぐ引き取ると
        移した先で新しいデッドを作りかねないので、本来の④からは外している。
      ・ここから「1〜3ヶ月後」の受取予定月つきで予約できる（本命の狙い）。
      ・④本体と同じ方式（st.dataframe + 行選択）・同じ2列構成（出し手の在庫数＋自店の在庫数）。
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
        df = pd.DataFrame([{k: v for k, v in r.items() if k not in hidden} for r in view_ref])
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


def reserved_section(view_reserved, my_store, backend, reservations, result, latest):
    """ 「予約中の品」を折りたたみで出し、選んで取り消せるようにする（②③の除外と同じ操作感）。
        ★2026-08-01：一番下に「引取依頼書（出し手店ごとのExcel）」を作るボタンを足した（④の改修）。 """
    with st.expander('予約中の品（%d件）＝自店が引き取ると押さえている薬' % len(view_reserved)):
        if not view_reserved:
            st.write('いまは1件もありません。')
            return
        st.caption('予約をやめる品を左端の□で選び、下のボタンを押してください。'
                   '取り消すと、その品はまたほかの店の一覧にも出るようになります。')
        hidden = ('_出し手店', '_key', '_受取予定月')
        df = pd.DataFrame([{k: v for k, v in r.items() if k not in hidden}
                           for r in view_reserved])
        event = st.dataframe(
            _style_expiry(df, paint=False), hide_index=True, width='stretch',
            on_select='rerun', selection_mode='multi-row', key='table_reserved')
        picked = _selected_rows(event)
        checked = [view_reserved[i] for i in picked if 0 <= i < len(view_reserved)]
        n = len(checked)
        if st.button('予約を取り消す（%d件）' % n, key='btn_unreserve', disabled=(n == 0)):
            keep = app_logic.cancel_reservations(reservations, my_store, checked, latest)
            backend.save_reservations(keep)
            st.session_state.pop('pickup_xls', None)   # 予約が変わったので作りかけの帳票は捨てる
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
            keep = [r for r in exclusions
                    if not (r['店名'] == my_store and r['除外キー'] in back)]
            backend.save_exclusions(keep)
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

    # 除外リスト（店が「この品は出さない」と外した品目）を読み、計算から外す
    try:
        exclusions = backend.load_exclusions()
    except Exception as e:
        show_gsheet_error(e, '除外リストを読めませんでした（除外なしで表示します）', 'warning')
        exclusions = []

    # 予約リスト（受け手の店が「うちが引き取る」と押さえた品目）を読む。
    #   ★対象年月が当月と一致するものだけを効かせる（reservation_map の中で絞る）＝
    #     月が変わったら前月の予約は自動で無効になる。
    try:
        reservations = backend.load_reservations()
    except Exception as e:
        show_gsheet_error(e, '予約リストを読めませんでした（予約なしで表示します）', 'warning')
        reservations = []
    reserved = app_logic.reservation_map(reservations, latest)

    # 全店ぶんを毎回まとめて再計算
    #   ※予約は出し手から品を取り除かない（件数・金額・自己検算は動かない）。印を付けるだけ。
    result = yuzu_core.compute_matching(stores, excluded=_exclusion_set(exclusions),
                                        reserved=reserved)

    # 滞留（同じ品が何ヶ月つづけて載っているか）を書き込む。
    #   ★compute_matching の直後・ビューを作る前に済ませる。ここで result に書き込むので、
    #     ②③④の画面・Excel・Gシートのすべてに同じ値が行き渡る。
    #   前月の記録がまだ無い月（この機能を入れた最初の月）は全部「今月から」になる＝
    #     画面は今までどおりの見た目で、2ヶ月目以降から色が付きはじめる。
    prev_map = load_prev_cached(backend, latest)
    app_logic.apply_stagnation(result, prev_map)

    base_ym_disp = ('%s年%s月' % (latest[:4], latest[4:6])) if latest else '不明'
    csv_base_disp = datetime.date.today().strftime('%Y/%m/%d')

    # 結果を保管庫にも書き戻す（次の月替わりで前月退避できるように）
    #   ★中身が前回と同じなら書かない。結果4タブの書き戻しは clear＋update で8回の
    #     API呼び出しになり、行を選ぶたびにこれをやるとGoogleの上限に当たって
    #     画面が固まる。結果が変わっていないなら書き直す意味もない。
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

            # ---- ②③④の切替ボタン。選ばれた1つだけを下に描く ----
            #   ④のボタンには「引き取れる薬の件数」を出す。予約中の件数は④の中で別に出す。
            chosen = view_switcher(len(view_a), len(view_expiry), len(view_receive))

            if chosen == VIEW_DEAD:
                # ---- ②（自店）のデッド品 ----
                st.subheader('②（%s）のデッド品' % my_store)
                # ②の注記：少額カット（在庫金額しきい値未満で非表示）の件数・金額は「自店分」で出す。
                #   全店合計ではなく small_by_store の自店分を使う。0件なら少額の注記は出さない。
                #   しきい値（1,500円）は result['min_supply_amount'] から取る（ベタ書きしない）。
                min_amt = result.get('min_supply_amount', 0)
                sbs = result.get('small_by_store', {}).get(my_store, {'count': 0, 'amt': 0.0})
                cap = ('デッド＝直近6ヶ月以上、出庫（払い出し）がない在庫です。'
                       '行が薄赤の品は有効期限まで5ヶ月以内です（期限切迫を兼ねています）。'
                       '『滞留』の欄は、その品が何ヶ月つづけてこの表に載っているかです'
                       '（色の意味は表の下に出ます）。'
                       'デッドストックリストに載せない医薬品は左端の□にチェックを入れて'
                       '「除外を保存」を押してください。'
                       '／在庫金額%s円以上のものだけを載せています。'
                       % '{:,}'.format(min_amt))
                if sbs.get('count'):
                    cap += '（少額のため非表示：%d件・計%s円）' % (
                        sbs['count'], '{:,.0f}'.format(sbs.get('amt', 0.0)))
                st.caption(cap)
                supply_editor(view_a, my_store, backend, exclusions, table_key='dead')
                excluded_section(my_store, backend, exclusions)

            elif chosen == VIEW_EXPIRY:
                # ---- ③（自店）の期限切迫品 ----
                st.subheader('③（%s）の期限切迫品' % my_store)
                st.caption('自店の期限切迫在庫（有効期限まで5ヶ月以内・デッドではないもの）と、その引取候補店。'
                           '期限が近い在庫なので、使ってくれる店へ早めに動かすのが有効です。'
                           '『滞留』の欄は、その品が何ヶ月つづけてこの表に載っているかです'
                           '（色の意味は表の下に出ます）。'
                           '（③は全行が期限切迫のため、背景は白のままにしています）')
                supply_editor(view_expiry, my_store, backend, exclusions, table_key='expiry',
                              paint_expiry=False)
                excluded_section(my_store, backend, exclusions)

            else:
                # ---- ④（自店）が引き取れる薬（他店のデッド・期限切迫）＝受け手ビュー ----
                receive_section(view_receive, my_store, backend, reservations, latest)
                # ④の下に別枠『いまは在庫があるが、先になら引き取れる薬』（③の改修）
                receive_ref_section(view_receive_ref, my_store, backend, reservations, latest)
                # さらに下に「予約中の品」＋引取依頼書ボタン（④の改修）
                reserved_section(view_reserved, my_store, backend, reservations, result, latest)

    # ---- Excelダウンロード（全店一覧・不足一覧は従来どおりこのExcelに入っている）----
    st.divider()
    try:
        xls = app_logic.excel_bytes(result, base_ym_disp, csv_base_disp)
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
    except Exception as e:
        st.warning('Excel生成に失敗しました：%s' % e)


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
