# -*- coding: utf-8 -*-
"""
店舗間 在庫融通アプリ（streamlit_app.py）＝店舗セルフアップロード型の画面本体

【何をするアプリか】
各店がブラウザから薬VANの在庫ファイル（.xls / .csv / .xlsx）をアップロードすると、
共通エンジン（yuzu_core）が全店ぶんを毎回まとめて再計算し、
  ・自店（②）のデッド品／（③）の期限切迫品（それぞれ欲しがる店つき）
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
  3. 薬VANの在庫ファイルを選んでアップロードする。
  4. 結果（自店のデッド品・期限切迫品／全店一覧／○/14）が表示される。

★数量・金額の定義（本間部長判断 2026-07-27）★
  「在庫数」＝その店がいま持っている全量、「在庫金額」＝在庫数×薬価（薬VANの薬価金額列）。
  旧仕様の「過剰数／過剰数金額」（＝安全在庫を超えた分だけ）は使いません。デッド品は
  在庫まるごとが動かす対象で、過剰数だと0になってしまう品が実在するためです。

★載せない品（本間部長判断 2026-07-27）★
  1) 在庫金額が 1,500円未満の少額品（yuzu_core.CONFIG['min_supply_amount'] で変更可）
  2) 店が「除外」にチェックを入れた品（保管庫の _除外 タブに保存。いつでも戻せる）
  どちらも自店の表だけでなく、全店一覧・他店の参考ビュー・Excel・Gシートから同時に消えます。

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
st.set_page_config(page_title='店舗間 在庫融通', page_icon='💊', layout='wide')


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
# 保管庫バックエンド（Gシート or ローカル）を用意する
#   ・両者は同じメソッド名（save_store_upload / load_current_month_stores）を持つ。
# ============================================================================
class _GSheetAdapter:
    """ gsheet_store（モジュール関数）を、ローカル保管庫と同じ形で呼べるようにする薄いラッパ。 """

    def __init__(self, sh):
        self.sh = sh

    def save_store_upload(self, store_name, ym, slim_rows, filename, format_ok):
        return gsheet_store.save_store_upload(self.sh, store_name, ym, slim_rows, filename, format_ok)

    def load_current_month_stores(self):
        return gsheet_store.load_current_month_stores(self.sh)

    def write_results(self, payload):
        gsheet_store.write_results(self.sh, payload)

    def load_exclusions(self):
        return gsheet_store.read_exclusions(self.sh)

    def save_exclusions(self, rows):
        gsheet_store.write_exclusions(self.sh, rows)


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

    st.title('💊 店舗間 在庫融通')
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
_NUM2_COLS = ['在庫数', '在庫金額', '安全在庫数', '不足数']


def _style_expiry(df, paint=True):
    """ 表を見やすく整えた pandas Styler を返す。
          ・数値列（在庫数・在庫金額など）は小数第2位まで（例 7.00 / 49,630.00・カンマ区切り）
          ・paint=True のとき、「期限切迫区分」が非空の行を【薄赤の背景＋黒文字】で行ごと目立たせる
            （②デッド一覧の中で"期限が近いデッド"が一目で分かるように）
          ・paint=False のときは背景を塗らない（③期限切迫品は全行が期限切迫で、
            塗ると表全体が赤くなるだけで情報量がゼロのため。数値フォーマットは②③とも効かせる）。
        ★文字色を黒（#000000）で明示するのは、ダークテーマだと白文字×薄赤で読めなくなるため。
          ライトテーマでは元から黒文字なので見た目は変わらない。
        Styler が使えない環境では素の DataFrame をそのまま返す。 """
    if df is None or df.empty:
        return df

    # 薄赤背景＋黒文字。行全体を塗るので有効期限も一緒に目立つ。
    # 色は #FFE3E6（旧 #FFC7CE より薄いピンク＝現行色と白のちょうど中間）。
    HIT_STYLE = 'background-color: #FFE3E6; color: #000000'

    def _paint(row):
        hit = str(row.get('期限切迫区分', '') or '').strip() != ''
        return [HIT_STYLE if hit else '' for _ in row]

    try:
        sty = df.style
        fmt = {c: '{:,.2f}' for c in _NUM2_COLS if c in df.columns}
        if fmt:
            sty = sty.format(fmt)
        if paint and '期限切迫区分' in df.columns:
            sty = sty.apply(_paint, axis=1)
        return sty
    except Exception:
        # 万一 Styler が使えない環境でも、表示自体は素のDataFrameで続行する
        return df


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
      paint_expiry  … 期限切迫を兼ねる行を薄赤に塗るか（②=True／③=False）。
                      ③は全行が期限切迫なので塗ると真っ赤になるだけ＝Falseで白のままにする。
    表示は8列（薬品名／単位／在庫数／在庫金額／有効期限／期限切迫区分／区分／引取候補店）。
    ・②では期限切迫を兼ねる行を pandas Styler で薄赤（背景 #FFE3E6・黒文字）に塗る。
      チェック欄つきの旧方式では行の色を付けられず ⚠ 記号で代替していたが、
      行選択方式なら Styler の背景色と左端の選択チェックが両立するので色を戻した。
    ・数値（在庫数・在庫金額）はカンマ区切り＋小数第2位で表示する（②③とも共通）。
    ・左端の□で行を選び「除外を保存」を押すと、その品目は融通提案から完全に消える
      （自店の表・全店一覧・他店の参考ビュー・Excel・Gシートのすべてから）。
    """
    if not rows:
        st.info('該当する品目はありません。')
        return

    disp_cols = ['薬品名', '単位', '在庫数', '在庫金額',
                 '有効期限', '期限切迫区分', '区分', '引取候補店']
    df = pd.DataFrame([{c: r[c] for c in disp_cols} for r in rows])

    event = st.dataframe(
        _style_expiry(df, paint=paint_expiry),
        hide_index=True,
        use_container_width=True,
        on_select='rerun',
        selection_mode='multi-row',
        key='table_%s' % table_key)

    # 選択された行を「行の位置」で拾い、位置で元の rows から引く（隠し列の値に頼らない確実な方法）
    picked = _selected_rows(event)
    checked = [rows[i] for i in picked if 0 <= i < len(rows)]
    n = len(checked)
    if st.button('除外を保存（%d件）' % n, key='btn_%s' % table_key, disabled=(n == 0)):
        now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')
        keep = list(exclusions)
        have = {(r['店名'], r['除外キー']) for r in keep}
        added = 0
        for d in checked:
            pair = (my_store, d['_key'])
            if pair in have:
                continue
            keep.append({'店名': my_store, '除外キー': d['_key'],
                         '薬品名': d['薬品名'], '除外日時': now})
            have.add(pair)
            added += 1
        backend.save_exclusions(keep)
        st.success('%d件を融通の対象から外しました。' % added)
        st.rerun()


def excluded_section(my_store, backend, exclusions):
    """ 「除外中の品目」を折りたたみで出し、選んで元に戻せるようにする（②③と同じ行選択方式）。
        操作方法が画面内で2種類あると迷うため、②③のチェックと同じ「左端の□で選ぶ」に統一している。 """
    mine = [r for r in exclusions if r['店名'] == my_store]
    with st.expander('除外中の品目（%d件）＝融通の対象から外している薬' % len(mine)):
        if not mine:
            st.write('いまは1件もありません。')
            return
        st.caption('融通の対象に戻したい品を左端の□で選び、下のボタンを押してください。')
        df = pd.DataFrame([{'薬品名': r['薬品名'], '除外日時': r['除外日時']} for r in mine])
        event = st.dataframe(
            df, hide_index=True, use_container_width=True,
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
            st.success('%d件を融通の対象に戻しました。' % n)
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
# アップロード欄
# ============================================================================
def upload_section(backend):
    st.subheader('① 自分の店を選んで、在庫ファイルをアップロード')

    # 店舗名の選択（先頭は「選択してください」。会社名（ソユーズ/内観堂）の付記は無し）
    options = [SENTINEL_STORE] + STORE_NAMES
    prev = st.session_state.get('my_store')
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
        st.caption('→ **%s** で提出します。' % my_store)
    else:
        # 未選択 → ラベルを赤の太字で強調
        label_ph.markdown(':red[**店舗名（必須）**]')
        my_store = None
        st.session_state['my_store'] = None
        st.caption('※ まず店舗名を選んでください。')

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

    can_upload = (up is not None) and (my_store is not None)
    if st.button('② この内容でアップロードする', type='primary', disabled=not can_upload):
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
            st.error('このファイルは在庫融通に必要な列が足りません（別様式の可能性）。'
                     '不足列：' + '、'.join(res['missing']))
            st.stop()

        backend.save_store_upload(my_store, ym, res['slim'], up.name, res['format_ok'])
        st.success('%s の %s年%s月分を受け付けました（%d行・前処理後）。'
                   % (my_store, ym[:4], ym[4:6], res['n_kept']))
        st.rerun()


# ============================================================================
# 結果表示（②自店のデッド品／③自店の期限切迫品＋Excel）
#   ※2026-07-27：画面から「参考ビューB（自店の不足を持つ店）」と「④全店一覧」を外した。
#     どちらも Excel（4シート）・Gシートには従来どおり出力している（画面が長すぎる対策）。
# ============================================================================
def results_section(backend):
    stores, latest, index = backend.load_current_month_stores()
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
        st.warning('除外リストを読めませんでした（除外なしで表示します）：%s' % e)
        exclusions = []

    # 全店ぶんを毎回まとめて再計算
    result = yuzu_core.compute_matching(stores, excluded=_exclusion_set(exclusions))
    base_ym_disp = ('%s年%s月' % (latest[:4], latest[4:6])) if latest else '不明'
    csv_base_disp = datetime.date.today().strftime('%Y/%m/%d')

    # 結果を保管庫にも書き戻す（次の月替わりで前月退避できるように）
    try:
        payload = gsheet_store.build_results_payload(result, base_ym_disp, csv_base_disp)
        if hasattr(backend, 'write_results'):
            backend.write_results(payload)
        elif hasattr(backend, 'save_results'):
            backend.save_results(payload)
    except Exception as e:
        st.warning('結果の保管庫への書き戻しに失敗しました（画面表示は続行します）：%s' % e)

    my_store = st.session_state.get('my_store')

    st.divider()
    if not my_store:
        st.subheader('②（自店）のデッド品')
        st.info('上で自店の店舗名を選ぶと、自店のデッド品（②）・期限切迫品（③）と、'
                'それぞれを欲しがっている店が表示されます。')
    else:
        my_uploaded = (my_store in status['uploaded'])
        # ---- ②（自店）のデッド品 ----
        st.subheader('②（%s）のデッド品' % my_store)
        if not my_uploaded:
            st.warning('まず自店（%s）の在庫ファイルをアップロードしてください。'
                       '自店のデッド・期限切迫が入って初めて、自店視点の提案が出ます。' % my_store)
        else:
            view_a = app_logic.build_view_a(result, my_store)       # 種別＝デッド
            view_expiry = app_logic.build_view_expiry(result, my_store)  # 種別＝期限切迫

            # ②の注記：少額カット（在庫金額しきい値未満で非表示）の件数・金額は「自店分」で出す。
            #   全店合計ではなく small_by_store の自店分を使う。0件なら少額の注記は出さない。
            #   しきい値（1,500円）は result['min_supply_amount'] から取る（ベタ書きしない）。
            min_amt = result.get('min_supply_amount', 0)
            sbs = result.get('small_by_store', {}).get(my_store, {'count': 0, 'amt': 0.0})
            cap = ('行が薄赤の品は期限切迫を兼ねています。'
                   '融通に出さない品は左端の□にチェックを入れて「除外を保存」を押してください。'
                   '／在庫金額%s円以上のものだけを載せています。'
                   % '{:,}'.format(min_amt))
            if sbs.get('count'):
                cap += '（少額のため非表示：%d件・計%s円）' % (
                    sbs['count'], '{:,.0f}'.format(sbs.get('amt', 0.0)))
            st.caption(cap)
            supply_editor(view_a, my_store, backend, exclusions, table_key='dead')

            # ---- ③（自店）の期限切迫品 ----
            st.subheader('③（%s）の期限切迫品' % my_store)
            st.caption('自店の期限切迫在庫（デッドではないもの）と、その引取候補店。'
                       '期限が近い在庫なので、使ってくれる店へ早めに動かすのが有効です。'
                       '（③は全行が期限切迫のため、背景は白のままにしています）')
            supply_editor(view_expiry, my_store, backend, exclusions, table_key='expiry',
                          paint_expiry=False)

            # ---- 除外中の品目（戻せる） ----
            excluded_section(my_store, backend, exclusions)

    # ---- ④ Excelダウンロード（全店一覧・不足一覧は従来どおりこのExcelに入っている）----
    st.divider()
    try:
        xls = app_logic.excel_bytes(result, base_ym_disp, csv_base_disp)
        st.download_button(
            '④ この結果をExcel（4シート）でダウンロード',
            data=xls,
            file_name='在庫融通リスト_%s-%s.xlsx' % (latest[:4], latest[4:6]) if latest else '在庫融通リスト.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        st.warning('Excel生成に失敗しました：%s' % e)


# ============================================================================
# メイン
# ============================================================================
def main():
    if not password_gate():
        return

    st.title('💊 店舗間 在庫融通')
    if not gsheet_configured():
        st.warning('（開発モード）Googleシート未接続のため、このブラウザのセッションにだけ保存します。'
                   '本番では Streamlit Cloud の Secrets を設定してください。')

    try:
        backend, is_cloud = get_backend()
    except Exception as e:
        st.error('保管庫（Googleシート）に接続できませんでした。Secrets の設定をご確認ください：%s' % e)
        return

    upload_section(backend)
    st.divider()
    results_section(backend)


main()
