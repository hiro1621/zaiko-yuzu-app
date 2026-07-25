# -*- coding: utf-8 -*-
"""
店舗間 在庫融通アプリ（streamlit_app.py）＝店舗セルフアップロード型の画面本体

【何をするアプリか】
各店がブラウザから薬VANの在庫ファイル（.xls / .csv / .xlsx）をアップロードすると、
共通エンジン（yuzu_core）が全店ぶんを毎回まとめて再計算し、
  ・自店視点の2ビュー（A：自店の過剰を欲しがる店／B：自店の不足を過剰に持つ店）
  ・全店の融通提案一覧（金額の大きい順）
  ・「現在 N/15店 アップ済み」
を返します。裏側の保管庫はGoogleスプレッドシート（サービスアカウント経由のみ）。
店舗はGシートを直接触らず、必ずこのアプリ経由で読み書きします。

【使い方（店舗）】
  1. 共有パスワードを入れる（1回だけ）。
  2. 自分の店をドロップダウンで選ぶ。
  3. 薬VANの在庫ファイルを選んでアップロードする。
  4. 結果（自店視点の2ビュー・全店一覧・○/15）が表示される。

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

    my_store = st.selectbox(
        '自分の店（必ず選んでください）',
        STORE_NAMES,
        index=STORE_NAMES.index(st.session_state.get('my_store', STORE_NAMES[0]))
        if st.session_state.get('my_store') in STORE_NAMES else 0,
        format_func=lambda n: '%s（%s）' % (n, COMPANY_OF.get(n, '')))
    st.session_state['my_store'] = my_store
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

    if st.button('② この内容でアップロードする', type='primary', disabled=(up is None)):
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
# 結果表示（自店視点2ビュー＋全店一覧）
# ============================================================================
def results_section(backend):
    stores, latest, index = backend.load_current_month_stores()
    status = app_logic.uploaded_status(index, latest, STORE_NAMES)

    st.subheader('現在の状況')
    show_upload_status(status, latest)

    if not stores:
        st.stop()

    # 全店ぶんを毎回まとめて再計算
    result = yuzu_core.compute_matching(stores)
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

    my_store = st.session_state.get('my_store', STORE_NAMES[0])
    my_uploaded = (my_store in status['uploaded'])

    st.divider()
    st.subheader('② 自店（%s）から見た融通' % my_store)
    if not my_uploaded:
        st.warning('まず自店（%s）の在庫ファイルをアップロードしてください。'
                   '自店の過剰・不足が入って初めて、自店視点の提案が出ます。' % my_store)
    else:
        view_a = app_logic.build_view_a(result, my_store)
        view_b = app_logic.build_view_b(result, my_store)

        st.markdown('#### ビューA：自店の《過剰・デッド》を欲しがっている店')
        st.caption('引取候補店（本命）＝ ①不足中 → ②使用中 の順。「参考:過剰だが使用中の店」は別枠（送っても余りが増えがち）。')
        st.dataframe(_df(view_a, [
            '薬品名', '単位', '過剰数', '過剰数金額', '有効期限', '期限切迫区分',
            '要記録警告', '引取候補店（本命）', '参考:過剰だが使用中の店', '医薬品CD']),
            use_container_width=True, hide_index=True)

        st.markdown('#### ビューB：自店が《不足》している薬を、過剰に持っている店')
        st.dataframe(_df(view_b, [
            '薬品名', '在庫数', '安全在庫数', '不足数', '過剰に持つ他店', '医薬品CD']),
            use_container_width=True, hide_index=True)

    st.divider()
    st.subheader('③ 全店の融通提案一覧（金額の大きい順）')
    all_rows = [{
        '出し手店': r['出し手店'], '薬品名': r['薬品名'], '単位': r['単位'],
        'メーカ名': r['メーカ名'], '過剰数': r['過剰数'], '過剰数金額': r['過剰数金額'],
        '有効期限': r['有効期限'], '期限切迫区分': r['期限切迫区分'],
        '要記録警告': r['要記録警告'], '引取候補店': r['引取候補店'],
        '参考:過剰だが使用中の店': r['参考:過剰だが使用中の店'], '医薬品CD': r['医薬品CD'],
    } for r in result['proposal_rows']]
    st.caption('全 %d 件' % len(all_rows))
    st.dataframe(_df(all_rows, [
        '出し手店', '薬品名', '単位', 'メーカ名', '過剰数', '過剰数金額', '有効期限',
        '期限切迫区分', '要記録警告', '引取候補店', '参考:過剰だが使用中の店', '医薬品CD']),
        use_container_width=True, hide_index=True)

    # Excelダウンロード
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
