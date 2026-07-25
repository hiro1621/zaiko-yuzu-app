# -*- coding: utf-8 -*-
"""
在庫融通アプリの「計算の下ごしらえ」＝画面に依存しないロジック（app_logic.py）

【何をするスクリプトか】
Streamlit画面（streamlit_app.py）から切り離した“純粋な処理”だけを集めたファイルです。
Streamlit を import しないので、そのまま単体テストできます（品質管理部の検証もしやすい）。
  ・アップロードされたバイトを読み、様式チェック→前処理→約35列へスリム化
  ・保管庫（Gシート or ローカル）から当月データを集めて compute_matching に渡す
  ・自店視点の2ビュー（A：自店の過剰を欲しがる店／B：自店の不足を過剰に持つ店）を作る
  ・「○/15店」の集計、Excelダウンロード用のバイト生成
  ・鍵が無いローカル検証用の「ローカル保管庫（LocalBackend）」（Gシートと同じ月替わり挙動）

【必要なライブラリのインストール（コマンドプロンプトで実行）】
    pip install openpyxl xlrd==2.0.1
    ※ Gシート連携を使うときだけ： pip install gspread google-auth

※ venv不要。Windows専用パス。コメント・メッセージはすべて日本語です。
"""

import io
import datetime

import yuzu_core
from yuzu_core import REQUIRED_COLS, KEEP_COLS


# ============================================================================
# アップロード1件の下ごしらえ（バイト → 様式チェック → 前処理 → スリム化）
# ============================================================================
def process_upload_bytes(filename, data):
    """
    アップロードされたファイル（バイト）を読み、次を返す辞書。
      header    … 見出し（列名リスト）
      missing   … 不足している必須列（空なら様式OK）
      format_ok … 必須列がすべて揃っているか（True/False）
      n_all     … 生の行数
      n_kept    … 前処理（削除/取扱/OTC除外）後に残った行数
      slim      … 保管庫に入れる約35列（KEEP_COLS）だけの辞書リスト（前処理済み）
    ※ 列は列名で特定するので、元ファイルの列の並び順は問いません。
    """
    header = yuzu_core.read_header_from_bytes(filename, data)
    missing = [c for c in REQUIRED_COLS if c not in header]
    format_ok = (len(missing) == 0)

    rows_all = yuzu_core.read_rows_from_bytes(filename, data)
    rows_kept = [r for r in rows_all if yuzu_core.preprocess_keep(r)]
    slim = yuzu_core.slim_rows(rows_kept)
    return {
        'header': header,
        'missing': missing,
        'format_ok': format_ok,
        'n_all': len(rows_all),
        'n_kept': len(rows_kept),
        'slim': slim,
    }


# ============================================================================
# 自店視点の2ビューを作る
# ============================================================================
def build_view_a(result, store_name):
    """
    ビューA：自店の過剰・デッドを、欲しがっている（＝引き取ってくれそうな）店。
      融通提案（result['proposal_rows']）のうち、出し手店が自店の行を取り出す。
      本命候補（①不足中→②使用中）＝『引取候補店』列、参考（③過剰だが使用中）は別枠。
    """
    out = []
    for r in result['proposal_rows']:
        if r['出し手店'] != store_name:
            continue
        out.append({
            '薬品名': r['薬品名'], '単位': r['単位'], '過剰数': r['過剰数'],
            '過剰数金額': r['過剰数金額'], '有効期限': r['有効期限'],
            '期限切迫区分': r['期限切迫区分'], '要記録警告': r['要記録警告'],
            '引取候補店（本命）': r['引取候補店'],
            '参考:過剰だが使用中の店': r['参考:過剰だが使用中の店'],
            '医薬品CD': r['医薬品CD'],
        })
    return out


def build_view_b(result, store_name):
    """
    ビューB：自店が不足している薬を、過剰に持っている店。
      不足品目一覧（result['shortage_rows']）のうち、店が自店の行を取り出す。
    """
    out = []
    for r in result['shortage_rows']:
        if r['店'] != store_name:
            continue
        out.append({
            '薬品名': r['薬品名'], '在庫数': r['在庫数'], '安全在庫数': r['安全在庫数'],
            '不足数': r['不足数'], '過剰に持つ他店': r['過剰に持つ他店'],
            '医薬品CD': r['医薬品CD'],
        })
    return out


# ============================================================================
# 「○/N店 アップ済み」の集計
# ============================================================================
def uploaded_status(index, latest, all_store_names):
    """
    index（_index の内容）と当月年月 latest から、アップ状況を返す。
      戻り値：{'uploaded':[当月アップ済み店], 'missing':[未アップ店], 'ng':[様式NG店], 'n':当月店数}
    """
    uploaded = []
    ng = []
    for name, e in index.items():
        if e.get('ym', '') != latest:
            continue
        if e.get('format', '').startswith('NG'):
            ng.append(name)
        else:
            uploaded.append(name)
    up_set = set(uploaded) | set(ng)
    missing = [n for n in all_store_names if n not in up_set]
    return {'uploaded': uploaded, 'missing': missing, 'ng': ng, 'n': len(uploaded)}


# ============================================================================
# Excelダウンロード用のバイト生成（yuzu_core.write_excel を BytesIO へ）
# ============================================================================
def excel_bytes(result, base_ym_disp, csv_base_disp):
    """ 結果をExcel（4シート）にして bytes で返す。画面のダウンロードボタン用。 """
    bio = io.BytesIO()
    yuzu_core.write_excel(
        bio, base_ym_disp, csv_base_disp,
        result['proposal_rows'], result['shortage_rows'],
        result['store_names'], result['matrix_rows'], result['summary_rows'])
    bio.seek(0)
    return bio.getvalue()


# ============================================================================
# ローカル保管庫（LocalBackend）… 鍵が無い環境（ローカル検証・Secrets未設定）用
#   Gシート保管庫（gsheet_store）と同じ「月替わり挙動・当月のみ計算」を、
#   メモリ上（辞書）で再現します。streamlit_app.py は Secrets の有無で
#   この LocalBackend か GSheet を選びます（インターフェースは同じ）。
# ============================================================================
class LocalBackend:
    """ メモリ上の保管庫。state 辞書を外から渡せば（例：st.session_state）永続化にも使える。 """

    def __init__(self, state=None):
        # state = {'index': {店名:{...}}, 'raw': {店名:[slim行,...]}, 'prev': {ym:...}}
        if state is None:
            state = {}
        state.setdefault('index', {})
        state.setdefault('raw', {})
        state.setdefault('prev', {})
        state.setdefault('results', None)
        self.state = state

    # --- gsheet_store と同じメソッド名・戻り値でそろえる ---
    def _latest_ym(self):
        yms = [e.get('ym', '') for e in self.state['index'].values() if e.get('ym', '')]
        return max(yms) if yms else None

    def save_store_upload(self, store_name, ym, slim_rows, filename, format_ok):
        old_latest = self._latest_ym()
        # 月替わり退避：新しい月が始まるなら、直前の結果を prev へ
        if old_latest is not None and ym > old_latest:
            if self.state.get('results') is not None:
                self.state['prev'][old_latest] = self.state['results']
        self.state['raw'][store_name] = list(slim_rows)
        self.state['index'][store_name] = {
            'ym': ym,
            'uploaded_at': datetime.datetime.now().strftime('%Y/%m/%d %H:%M'),
            'rows': str(len(slim_rows)),
            'format': 'OK' if format_ok else 'NG(別様式)',
            'filename': filename,
        }
        return self.state['index']

    def load_current_month_stores(self):
        index = self.state['index']
        latest = self._latest_ym()
        stores = []
        if latest is None:
            return stores, latest, index
        for name, e in index.items():
            if e.get('ym', '') != latest:
                continue
            if e.get('format', '').startswith('NG'):
                continue
            rows = self.state['raw'].get(name, [])
            # KEEP_COLS に無い列があれば空文字で補う（Gシート往復と挙動をそろえる）
            fixed = []
            for r in rows:
                d = dict(r)
                for c in KEEP_COLS:
                    d.setdefault(c, '')
                fixed.append(d)
            y, m = int(latest[:4]), int(latest[4:6])
            base_date = datetime.date(y, m, 1)
            stores.append({'name': name, 'ym': latest, 'base_date': base_date, 'rows': fixed})
        return stores, latest, index

    def save_results(self, payload):
        """ 結果を保持（次の月替わりで prev へ退避できるように）。 """
        self.state['results'] = payload
