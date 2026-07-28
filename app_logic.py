# -*- coding: utf-8 -*-
"""
在庫融通アプリの「計算の下ごしらえ」＝画面に依存しないロジック（app_logic.py）

【何をするスクリプトか】
Streamlit画面（streamlit_app.py）から切り離した“純粋な処理”だけを集めたファイルです。
Streamlit を import しないので、そのまま単体テストできます（品質管理部の検証もしやすい）。
  ・アップロードされたバイトを読み、様式チェック→前処理→約35列へスリム化
  ・保管庫（Gシート or ローカル）から当月データを集めて compute_matching に渡す
  ・自店視点のビュー（デッド品／期限切迫品／自店の不足をデッド・期限切迫で持つ店）を作る
  ・「○/15店」の集計、Excelダウンロード用のバイト生成
  ・鍵が無いローカル検証用の「ローカル保管庫（LocalBackend）」（Gシートと同じ月替わり挙動）

  【2026-07-27 追加】④受け手ビュー build_view_receive() を追加しました。
    他店がデッド・期限切迫で持っていて、自店（引取候補店＝自店）が引き取れば活かせる品の一覧を作ります。
    店名の突合は完全一致（==）のみ（和光／さと和光は別会社のため in は使わない）。

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
# 自店視点のビューを作る
#   ・build_view_a       … 自店の「デッド品」（種別＝デッド）を欲しがる店
#   ・build_view_expiry  … 自店の「期限切迫品」（種別＝期限切迫）を欲しがる店
#   ・build_view_b       … 自店が不足している薬を、デッド／期限切迫で持っている店
# ============================================================================
def _view_supply(result, store_name, category):
    """
    融通提案（result['proposal_rows']）のうち、出し手店＝自店 かつ 種別＝category の行を、
    自店視点の表（引取候補店つき）に整えて返す。build_view_a / build_view_expiry の共通処理。

    ★『引取候補店』は店名だけの版（proposal_rows の '引取候補店（店名のみ）'）を使う
      （2026-07-28 本間部長指示。tier・消化目安つきだと横に長く、肝心の店名が読みづらいため）。
      Excel／Gシートの『引取候補店』は従来どおり詳細つきのまま。
    """
    out = []
    for r in result['proposal_rows']:
        if r['出し手店'] != store_name:
            continue
        if r.get('種別') != category:
            continue
        out.append({
            '薬品名': r['薬品名'], '単位': r['単位'], '在庫数': r['在庫数'],
            '在庫金額': r['在庫金額'], '有効期限': r['有効期限'],
            '期限切迫区分': r['期限切迫区分'], '区分': r['区分'],
            '引取候補店': r['引取候補店（店名のみ）'],
            # 除外チェック欄で使う内部キー（画面には出さない）
            '_key': r['_ex_key'],
        })
    return out


def build_view_a(result, store_name):
    """
    ②デッド品ビュー：自店の《デッド（不動）》在庫を、欲しがっている店。
      融通提案のうち、出し手店＝自店 かつ 種別＝デッド の行を取り出す。
      候補（①不足中→②使用中）＝『引取候補店』列。
      ※ デッドかつ期限切迫の品もここ（デッド側）に入り、期限切迫区分の値は列で表示する。
    """
    return _view_supply(result, store_name, 'デッド')


def build_view_expiry(result, store_name):
    """
    ③期限切迫品ビュー：自店の《期限切迫》在庫（＝デッドではない純粋な期限切迫）を、欲しがっている店。
      融通提案のうち、出し手店＝自店 かつ 種別＝期限切迫 の行を取り出す。列構成は build_view_a と同じ。
    """
    return _view_supply(result, store_name, '期限切迫')


def build_view_b(result, store_name):
    """
    ビューB：自店が不足している薬を、デッド／期限切迫で持っている店。
      不足品目一覧（result['shortage_rows']）のうち、店が自店の行を取り出す。
      供給元は「デッドまたは期限切迫で持つ他店」に統一（A案。旧・過剰保有基準は廃止）。
    ※2026-07-27 に画面表示からは外した。Excel／Gシートには従来どおり出力している。
    """
    out = []
    for r in result['shortage_rows']:
        if r['店'] != store_name:
            continue
        out.append({
            '薬品名': r['薬品名'], '在庫数': r['在庫数'], '安全在庫数': r['安全在庫数'],
            '不足数': r['不足数'], 'デッド/期限切迫で持つ他店': r['デッド/期限切迫で持つ他店'],
            '医薬品CD': r['医薬品CD'],
        })
    return out


def build_view_receive(result, store_name):
    """
    ④受け手ビュー：他店がデッド／期限切迫で持っていて、自店が引き取れば活かせる品の一覧。
      compute_matching が作る result['candidate_rows']（引取候補店つき・丸め無しの構造化データ）から、
      引取候補店＝自店 の行だけを残す。見るだけの一覧（申込みの仕組みは無し。連絡は電話・デスクネッツ）。

      ★突き合わせは必ず完全一致（==）で行う。in（部分一致）は絶対に使わない。
        理由：「和光」（ソユーズ和光）は「さと和光」（内観堂さと和光・別会社）に文字列として
        含まれるため、in を使うと別会社の品が自店に混ざってしまう。

      並べ替え：不足中（_tier_order=0）を上・使用中（=1）を下。各グループ内は有効期限が近い順。
        有効期限が空（_remain が None）の行は最後にまわす。
        ★『なぜ候補か』列は下のとおり画面から外したが、並べ替えには引き続き _tier_order を使う
          （不足中の品が上に来る優先順位はそのまま）。

      画面用の9列だけに整えて返す（内部用の _tier_order・_remain・引取候補店 は返さない）：
        出し手店／薬品名／単位／在庫数／在庫金額／有効期限／消化目安／区分／医薬品CD
      ※2026-07-28 本間部長指示で『なぜ候補か』（不足中／使用中）列を削除した。
        引き取る側にとっては「自店で使える品かどうか」は消化目安を見れば足り、
        出し手側の tier 区分まで出す必要がないため。
    """
    rows = []
    for r in result.get('candidate_rows', []):
        # ★完全一致（==）のみ。in は使わない（和光／さと和光は別会社で、文字列として一方が他方を含むため）
        if r['引取候補店'] == store_name:
            rows.append(r)
    # 不足中→使用中の順、各グループ内は有効期限の近い順、有効期限が空（_remain=None）の行は最後
    rows.sort(key=lambda r: (r['_tier_order'],
                             r['_remain'] is None,
                             r['_remain'] if r['_remain'] is not None else 0))
    disp_cols = ['出し手店', '薬品名', '単位', '在庫数', '在庫金額', '有効期限',
                 '消化目安', '区分', '医薬品CD']
    return [{c: r[c] for c in disp_cols} for r in rows]


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
        state.setdefault('exclusions', [])   # 店が「出さない」と外した品目
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

    def load_exclusions(self):
        """ 店が「融通に出さない」と外した品目の一覧を返す（Gシート版と同じ形）。 """
        return list(self.state['exclusions'])

    def save_exclusions(self, rows):
        """ 除外リストを丸ごと入れ替える（Gシート版と同じ挙動）。 """
        self.state['exclusions'] = list(rows)
