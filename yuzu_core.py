# -*- coding: utf-8 -*-
"""
在庫融通 マッチング共通エンジン（yuzu_core.py）

【何をするスクリプトか】
店舗間 在庫融通ツールの「頭脳」だけを1ファイルにまとめた共通エンジンです。
  ・薬VANの在庫マスタ（.csv / .xls / .xlsx）を「バイト＋ファイル名」で読み込む道具
  ・過剰／不足の判定、引取候補店（①不足中→②使用中→③参考:過剰）の組み立て
  ・全店ぶんを一括計算する compute_matching(stores)（4シート相当データ＋自己検算を返す）
  ・Excel出力 write_excel()
を提供します。

このファイルは「単一の真実」です。次の2つが必ず同じ計算をするよう、両者ともここを呼びます。
  ・コマンド版：create_yuzu_list.py（input フォルダのファイルをパスから読む）
  ・アプリ版　：streamlit_app.py（店がアップロードしたファイルのバイトをそのまま読む）
ファイル読み込みを「バイト＋ファイル名」に一般化してあるので、
コマンド版はパスからバイトを読んで、アプリ版はアップロード物のバイトを、同じ関数に渡せます。

【必要なライブラリのインストール（コマンドプロンプトで実行）】
    pip install openpyxl xlrd==2.0.1
    ※ xlrd は薬VANが .xls（古いExcel＝BIFF/OLE2形式）で出力したファイルを読むために使います。
      xlrd 2系は .xls 専用（.xlsx は読みません）。.xlsx は openpyxl で読みます。
    ※ .csv だけしか使わないなら xlrd は無くても動きます。

※ venv不要。Windows専用パス。コメント・ログ・エラーメッセージはすべて日本語です。
"""

import os
import csv
import io
import datetime
from collections import defaultdict, Counter

# openpyxl（Excel出力・.xlsx読み込み）
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# 設定（CONFIG）… 抽出条件・列名はここに集約。ロジック本体に値を直書きしない。
#   ※ 入力／出力フォルダやGシート設定ファイルのパスは、呼び出し側（コマンド版）が持ちます。
#     ここには「計算のルール」だけを置きます。
# ============================================================================
CONFIG = {
    # --- 前処理の除外（True で除外を有効化）---
    'exclude_deleted':   True,   # 削除フラグ = TRUE の行を除外
    'exclude_untreated': True,   # 取扱フラグ = FALSE の行を除外
    'exclude_otc':       True,   # OTCフラグ = TRUE の行を除外

    # --- キー列（店舗間の突合キー。位置ではなく列名で特定する）---
    'col_key':          '個別医薬品CD',   # 第一キー（YJコード）
    'col_key_fallback': 'レセプト電算CD', # キーが空のときの代用

    # --- 直近6ヶ月の出庫（当月は未確定なので含めず、前月〜六ヶ月前の6つを使う）---
    # ★重要★ 薬VANの「出庫数」は、実際に調剤して払い出すと【マイナス（負値）】で記録されます。
    #   （例：アリピプラゾール錠６ｍｇ「ＹＤ」は6ヶ月の出庫数合計 -2778 ＝バリバリ使用中）
    #   一方「出庫回数」は払い出した回数で【プラス（正値）】です。
    #   → 「使っているか？」の判定は符号に左右されない“出庫回数合計>0”で行い、
    #      月あたりの消化量は“出庫数合計の絶対値÷6”で求めます（下の usage_* 関数・pace 参照）。
    'outflow_qty_cols_6m': ['前月出庫数', '前々月出庫数', '前々々月出庫数',
                            '四ヶ月前出庫数', '五ヶ月前出庫数', '六ヶ月前出庫数'],
    'outflow_cnt_cols_6m': ['前月出庫回数', '前々月出庫回数', '前々々月出庫回数',
                            '四ヶ月前出庫回数', '五ヶ月前出庫回数', '六ヶ月前出庫回数'],

    # --- 引取候補の3段階（tier）判定 ---
    #   ① 不足中          ：受け取り側が 安全在庫数>0 かつ 在庫数<安全在庫数（is_shortage）
    #   ② 使用中(適正)    ：不足中でなく、6ヶ月出庫回数合計>0 で、かつ受け取り側がその品目を
    #                        「過剰保有」していない（＝ここに送ると有効活用される本命候補）
    #   ③ 参考:使用中(過剰)：6ヶ月出庫回数合計>0 だが、受け取り側も「過剰保有」している
    #                        （＝送っても余りを増やすだけなので“参考”として下位・別枠に表示）
    #
    # 「過剰保有（holds_excess）」の定義 ＝ 受け取り側がすでにその品目を余らせているか：
    #   過剰在庫区分が非空 または 不動区分が非空 または 過剰数>0
    #   ※ 期限切迫は「使えば消える」ので過剰保有には含めない
    #     （出し手判定 is_supplier＝デッド／期限切迫 とは別定義であることに注意。
    #       holds_excess は“受け取り側が余らせているか”の判定なので、A案でも変更しない）
    'holds_excess_flag_cols': ['過剰在庫区分', '不動区分'],  # このどれかが非空なら過剰保有
    'holds_excess_qty_col':   '過剰数',                     # この数量が >0 なら過剰保有
    # tierごとの表示ラベル
    'tier_labels': {'short': '不足中', 'use_ok': '使用中', 'use_ref': '参考:過剰だが使用中'},

    # --- 区分の扱い ---
    # 過剰/不動/期限切迫/発注候補：空欄＝非該当、それ以外（R/B/O 等）＝該当
    # 法規制（麻薬/覚醒剤/向精神薬/毒薬/劇薬）：'0' か空欄＝非該当、それ以外＝該当
    #   ※ 実データでは法規制列は「該当しない」を '0' で表す（空欄ではない）ため区別が必要
    'legal_not_flagged_values': ['', '0'],

    # 融通提案から完全に除外する法規制列（麻薬・覚醒剤）
    'legal_exclude_cols': ['麻薬区分', '覚醒剤区分'],
    # 除外はしないが「要記録」警告を出す法規制列（列名 → 表示ラベル）
    'legal_warn_cols': {'向精神薬区分': '向精神薬', '毒薬区分': '毒薬', '劇薬区分': '劇薬'},

    # 融通提案に載せる最低の在庫金額（円）。これ未満の少額品は載せない。
    #   （本間部長判断 2026-07-27：少額品まで並ぶと本当に動かすべき品が埋もれるため）
    'min_supply_amount': 1500,

    # 引取候補店の最大表示数（あふれた分は「他N店」と表示）
    'max_candidates': 5,
    # 有効期限「1年以内」を黄色でハイライトする閾値（月）
    'expiry_yellow_months': 12,
}

# 必須列（このどれかが欠けている店はスキップ）
REQUIRED_COLS = ['個別医薬品CD', 'レセプト電算CD', '薬品名', '単位', 'メーカ名',
                 '在庫数', '全在庫数', '安全在庫数', '過剰数', '過剰数金額', '薬価金額',
                 '過剰在庫区分', '不動区分', '期限切迫区分', '有効期限', 'ロットNO',
                 '最終出庫日', '削除フラグ', '取扱フラグ', 'OTCフラグ']

# 区分値の分布ログに出す列
DIST_COLS = ['過剰在庫区分', '不動区分', '期限切迫区分', '発注候補区分']

# アプリ版でGシート保管庫へ入れる前に残す「マッチングに要る約35列」。
#   薬VANは243〜300列と巨大で、Gシート（1ブック1000万セル上限）を圧迫するため、
#   エンジンが実際に使う列だけへ圧縮してから保管する（下の slim_row / slim_rows 参照）。
KEEP_COLS = [
    # 突合キー・基本情報
    '個別医薬品CD', 'レセプト電算CD', '薬品名', '単位', 'メーカ名',
    # 在庫・過剰・安全在庫（薬価/薬価金額＝在庫金額を出すのに使う）
    '在庫数', '全在庫数', '安全在庫数', '過剰数', '過剰数金額', '薬価', '薬価金額',
    # 区分（過剰・不動・期限切迫・発注候補）
    '過剰在庫区分', '不動区分', '期限切迫区分', '発注候補区分',
    # 期限・ロット・最終出庫
    '有効期限', 'ロットNO', '最終出庫日',
    # 前処理フラグ（削除・取扱・OTC）
    '削除フラグ', '取扱フラグ', 'OTCフラグ',
    # 法規制区分（除外・警告）
    '麻薬区分', '覚醒剤区分', '向精神薬区分', '毒薬区分', '劇薬区分',
    # 直近6ヶ月の出庫数（負値＝払い出し）
    '前月出庫数', '前々月出庫数', '前々々月出庫数',
    '四ヶ月前出庫数', '五ヶ月前出庫数', '六ヶ月前出庫数',
    # 直近6ヶ月の出庫回数（正値＝払い出し回数）
    '前月出庫回数', '前々月出庫回数', '前々々月出庫回数',
    '四ヶ月前出庫回数', '五ヶ月前出庫回数', '六ヶ月前出庫回数',
]


# ============================================================================
# 小さな道具（数値・日付・区分の判定）
# ============================================================================
def parse_num(s):
    """ "1,428.00" のような桁区切りカンマ＋文字列を float にする。空欄・変換不可は 0.0 """
    if s is None:
        return 0.0
    s = str(s).strip().replace(',', '')
    if s == '':
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(s):
    """ 'YYYY/M/D' 等の日付文字列を date にする。読めなければ None """
    s = (s or '').strip()
    if not s:
        return None
    for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%Y/%m', '%Y%m%d'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def fmt_date(d):
    return d.strftime('%Y/%m/%d') if d else ''


def fmt_ym(d):
    return d.strftime('%Y/%m') if d else '期限不明'


def month_diff(a, b):
    """ a から b までのおおよその月数（bが未来なら正） """
    return (b.year - a.year) * 12 + (b.month - a.month) + (b.day - a.day) / 30.0


def fmt_months(m):
    """ 消化目安の月数を読みやすく（小数1桁、整数なら整数で） """
    m = round(m, 1)
    return str(int(m)) if m == int(m) else str(m)


def fmt_qty(x):
    """ 数量を読みやすく（整数なら整数、端数があれば小数2桁） """
    if x == int(x):
        return str(int(x))
    return str(round(x, 2))


def is_text_flag(v):
    """ 過剰/不動/期限切迫など：空欄でなければ「該当」 """
    return bool(str(v).strip())


def is_legal_flag(v):
    """ 法規制区分：'0' か空欄なら非該当、それ以外なら「該当」 """
    return str(v).strip() not in CONFIG['legal_not_flagged_values']


def g(row, col):
    """ 行から列を安全に取り出す（無ければ空文字） """
    return (row.get(col) or '').strip() if isinstance(row.get(col), str) else (row.get(col) or '')


# ============================================================================
# 行レベルの判定
# ============================================================================
# ★出し手（融通候補）の定義（A案・本間部長承認 2026-07-27）★
#   従来は「過剰在庫区分/不動区分/期限切迫区分のいずれか非空、または過剰数>0」（＝is_overstock）
#   でしたが、純粋な過剰（デッドでも期限切迫でもない、ただ在庫が多いだけの品）を提案から外し、
#   本当に動かすべき「デッド（不動）」と「期限切迫」だけを出し手とすることにしました。
#     ・is_dead(row)   … 不動区分が非空（＝デッド在庫）
#     ・is_expiry(row) … 期限切迫区分が非空
#     ・is_supplier(row)＝ is_dead または is_expiry （＝新しい出し手）
#   カテゴリは排他・デッド優先：デッドなら「デッド」、デッドでなく期限切迫なら「期限切迫」。
#   （デッドかつ期限切迫の品は「デッド」に入れる。ただし期限切迫区分の値は各行に残し、
#     Excelでは赤ハイライトを維持する。）
def is_dead(row):
    """ デッド（不動）在庫か：不動区分が非空なら True """
    return is_text_flag(g(row, '不動区分'))


def is_expiry(row):
    """ 期限切迫か：期限切迫区分が非空なら True """
    return is_text_flag(g(row, '期限切迫区分'))


def is_supplier(row):
    """ 出し手（融通候補）か：デッド または 期限切迫 なら True。
        ※ 純粋な過剰（過剰在庫区分だけ／過剰数>0だけで、不動でも期限切迫でもない品）は
          出し手に含めない（A案）。 """
    return is_dead(row) or is_expiry(row)


def supplier_category(row):
    """ 出し手のカテゴリを返す（排他・デッド優先）：
        デッドなら 'デッド'、デッドでなく期限切迫なら '期限切迫'、どちらでもなければ ''（＝出し手でない）。 """
    if is_dead(row):
        return 'デッド'
    if is_expiry(row):
        return '期限切迫'
    return ''


def is_below_min_amount(row):
    """ 在庫金額が少額（既定 1,500円未満）で、融通提案に載せない品か。
        少額品まで並ぶと、本当に動かすべき品が埋もれてしまうため。
        しきい値は CONFIG['min_supply_amount'] で変えられる。 """
    return stock_amount(row) < CONFIG['min_supply_amount']


def exclusion_key(row):
    """ 除外リスト（店が「この品は出さない」と外した品目）で1品を特定するキー。
        個別医薬品CDが取れればそれを使い、取れない品だけ『名:薬品名』で代用する。
        ※ 保存側（画面）と判定側（ここ）で必ず同じ関数を使うこと。 """
    key, _ = row_key(row)
    return key or ('名:' + g(row, '薬品名'))


def stock_qty(row):
    """ 表示・提案に使う数量＝『在庫数』（その店がいま持っている全量）。
        ※ 旧仕様は『過剰数』（安全在庫を超えた分）でしたが、デッド（不動）品は
          在庫まるごとが動かす対象なので、在庫数に統一しました（本間部長判断 2026-07-27）。
          実データでは304件中141件で 過剰数≠在庫数、過剰数0（＝金額0で最下位に沈む）の
          デッド品も37件あり、在庫数のほうが実態に合います。 """
    return parse_num(g(row, '在庫数'))


def stock_amount(row):
    """ 表示・提案に使う金額＝『在庫金額』＝ 在庫数 × 薬価。
        薬VANの『薬価金額』列がまさにこの値（実データ1,431行中1,430行で一致。
        残る1行は薬VAN側の丸め誤差0.52円）。列が無い・空のときだけ 在庫数×薬価 で補う。
        ※ 旧仕様の『過剰数金額』＝ 過剰数 × 薬価 とは別物です。 """
    a = g(row, '薬価金額')
    if str(a).strip() != '':
        return parse_num(a)
    return parse_num(g(row, '在庫数')) * parse_num(g(row, '薬価'))


def is_shortage(row):
    """ 受け手（不足候補）か：安全在庫数>0 かつ 在庫数<安全在庫数 """
    safe = parse_num(g(row, '安全在庫数'))
    zaiko = parse_num(g(row, '在庫数'))
    return safe > 0 and zaiko < safe


def holds_excess(row):
    """ 受け取り側が『すでに過剰保有』か（引取候補 tier②適正 と tier③参考 の分かれ目）。
        過剰在庫区分が非空 または 不動区分が非空 または 過剰数>0 なら True。
        ※ 期限切迫は「使えば消える」ため、ここでは過剰保有に含めない（出し手判定 is_supplier とは別定義）。 """
    for c in CONFIG['holds_excess_flag_cols']:
        if is_text_flag(g(row, c)):
            return True
    if parse_num(g(row, CONFIG['holds_excess_qty_col'])) > 0:
        return True
    return False


def is_legal_excluded(row):
    """ 麻薬・覚醒剤に該当するか（該当なら融通提案から除外） """
    for c in CONFIG['legal_exclude_cols']:
        if is_legal_flag(g(row, c)):
            return True
    return False


def warn_labels(row):
    """ 向精神薬・毒薬・劇薬など「要記録」警告の文字列を作る """
    out = []
    for c, label in CONFIG['legal_warn_cols'].items():
        if is_legal_flag(g(row, c)):
            out.append(label)
    return '・'.join(out)


def row_key(row):
    """ 突合キーを返す（個別医薬品CD優先、無ければレセプト電算CD、両方空なら None） """
    k = g(row, CONFIG['col_key'])
    if k:
        return k, CONFIG['col_key']
    k2 = g(row, CONFIG['col_key_fallback'])
    if k2:
        return k2, CONFIG['col_key_fallback']
    return None, None


def usage_qty6(row):
    """ 直近6ヶ月の出庫数合計。
    薬VANの出庫数は「払い出し＝マイナス」で記録されるため、この合計は通常マイナスになる。
    消化ペース（月あたり消化量）は、この値の絶対値を使って求める（build_candidates の pace 参照）。
    ※ 返品等でまれに正値が混じっても、使用中の判定は“出庫回数”で行うため実害は出ない。 """
    return sum(parse_num(g(row, c)) for c in CONFIG['outflow_qty_cols_6m'])


def usage_cnt6(row):
    """ 直近6ヶ月の出庫回数合計（払い出した回数＝プラス。使用中かどうかの判定に使う） """
    return sum(parse_num(g(row, c)) for c in CONFIG['outflow_cnt_cols_6m'])


def preprocess_keep(row):
    """ 前処理：削除/取扱/OTC で除外すべきなら False（=残さない） """
    if CONFIG['exclude_deleted'] and g(row, '削除フラグ').upper() == 'TRUE':
        return False
    if CONFIG['exclude_untreated'] and g(row, '取扱フラグ').upper() == 'FALSE':
        return False
    if CONFIG['exclude_otc'] and g(row, 'OTCフラグ').upper() == 'TRUE':
        return False
    return True


def slim_row(row):
    """ アプリ版でGシートへ入れる前に、マッチングに要る約35列（KEEP_COLS）だけへ圧縮する。
        列は列名で特定するので、元ファイルの列順は問わない。無い列は空文字で補う。 """
    return {c: g(row, c) for c in KEEP_COLS}


def slim_rows(rows):
    """ 行のリストをまとめて slim_row する。 """
    return [slim_row(r) for r in rows]


# ============================================================================
# ファイル名の解釈
# ============================================================================
def parse_filename(path_or_name):
    """ 「店舗名_YYYYMM.csv」から 店舗名 と YYYYMM を取り出す（パスでもファイル名でも可）。 """
    base = os.path.splitext(os.path.basename(path_or_name))[0]
    if '_' in base:
        name, tail = base.rsplit('_', 1)
        if len(tail) == 6 and tail.isdigit():
            return name, tail
        # 末尾がYYYYMMでない場合は全体を店名扱い
        return base, None
    return base, None


# ============================================================================
# ファイル読み込み（「バイト＋ファイル名」を入口に一般化）
#   ・コマンド版：パスからバイトを読み、read_store_header / read_store_rows を呼ぶ
#   ・アプリ版　：アップロード物のバイトをそのまま read_*_from_bytes へ渡す
#   どちらも下の共通ロジックに合流するので、計算は完全に一致する。
# ============================================================================
def _norm_number(v):
    """ 数値を、CSV由来の値と揃う文字列にする。
        整数はそのまま（例 202.0→'202'）、端数があれば余計な末尾ゼロを落とす（例 14649.60→'14649.6'）。
        後段の parse_num が float 化するので、多少の体裁差は結果に影響しない。 """
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v).strip()
    if f != f:  # NaN
        return ''
    if f.is_integer():
        return str(int(f))
    # 10桁で丸めて浮動小数点の“ゴミ”を消してから末尾ゼロ・小数点を除去
    return ('%.10f' % f).rstrip('0').rstrip('.')


def _norm_xldate(v, datemode):
    """ Excelのシリアル値（例 46419.0）を 'YYYY/M/D'（ゼロ埋めなし＝CSVと同じ体裁）に変換する。
        有効期限・最終出庫日などの日付列はこれで既存処理（parse_date／そのまま表示）に載る。 """
    import xlrd
    dt = xlrd.xldate.xldate_as_datetime(v, datemode)
    return '%d/%d/%d' % (dt.year, dt.month, dt.day)


def _xls_cell_to_str(cell, datemode):
    """ xlrd のセル1つを、CSV DictReader と同じ「文字列」に正規化する。
        ・空欄／エラー → ''
        ・文字列 → そのまま
        ・数値 → 文字列化（_norm_number）
        ・日付（シリアル値）→ 'YYYY/M/D'（_norm_xldate）
        ・真偽（削除／取扱／OTCフラグ等）→ CSVに合わせて 'TRUE' / 'FALSE' """
    import xlrd
    t = cell.ctype
    v = cell.value
    if t == xlrd.XL_CELL_TEXT:            # 1: 文字列
        return str(v).strip()
    if t == xlrd.XL_CELL_NUMBER:          # 2: 数値
        return _norm_number(v)
    if t == xlrd.XL_CELL_DATE:            # 3: 日付（Excelシリアル値）
        return _norm_xldate(v, datemode)
    if t == xlrd.XL_CELL_BOOLEAN:         # 4: 真偽（薬VANのフラグ列はここ）
        return 'TRUE' if v else 'FALSE'
    # 0:空欄, 5:エラー, 6:ブランク など
    return ''


def _norm_openpyxl_value(v):
    """ openpyxl（.xlsx）のセル値を、CSV DictReader と同じ「文字列」に正規化する。 """
    import datetime as _dt
    if v is None:
        return ''
    if isinstance(v, bool):               # ※bool は int より先に判定する
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, (int, float)):
        return _norm_number(v)
    if isinstance(v, (_dt.datetime, _dt.date)):
        return '%d/%d/%d' % (v.year, v.month, v.day)
    return str(v).strip()


# --- .csv（バイト → utf-8-sig で復号。newline='' で開いていた挙動に合わせる）---
def _read_csv_header_bytes(data):
    text = data.decode('utf-8-sig')
    return list(next(csv.reader(io.StringIO(text, newline=''))))


def _read_csv_rows_bytes(data):
    text = data.decode('utf-8-sig')
    return list(csv.DictReader(io.StringIO(text, newline='')))


# --- .xls（xlrd。file_contents= でバイトから直接開く。シートは "QFORM" 優先、無ければ先頭）---
def _open_xls_sheet_bytes(data):
    import xlrd
    book = xlrd.open_workbook(file_contents=data)
    names = book.sheet_names()
    sheet = book.sheet_by_name('QFORM') if 'QFORM' in names else book.sheet_by_index(0)
    return book, sheet


def _read_xls_header_bytes(data):
    book, sh = _open_xls_sheet_bytes(data)
    if sh.nrows < 1:
        return []
    return [str(sh.cell_value(0, c)) for c in range(sh.ncols)]


def _read_xls_rows_bytes(data):
    book, sh = _open_xls_sheet_bytes(data)
    if sh.nrows < 1:
        return []
    headers = [str(sh.cell_value(0, c)) for c in range(sh.ncols)]
    rows = []
    for rr in range(1, sh.nrows):
        d = {}
        for ci, h in enumerate(headers):
            d[h] = _xls_cell_to_str(sh.cell(rr, ci), book.datemode)
        rows.append(d)
    return rows


# --- .xlsx（openpyxl。BytesIO でバイトから直接開く。シートは "QFORM" 優先、無ければ先頭）---
def _open_xlsx_sheet_bytes(data):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb['QFORM'] if 'QFORM' in wb.sheetnames else wb[wb.sheetnames[0]]
    return wb, ws


def _read_xlsx_header_bytes(data):
    wb, ws = _open_xlsx_sheet_bytes(data)
    try:
        first = next(ws.iter_rows(values_only=True))
    except StopIteration:
        wb.close()
        return []
    header = [('' if h is None else str(h)) for h in first]
    wb.close()
    return header


def _read_xlsx_rows_bytes(data):
    wb, ws = _open_xlsx_sheet_bytes(data)
    it = ws.iter_rows(values_only=True)
    try:
        first = next(it)
    except StopIteration:
        wb.close()
        return []
    headers = [('' if h is None else str(h)) for h in first]
    rows = []
    for raw in it:
        d = {}
        for ci, h in enumerate(headers):
            v = raw[ci] if ci < len(raw) else None
            d[h] = _norm_openpyxl_value(v)
        rows.append(d)
    wb.close()
    return rows


def _ext_of(name):
    return os.path.splitext(name)[1].lower()


def read_header_from_bytes(filename, data):
    """ 入力ファイルの見出し行（列名のリスト）を、形式に依らず返す（バイト入力）。 """
    ext = _ext_of(filename)
    if ext == '.csv':
        return _read_csv_header_bytes(data)
    if ext == '.xls':
        return _read_xls_header_bytes(data)
    if ext == '.xlsx':
        return _read_xlsx_header_bytes(data)
    raise ValueError('対応していない拡張子です（.csv / .xls / .xlsx のみ対応）：%s' % ext)


def read_rows_from_bytes(filename, data):
    """ 入力ファイルの全データ行を、形式に依らず「列名→文字列」の辞書のリストで返す（バイト入力）。 """
    ext = _ext_of(filename)
    if ext == '.csv':
        return _read_csv_rows_bytes(data)
    if ext == '.xls':
        return _read_xls_rows_bytes(data)
    if ext == '.xlsx':
        return _read_xlsx_rows_bytes(data)
    raise ValueError('対応していない拡張子です（.csv / .xls / .xlsx のみ対応）：%s' % ext)


# --- パス入口（コマンド版はここを呼ぶ。中身は「バイトを読んで上の関数へ」）---
def read_store_header(path):
    """ 入力ファイルの見出し行（列名のリスト）を、形式に依らず返す（パス入力）。 """
    with open(path, 'rb') as f:
        data = f.read()
    return read_header_from_bytes(os.path.basename(path), data)


def read_store_rows(path):
    """ 入力ファイルの全データ行を、形式に依らず辞書のリストで返す（パス入力）。 """
    with open(path, 'rb') as f:
        data = f.read()
    return read_rows_from_bytes(os.path.basename(path), data)


# ============================================================================
# 引取候補店の文字列を作る
# ============================================================================
def _join_with_overflow(labels):
    """ 候補ラベルを最大 max_candidates 件まで ' / ' 連結し、あふれは「他N店」でまとめる。 """
    if not labels:
        return ''
    cap = CONFIG['max_candidates']
    shown = labels[:cap]
    if len(labels) > cap:
        shown.append('他%d店' % (len(labels) - cap))
    return ' / '.join(shown)


def build_candidates(by_key, key, source_store, supply_qty, exp_date, base_date):
    """ 出し手（デッド／期限切迫）1件について、引取候補店を①②③のtierに分けて
        (本命候補の文字列, 参考(過剰)候補の文字列) の2つを返す。
          本命 ＝ ①不足中 → ②使用中(適正) の順（最大 max_candidates 件、あふれは「他N店」）
          参考 ＝ ③使用中(過剰)。本命とは別の列に出すことで、本命候補と混ざらないようにする。 """
    if not key:
        return '（医薬品コード無し・突合対象外）', ''
    entries = by_key.get(key, [])
    short = []    # ① 不足中
    use_ok = []   # ② 使用中(適正)：使っていて、かつ受け取り側が過剰保有していない＝本命
    use_ref = []  # ③ 使用中(過剰)：使っているが受け取り側も過剰保有＝参考（下位・別枠）
    for e in entries:
        if e['store'] == source_store:
            continue
        if e['shortage']:
            short.append(e)
        elif e['usage_cnt6'] > 0:
            # 使用中：直近6ヶ月に出庫回数がある（＝その薬を実際に使っている店）。
            #   判定は符号に左右されない“出庫回数合計>0”で行う。
            #   さらに、受け取り側が既に過剰保有かどうかで「適正(本命)」と「過剰(参考)」に分ける。
            if e['holds_excess']:
                use_ref.append(e)
            else:
                use_ok.append(e)

    def pace(e):
        # 月あたり消化量。薬VANの出庫数はマイナス＝出庫なので、絶対値をとってから6で割る。
        return abs(e['usage_qty6']) / 6.0

    def detail_for(e):
        p = pace(e)
        if p <= 0:
            # 出庫回数はあるが出庫数の正味がゼロ（返品と相殺等）＝消化量が読めないケース
            return '消化ペース不明'
        # 消化目安＝渡す数量（在庫数）÷ 相手の月あたり消化量
        months = supply_qty / p
        if exp_date and base_date:
            remain = month_diff(base_date, exp_date)
            if months <= remain:
                return '約%sヶ月で消化可' % fmt_months(months)
            return '約%sヶ月・期限内消化不可' % fmt_months(months)
        return '約%sヶ月' % fmt_months(months)

    short.sort(key=lambda e: -pace(e))
    use_ok.sort(key=lambda e: -pace(e))
    use_ref.sort(key=lambda e: -pace(e))

    L = CONFIG['tier_labels']
    # 本命（①不足中 → ②使用中(適正)）
    main_ordered = [(L['short'], e) for e in short] + [(L['use_ok'], e) for e in use_ok]
    main_labels = ['%s（%s・%s）' % (e['store'], tag, detail_for(e)) for tag, e in main_ordered]
    main_str = _join_with_overflow(main_labels) or '該当なし'

    # 参考（③使用中(過剰)）… 必ず末尾・別枠。ラベルで本命と一目で区別できるようにする。
    ref_labels = ['%s（%s・%s）' % (e['store'], L['use_ref'], detail_for(e)) for e in use_ref]
    ref_str = _join_with_overflow(ref_labels)

    return main_str, ref_str


# ============================================================================
# 全店一括計算 compute_matching(stores)
#   入力 stores = [{'name': 店名, 'ym': 'YYYYMM'|None, 'base_date': date, 'rows': [dict,...]}, ...]
#        （rows は前処理 preprocess_keep 済みの行。列名→文字列の辞書。）
#   戻り値 dict … 4シート相当データ＋自己検算＋ログ用の集計。
#   ※ 現在 create_yuzu_list.main() にベタ書きされていた計算を、そのまま関数化したものです。
#     計算内容は一切変えていません（回帰で1円も変わらないことを保証する目的）。
# ============================================================================
def compute_matching(stores, excluded=None):
    """ excluded … 店が「融通に出さない」と外した品目の集合。
        {(店名, exclusion_key(row)), ...} の形。ここで出し手から完全に取り除くので、
        自店の②③だけでなく全店一覧・マトリクス・受け手の供給元・Excel・サマリ・
        自己検算のすべてから同時に消える（どこかに残って他店から問い合わせが来る事故を防ぐ）。 """
    excluded = set(excluded or ())

    def is_excluded(store_name, row):
        return bool(excluded) and (store_name, exclusion_key(row)) in excluded

    # --- 区分値の分布（ログ用・全店合算）---
    dist = {}
    for col in DIST_COLS:
        cnt = Counter()
        for s in stores:
            for row in s['rows']:
                cnt[g(row, col) or '（空欄）'] += 1
        dist[col] = cnt

    # --- (店,キー)ごとに集約 → by_key を作る ---
    store_key_info = {}
    nokey_rows = 0
    for s in stores:
        for row in s['rows']:
            key, _ = row_key(row)
            if key is None:
                nokey_rows += 1
                continue
            kk = (s['name'], key)
            # 出し手から外す2条件：店が「出さない」と外した品／少額（1,500円未満）の品
            ex = is_excluded(s['name'], row) or is_below_min_amount(row)
            sh = is_shortage(row)
            dead = is_dead(row) and not ex     # デッド（不動区分が非空）
            expf = is_expiry(row) and not ex   # 期限切迫（期限切迫区分が非空）
            supp = dead or expf                # 出し手（デッド or 期限切迫）
            he = holds_excess(row)   # 受け取り側の「過剰保有」（tier②適正/③参考の分かれ目）
            uq = usage_qty6(row)
            uc = usage_cnt6(row)
            over_qty = stock_qty(row)      # 在庫数（旧：過剰数）
            over_amt = stock_amount(row)   # 在庫金額＝在庫数×薬価（旧：過剰数金額）
            exp = parse_date(g(row, '有効期限'))
            name = g(row, '薬品名')
            zaiko = parse_num(g(row, '在庫数'))
            safe = parse_num(g(row, '安全在庫数'))
            e = store_key_info.get(kk)
            if e is None:
                store_key_info[kk] = {
                    'store': s['name'], 'shortage': sh,
                    'dead': dead, 'expiry': expf, 'supplier': supp,
                    'holds_excess': he,
                    'usage_qty6': uq, 'usage_cnt6': uc, 'over_qty': over_qty,
                    'over_amt': over_amt, 'exp': exp, 'name': name,
                    'zaiko': zaiko, 'safe': safe}
            else:
                e['shortage'] = e['shortage'] or sh
                e['dead'] = e['dead'] or dead
                e['expiry'] = e['expiry'] or expf
                e['supplier'] = e['supplier'] or supp
                e['holds_excess'] = e['holds_excess'] or he
                e['usage_qty6'] += uq
                e['usage_cnt6'] += uc
                e['over_qty'] += over_qty
                e['over_amt'] += over_amt
                if exp and (e['exp'] is None or exp < e['exp']):
                    e['exp'] = exp

    by_key = defaultdict(list)
    for (store, key), e in store_key_info.items():
        by_key[key].append(e)

    # --- 出し手（融通提案）を1行ずつ作る ---
    #   出し手＝デッド または 期限切迫（A案）。純粋な過剰は提案に載せない。
    #   各行に「種別」（デッド／期限切迫・排他デッド優先）を持たせる。
    proposal_rows = []
    legal_excluded_count = 0
    legal_excluded_amt = 0.0
    nokey_overstock = 0
    user_excluded_count = 0
    small_excluded_count = 0
    small_excluded_amt = 0.0
    for s in stores:
        for row in s['rows']:
            if not is_supplier(row):
                continue
            if is_excluded(s['name'], row):
                # 店が「この品は出さない」と外したもの
                user_excluded_count += 1
                continue
            if is_below_min_amount(row):
                # 少額（既定1,500円未満）で載せない品
                small_excluded_count += 1
                small_excluded_amt += stock_amount(row)
                continue
            if is_legal_excluded(row):
                legal_excluded_count += 1
                legal_excluded_amt += stock_amount(row)
                continue
            key, _ = row_key(row)
            if key is None:
                nokey_overstock += 1
            over_qty = stock_qty(row)      # 在庫数
            over_amt = stock_amount(row)   # 在庫金額
            exp = parse_date(g(row, '有効期限'))
            cand_main, cand_ref = build_candidates(
                by_key, key, s['name'], over_qty, exp, s['base_date'])
            remain = month_diff(s['base_date'], exp) if exp else None
            proposal_rows.append({
                '出し手店': s['name'], '種別': supplier_category(row),
                '薬品名': g(row, '薬品名'), '単位': g(row, '単位'),
                'メーカ名': g(row, 'メーカ名'), '在庫数': round(over_qty, 2),
                '在庫金額': round(over_amt, 2),
                '過剰在庫区分': g(row, '過剰在庫区分'), '不動区分': g(row, '不動区分'),
                '期限切迫区分': g(row, '期限切迫区分'), '有効期限': fmt_date(exp),
                'ロットNO': g(row, 'ロットNO'), '最終出庫日': g(row, '最終出庫日'),
                '区分': warn_labels(row),
                '引取候補店': cand_main, '参考:過剰だが使用中の店': cand_ref,
                '6ヶ月出庫回数': round(usage_cnt6(row), 2), '医薬品CD': key or '',
                '_ex_key': exclusion_key(row),
                '_expiry_flag': is_text_flag(g(row, '期限切迫区分')),
                '_remain': remain, '_amt_raw': over_amt})
    proposal_rows.sort(key=lambda r: -r['在庫金額'])

    # --- 受け手（不足品目一覧）を作る ---
    shortage_rows = []
    for s in stores:
        for row in s['rows']:
            if not is_shortage(row):
                continue
            key, _ = row_key(row)
            safe = parse_num(g(row, '安全在庫数'))
            zaiko = parse_num(g(row, '在庫数'))
            # 供給元は「デッドまたは期限切迫で持つ他店」に統一（A案。旧・過剰保有基準は廃止）
            holders = []
            if key:
                for e in by_key.get(key, []):
                    if e['store'] == s['name']:
                        continue
                    if e['supplier']:
                        cat = 'デッド' if e['dead'] else '期限切迫'
                        holders.append('%s（%s・数量%s・期限%s）'
                                       % (e['store'], cat, fmt_qty(e['over_qty']), fmt_ym(e['exp'])))
            shortage_rows.append({
                '店': s['name'], '薬品名': g(row, '薬品名'), '在庫数': round(zaiko, 2),
                '安全在庫数': round(safe, 2), '不足数': round(safe - zaiko, 2),
                '医薬品CD': key or '',
                'デッド/期限切迫で持つ他店': ' / '.join(holders) if holders else '（なし）'})
    shortage_rows.sort(key=lambda r: (r['店'], -r['不足数']))

    # --- 品目×店舗マトリクス ---
    #   行に出す品目＝出し手（デッド／期限切迫）がある品目。セルは：
    #     デX＝デッドで数量X ／ 限X＝期限切迫で数量X ／ 不＝安全在庫割れ ／ 使＝直近6ヶ月出庫あり
    over_keys = {}
    for (store, key), e in store_key_info.items():
        if e['supplier']:
            over_keys.setdefault(key, e['name'])
    matrix_keys = sorted(over_keys.keys(), key=lambda k: over_keys[k])
    store_names = [s['name'] for s in stores]
    matrix_rows = []
    for key in matrix_keys:
        cells = [over_keys[key], key]
        for sn in store_names:
            e = store_key_info.get((sn, key))
            if e is None:
                cells.append('')
            elif e['supplier']:
                # デッド優先（デッドかつ期限切迫なら「デ」）。純粋な期限切迫だけ「限」。
                mark = 'デ' if e['dead'] else '限'
                cells.append('%s%s' % (mark, fmt_qty(e['over_qty'])))
            elif e['shortage']:
                cells.append('不')
            elif e['usage_cnt6'] > 0:
                cells.append('使')
            else:
                cells.append('')
        matrix_rows.append(cells)

    # --- 店舗別サマリ ---
    #   出し手（デッド／期限切迫）を、排他・デッド優先で数える（proposal と同じ数え方）。
    #   金額は「在庫金額（在庫数×薬価）」を使う。麻薬・覚醒剤（法規制除外）は数えない。
    summary_rows = []
    for s in stores:
        dead_cnt = 0
        dead_amt = 0.0
        exp_cnt = 0
        exp_amt = 0.0
        short_cnt = 0
        for row in s['rows']:
            if (is_supplier(row) and not is_legal_excluded(row)
                    and not is_excluded(s['name'], row)
                    and not is_below_min_amount(row)):
                amt = stock_amount(row)
                if supplier_category(row) == 'デッド':
                    dead_cnt += 1
                    dead_amt += amt
                else:  # 期限切迫（デッドでない）
                    exp_cnt += 1
                    exp_amt += amt
            if is_shortage(row):
                short_cnt += 1
        summary_rows.append({
            '店': s['name'], 'デッド品目数': dead_cnt, 'デッド金額計': round(dead_amt, 2),
            '期限切迫品目数': exp_cnt, '期限切迫金額計': round(exp_amt, 2),
            '不足品目数': short_cnt})

    # --- 自己検算：出し手（デッド＋期限切迫・法規制除外後）の在庫金額合計 と 融通提案シートの合計 ---
    checkA = 0.0
    for s in stores:
        for row in s['rows']:
            if (is_supplier(row) and not is_legal_excluded(row)
                    and not is_excluded(s['name'], row)
                    and not is_below_min_amount(row)):
                checkA += stock_amount(row)
    checkB = sum(r['_amt_raw'] for r in proposal_rows)
    diff = abs(checkA - checkB)
    check_ok = diff < 0.01

    return {
        'proposal_rows': proposal_rows,
        'shortage_rows': shortage_rows,
        'store_names': store_names,
        'matrix_keys': matrix_keys,
        'matrix_rows': matrix_rows,
        'summary_rows': summary_rows,
        'dist': dist,
        'nokey_rows': nokey_rows,
        'nokey_overstock': nokey_overstock,
        'legal_excluded_count': legal_excluded_count,
        'legal_excluded_amt': legal_excluded_amt,
        'user_excluded_count': user_excluded_count,
        'small_excluded_count': small_excluded_count,
        'small_excluded_amt': small_excluded_amt,
        'min_supply_amount': CONFIG['min_supply_amount'],
        'checkA': checkA,
        'checkB': checkB,
        'diff': diff,
        'check_ok': check_ok,
    }


# ============================================================================
# Excel出力（コマンド版のExcelダウンロードと、アプリ版のExcelダウンロードで共用）
# ============================================================================
RED_FILL = PatternFill('solid', fgColor='FFC7CE')     # 期限切迫（赤系）
DEAD_FILL = PatternFill('solid', fgColor='FCE4D6')    # デッド（薄オレンジ系）※マトリクスの「デX」用
YELLOW_FILL = PatternFill('solid', fgColor='FFEB9C')  # 期限1年以内（黄系）
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')   # 見出し
NOTE_FILL = PatternFill('solid', fgColor='FFF2CC')     # 注記
REF_FILL = PatternFill('solid', fgColor='EDEDED')      # 参考(過剰だが使用中)＝本命と区別する薄グレー
REF_HEADER_FILL = PatternFill('solid', fgColor='D0CECE')  # 参考列の見出し（濃いめグレー）
HEADER_FONT = Font(bold=True)
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _auto_width(ws, headers, max_width=48):
    """ 列幅をざっくり自動調整 """
    for ci, h in enumerate(headers, start=1):
        col = get_column_letter(ci)
        maxlen = len(str(h))
        for cell in ws[col]:
            if cell.value is not None:
                # 全角をやや広めに数える
                s = str(cell.value)
                w = sum(2 if ord(ch) > 0x2E7F else 1 for ch in s)
                maxlen = max(maxlen, w)
        ws.column_dimensions[col].width = min(maxlen + 2, max_width)


def write_excel(path, base_ym_disp, csv_base_disp,
                proposal_rows, shortage_rows,
                matrix_store_names, matrix_rows,
                summary_rows):
    wb = Workbook()

    # ---------- シート1：融通提案 ----------
    ws = wb.active
    ws.title = '融通提案'
    note = ('基準年月：%s ／ CSV基準日：%s ／ '
            '※これは月初スナップショットであり、現在庫と異なる場合があります。'
            % (base_ym_disp, csv_base_disp))
    ws.cell(row=1, column=1, value=note).fill = NOTE_FILL
    ws.cell(row=1, column=1).font = Font(bold=True, color='7F6000')

    headers = ['出し手店', '種別', '薬品名', '単位', 'メーカ名', '在庫数', '在庫金額',
               '過剰在庫区分', '不動区分', '期限切迫区分', '有効期限', 'ロットNO',
               '最終出庫日', '区分', '引取候補店', '参考:過剰だが使用中の店',
               '6ヶ月出庫回数', '医薬品CD']
    # 「参考」列は本命（引取候補店）とはっきり区別するため、見出し・セルをグレーで塗る
    ref_col_idx = headers.index('参考:過剰だが使用中の店') + 1  # 1始まり
    hr = 2
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=hr, column=ci, value=h)
        c.fill = REF_HEADER_FILL if ci == ref_col_idx else HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    r = hr + 1
    for pr in proposal_rows:
        vals = [pr['出し手店'], pr['種別'], pr['薬品名'], pr['単位'], pr['メーカ名'], pr['在庫数'],
                pr['在庫金額'], pr['過剰在庫区分'], pr['不動区分'], pr['期限切迫区分'],
                pr['有効期限'], pr['ロットNO'], pr['最終出庫日'], pr['区分'],
                pr['引取候補店'], pr['参考:過剰だが使用中の店'],
                pr['6ヶ月出庫回数'], pr['医薬品CD']]
        for ci, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=ci, value=v)
            cell.border = BORDER
            # 在庫数（6列目）・在庫金額（7列目）は小数第2位まで表示する
            if ci in (6, 7):
                cell.number_format = '#,##0.00'
        # 塗り分け：期限切迫は赤、そうでなく1年以内は黄（ただし「参考」列は対象外）
        fill = None
        if pr['_expiry_flag']:
            fill = RED_FILL
        elif pr['_remain'] is not None and pr['_remain'] <= CONFIG['expiry_yellow_months']:
            fill = YELLOW_FILL
        if fill:
            for ci in range(1, len(headers) + 1):
                if ci == ref_col_idx:
                    continue
                ws.cell(row=r, column=ci).fill = fill
        # 「参考」列は行の色に関係なく常にグレー＝本命候補と一目で区別できるようにする
        ws.cell(row=r, column=ref_col_idx).fill = REF_FILL
        r += 1
    ws.auto_filter.ref = '%s%d:%s%d' % ('A', hr, get_column_letter(len(headers)), max(hr, r - 1))
    ws.freeze_panes = 'A%d' % (hr + 1)
    _auto_width(ws, headers)

    # ---------- シート2：不足品目一覧 ----------
    ws2 = wb.create_sheet('不足品目一覧')
    headers2 = ['店', '薬品名', '在庫数', '安全在庫数', '不足数', '医薬品CD', 'デッド/期限切迫で持つ他店']
    for ci, h in enumerate(headers2, start=1):
        c = ws2.cell(row=1, column=ci, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    r = 2
    for sr in shortage_rows:
        vals = [sr['店'], sr['薬品名'], sr['在庫数'], sr['安全在庫数'], sr['不足数'],
                sr['医薬品CD'], sr['デッド/期限切迫で持つ他店']]
        for ci, v in enumerate(vals, start=1):
            ws2.cell(row=r, column=ci, value=v).border = BORDER
        r += 1
    ws2.auto_filter.ref = 'A1:%s%d' % (get_column_letter(len(headers2)), max(1, r - 1))
    ws2.freeze_panes = 'A2'
    _auto_width(ws2, headers2)

    # ---------- シート3：品目×店舗マトリクス ----------
    ws3 = wb.create_sheet('品目×店舗マトリクス')
    headers3 = ['薬品名', '医薬品CD'] + matrix_store_names
    for ci, h in enumerate(headers3, start=1):
        c = ws3.cell(row=1, column=ci, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    r = 2
    for mrow in matrix_rows:
        for ci, v in enumerate(mrow, start=1):
            cell = ws3.cell(row=r, column=ci, value=v)
            cell.border = BORDER
            if ci >= 3 and isinstance(v, str):
                if v.startswith('限'):
                    cell.fill = RED_FILL       # 期限切迫（赤）
                elif v.startswith('デ'):
                    cell.fill = DEAD_FILL      # デッド（薄オレンジ）
                elif v == '不':
                    cell.fill = YELLOW_FILL
        r += 1
    # 凡例
    ws3.cell(row=r + 1, column=1,
             value='凡例：デX=デッド（不動・数量X） ／ 限X=期限切迫（数量X） ／ '
                   '不=安全在庫割れ ／ 使=直近6ヶ月に出庫実績あり ／ 空欄=在庫なし')
    ws3.freeze_panes = 'C2'
    _auto_width(ws3, headers3, max_width=14)
    ws3.column_dimensions['A'].width = 32

    # ---------- シート4：店舗別サマリ ----------
    ws4 = wb.create_sheet('店舗別サマリ')
    headers4 = ['店', 'デッド品目数', 'デッド金額計', '期限切迫品目数', '期限切迫金額計', '不足品目数']
    for ci, h in enumerate(headers4, start=1):
        c = ws4.cell(row=1, column=ci, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    r = 2
    for su in summary_rows:
        vals = [su['店'], su['デッド品目数'], su['デッド金額計'], su['期限切迫品目数'],
                su['期限切迫金額計'], su['不足品目数']]
        for ci, v in enumerate(vals, start=1):
            cell = ws4.cell(row=r, column=ci, value=v)
            cell.border = BORDER
            # デッド金額計（3列目）・期限切迫金額計（5列目）は小数第2位まで
            if ci in (3, 5):
                cell.number_format = '#,##0.00'
        r += 1
    ws4.freeze_panes = 'A2'
    _auto_width(ws4, headers4)

    wb.save(path)
