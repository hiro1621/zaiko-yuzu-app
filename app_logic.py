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
import uuid
import hashlib
import datetime

import jst
import yuzu_core
from yuzu_core import REQUIRED_COLS, KEEP_COLS


# ============================================================================
# 月次スケジュール（10日締切・11日予約開始）… 本間部長確定 2026-08-12
#   月を2つに割る：
#     1〜10日  … 集める期間（全店がアップし、自店のデッド確認・除外・出庫可能数の調整をする）。
#                新規予約はできない。
#     11日〜月末 … 予約する期間（全店そろった一覧に対して、同じスタートラインから予約する）。
#
#   ★締切日（10日）と解禁日（11日）を別々の定数で持つのは、あとで「締切15日・解禁20日」の
#     ように離したくなったとき、この1か所を直すだけで済むようにするため。
#   ★月の日数は見ない（31日でも28日でも「10日まで／11日から」で同じ）。
# ============================================================================
SCHEDULE = {
    'upload_deadline_day': 10,   # この日まではアップロード期間（この日を含む）
    'reserve_open_day': 11,      # この日から予約できる
}


def schedule_state(today, schedule=None):
    """ 今日が「集める期間」か「予約する期間」かを返す純関数。
        ★「今日」を引数で受け取る（画面側で jst.today() を1回だけ呼んで渡す）。
          こうすると、テストから任意の日付を流し込んで挙動を確かめられる（base_date と同じ方針）。

        引数：
          today    … 判定したい日付（datetime.date）。日（.day）だけを見る。
          schedule … 締切日・解禁日の設定（省略時は上の SCHEDULE を使う）。

        戻り値（辞書）：
          'phase'          … 'collecting'（集める期間・1〜10日）か 'open'（予約する期間・11日〜）
          'can_reserve'    … 予約してよいか（11日以降 True）
          'is_upload_late' … アップロードの締切（10日）を過ぎているか（11日以降 True）
          'deadline_day'   … 締切日（既定10）
          'open_day'       … 解禁日（既定11） """
    sch = schedule or SCHEDULE
    deadline = sch['upload_deadline_day']
    open_day = sch['reserve_open_day']
    day = today.day
    can_reserve = (day >= open_day)      # 解禁日を含む（11日から予約できる）
    is_upload_late = (day > deadline)    # 締切日は含まない（10日まではアップ期間）
    return {
        'phase': 'open' if can_reserve else 'collecting',
        'can_reserve': can_reserve,
        'is_upload_late': is_upload_late,
        'deadline_day': deadline,
        'open_day': open_day,
    }


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
        # 予約が入っている品は、候補店を並べる意味がもう無い（相手が決まっている）ので、
        #   『引取候補店』の欄をそのまま「◯◯が引取予定」に置き換える。
        #   出し手はこの欄を見て、どこへいつ送ればよいかが分かる。
        #   ★2026-08-01：受取時期（Nヶ月後）が「今すぐ」でないときだけ「（受取：Nヶ月後）」を足す。
        #     今すぐの予約は従来どおり「✅ ◯◯ が引取予定」のまま（文言を自然に保つ・本間部長裁量）。
        taker = r.get('_予約店', '')
        pickup = r.get('_予約受取', '')
        if taker:
            if pickup and pickup != '今すぐ':
                taker_label = '✅ %s が引取予定（受取：%s）' % (taker, pickup)
            else:
                taker_label = '✅ %s が引取予定' % taker
        else:
            taker_label = r['引取候補店（店名のみ）']
        out.append({
            '薬品名': r['薬品名'], '単位': r['単位'],
            # ★2026-08-10（第2弾・本間部長指示）：『在庫数』と『出せる数』の2列を
            #   『出庫可能数』の1列に統一した。数量を指定していない品では両者が必ず同じ値になり、
            #   2列並べても見づらいだけだったため。
            #   ・画面に出すのは『出庫可能数』＝実際に出す数（実効数量。proposal_rows の『在庫数』）。
            #   ・ただし全量（在庫数（全量））は number_input の max_value と『全量に戻す』に要るので、
            #     画面には出さない内部キー『在庫数』として必ず保持する（disp_cols に入れないので表には出ない）。
            '出庫可能数': r['在庫数'],                       # 画面に出す（実効数量＝旧『出せる数』の値）
            '在庫数': r.get('在庫数（全量）', r['在庫数']),   # 内部保持（全量）。画面には出さない
            # 出せる数の指定が入っているか（bool）。画面には出さないが将来の絞り込み等に使える。
            '_数量指定': r.get('_数量指定', False),
            '在庫金額': r['在庫金額'], '有効期限': r['有効期限'],
            '期限切迫区分': r['期限切迫区分'], '区分': r['区分'],
            '引取候補店': taker_label,
            # 期限切迫フラグ（有効期限まで5ヶ月以内か）。②の薄赤ハイライトに使う。
            #   ★判定は yuzu_core（compute_matching）で1度だけ行い、画面へはその結果を運ぶだけにして
            #     判定の出どころを一本化する（画面側で期限切迫区分の非空を見ると定義がズレるため）。
            '_expiry_flag': r.get('_expiry_flag', False),
            # 除外・出せる数の指定で使う内部キー（画面には出さない）
            '_key': r['_ex_key'],
            # 予約済みの品を誤って除外しないよう、店名を持たせておく（画面には出さない）
            '_予約店': taker,
            # 滞留（何ヶ月つづけて載っているか）。'滞留' が画面に出す文字、
            #   '_滞留区分' は色を決めるための区分（画面には出さない）。
            '滞留': r.get('滞留', ''), '_滞留区分': r.get('_滞留区分', 'new'),
            '_滞留月数': r.get('_滞留月数', 1),
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

      画面用の列に整えて返す（内部用の _tier_order・_remain・引取候補店 は返さない）：
        出し手店／薬品名／滞留／単位／出し手の出庫可能数／自店の在庫数／在庫金額／有効期限／消化目安／区分／医薬品CD
      ※2026-07-28 本間部長指示で『なぜ候補か』（不足中／使用中）列を削除した。
      ※2026-08-01（本間部長確定）に列を2つ整理した（①の改修）：
        ・既存の『在庫数』（＝出し手がいま持っている数）を『出し手の在庫数』に改名（取り違え防止）。
        ・その隣に『自店の在庫数』（＝自店がいまその品を何個持っているか）を新設。
      ※2026-08-10：『出し手の在庫数』→『出し手が出せる数』→『出し手の出庫可能数』に改名（値＝出し手の実効数量。
        第2弾で②③の『出庫可能数』と呼び名をそろえた）。
        ★これらの改名は【④の画面表示だけ】。Excel／Gシートの列名『在庫数』には一切触っていない。
    """
    rows = []
    for r in result.get('candidate_rows', []):
        # ★完全一致（==）のみ。in は使わない（和光／さと和光は別会社で、文字列として一方が他方を含むため）
        if r['引取候補店'] != store_name:
            continue
        # 予約済みの品は「予約した店の④」からも消す（下の build_view_reserved が
        #   『予約中の品』として別に出すので、同じ品が2か所に出ないようにする）。
        #   ほかの店の④からも当然消える＝早い者勝ちが画面にそのまま出る。
        if r.get('_予約店', ''):
            continue
        rows.append(r)
    # 並べ替え：まず滞留の長い品を上に、そのあと従来どおり 不足中→使用中／有効期限の近い順。
    #   ★滞留を最優先にしているのは、この一覧の目的が「引取候補店がいるのに動いていない品を
    #     引き取ってもらう」ことだから（2026-07-30 本間部長指摘）。
    #     長く残っている品ほど上に来るので、下までスクロールしなくても目に入る。
    #     ただし「先月予約されたのに残っている品」は月数に関係なく最上段に置く
    #     （受け渡しの取りこぼし＝いちばん先に確認してほしいもの）。
    rows.sort(key=lambda r: (0 if r.get('_滞留区分') == 'booked' else 1,
                             -int(r.get('_滞留月数', 1) or 1),
                             r['_tier_order'],
                             r['_remain'] is None,
                             r['_remain'] if r['_remain'] is not None else 0))
    return [_receive_row_to_view(r) for r in rows]


def _receive_row_to_view(r):
    """ ④（および別枠）の候補行 r を、画面表示用の辞書に整える共通処理。
        ★『在庫数』→『出し手の出庫可能数』の改名と『自店の在庫数』の追加はここで一元的に行う
          （④本体と別枠 build_view_receive_ref で同じ整え方を使い、二重管理しないため）。
        ★2026-08-10（第2弾）：呼び名を②③の『出庫可能数』にそろえ、
          『出し手が出せる数』→『出し手の出庫可能数』に改名した（値＝出し手の実効数量。
          出し手が「N錠だけ出す」と指定していればその数になる）。これは【④の画面表示だけ】の改名で、
          Excel／Gシートの列名『在庫数』には一切触っていない。 """
    return {
        '出し手店': r.get('出し手店', ''), '薬品名': r.get('薬品名', ''),
        '滞留': r.get('滞留', ''), '単位': r.get('単位', ''),
        # 既存『在庫数』を『出し手の出庫可能数』に改名（値＝出し手の実効数量）＋『自店の在庫数』を新設
        '出し手の出庫可能数': r.get('在庫数', ''), '自店の在庫数': r.get('自店の在庫数', ''),
        '在庫金額': r.get('在庫金額', ''), '有効期限': r.get('有効期限', ''),
        '消化目安': r.get('消化目安', ''), '区分': r.get('区分', ''),
        '医薬品CD': r.get('医薬品CD', ''),
        # ★予約を保存するためのキーを2つ持たせる（画面では隠す）。_出し手店・_key の組が予約1件を特定する。
        #   ここを落とすと予約が空キーで保存され、読み込み時に捨てられて「予約したのに消える」事故になる
        #   （2026-07 に build_view_receive で実際に起きた不具合。別枠でも必ず持たせる）。
        '_出し手店': r.get('出し手店', ''), '_key': r.get('_ex_key', ''),
        '_滞留区分': r.get('_滞留区分', 'new'), '_滞留月数': r.get('_滞留月数', 1),
        # 有効期限キャップの判定に使う（受取予定月をこの月より先に選ばせない）。画面では隠す。
        '_有効期限': r.get('有効期限', ''),
    }


def build_view_receive_ref(result, store_name):
    """
    ④の別枠『いまは在庫があるが、先になら引き取れる薬』（2026-08-01 追加・③の改修）。

      他店がデッド／期限切迫で出していて、自店も同じ品を使ってはいるが在庫を余らせている
      （＝tier③参考／use_ref）ため、本来の④（今すぐ引き取れる薬）からは意図的に外している品を、
      別枠で拾う。狙いは「今は在庫があるが、いまの在庫を使い切る先（1〜3ヶ月後）なら受け取れる」品を
      拾い、受取予定月つきで予約できるようにすること（本間部長の本命）。

      ★元データは result['candidate_rows_ref']（compute_matching が別リストで作った丸め無しの構造化データ）。
        『引取候補店』の文字列からは作らない（上位5店で丸められて候補漏れが起きるため）。
      ★突き合わせは完全一致（==）のみ（和光／さと和光は別会社なので in は使わない）。
      ★別枠の行にも _出し手店・_key を必ず持たせる（予約が空キーで消える過去の不具合を再発させない）。

      列は④本体とまったく同じ（出し手の出庫可能数＋自店の在庫数の2列構成）。
    """
    rows = []
    for r in result.get('candidate_rows_ref', []):
        if r['引取候補店'] != store_name:
            continue
        if r.get('_予約店', ''):
            continue
        rows.append(r)
    # 並べ替え：滞留の長い品を上に、そのあと有効期限の近い順・薬品名順（④本体と同じ考え方）。
    rows.sort(key=lambda r: (0 if r.get('_滞留区分') == 'booked' else 1,
                             -int(r.get('_滞留月数', 1) or 1),
                             r['_remain'] is None,
                             r['_remain'] if r['_remain'] is not None else 0,
                             r.get('薬品名', '')))
    return [_receive_row_to_view(r) for r in rows]


# ============================================================================
# 滞留（同じ品が何ヶ月つづけてリストに載っているか）
#   ・前月の結果（gsheet_store.read_prev_proposal の戻り）と、今月の計算結果を突き合わせる。
#   ・色とラベルの決め方は yuzu_core.stagnation_view に一本化してある（画面・Excel共通）。
# ============================================================================
def apply_stagnation(result, prev_map):
    """
    今月の計算結果に「滞留」の情報を書き込む（result をその場で書き換える）。

      prev_map … {(出し手店, 品目キー): {'滞留月数','引取候補店','予約'}}
                 gsheet_store.read_prev_proposal の戻り。空なら全部「今月から」になる。

    数え方：
      前月に同じ (出し手店, 品目キー) が載っていれば 前月の月数＋1、載っていなければ 1。
      ★品目キーは除外・予約と同じ _ex_key を使う（品目の呼び名を1つに保つ）。

    ★前月に一部の店がアップし忘れていると、その店の品は前月の結果に入っていないため
      滞留が1に戻る（実際より短く出る）。数え落とす側に倒しているのは、
      載っていないものを「滞留していた」と決めつけるより安全なため。

    戻り値：滞留区分ごとの件数（{'m2': 3, 'booked': 1, ...}）。画面の注記に使う。
    """
    prev_map = prev_map or {}
    counts = {}

    # --- 出し手の提案行に滞留を書き込む ---
    for pr in result.get('proposal_rows', []):
        prev = prev_map.get((pr['出し手店'], pr['_ex_key']))
        months = (prev.get('滞留月数', 1) + 1) if prev else 1
        was_reserved = bool(prev and str(prev.get('予約', '') or '').strip())
        # ★2026-08-04：いま有効な予約が生きているか（＝提案行の _予約店 が空でない）を渡す。
        #   持ち越し予約（受取予定月が先）は was_reserved も True になるので、now を先に見ないと
        #   約束どおり待っている品まで booked（紫）で警告され続ける誤検知になる。
        now_reserved = bool(str(pr.get('_予約店', '') or '').strip())
        level, label = yuzu_core.stagnation_view(
            months, pr.get('_候補あり', False), was_reserved, now_reserved)
        pr['_滞留月数'] = months
        pr['_先月予約'] = was_reserved
        pr['_滞留区分'] = level
        pr['滞留'] = label   # 画面・Excel の表示（reserved は '' ＝塗らない・文字なし）
        # ★_滞留persist … Gシート『融通提案』タブの滞留列へ書き出す“翌月チェーン用”の値。
        #   表示は空でも、月数を来月へ引き継げるよう「予約中Nヶ月目」を残す
        #   （read_prev_proposal は滞留列の数字を拾ってチェーンをつなぐため、空だとチェーンが切れて
        #     予約が明けたときに月数が実際より若く出てしまう）。表示（滞留列）と別値なのは、
        #     滞留列はもともと「翌月の判定材料」として設計されている列だから（build_results_payload 参照）。
        pr['_滞留persist'] = ('予約中%dヶ月目' % months) if level == 'reserved' else label
        counts[level] = counts.get(level, 0) + 1

    # --- ④受け手ビューの土台にも同じ値を配る ---
    #   candidate_rows は「1つの出し手行 × 候補店の数」だけ複製されているので、
    #   (出し手店, _ex_key) で引き当てて同じ滞留を持たせる。
    by_key = {(pr['出し手店'], pr['_ex_key']): pr for pr in result.get('proposal_rows', [])}
    for cr in result.get('candidate_rows', []):
        pr = by_key.get((cr['出し手店'], cr['_ex_key']))
        if not pr:
            continue
        for k in ('滞留', '_滞留月数', '_滞留区分', '_先月予約'):
            cr[k] = pr[k]

    result['stagnation_counts'] = counts
    return counts


def stagnation_summary(rows):
    """ 表示中の行リストから、滞留区分ごとの件数を数える（画面の注記用）。
        rows … build_view_a / build_view_expiry / build_view_receive の戻り。
        ★'new'（今月から）と 'reserved'（予約が生きていて塗らない）は凡例に出さないので数えない
          （数えると凡例ヘッダーだけ出て中身が空、という矛盾になる。凡例＝STAGNATION_LEGEND_ORDER とそろえる）。 """
    counts = {}
    for r in (rows or []):
        lv = r.get('_滞留区分', 'new')
        if lv and lv not in ('new', 'reserved'):
            counts[lv] = counts.get(lv, 0) + 1
    return counts


# ============================================================================
# 予約（受け手の店が「この品はうちが引き取ります」と押さえた印）
#   ・保管庫の _予約 タブ（辞書のリスト）と、compute_matching が欲しい形（辞書）を橋渡しする。
#   ・「品目まるごと」を押さえる方式。数量の指定はしない（本間部長判断 2026-07-28）。
# ============================================================================
# ----------------------------------------------------------------------------
# 予約の「有効期間」判定（A案・据え置き＋実在チェック／2026-08-01）
#   旧：対象年月＝当月のときだけ有効（月が変われば前月の予約は自動失効）。
#   新：受取予定月を持たせ、「当月 ≤ 受取予定月」なら有効＝最大3ヶ月まで予約を持ち越せる。
#
#   ★実在チェックは特別な処理を足さない。品が当月の一覧（proposal_rows）から消えれば、
#     その予約に紐づく表示行がそもそも存在しないので、予約は自動的に「何もしない状態」になる。
#     ここがA案の安全装置の肝：「実在しない品を押さえたまま」を、余計なコードなしで防ぐ。
# ----------------------------------------------------------------------------
def _reservation_effective_ym(r):
    """ その予約の『受取予定月』(YYYYMM)。空なら後方互換で『対象年月』（＝今すぐ）とみなす。
        ★古い行（受取予定月の列が無かった頃の予約）は受取予定月が空なので、
          対象年月と同じ月＝「今すぐ受け取る予定だった」と解釈する。 """
    return (r.get('受取予定月', '') or '').strip() or (r.get('対象年月', '') or '').strip()


def _reservation_active(r, ym):
    """ 予約が当月 ym に有効か（当月 ≤ 受取予定月）。ym が空なら常に有効扱い。
        受取予定月・対象年月がどちらも空の壊れた行は、無効として扱う（旧挙動＝当月不一致は捨てる、に合わせる）。 """
    if not ym:
        return True
    eff = _reservation_effective_ym(r)
    if not eff:
        return False
    return str(ym) <= eff


def _pickup_display(eff_ym, ym):
    """ 受取予定月(YYYYMM)を画面・帳票に出す読みやすい文字列にする。
        当月と同じ（またはどちらか空）なら『今すぐ』、先なら『YYYY/MM（Nヶ月後）』。 """
    if not eff_ym or not ym:
        return '今すぐ'
    off = yuzu_core.ym_offset(ym, eff_ym)
    if off <= 0:
        return '今すぐ'
    return '%s/%s（%dヶ月後）' % (eff_ym[:4], eff_ym[4:6], off)


def pickup_cap(expiry_str, ym):
    """ 受取予定月の上限オフセット（今から何ヶ月後まで選べるか、0〜3）を返す純関数。
        ★有効期限のある品は、その期限の月より先の受取予定月を選ばせない
          （期限切れ後に受け取っても使えず、その間ほかの店も引き取れなくなるため）。
        ★有効期限が空の品は3ヶ月後まで選べる（最大3ヶ月・本間部長確定）。
          expiry_str … 'YYYY/MM/DD' などの日付文字列（④の行の『有効期限』）。 """
    d = yuzu_core.parse_date(expiry_str)
    if d is None:
        return 3
    return max(0, min(3, yuzu_core.ym_offset(ym, d.strftime('%Y%m'))))


def reservation_map(reservations, ym):
    """ _予約 の行リストを compute_matching(reserved=...) に渡す形へ変換する。
          {(出し手店, 予約キー): {'店': 予約した店, '日時': 予約日時,
                                  '受取予定月': YYYYMM, '受取ラベル': '3ヶ月後'|'今すぐ'}, ...}

        ★2026-08-01：有効判定を「対象年月＝当月」→「当月 ≤ 受取予定月」に変更（最大3ヶ月持ち越し）。
          受取ラベル（当月からの月数）は当月 ym を知っているここで作って渡す
          （yuzu_core 側は月数計算をせず、このラベルを受け取って表示するだけにする）。
        同じ (出し手店, 予約キー) が複数あったら、先に書かれている行を優先する
        （＝先に予約した人が勝つ。保存側でも重複は止めているが、念のためここでも守る）。 """
    out = {}
    for r in (reservations or []):
        if not _reservation_active(r, ym):
            continue
        k = (r.get('出し手店', ''), r.get('予約キー', ''))
        if not k[0] or not k[1] or k in out:
            continue
        eff = _reservation_effective_ym(r)
        offset = yuzu_core.ym_offset(ym, eff) if (ym and eff) else 0
        out[k] = {'店': r.get('予約した店', ''), '日時': r.get('予約日時', ''),
                  '受取予定月': eff, '受取ラベル': yuzu_core.pickup_label(offset)}
    return out


def plan_reservations(current_rows, my_store, ym, picked):
    """
    「選ばれた品を予約する」ときの結果を計算する純関数（画面部品に依存しない＝テストできる）。

      current_rows … いま保管庫に入っている予約の行リスト（★保存直前に読み直したもの）
      picked       … 予約したい品のリスト。各要素は {'_出し手店','_key','薬品名'} を持つ
      戻り値 … {'keep': 保存すべき新しい行リスト,
                'added': 新しく予約できた件数,
                'already': すでに自店が予約済みだった件数,
                'conflicts': [{'薬品名','出し手店','予約した店'}, ...],
                'broken': キーが空で保存できなかった件数（通常は0。0以外なら不具合）}

    ★ほかの店がすでに押さえていた品は予約しない（conflicts に入れて呼び出し側が知らせる）。
      Googleシートは同時書き込みを止められないため、画面を開いてからボタンを押すまでの間に
      別の店が予約している可能性がある。読み直さずに書くと先の予約を上書きしてしまい、
      両方の店が「自分が予約できた」と思い込む事故になる。
    ★重複していない品はそのまま予約する（1件ぶつかっただけで全部やり直しにはしない）。
    ★2026-08-01：重複チェックを「当月」→「有効期間内の全予約（当月 ≤ 受取予定月）」に広げた。
      予約を持ち越せるようになったので、来月・再来月受取の予約とも品目がぶつかりうるため。
    ★picked の各要素は '_受取予定月'（解決済みのYYYYMM）を持つ。無ければ当月（＝今すぐ）とみなす。
    """
    taken = {}      # (出し手店, 予約キー) → いま予約している店
    for r in (current_rows or []):
        # ★有効期間内の予約だけを『取られている』とみなす（持ち越し予約ともぶつかる）
        if not _reservation_active(r, ym):
            continue
        taken[(r.get('出し手店', ''), r.get('予約キー', ''))] = r.get('予約した店', '')

    keep = list(current_rows or [])
    added, already, conflicts, broken = 0, 0, [], 0
    for d in (picked or []):
        k = (d.get('_出し手店', ''), d.get('_key', ''))
        # ★キーが空の行は保存しない。書いても読み込み時に捨てられ、
        #   「予約したのに消えている」という気づけない不具合になるため、ここで止めて数える。
        if not k[0] or not k[1]:
            broken += 1
            continue
        holder = taken.get(k)
        if holder and holder != my_store:
            conflicts.append({'薬品名': d.get('薬品名', ''), '出し手店': k[0], '予約した店': holder})
            continue
        if holder == my_store:
            already += 1        # すでに自店が予約済み＝二重に足さない
            continue
        # 受取予定月（呼び出し側で有効期限キャップまで反映して解決済み）。無ければ当月＝今すぐ。
        pickup_ym = (d.get('_受取予定月', '') or '').strip() or (ym or '')
        keep.append({'予約した店': my_store, '出し手店': k[0], '対象年月': ym or '',
                     '受取予定月': pickup_ym, '予約キー': k[1], '薬品名': d.get('薬品名', ''),
                     '予約日時': d.get('_now', '')})
        taken[k] = my_store
        added += 1

    # ★掃除（reservation_autoclean）がON のときだけ、失効予約（当月＞受取予定月）を
    #   この保存に相乗りして落とす。掃除専用の書き込みは作らない（Google API を増やさないため）。
    #   既定はOFF＝失効予約は _予約 タブに残す（画面内の有効判定だけで無害化する）。
    if yuzu_core.CONFIG.get('reservation_autoclean'):
        keep = [r for r in keep if _reservation_active(r, ym)]

    return {'keep': keep, 'added': added, 'already': already,
            'conflicts': conflicts, 'broken': broken}


def cancel_reservations(current_rows, my_store, picked, ym=None):
    """ 「選んだ予約を取り消す」あとの行リストを返す純関数。
        ★自店が予約した行だけを消す（他店の予約は絶対に触らない）。
        ★ym を渡し、掃除（reservation_autoclean）がON のときは、取消の保存に相乗りして
          失効予約（当月＞受取予定月）も一緒に落とす（掃除専用の書き込みは作らない）。既定OFF。 """
    drop = {(d.get('_出し手店', ''), d.get('_key', '')) for d in (picked or [])}
    keep = [r for r in (current_rows or [])
            if not (r.get('予約した店', '') == my_store
                    and (r.get('出し手店', ''), r.get('予約キー', '')) in drop)]
    if ym and yuzu_core.CONFIG.get('reservation_autoclean'):
        keep = [r for r in keep if _reservation_active(r, ym)]
    return keep


def build_view_reserved(result, store_name, reservations, ym):
    """
    「予約中の品」＝自店が押さえている品の一覧（取り消し用）。

    ★一覧の元は④の候補ではなく《予約の記録そのもの》にする。
      候補から作ると、予約したあとに条件が変わって候補から外れた品が黙って消え、
      予約したことを本人が確認できなくなるため。
      在庫数などの中身は、出し手の提案行（proposal_rows）から借りて表示する。
      出し手の一覧から品が消えていた場合は『状態』にそう書く（黙って落とさない）。

    画面用の列：出し手店／薬品名／単位／在庫数／在庫金額／有効期限／受取予定月／区分／医薬品CD／予約日時／状態
    ※2026-08-01：有効判定を「当月 ≤ 受取予定月」に変更し、『受取予定月』列を足した（②の改修）。
      内部用に _受取予定月（YYYYMMの生値）も持たせる（帳票の並べ替えに使う）。
    """
    # 出し手の提案行を (出し手店, 品目キー) で引けるようにしておく
    by_key = {(r['出し手店'], r['_ex_key']): r for r in result.get('proposal_rows', [])}
    out = []
    for rv in (reservations or []):
        if rv.get('予約した店', '') != store_name:
            continue
        # ★当月 ≤ 受取予定月 なら表示（旧：対象年月＝当月）。持ち越し中の予約もここに出る。
        if not _reservation_active(rv, ym):
            continue
        k = (rv.get('出し手店', ''), rv.get('予約キー', ''))
        eff = _reservation_effective_ym(rv)
        pickup_disp = _pickup_display(eff, ym)
        pr = by_key.get(k)
        if pr is not None:
            out.append({
                '出し手店': pr['出し手店'], '薬品名': pr['薬品名'], '単位': pr['単位'],
                '在庫数': pr['在庫数'], '在庫金額': pr['在庫金額'],
                '有効期限': pr['有効期限'], '受取予定月': pickup_disp,
                '区分': pr['区分'], '医薬品CD': pr['医薬品CD'],
                '予約日時': rv.get('予約日時', ''), '状態': '出し手が掲載中',
                '_出し手店': k[0], '_key': k[1], '_受取予定月': eff,
            })
        else:
            # 出し手が差し替え・除外などでこの品を出さなくなった＝予約だけが残っている状態
            out.append({
                '出し手店': k[0], '薬品名': rv.get('薬品名', ''), '単位': '',
                '在庫数': '', '在庫金額': '', '有効期限': '', '受取予定月': pickup_disp,
                '区分': '', '医薬品CD': '',
                '予約日時': rv.get('予約日時', ''), '状態': '出し手の一覧から外れました（要確認）',
                '_出し手店': k[0], '_key': k[1], '_受取予定月': eff,
            })
    # 出し手店ごとにまとめ、その中は薬品名順（電話をかける単位で並ぶようにする）
    out.sort(key=lambda r: (r['出し手店'], r['薬品名']))
    return out


# ============================================================================
# 出せる数（提供数量）… 出し手の店が「この品はN錠だけ出す」と決めた数量
#   2026-08-10 追加（第1弾）。保管庫の _提供数量 タブ（辞書のリスト）と、
#   compute_matching が欲しい形（辞書）を橋渡しする。品目キーは exclusion_key（除外・予約と同一）。
# ============================================================================
def supply_qty_map(rows):
    """ _提供数量 の行リストを compute_matching(supply_qty=...) に渡す形へ変換する。
          {(店名, 品目キー): 出せる数(float)}
        数値に直せない・0以下の行は無視する（＝指定なし＝全量扱い）。 """
    out = {}
    for r in (rows or []):
        store = (r.get('店名', '') or '').strip()
        key = (r.get('品目キー', '') or '').strip()
        if not store or not key:
            continue
        try:
            q = float(str(r.get('出せる数', '')).replace(',', '').strip())
        except (TypeError, ValueError):
            continue
        if q <= 0:
            continue
        out[(store, key)] = q
    return out


def apply_supply_cap(rows, stock_by_key):
    """ 月替わりの頭打ち。_提供数量 の各行の『出せる数』を、その品の当月在庫数まで自動で下げる。
          rows         … _提供数量 の行リスト（read_supply_qty と同じ形）
          stock_by_key … {(店名, 品目キー): 当月の在庫数(float)}
        ・在庫数を超える指定は在庫数まで下げる（「60錠出す」と決めた品の在庫が翌月20錠なら20錠＝実質全量）。
        ・当月の在庫が分からない品（stock_by_key に無い）はそのまま残す（勝手に消さない）。
        戻り値 … 頭打ち後の新しい行リスト（元は変更しない）。 """
    out = []
    for r in (rows or []):
        nr = dict(r)
        stock = stock_by_key.get((nr.get('店名', ''), nr.get('品目キー', '')))
        if stock is not None:
            try:
                qty = float(str(nr.get('出せる数', '')).replace(',', '').strip())
            except (TypeError, ValueError):
                qty = None
            if qty is not None and qty > stock:
                nr['出せる数'] = yuzu_core.fmt_qty(stock)
        out.append(nr)
    return out


def plan_supply_qty(current_rows, my_store, picked):
    """
    「選んだ品の『出せる数』を保存する」ときの結果を計算する純関数（画面部品に依存しない＝テスト可能）。

      current_rows … いま保管庫に入っている _提供数量 の行リスト（★保存直前に読み直したもの）
      my_store     … 自店名
      picked       … 保存したい品のリスト。各要素は
                     {'_key': 品目キー, '薬品名': 薬品名, '出せる数': float(出す数),
                      '在庫数': float(その品の在庫数＝全量), '_予約店': 予約店(あれば), '_now': 日時}
      戻り値 …
        {'keep':    保存すべき新しい行リスト（_提供数量 タブへ丸ごと書き戻す形）,
         'added':   新しく指定した件数,
         'updated': 指定を変更した件数,
         'removed': 指定を消した（＝全量に戻した）件数,
         'blocked': [{'薬品名','予約店'}, ...]（予約が入っていて数量を変えなかった品）}

    ルール：
      ・予約が入っている品（_予約店 が非空）は数量を変えない（blocked に入れて呼び出し側が知らせる）。
        黙って変えると、引き取るつもりの店の見込みが理由も分からず狂うため。
      ・『出せる数』が『在庫数（全量）』以上なら指定を削除する＝全量に戻す
        （在庫まるごと出すなら指定を持たない。月替わりで在庫が減っても自動で全量に追随できる）。
      ・それ以外は追加（無ければ）または更新（値が変わったら）。同じ値なら何もしない。
    """
    keep = list(current_rows or [])

    def _find(store, key):
        for i, r in enumerate(keep):
            if r.get('店名', '') == store and r.get('品目キー', '') == key:
                return i
        return -1

    added = updated = removed = 0
    blocked = []
    for d in (picked or []):
        key = (d.get('_key', '') or '').strip()
        if not key:
            continue
        # 予約が入っている品は触らない
        if (d.get('_予約店', '') or '').strip():
            blocked.append({'薬品名': d.get('薬品名', ''), '予約店': d.get('_予約店', '')})
            continue
        try:
            qty = float(str(d.get('出せる数', 0)).replace(',', '').strip() or 0)
        except (TypeError, ValueError):
            continue
        try:
            stock = float(str(d.get('在庫数', 0)).replace(',', '').strip() or 0)
        except (TypeError, ValueError):
            stock = 0.0
        i = _find(my_store, key)
        # 在庫数（全量）以上＝全量に戻す → 指定を消す
        if stock > 0 and round(qty, 2) >= round(stock, 2):
            if i >= 0:
                keep.pop(i)
                removed += 1
            continue
        if qty <= 0:
            # 0以下は指定しない（画面側で 0.01 以上に制限しているが、念のためここでも無視）
            continue
        new_row = {'店名': my_store, '品目キー': key, '薬品名': d.get('薬品名', ''),
                   '出せる数': yuzu_core.fmt_qty(qty), '更新日時': d.get('_now', '')}
        if i >= 0:
            old = (keep[i].get('出せる数', '') or '').strip()
            if old != new_row['出せる数']:
                keep[i] = new_row
                updated += 1
        else:
            keep.append(new_row)
            added += 1
    return {'keep': keep, 'added': added, 'updated': updated,
            'removed': removed, 'blocked': blocked}


# ============================================================================
# 引取依頼書（帳票）＝受け手（引き取る側）が予約した品を、出し手店ごとにまとめる
#   2026-08-01 追加（④の改修）。デスクネッツ貼付・FAX・電話しながらの参照用。
#   ・build_view_reserved（《予約の記録そのもの》から作る既存関数）を土台にする。
#   ・build_view_reserved 自体は帳票のために作り替えない（同じ値を2箇所で作らない原則）。
#     ロットNO は build_view_reserved が返さないので、出し手の提案行から引き当てて借りる。
# ============================================================================
def build_pickup_request(result, my_store, reservations, ym):
    """
    引取依頼書のデータを組み立てる純関数（Excel化は yuzu_core.write_pickup_request_excel が担当）。

      戻り値：
        {'my_store': 自店名, 'ym': 'YYYYMM',
         'sheets': [{'出し手店': 店名,
                     'rows': [{'薬品名','単位','数量','有効期限','ロットNO','医薬品CD',
                               '受取予定月','区分','状態','_期限強調','_受取予定月'}, ...]}, ...]}

      ・対象＝自店が予約している品のうち有効期間内（受取予定月をまだ過ぎていない）もの
        ＝build_view_reserved の戻りそのもの（有効判定はあちらに一本化済み）。
      ・数量＝在庫まるごと（出し手の在庫数＝提案行の在庫数）を初期値。手で書き換える前提。
      ・ロットNO＝build_view_reserved は返さないので、出し手の提案行から (出し手店, 品目キー) で引き当てる。
      ・『出し手の一覧から外れました（要確認）』の品も黙って落とさず載せる
        （数量・ロットNO・有効期限は空欄、状態欄に明記）。このツールの一貫原則。
      ・受取予定月は1枚のシートの中に列で出し、近い順（同月内は薬品名順）に並べる。
    """
    reserved = build_view_reserved(result, my_store, reservations, ym)
    # ロットNO を出し手の提案行から借りる（同じ値を2箇所で作らない）。掲載が外れた品はキーが無い＝空。
    lot_by_key = {(r['出し手店'], r['_ex_key']): r.get('ロットNO', '')
                  for r in result.get('proposal_rows', [])}
    # 有効期限が近い品を帳票で太字強調するための基準日（当月1日）。ym から作る。
    base_date = None
    if ym and len(str(ym)) >= 6:
        try:
            base_date = datetime.date(int(str(ym)[:4]), int(str(ym)[4:6]), 1)
        except ValueError:
            base_date = None

    by_store = {}
    for rv in reserved:
        listed = (rv.get('状態') == '出し手が掲載中')
        lot = lot_by_key.get((rv['_出し手店'], rv['_key']), '') if listed else ''
        # 期限強調：有効期限があり、当月から expiry_yellow_months（既定12ヶ月）以内なら太字にする
        exp_d = yuzu_core.parse_date(rv.get('有効期限', ''))
        near = bool(exp_d and base_date
                    and yuzu_core.month_diff(base_date, exp_d) <= yuzu_core.CONFIG['expiry_yellow_months'])
        detail = {
            '薬品名': rv.get('薬品名', ''),
            '単位': rv.get('単位', ''),
            '数量': rv.get('在庫数', ''),     # ＝出し手の在庫数（在庫まるごと）を初期値
            '有効期限': rv.get('有効期限', ''),
            'ロットNO': lot,
            '医薬品CD': rv.get('医薬品CD', ''),
            '受取予定月': rv.get('受取予定月', ''),
            '区分': rv.get('区分', ''),
            '状態': rv.get('状態', ''),
            '_期限強調': near,
            '_受取予定月': rv.get('_受取予定月', ''),   # 並べ替え用の生値（YYYYMM）
        }
        by_store.setdefault(rv['出し手店'], []).append(detail)

    sheets = []
    for store in sorted(by_store.keys()):
        rows = by_store[store]
        # 受取予定月の近い順（空＝末尾）→ 同月内は薬品名順
        rows.sort(key=lambda d: (d.get('_受取予定月', '') or '999999', d.get('薬品名', '')))
        sheets.append({'出し手店': store, 'rows': rows})
    return {'my_store': my_store, 'ym': ym or '', 'sheets': sheets}


def pickup_request_bytes(result, my_store, reservations, ym):
    """ 引取依頼書を Excel にして bytes で返す（画面のダウンロードボタン用）。 """
    data = build_pickup_request(result, my_store, reservations, ym)
    bio = io.BytesIO()
    yuzu_core.write_pickup_request_excel(bio, data)
    bio.seek(0)
    return bio.getvalue()


# ============================================================================
# 店舗間のやり取り（掲示板）… 第2弾・本間部長確定 2026-08-10
#   ・スレッドの単位は「相手店ごとに1本」。予約1件ごとではない
#     （東立石⇄みずほ台 は、みずほ台のデッド品を4件予約していても会話の場は1つ）。
#   ・スレッドのキーは店名2つを sorted() で固定（店A／店B）。どちらが出し手／受け手かでは分けない。
#   ・個別の薬の話は、投稿ごとに任意で『薬品名』を添えて解決する。
#   ・ここは Streamlit に依存しない純関数だけ（品質管理部がテストしやすいように）。
# ============================================================================
def thread_pair(a, b):
    """ 店名2つを sorted() して (店A, 店B) の順に固定して返す。
        同じ相手との会話を必ず1本のスレッドにまとめるためのキー。前後の空白は落とす。 """
    pair = sorted([str(a or '').strip(), str(b or '').strip()])
    return (pair[0], pair[1])


def _last_read_at(my_store, store_a, store_b, msg_reads):
    """ 自店(my_store)が、(店A,店B)のスレッドを最後に確認した日時を返す。無ければ空文字。 """
    a, b = thread_pair(store_a, store_b)
    my = str(my_store or '').strip()
    for r in (msg_reads or []):
        if (str(r.get('店名', '') or '').strip() == my
                and str(r.get('店A', '') or '').strip() == a
                and str(r.get('店B', '') or '').strip() == b):
            return str(r.get('最終確認日時', '') or '')
    return ''


def unread_count(my_store, thread, msg_reads):
    """ スレッドの未読件数を返す純関数。
        未読＝その投稿が『自分以外の店』のもので、かつ『自分の既読日時より新しい』もの。
          my_store  … 自店名
          thread    … build_threads が返すスレッド辞書（'messages' に投稿のリストを持つ）
          msg_reads … _やり取り既読 の行リスト（read_msg_reads と同じ形）
        ★日時は 'YYYY/MM/DD HH:MM:SS'（ゼロ詰め固定幅）なので、文字列の大小比較でそのまま時系列になる。 """
    my = str(my_store or '').strip()
    last_read = _last_read_at(my, thread.get('店A', ''), thread.get('店B', ''), msg_reads)
    n = 0
    for m in thread.get('messages', []):
        if str(m.get('投稿店', '') or '').strip() == my:
            continue                                  # 自分の投稿は未読にならない
        if str(m.get('投稿日時', '') or '') > last_read:
            n += 1
    return n


def mark_thread_read(msg_reads, my_store, store_a, store_b, now):
    """ (店名=my_store, 店A, 店B) の既読日時を now に更新した新しい行リストを返す純関数。
        既存の自分の行があれば更新、無ければ追加する。ほかの店・ほかのスレッドの行は一切触らない。 """
    a, b = thread_pair(store_a, store_b)
    my = str(my_store or '').strip()
    out = []
    replaced = False
    for r in (msg_reads or []):
        if (str(r.get('店名', '') or '').strip() == my
                and str(r.get('店A', '') or '').strip() == a
                and str(r.get('店B', '') or '').strip() == b):
            out.append({'店名': my, '店A': a, '店B': b, '最終確認日時': now})
            replaced = True
        else:
            out.append(dict(r))
    if not replaced:
        out.append({'店名': my, '店A': a, '店B': b, '最終確認日時': now})
    return out


# ============================================================================
# 全店へのお知らせ板（全店板）の未読・既読を数える純関数（第3弾）
#   ・1対地の“放送”なので相手店（店A/店B）で分けない。1つの板を全店で共有する。
#   ・未読の考え方は 1対1 の unread_count と同じ＝「自分以外の店の投稿で、
#     自分がこの板を最後に見た日時より後のもの」を数える。
#   ・Streamlit に依存しない純関数だけ（品質管理部がテストしやすいように）。
# ============================================================================
def _allboard_last_read_at(my_store, reads):
    """ 自店(my_store)が、全店板を最後に確認した日時を返す。無ければ空文字。 """
    my = str(my_store or '').strip()
    for r in (reads or []):
        if str(r.get('店名', '') or '').strip() == my:
            return str(r.get('最終確認日時', '') or '')
    return ''


def allboard_unread_count(my_store, posts, reads):
    """ 全店板の未読件数を返す純関数。
        未読＝その行が『自分以外の店』のもので、かつ『自分の既読日時より新しい』もの。
          my_store … 自店名
          posts    … 全店板の投稿リスト（read_allboard と同じ形＝'投稿日時','投稿店','本文','投稿ID','親ID','種別'）
          reads    … 全店板既読の行リスト（read_allboard_reads と同じ形＝'店名','最終確認日時'）
        ★第4弾（2026-08-22）：未読は「投稿＋返信」で数える（返信も他店の新着として拾う）。
          種別「状態」（解決済み／未解決の切替）は“新着”に数えない＝バッジのために板を開かせる
          必要がないため。旧3列の投稿は種別が空なので、これまでどおり投稿として数える。
        ★日時は 'YYYY/MM/DD HH:MM:SS'（ゼロ詰め固定幅）なので、文字列の大小比較でそのまま時系列になる。 """
    my = str(my_store or '').strip()
    last_read = _allboard_last_read_at(my, reads)
    n = 0
    for m in (posts or []):
        if str(m.get('種別', '') or '').strip() == '状態':
            continue                                  # 状態変更（解決済み等）は新着に数えない
        if str(m.get('投稿店', '') or '').strip() == my:
            continue                                  # 自分の投稿・返信は未読にならない
        if str(m.get('投稿日時', '') or '') > last_read:
            n += 1
    return n


def allboard_mark_read(reads, my_store, now):
    """ 自店(my_store)の全店板の既読日時を now に更新した新しい行リストを返す純関数。
        既存の自分の行があれば更新、無ければ追加する。ほかの店の行は一切触らない。 """
    my = str(my_store or '').strip()
    out = []
    replaced = False
    for r in (reads or []):
        if str(r.get('店名', '') or '').strip() == my:
            out.append({'店名': my, '最終確認日時': now})
            replaced = True
        else:
            out.append(dict(r))
    if not replaced:
        out.append({'店名': my, '最終確認日時': now})
    return out


# ============================================================================
# 全店へのお知らせ板：ツリー返信・解決済みを組み立てる純関数（第4弾・2026-08-22）
#   ・板の中で「どの投稿への返信か」を親子で表す。稼働中の1対1やり取りには一切触らない。
#   ・返信への返信（多段）も許す。解決済みは投稿した店だけが立てられる（決定6）。
#   ・Streamlit に依存しない純関数だけ（品質管理部がテストしやすいように）。
# ============================================================================
def new_allboard_id():
    """ 全店板の投稿・返信・状態行に振る一意な符号（UUID の先頭12桁）。
        日時・店名・行番号に依存しないので、同じ店が同じ秒に投稿しても、2店が同時でも重複しない。
        ★これは非決定（呼ぶたび違う値）。テストでは『重複しないこと・桁数』だけを確かめる。 """
    return uuid.uuid4().hex[:12]


def _legacy_allboard_id(row):
    """ 投稿IDが空の“古い投稿行”（UUID採番前の3列時代）に、毎回同じ値で振る仮の符号。
        『投稿日時＋投稿店＋本文』から決める＝古い投稿は追記のみで中身が変わらないので毎回同じ値に
        なり、返信を保存しても後で同じ親に正しく結び付く（legacy-<短いハッシュ>）。
        ★本番シートは手で書き換えない。読み手側でこの符号を計算して割り当てるだけ（決定どおり）。 """
    base = '%s|%s|%s' % (
        str(row.get('投稿日時', '') or ''),
        str(row.get('投稿店', '') or ''),
        str(row.get('本文', '') or ''))
    return 'legacy-' + hashlib.md5(base.encode('utf-8')).hexdigest()[:12]


def build_allboard_tree(rows):
    """ 全店板の行リスト（read_allboard と同じ形）から、画面に出す“表示順リスト”を作る純関数。

      入力  rows … [{'投稿日時','投稿店','本文','投稿ID','親ID','種別'}, ...]
      出力  [ {'投稿ID','投稿日時','投稿店','本文','親ID','種別'（元の値）,
                '深さ'（0＝大元・返信ごとに+1）,
                'ルートID'（属する大元の投稿ID）,
                'ルート解決済み'（その大元が解決済みか＝バッジ＋折りたたみの判定に使う）}, ... ]
            ★状態行（種別「状態」）は“解決済みかどうか”に消費し、表示行としては返さない。

      手順：
        (1) 投稿IDが空の古い行に legacy-<hash> を補う（毎回同じ値）。
        (2) 投稿・返信ノードと状態行に仕分ける（種別「状態」＝状態行、それ以外＝表示ノード）。
        (3) 各投稿の“持ち主の店”を把握する（決定6：状態を立てられるのは投稿した店だけ）。
        (4) 状態行から、各対象投稿の『最新の有効な状態』を確定する。
            有効＝対象投稿の持ち主の店が書いた状態行だけ（別店の状態行は計算側でも無効にする＝
            画面でボタンを出さないのと合わせて“二重の歯止め”）。最新＝投稿日時の文字列比較で最大。
        (5) 親IDで親子を組み立てる。★親が見つからない返信は大元扱いにして必ず画面に出す
            （黙って消すのが最悪）。★自分自身が親（自己参照）も大元扱い。
        (6) 深さ優先で表示順に並べる。★循環参照（A→B→A 等）でも無限ループしないよう visited で止める。
        (7) ★安全網：どのルートからも辿れなかったノード（循環だけで閉じた塊）を大元扱いで拾い残さない。
    """
    rows = rows or []

    # (1) 投稿IDが空の行に仮ID（legacy-<hash>）を補う。親ID・種別は前後空白を落として正規化。
    prepared = []
    for r in rows:
        d = dict(r)
        pid = str(d.get('投稿ID', '') or '').strip()
        d['投稿ID'] = pid if pid else _legacy_allboard_id(d)
        d['親ID'] = str(d.get('親ID', '') or '').strip()
        d['種別'] = str(d.get('種別', '') or '').strip()
        prepared.append(d)

    # (2) 状態行と表示ノード（投稿・返信）に仕分ける。
    posts = []      # 表示ノード（投稿・返信）
    states = []     # 状態行（解決済み／未解決）
    for d in prepared:
        (states if d['種別'] == '状態' else posts).append(d)

    # 投稿ID → ノード（同じIDが複数あれば最初のを採用＝表示は1件に保つ）
    by_id = {}
    for d in posts:
        by_id.setdefault(d['投稿ID'], d)

    # (3) 各投稿の“持ち主の店”
    owner_of = {pid: str(d.get('投稿店', '') or '').strip() for pid, d in by_id.items()}

    # (4) 状態行 → 各対象投稿の『最新の有効な状態』
    latest_state = {}   # 対象投稿ID → (最新日時, 状態値)
    for s in states:
        target = str(s.get('親ID', '') or '').strip()
        if not target or target not in owner_of:
            continue
        writer = str(s.get('投稿店', '') or '').strip()
        if writer != owner_of[target]:
            continue    # ★対象投稿の持ち主以外が立てた状態は無効（決定6の“計算側の歯止め”）
        ts = str(s.get('投稿日時', '') or '')
        if (target not in latest_state) or (ts >= latest_state[target][0]):
            latest_state[target] = (ts, str(s.get('本文', '') or '').strip())
    resolved_of = {pid: (v[1] == '解決済み') for pid, v in latest_state.items()}

    # (5) 親子を組み立てる。返信の親が見つからなければ大元扱い（★黙って落とさない）。
    children = {}    # 親ID → [子ノード...]
    roots = []
    for d in posts:
        parent = d['親ID']
        if parent and parent in by_id and parent != d['投稿ID']:
            children.setdefault(parent, []).append(d)
        else:
            roots.append(d)   # 親が空／親が見つからない／自己参照 → 大元として出す

    # 兄弟は投稿日時の昇順（古い→新しい）。同時刻でも安定するよう投稿IDを第2キーにする。
    def _sort_key(x):
        return (str(x.get('投稿日時', '') or ''), str(x.get('投稿ID', '') or ''))
    roots.sort(key=_sort_key)
    for k in children:
        children[k].sort(key=_sort_key)

    # (6) 深さ優先で表示順リストを作る。★循環・重複は visited で1回だけに抑える。
    out = []
    visited = set()

    def _emit(node, depth, root_id, root_resolved):
        nid = node['投稿ID']
        if nid in visited:
            return                    # 循環参照・重複IDはここで打ち切る（無限ループ防止）
        visited.add(nid)
        out.append({
            '投稿ID': nid,
            '投稿日時': node.get('投稿日時', ''),
            '投稿店': node.get('投稿店', ''),
            '本文': node.get('本文', ''),
            '親ID': node.get('親ID', ''),
            '種別': node.get('種別', ''),
            '深さ': depth,
            'ルートID': root_id,
            'ルート解決済み': root_resolved,
        })
        for c in children.get(nid, []):
            _emit(c, depth + 1, root_id, root_resolved)

    for r in roots:
        rid = r['投稿ID']
        _emit(r, 0, rid, resolved_of.get(rid, False))

    # (7) 安全網：循環だけで閉じてどのルートからも辿れなかったノードを大元扱いで必ず出す。
    for d in posts:
        if d['投稿ID'] not in visited:
            rid = d['投稿ID']
            _emit(d, 0, rid, resolved_of.get(rid, False))

    return out


def build_threads(my_store, messages, reservations, qty_by_key=None):
    """
    自店(my_store)が関わるスレッド（相手店ごとに1本）の一覧を返す純関数。

      messages     … _やり取り の行リスト（read_messages と同じ形）
      reservations … _予約 の行リスト（read_reservations と同じ形）
      qty_by_key   … {(出し手店, 予約キー): '20錠'} の早見表（2026-08-10 追加・②の改修）。
                     渡すと『予約中の品』の各薬品名に数量＋単位を添える（例『ワイドシリン… 20錠』）。
                     ★既定 None なら今までどおり薬品名だけ（純関数のテストが壊れないように）。
                     ★引けなかった品（出し手が除外した／当月の提案から消えた）は薬品名だけにする
                       ＝勝手に「0錠」等と表示して誤解させない（本間部長指示）。
                     予約キーは reservations 行の『予約キー』＝提案行の yuzu_core.exclusion_key(row)
                     （除外・予約と共通の規則）。

    スレッドが一覧に出る条件（本間部長確定）：
      「自店が関わる予約が1件以上ある相手店」または「過去に1件でも投稿がある相手店」。
      → 予約が取り消されても、月をまたいでも、投稿さえ残っていれば会話は消えない
        （言った言わないの元になるため、会話ログは残す）。

    各要素（スレッド辞書）が持つもの：
      '相手店名'／'予約中の品'（薬品名＋数量のリスト・重複除去）／
      '予約中の品名'（数量を付けない“素の薬品名”のリスト。★『どの薬の話』の投稿タグに使うので
                      数量を混ぜない＝ログに数量が残らないようにするため）／
      '最終投稿日時'／'最終投稿店'／'件数'（投稿数）／
      '店A'・'店B'（sorted 済みのスレッドキー）／'messages'（時系列の投稿リスト）
    """
    my = str(my_store or '').strip()
    messages = messages or []
    reservations = reservations or []
    qty_by_key = qty_by_key or {}

    # 相手店 → 予約中の品・投稿 をためる箱
    others = {}   # 相手店名 → {'reserved_names':[],'reserved_clean':[],'reserved_seen':set(),'messages':[]}

    def _box(other):
        o = str(other or '').strip()
        if not o or o == my:
            return None
        if o not in others:
            others[o] = {'reserved_names': [], 'reserved_clean': [],
                         'reserved_seen': set(), 'messages': []}
        return o

    # (1) 予約から「相手店」と「予約中の品（薬品名）」を集める（出し手／受け手の両方向）
    for r in reservations:
        booker = str(r.get('予約した店', '') or '').strip()      # 予約した店（受け手）
        supplier = str(r.get('出し手店', '') or '').strip()      # 出し手店
        if my == booker:
            other = supplier
        elif my == supplier:
            other = booker
        else:
            continue
        o = _box(other)
        if o is None:
            continue
        name = str(r.get('薬品名', '') or '').strip()
        if name and name not in others[o]['reserved_seen']:
            others[o]['reserved_seen'].add(name)
            others[o]['reserved_clean'].append(name)        # 素の薬品名（投稿タグ用）
            # 数量が引ければ『薬品名 20錠』、引けなければ薬品名だけ（0錠と誤解させない）
            qkey = (supplier, str(r.get('予約キー', '') or '').strip())
            qty = qty_by_key.get(qkey)
            others[o]['reserved_names'].append(
                ('%s %s' % (name, qty)) if qty else name)

    # (2) 投稿から「相手店」と「その相手との投稿」を集める（店A/店Bは sorted 済み）
    for m in messages:
        a = str(m.get('店A', '') or '').strip()
        b = str(m.get('店B', '') or '').strip()
        if my != a and my != b:
            continue
        other = b if my == a else a
        o = _box(other)
        if o is None:
            continue
        others[o]['messages'].append(m)

    threads = []
    for other, info in others.items():
        # 投稿は時系列（投稿日時の昇順）に並べる。日時が同じ・空でも安定するよう元の順を保つ
        msgs = sorted(info['messages'], key=lambda x: str(x.get('投稿日時', '') or ''))
        a, b = thread_pair(my, other)
        last_at = msgs[-1].get('投稿日時', '') if msgs else ''
        last_by = msgs[-1].get('投稿店', '') if msgs else ''
        threads.append({
            '相手店名': other,
            '予約中の品': list(info['reserved_names']),      # 数量つき（画面『予約中：』用）
            '予約中の品名': list(info['reserved_clean']),    # 素の薬品名（投稿タグ『どの薬の話』用）
            '最終投稿日時': last_at,
            '最終投稿店': last_by,
            '件数': len(msgs),
            '店A': a, '店B': b,
            'messages': msgs,
        })
    # 並び：最終投稿が新しい順（投稿の無いスレッドは末尾）→ 相手店名の五十音
    threads.sort(key=lambda t: t['相手店名'])
    threads.sort(key=lambda t: t['最終投稿日時'], reverse=True)
    return threads


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


def results_signature(index, latest, exclusions, reservations, supply_rows, csv_base_disp=''):
    """
    ②案1（2026-08-12）用：照合（compute_matching）とExcel（excel_bytes）を作り直すかどうかを
    判断するための『入力の署名』を作る純関数（st に依存しない＝単体テスト可能）。

    材料が1つでも変われば署名が変わる＝古い結果を見せない。混ぜるのは次のぜんぶ：
      ・index（各店のアップ記録：対象年月・アップ日時・行数・様式・ファイル名）… ★下記の理由で
        「行の内容そのもの」の代わりに使う『店データの指紋』。
      ・latest（対象年月）… 月替わりで必ず変わる。
      ・exclusions（除外リスト）／reservations（予約リスト）／supply_rows（出庫可能数の生の指定リスト）。
      ・csv_base_disp（Excelに刷る「本日」の日付）… 日付が変わればExcelを刷り直す。

    ★なぜ stores の『行の内容』ではなく index（アップ記録）を指紋に使うのか：
      在庫本体（raw_<店>）は「アップロード」以外では絶対に変わらず、アップロードのたびに index の
      アップ日時が必ず書き換わる（save_store_upload）。よって index が1文字も変わっていなければ
      行の内容も変わっていない、と言い切れる（load_stores_cached が読み直しの要否を判断するのと同じ理屈）。
      逆に、行を毎回ぜんぶ突き合わせて指紋を取ると、店数に比例して重くなり（14店で0.15秒/回）、
      「チェック1つで重い」を直すという今回の目的そのものを損なう。だからアップ記録を指紋にする。
    """
    def _norm(rows):
        # 辞書の並び順ゆらぎに影響されないよう、(キー,値) をソートしてから連結する。
        return [tuple(sorted((str(k), str(v)) for k, v in (r or {}).items())) for r in (rows or [])]
    material = (
        # 各店のアップ記録（店名でソート・各記録の中身もソートして並び順非依存にする）
        sorted((str(k), tuple(sorted((str(kk), str(vv)) for kk, vv in (v or {}).items())))
               for k, v in (index or {}).items()),
        str(latest),
        _norm(exclusions), _norm(reservations), _norm(supply_rows),
        str(csv_base_disp),
    )
    return hashlib.md5(repr(material).encode('utf-8', 'replace')).hexdigest()


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
        state.setdefault('reservations', [])  # 受け手の店が「うちが引き取る」と押さえた品目
        state.setdefault('supply_qty', [])   # 出し手の店が「この品はN錠だけ出す」と決めた数量
        state.setdefault('messages', [])     # 店舗間のやり取り（掲示板）＝1件1投稿
        state.setdefault('msg_reads', [])    # どの店がどのスレッドをいつまで読んだか（未読判定用）
        state.setdefault('allboard', [])     # 全店へのお知らせ板（放送）＝1件1投稿（第3弾）
        state.setdefault('allboard_reads', [])  # どの店が全店板をいつまで読んだか（未読判定用）
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
            'uploaded_at': jst.now().strftime('%Y/%m/%d %H:%M'),   # 日本時間（UTCずれ対策）
            'rows': str(len(slim_rows)),
            'format': 'OK' if format_ok else 'NG(別様式)',
            'filename': filename,
        }
        return self.state['index']

    def load_index(self):
        """ _index だけを返す（Gシート版と同じインターフェース）。 """
        return dict(self.state['index'])

    def load_current_month_stores(self, index=None):
        # index は Gシート版と引数をそろえるためだけに受ける（ローカルは読み直しの費用が無い）
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

    def load_prev_proposal(self, current_ym=None):
        """ 前月の結果（滞留の判定材料）を返す。
            ローカル保管庫は検証・お試し用で月をまたいだ運用をしないため、
            state['prev_proposal'] に手で入れたものだけを返す（既定は空＝全部「今月から」）。
            ★テストから滞留の挙動を確かめられるように、入口だけは用意しておく。 """
        return dict(self.state.get('prev_proposal', {}))

    def load_exclusions(self):
        """ 店が「融通に出さない」と外した品目の一覧を返す（Gシート版と同じ形）。 """
        return list(self.state['exclusions'])

    def save_exclusions(self, rows):
        """ 除外リストを丸ごと入れ替える（Gシート版と同じ挙動）。 """
        self.state['exclusions'] = list(rows)

    def load_reservations(self):
        """ 受け手の店が押さえた品目の一覧を返す（Gシート版と同じ形）。 """
        return list(self.state['reservations'])

    def save_reservations(self, rows):
        """ 予約リストを丸ごと入れ替える（Gシート版と同じ挙動）。 """
        self.state['reservations'] = list(rows)

    def load_supply_qty(self):
        """ 出し手の店が「N錠だけ出す」と決めた数量の一覧を返す（Gシート版と同じ形）。 """
        return list(self.state['supply_qty'])

    def save_supply_qty(self, rows):
        """ 提供数量リストを丸ごと入れ替える（Gシート版と同じ挙動）。 """
        self.state['supply_qty'] = list(rows)

    def load_messages(self):
        """ 店舗間のやり取り（掲示板）を全部返す（Gシート版 read_messages と同じ形）。 """
        return list(self.state['messages'])

    def append_message(self, row):
        """ やり取りを1件だけ追記する（★追記のみ＝過去の投稿を消さない・Gシート版と同じ挙動）。 """
        self.state['messages'].append(dict(row))

    def load_msg_reads(self):
        """ どの店がどのスレッドをいつまで読んだかを返す（Gシート版 read_msg_reads と同じ形）。 """
        return list(self.state['msg_reads'])

    def save_msg_reads(self, rows):
        """ 既読リストを丸ごと入れ替える（Gシート版 write_msg_reads と同じ挙動）。 """
        self.state['msg_reads'] = list(rows)

    def load_allboard(self):
        """ 全店へのお知らせ板の投稿を全部返す（Gシート版 read_allboard と同じ形）。 """
        return list(self.state['allboard'])

    def append_allboard(self, row):
        """ 全店板へ1件だけ追記する（★追記のみ＝過去の投稿を消さない・Gシート版と同じ挙動）。 """
        self.state['allboard'].append(dict(row))

    def load_allboard_reads(self):
        """ どの店が全店板をいつまで読んだかを返す（Gシート版 read_allboard_reads と同じ形）。 """
        return list(self.state['allboard_reads'])

    def save_allboard_reads(self, rows):
        """ 全店板の既読リストを丸ごと入れ替える（Gシート版 write_allboard_reads と同じ挙動）。 """
        self.state['allboard_reads'] = list(rows)
