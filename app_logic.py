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
            '薬品名': r['薬品名'], '単位': r['単位'], '在庫数': r['在庫数'],
            '在庫金額': r['在庫金額'], '有効期限': r['有効期限'],
            '期限切迫区分': r['期限切迫区分'], '区分': r['区分'],
            '引取候補店': taker_label,
            # 期限切迫フラグ（有効期限まで5ヶ月以内か）。②の薄赤ハイライトに使う。
            #   ★判定は yuzu_core（compute_matching）で1度だけ行い、画面へはその結果を運ぶだけにして
            #     判定の出どころを一本化する（画面側で期限切迫区分の非空を見ると定義がズレるため）。
            '_expiry_flag': r.get('_expiry_flag', False),
            # 除外チェック欄で使う内部キー（画面には出さない）
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
        出し手店／薬品名／滞留／単位／出し手の在庫数／自店の在庫数／在庫金額／有効期限／消化目安／区分／医薬品CD
      ※2026-07-28 本間部長指示で『なぜ候補か』（不足中／使用中）列を削除した。
      ※2026-08-01（本間部長確定）に列を2つ整理した（①の改修）：
        ・既存の『在庫数』（＝出し手がいま持っている数）を『出し手の在庫数』に改名（取り違え防止）。
        ・その隣に『自店の在庫数』（＝自店がいまその品を何個持っているか）を新設。
        ★この改名は【④の画面表示だけ】。Excel／Gシートの列名『在庫数』には一切触っていない。
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
        ★『在庫数』→『出し手の在庫数』の改名と『自店の在庫数』の追加はここで一元的に行う
          （④本体と別枠 build_view_receive_ref で同じ整え方を使い、二重管理しないため）。 """
    return {
        '出し手店': r.get('出し手店', ''), '薬品名': r.get('薬品名', ''),
        '滞留': r.get('滞留', ''), '単位': r.get('単位', ''),
        # ①の改修：既存『在庫数』を『出し手の在庫数』に改名（値は同じ）＋『自店の在庫数』を新設
        '出し手の在庫数': r.get('在庫数', ''), '自店の在庫数': r.get('自店の在庫数', ''),
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

      列は④本体とまったく同じ（出し手の在庫数＋自店の在庫数の2列構成）。
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
        state.setdefault('reservations', [])  # 受け手の店が「うちが引き取る」と押さえた品目
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
