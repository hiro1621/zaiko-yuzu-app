# -*- coding: utf-8 -*-
"""
append_allboard／append_message の「見出し行の穴」修正（2026-08-21）の検証テスト。

【何をするスクリプトか】
本番の Google スプレッドシートには一切つながず、ダミーの ws（ワークシート）を使って
append_allboard／append_message のロジックだけを確かめる。
gsheet_store の低レベル関数（_get_or_create_ws・_values・_update・_find_ws）を
ダミーに差し替え、書き込みが実際に2次元グリッドへどう反映されるかを再現したうえで、
本物の read_allboard／read_messages を通して「投稿が画面に出るか」まで確認する。

再現する障害：タブを新規作成した直後 get_all_values() が「空の1行」を返すと、
旧コードの `if not values:` が False になり、見出しを書かずに2行目から投稿を書いて
1行目が空のまま固定され、read 側が列名を引けず「まだ投稿はありません。」になっていた。

【実行方法（このフォルダで）】
    python test_append_header_fix_20260821.py
※ Google API は叩かない。本番シートには書き込まない。追加ライブラリ不要。
"""

import sys
import os

# このフォルダを import パスに入れる（gsheet_store / yuzu_core / jst を読むため）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gsheet_store as gs


# ============================================================================
# ダミーのワークシート（本物の gspread ws の代わり）
# ============================================================================
class FakeWS:
    """ 2次元グリッド grid を1枚だけ持つダミーの ws。
        _update の書き込みを記録（updates）しつつ、grid にも反映する。 """
    def __init__(self, grid):
        # grid は行のリスト（各行はセルのリスト）。get_all_values の返り値を模す。
        self.grid = [list(r) for r in grid]
        self.updates = []   # 記録: (range_name, 書き込んだ2次元配列)

    def append_rows(self, *a, **k):
        # ★万一 append_rows（values.append）が呼ばれたら即失敗させる（2026-08-12の禁止事項）
        raise AssertionError('append_rows は使ってはいけない（列ズレ障害の再発防止）')


def _col_row_from_a1(range_name):
    """ 'A1' / 'A5' などから (0始まりの開始行, 0始まりの開始列) を返す。列はAのみ想定。 """
    letters = ''.join(ch for ch in range_name if ch.isalpha())
    digits = ''.join(ch for ch in range_name if ch.isdigit())
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch.upper()) - ord('A') + 1)
    col -= 1
    row = int(digits) - 1
    return row, col


def _install_fakes(initial_grid):
    """ gsheet_store の低レベル関数をダミーに差し替え、対象の FakeWS を返す。 """
    ws = FakeWS(initial_grid)

    def fake_get_or_create(sh, title, **kw):
        return ws

    def fake_find(sh, title, **kw):
        return ws

    def fake_values(w):
        # get_all_values 相当：現在の grid を複製して返す
        return [list(r) for r in w.grid]

    def fake_update(w, values, range_name='A1'):
        w.updates.append((range_name, [list(r) for r in values]))
        # 実際のグリッドへ反映（指定範囲だけ上書き。ほかの行には触れない）
        start_row, start_col = _col_row_from_a1(range_name)
        for i, rowvals in enumerate(values):
            r = start_row + i
            while len(w.grid) <= r:
                w.grid.append([])
            # 必要な列まで空文字で伸ばす
            need = start_col + len(rowvals)
            while len(w.grid[r]) < need:
                w.grid[r].append('')
            for j, v in enumerate(rowvals):
                w.grid[r][start_col + j] = v

    gs._get_or_create_ws = fake_get_or_create
    gs._find_ws = fake_find
    gs._values = fake_values
    gs._update = fake_update
    return ws


# ============================================================================
# テスト本体
# ============================================================================
_passed = 0
_failed = 0


def check(name, cond):
    global _passed, _failed
    if cond:
        _passed += 1
        print('  [OK] ' + name)
    else:
        _failed += 1
        print('  [NG] ' + name)


def a1_updates(ws):
    """ ws への書き込みのうち、range_name が 'A1' のものだけ返す。 """
    return [u for u in ws.updates if u[0] == 'A1']


# ---- _全店板（append_allboard）--------------------------------------------
def t_allboard_truly_empty():
    print('■ 全店板：本当に空（[]）')
    ws = _install_fakes([])
    gs.append_allboard(None, {'投稿日時': 't1', '投稿店': '和光', '本文': 'お知らせ'})
    posts = gs.read_allboard(None)
    check('見出しがA1に入る', ws.grid[0] == list(gs.ALLBOARD_HEADERS))
    check('投稿がA2に入る', ws.grid[1][:3] == ['t1', '和光', 'お知らせ'])
    check('read で1件見える', len(posts) == 1 and posts[0]['本文'] == 'お知らせ')


def t_allboard_spurious_blank_row():
    print("■ 全店板：新規作成直後の空1行 [['','','']]（★今回の障害トリガ）")
    ws = _install_fakes([['', '', '']])
    gs.append_allboard(None, {'投稿日時': 't1', '投稿店': '和光', '本文': 'お知らせ'})
    posts = gs.read_allboard(None)
    check('見出しがA1に入る（旧コードは入れ損ねていた）', ws.grid[0] == list(gs.ALLBOARD_HEADERS))
    check('投稿がA2に入る', ws.grid[1][:3] == ['t1', '和光', 'お知らせ'])
    check('read で1件見える（旧コードは0件だった）', len(posts) == 1 and posts[0]['投稿店'] == '和光')


def t_allboard_missing_header_with_data():
    print('■ 全店板：見出し欠け＋データ有り（本番で実際に壊れていた状態）')
    grid = [
        ['', '', ''],
        ['2026/08/21 14:57:15', '和光', '【テストです】'],
        ['2026/08/21 15:22:08', '和光', '【お試しのテストです】'],
        ['2026/08/21 15:34:41', '徳丸', 'テスト'],
    ]
    ws = _install_fakes(grid)
    before_rows = [list(r) for r in grid[1:]]   # 既存データ行の控え
    gs.append_allboard(None, {'投稿日時': 't-new', '投稿店': '志木', '本文': '新規'})
    posts = gs.read_allboard(None)
    check('A1に見出しが入り直る', ws.grid[0] == list(gs.ALLBOARD_HEADERS))
    check('既存データ行（2〜4行目）は一切変わらない',
          [ws.grid[1][:3], ws.grid[2][:3], ws.grid[3][:3]] == before_rows)
    check('新規投稿は5行目に入る', ws.grid[4][:3] == ['t-new', '志木', '新規'])
    check('read で4件見える（既存3＋新規1）', len(posts) == 4)
    check('A1への書き込みは見出しのみ（データを上書きしない）',
          all(u[1] == [list(gs.ALLBOARD_HEADERS)] for u in a1_updates(ws)))


def t_allboard_normal_no_overwrite():
    print('■ 全店板：正常タブ（1行目に見出しがある）→ A1を絶対に上書きしない')
    grid = [
        list(gs.ALLBOARD_HEADERS),   # ★第4弾で見出しは6列。3列だと『旧見出し』とみなし移行される
        ['t1', '和光', 'a'],
    ]
    ws = _install_fakes(grid)
    gs.append_allboard(None, {'投稿日時': 't2', '投稿店': '徳丸', '本文': 'b'})
    posts = gs.read_allboard(None)
    check('A1への書き込みが1度も無い', len(a1_updates(ws)) == 0)
    check('投稿は3行目に入る', ws.grid[2][:3] == ['t2', '徳丸', 'b'])
    check('read で2件見える', len(posts) == 2)


def t_allboard_trailing_blank():
    print('■ 全店板：末尾に空行がぶら下がる正常タブ → 詰めた位置に書く')
    grid = [
        list(gs.ALLBOARD_HEADERS),   # ★第4弾で見出しは6列。3列だと『旧見出し』とみなし移行される
        ['t1', '和光', 'a'],
        ['', '', ''],
        ['', '', ''],
    ]
    ws = _install_fakes(grid)
    gs.append_allboard(None, {'投稿日時': 't2', '投稿店': '徳丸', '本文': 'b'})
    posts = gs.read_allboard(None)
    check('A1への書き込みが無い（見出しは正常）', len(a1_updates(ws)) == 0)
    check('投稿は3行目（末尾空行を無視した本当の行数+1）に入る', ws.grid[2][:3] == ['t2', '徳丸', 'b'])
    check('read で2件見える', len(posts) == 2)


def t_allboard_inner_blank_kept():
    print('■ 全店板：途中の空行は詰めない（末尾側だけ削る）')
    grid = [
        list(gs.ALLBOARD_HEADERS),   # ★第4弾で見出しは6列。3列だと『旧見出し』とみなし移行される
        ['t1', '和光', 'a'],
        ['', '', ''],          # ← 途中の空行
        ['t2', '徳丸', 'b'],
    ]
    ws = _install_fakes(grid)
    gs.append_allboard(None, {'投稿日時': 't3', '投稿店': '志木', '本文': 'c'})
    posts = gs.read_allboard(None)
    check('途中の空行（3行目）はそのまま残る', ws.grid[2][:3] == ['', '', ''])
    check('投稿は5行目に入る（途中の空行を詰めていない）', ws.grid[4][:3] == ['t3', '志木', 'c'])
    check('read で3件見える（空行はread側でスキップ）', len(posts) == 3)


# ---- _やり取り（append_message）：1対1の既存動作が変わらないこと -------------
def t_message_empty_drug_no_shift():
    print('■ やり取り：薬品名が空でも列ズレしない（新規タブ空1行から）')
    ws = _install_fakes([['', '', '', '', '', '']])
    gs.append_message(None, {'投稿日時': 't1', '店A': '和光', '店B': '徳丸',
                             '薬品名': '', '投稿店': '和光', '本文': '在庫ありますか'})
    msgs = gs.read_messages(None)
    check('見出しがA1に入る', ws.grid[0] == list(gs.MESSAGE_HEADERS))
    check('投稿がA列先頭から入る（E列やI列にズレない）',
          ws.grid[1][:6] == ['t1', '和光', '徳丸', '', '和光', '在庫ありますか'])
    check('read で1件見える', len(msgs) == 1)
    check('店A/店B/本文が正しく引ける',
          msgs[0]['店A'] == '和光' and msgs[0]['店B'] == '徳丸' and msgs[0]['本文'] == '在庫ありますか')
    check('薬品名は空のまま読める', msgs[0]['薬品名'] == '')


def t_message_append_to_existing():
    print('■ やり取り：既存スレッドへの追記（既存投稿を消さない・正常タブはA1を触らない）')
    grid = [
        ['投稿日時', '店A', '店B', '薬品名', '投稿店', '本文'],
        ['t1', '和光', '徳丸', 'アムロジン', '和光', 'ありますか'],
        ['t2', '和光', '徳丸', '', '徳丸', 'あります'],
    ]
    ws = _install_fakes(grid)
    before = [list(r) for r in grid]
    gs.append_message(None, {'投稿日時': 't3', '店A': '和光', '店B': '徳丸',
                             '薬品名': '', '投稿店': '和光', '本文': 'ありがとう'})
    msgs = gs.read_messages(None)
    check('A1への書き込みが無い（見出し正常）', len(a1_updates(ws)) == 0)
    check('既存の見出し・2投稿は変わらない', ws.grid[0:3] == before)
    check('新規は4行目に入る', ws.grid[3][:6] == ['t3', '和光', '徳丸', '', '和光', 'ありがとう'])
    check('read で3件見える', len(msgs) == 3)


def t_message_missing_header_recover():
    print('■ やり取り：見出し欠け＋データ有り → 見出しだけ直し、既存投稿は温存')
    grid = [
        ['', '', '', '', '', ''],
        ['t1', '和光', '徳丸', '', '和光', 'ありますか'],
    ]
    ws = _install_fakes(grid)
    gs.append_message(None, {'投稿日時': 't2', '店A': '和光', '店B': '徳丸',
                             '薬品名': '', '投稿店': '徳丸', '本文': 'あります'})
    msgs = gs.read_messages(None)
    check('A1に見出しが入り直る', ws.grid[0] == list(gs.MESSAGE_HEADERS))
    check('既存投稿（2行目）は変わらない', ws.grid[1][:6] == ['t1', '和光', '徳丸', '', '和光', 'ありますか'])
    check('新規は3行目に入る', ws.grid[2][:6] == ['t2', '和光', '徳丸', '', '徳丸', 'あります'])
    check('read で2件見える', len(msgs) == 2)


def main():
    print('=== append_allboard / append_message 見出し穴修正テスト（2026-08-21） ===\n')
    for fn in [
        t_allboard_truly_empty,
        t_allboard_spurious_blank_row,
        t_allboard_missing_header_with_data,
        t_allboard_normal_no_overwrite,
        t_allboard_trailing_blank,
        t_allboard_inner_blank_kept,
        t_message_empty_drug_no_shift,
        t_message_append_to_existing,
        t_message_missing_header_recover,
    ]:
        fn()
        print('')
    print('=== 結果: %d 件成功 / %d 件失敗 ===' % (_passed, _failed))
    sys.exit(1 if _failed else 0)


if __name__ == '__main__':
    main()
