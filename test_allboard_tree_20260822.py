# -*- coding: utf-8 -*-
"""
全店板の「ツリー返信・解決済み」（第4弾・2026-08-22）の検証テスト。

【何をするスクリプトか】
本番の Google スプレッドシートには一切つながず、ダミーの ws（ワークシート）と純関数だけで
第4弾の追加ぶんを確かめる。確認する対象は3ファイル：
  ・gsheet_store … _全店板 の見出し6列化・見出しの自動移行（旧3列→6列）・新3列の読み取り・
                   read_allboard のフィルタが返信/状態行を落とさないこと・同時書き込み対策（決定8）。
  ・app_logic   … build_allboard_tree（親子ツリー・多段・親なし返信・循環・解決済み・仮ID）と
                   allboard_unread_count の拡張（投稿＋返信を数え、状態は数えない）。
  ・mailer      … build_notification の allboard_reply 文面と notify_allboard_reply の宛先1店・自返信抑止。

【実行方法（このフォルダで）】
    python test_allboard_tree_20260822.py
※ Google API は叩かない。本番シートには書き込まない。メールも実送信しない。追加ライブラリ不要。
"""

import sys
import os

# このフォルダを import パスに入れる（gsheet_store / app_logic / mailer / yuzu_core / jst を読むため）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gsheet_store as gs
import app_logic
import mailer


# ============================================================================
# ダミーのワークシート（本物の gspread ws の代わり）… 既存 test の FakeWS を踏襲
# ============================================================================
class FakeWS:
    """ 2次元グリッド grid を1枚だけ持つダミーの ws。
        _update の書き込みを記録（updates）しつつ、grid にも反映する。
        drop_first_data_write=True のとき、最初の“データ行への書き込み（A1以外）”を1回だけ握りつぶす
        ＝2店が同じ行番号を計算して片方が上書き（消失）した状況を再現する（決定8の検証用）。 """
    def __init__(self, grid, drop_first_data_write=False):
        self.grid = [list(r) for r in grid]
        self.updates = []   # 記録: (range_name, 書き込んだ2次元配列)
        self.drop_first_data_write = drop_first_data_write

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


def _install_fakes(initial_grid, drop_first_data_write=False):
    """ gsheet_store の低レベル関数をダミーに差し替え、対象の FakeWS を返す。 """
    ws = FakeWS(initial_grid, drop_first_data_write=drop_first_data_write)

    def fake_get_or_create(sh, title, **kw):
        return ws

    def fake_find(sh, title, **kw):
        return ws

    def fake_values(w):
        return [list(r) for r in w.grid]

    def fake_update(w, values, range_name='A1'):
        w.updates.append((range_name, [list(r) for r in values]))
        # 同時書き込みの再現：最初のデータ行（A1以外）への書き込みを1回だけ捨てる
        if w.drop_first_data_write and range_name != 'A1':
            w.drop_first_data_write = False
            return
        start_row, start_col = _col_row_from_a1(range_name)
        for i, rowvals in enumerate(values):
            r = start_row + i
            while len(w.grid) <= r:
                w.grid.append([])
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
# テスト集計
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
    return [u for u in ws.updates if u[0] == 'A1']


# ============================================================================
# 1) gsheet_store：見出し6列化・移行・新3列の読み取り・同時書き込み対策
# ============================================================================
def t_headers_are_six():
    print('■ gsheet：ALLBOARD_HEADERS が6列（投稿ID・親ID・種別を追加）')
    check('見出しは6列', list(gs.ALLBOARD_HEADERS) ==
          ['投稿日時', '投稿店', '本文', '投稿ID', '親ID', '種別'])


def t_append_new_post_six_cols():
    print('■ gsheet：新規タブへ6列の投稿を追記→read で6キー引ける')
    ws = _install_fakes([])
    gs.append_allboard(None, {'投稿日時': 't1', '投稿店': '和光', '本文': 'お知らせ',
                              '投稿ID': 'p1', '親ID': '', '種別': '投稿'})
    posts = gs.read_allboard(None)
    check('見出しがA1に6列で入る', ws.grid[0] == list(gs.ALLBOARD_HEADERS))
    check('投稿がA2に6列で入る', ws.grid[1][:6] == ['t1', '和光', 'お知らせ', 'p1', '', '投稿'])
    check('read で投稿IDが引ける', len(posts) == 1 and posts[0]['投稿ID'] == 'p1')
    check('read で種別が引ける', posts[0]['種別'] == '投稿')


def t_header_migration_3to6():
    print('■ gsheet：旧3列の見出し＋旧データ→1行目だけ6列へ移行・データ行は不変')
    grid = [
        ['投稿日時', '投稿店', '本文'],                      # 旧3列の見出し
        ['2026/08/21 14:57:15', '和光', '旧投稿1'],
        ['2026/08/21 15:22:08', '徳丸', '旧投稿2'],
    ]
    ws = _install_fakes(grid)
    before_data = [list(r) for r in grid[1:]]
    gs.append_allboard(None, {'投稿日時': 't-new', '投稿店': '志木', '本文': '新規',
                              '投稿ID': 'p9', '親ID': '', '種別': '投稿'})
    posts = gs.read_allboard(None)
    check('A1が6列の見出しへ入り直る', ws.grid[0] == list(gs.ALLBOARD_HEADERS))
    check('旧データ行（2〜3行目）は3列のまま一切変わらない',
          [ws.grid[1][:3], ws.grid[2][:3]] == before_data)
    check('新規は4行目に6列で入る', ws.grid[3][:6] == ['t-new', '志木', '新規', 'p9', '', '投稿'])
    check('read は3件（旧2＋新1）', len(posts) == 3)
    check('旧投稿は新3列が空で読める（投稿ID空）',
          posts[0]['投稿ID'] == '' and posts[0]['親ID'] == '' and posts[0]['種別'] == '')
    check('A1への書き込みは見出しのみ（データを上書きしない）',
          all(u[1] == [list(gs.ALLBOARD_HEADERS)] for u in a1_updates(ws)))


def t_six_col_header_not_overwritten():
    print('■ gsheet：正常な6列見出し→A1を絶対に上書きしない')
    grid = [
        list(gs.ALLBOARD_HEADERS),
        ['t1', '和光', 'a', 'p1', '', '投稿'],
    ]
    ws = _install_fakes(grid)
    gs.append_allboard(None, {'投稿日時': 't2', '投稿店': '徳丸', '本文': 'b',
                              '投稿ID': 'p2', '親ID': '', '種別': '投稿'})
    check('A1への書き込みが1度も無い', len(a1_updates(ws)) == 0)
    check('投稿は3行目に入る', ws.grid[2][:6] == ['t2', '徳丸', 'b', 'p2', '', '投稿'])


def t_read_keeps_reply_and_state_rows():
    print('■ gsheet：read_allboard のフィルタが「返信・状態」行を落とさない')
    grid = [
        list(gs.ALLBOARD_HEADERS),
        ['t1', '和光', '大元', 'p1', '', '投稿'],
        ['t2', '徳丸', '返信します', 'r1', 'p1', '返信'],
        ['t3', '和光', '解決済み', 's1', 'p1', '状態'],       # 状態行（本文＝状態値）
        ['', '', '', '', '', ''],                              # 空行は落ちる
    ]
    _install_fakes(grid)
    posts = gs.read_allboard(None)
    kinds = [p['種別'] for p in posts]
    check('投稿・返信・状態の3行が残る（空行だけ落ちる）', len(posts) == 3)
    check('返信行が残る', '返信' in kinds)
    check('状態行が残る（本文＝解決済みで非空なので落ちない）', '状態' in kinds)


def t_has_post_id_helper():
    print('■ gsheet：_allboard_has_post_id が見出し名で投稿ID列を引いて判定')
    values = [list(gs.ALLBOARD_HEADERS),
              ['t1', '和光', 'a', 'p1', '', '投稿']]
    check('存在する投稿IDは True', gs._allboard_has_post_id(values, 'p1') is True)
    check('存在しない投稿IDは False', gs._allboard_has_post_id(values, 'zzz') is False)
    check('空IDは False（無限リトライ防止）', gs._allboard_has_post_id(values, '') is False)


def t_concurrent_write_rewrites_once():
    print('■ gsheet：同時書き込みで自分の行が消えたら【1回だけ】書き直す（決定8）')
    grid = [
        list(gs.ALLBOARD_HEADERS),
        ['t1', '和光', '既存', 'p1', '', '投稿'],
    ]
    # 最初のデータ行書き込みを握りつぶす＝別店に上書きされた状況
    ws = _install_fakes(grid, drop_first_data_write=True)
    gs.append_allboard(None, {'投稿日時': 't2', '投稿店': '徳丸', '本文': '新規',
                              '投稿ID': 'abc123', '親ID': '', '種別': '返信'})
    posts = gs.read_allboard(None)
    ids = [p['投稿ID'] for p in posts]
    check('消えた行が書き直されて read に現れる', 'abc123' in ids)
    check('既存投稿は残る', 'p1' in ids)
    # データ行（A1以外）への _update 呼び出しが2回（握りつぶし1回＋書き直し1回）
    data_writes = [u for u in ws.updates if u[0] != 'A1']
    check('データ行への書き込みはちょうど2回（1回だけ書き直す）', len(data_writes) == 2)


def t_concurrent_no_rewrite_when_present():
    print('■ gsheet：ふだん（自分の行が残っている）は書き直さない＝二重書き込みしない')
    grid = [list(gs.ALLBOARD_HEADERS), ['t1', '和光', '既存', 'p1', '', '投稿']]
    ws = _install_fakes(grid)   # 握りつぶさない＝正常
    gs.append_allboard(None, {'投稿日時': 't2', '投稿店': '徳丸', '本文': '新規',
                              '投稿ID': 'abc123', '親ID': '', '種別': '返信'})
    data_writes = [u for u in ws.updates if u[0] != 'A1']
    check('データ行への書き込みは1回だけ', len(data_writes) == 1)


# ============================================================================
# 2) app_logic：build_allboard_tree / allboard_unread_count / 仮ID・採番
# ============================================================================
def _P(dt, store, body, pid, parent='', kind='投稿'):
    return {'投稿日時': dt, '投稿店': store, '本文': body,
            '投稿ID': pid, '親ID': parent, '種別': kind}


def t_tree_basic():
    print('■ tree：投稿＋返信＝親の下にぶら下がる（深さ）')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:05:00', '徳丸', 'r1', 'r1', 'p1', '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    check('2件が表示順に並ぶ', [n['投稿ID'] for n in t] == ['p1', 'r1'])
    check('大元の深さは0', t[0]['深さ'] == 0)
    check('返信の深さは1', t[1]['深さ'] == 1)
    check('返信のルートIDは大元', t[1]['ルートID'] == 'p1')


def t_tree_multilevel():
    print('■ tree：返信への返信（多段）も許す（決定3）')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:05:00', '徳丸', 'r1', 'r1', 'p1', '返信'),
        _P('2026/08/22 10:10:00', '和光', 'r2', 'r2', 'r1', '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    depth = {n['投稿ID']: n['深さ'] for n in t}
    check('表示順は p1→r1→r2', [n['投稿ID'] for n in t] == ['p1', 'r1', 'r2'])
    check('孫返信の深さは2', depth['r2'] == 2)
    check('孫返信のルートは p1', {n['投稿ID']: n['ルートID'] for n in t}['r2'] == 'p1')


def t_tree_sibling_order():
    print('■ tree：兄弟の返信は投稿日時の昇順で並ぶ')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:09:00', '徳丸', 'later', 'rB', 'p1', '返信'),
        _P('2026/08/22 10:03:00', '志木', 'earlier', 'rA', 'p1', '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    check('p1→rA(早)→rB(遅) の順', [n['投稿ID'] for n in t] == ['p1', 'rA', 'rB'])


def t_tree_orphan_reply_becomes_root():
    print('■ tree：親が見つからない返信は大元扱いで必ず出す（黙って消さない）')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:05:00', '徳丸', '迷子', 'x1', 'NOPE', '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    ids = [n['投稿ID'] for n in t]
    check('迷子の返信も表示される', 'x1' in ids and len(t) == 2)
    check('迷子は深さ0（大元扱い）', {n['投稿ID']: n['深さ'] for n in t}['x1'] == 0)


def t_tree_cycle_no_infinite_loop():
    print('■ tree：循環参照（A→B→A）でも無限ループせず両方出す')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'a', 'b', '返信'),
        _P('2026/08/22 10:01:00', '徳丸', 'B', 'b', 'a', '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    ids = sorted(n['投稿ID'] for n in t)
    check('両ノードとも1回ずつ出る（消えない・重複しない）', ids == ['a', 'b'])


def t_tree_legacy_id_stable_and_linkable():
    print('■ tree：旧3列の投稿に仮IDを毎回同じ値で振り、返信が結び付く')
    legacy = _P('2026/08/20 09:00:00', '和光', '旧デッド品', '', '', '')  # 投稿ID空＝旧行
    expect_id = app_logic._legacy_allboard_id(legacy)
    rows = [
        legacy,
        _P('2026/08/22 10:00:00', '徳丸', '引き取ります', 'r1', expect_id, '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    check('仮IDは legacy- で始まる', expect_id.startswith('legacy-'))
    check('旧投稿は仮IDで大元として出る', t[0]['投稿ID'] == expect_id and t[0]['深さ'] == 0)
    check('返信は旧投稿にぶら下がる', t[1]['深さ'] == 1 and t[1]['ルートID'] == expect_id)
    # 2回計算しても同じ値（追記のみで中身不変＝毎回同じ）
    check('仮IDは毎回同じ値', app_logic._legacy_allboard_id(legacy) == expect_id)


def t_tree_resolved_owner_only():
    print('■ tree：解決済みは「投稿した店」の状態行だけ有効（決定6）')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:05:00', '徳丸', 'r1', 'r1', 'p1', '返信'),
        _P('2026/08/22 10:06:00', '和光', '解決済み', 's1', 'p1', '状態'),   # 持ち主が立てた
    ]
    t = app_logic.build_allboard_tree(rows)
    root = t[0]
    check('大元は解決済み（ルート解決済み=True）', root['ルート解決済み'] is True)
    check('返信にもルート解決済み=Trueが伝わる（折りたたみ判定用）',
          t[1]['ルート解決済み'] is True)
    check('状態行は表示ノードに含まれない', all(n['種別'] != '状態' for n in t))


def t_tree_resolved_other_store_ignored():
    print('■ tree：別店が立てた状態行は無効（計算側でも無視＝二重の歯止め）')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:06:00', '徳丸', '解決済み', 's1', 'p1', '状態'),   # 別店が立てた
    ]
    t = app_logic.build_allboard_tree(rows)
    check('別店の状態は効かず未解決のまま', t[0]['ルート解決済み'] is False)


def t_tree_resolved_latest_wins():
    print('■ tree：状態は最新が勝つ（解決済み→未解決に戻す）')
    rows = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:06:00', '和光', '解決済み', 's1', 'p1', '状態'),
        _P('2026/08/22 10:20:00', '和光', '未解決', 's2', 'p1', '状態'),     # あとから未解決へ
    ]
    t = app_logic.build_allboard_tree(rows)
    check('最新が未解決なので未解決', t[0]['ルート解決済み'] is False)
    # 逆順（あとから解決済み）
    rows2 = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),
        _P('2026/08/22 10:06:00', '和光', '未解決', 's1', 'p1', '状態'),
        _P('2026/08/22 10:20:00', '和光', '解決済み', 's2', 'p1', '状態'),
    ]
    t2 = app_logic.build_allboard_tree(rows2)
    check('最新が解決済みなので解決済み', t2[0]['ルート解決済み'] is True)


def t_unread_counts_reply_not_state():
    print('■ unread：未読は投稿＋返信を数え、状態は数えない・自分の投稿は数えない')
    reads = [{'店名': '徳丸', '最終確認日時': '2026/08/22 09:00:00'}]
    posts = [
        _P('2026/08/22 10:00:00', '和光', 'A', 'p1'),                        # 他店の投稿→未読
        _P('2026/08/22 10:05:00', '志木', 'r1', 'r1', 'p1', '返信'),         # 他店の返信→未読
        _P('2026/08/22 10:06:00', '和光', '解決済み', 's1', 'p1', '状態'),   # 状態→数えない
        _P('2026/08/22 10:07:00', '徳丸', '自分の返信', 'r2', 'p1', '返信'), # 自店→数えない
        _P('2026/08/22 08:00:00', '和光', '既読より前', 'p0'),               # 既読より前→数えない
    ]
    n = app_logic.allboard_unread_count('徳丸', posts, reads)
    check('未読は2件（他店の投稿1＋他店の返信1）', n == 2)


def t_new_id_unique():
    print('■ 採番：new_allboard_id は12桁・毎回ユニーク')
    ids = {app_logic.new_allboard_id() for _ in range(2000)}
    check('2000回でも重複しない', len(ids) == 2000)
    check('すべて12桁', all(len(x) == 12 for x in ids))


# ---- 新着順（2026-08-22）：大元は新しい順・子は古い順のまま ---------------------
def t_tree_roots_desc_children_asc():
    print('■ tree：大元は新しい順（降順）・返信の子は古い順（昇順）＝多段でも子は昇順')
    rows = [
        _P('2026/08/22 10:00:00', '和光', '古い大元', 'p_old'),
        _P('2026/08/22 12:00:00', '徳丸', '新しい大元', 'p_new'),
        # p_old にぶら下がる返信（子は古い→新しい）。孫返信もぶら下げて多段にする。
        _P('2026/08/22 10:10:00', '志木', '子1（早）', 'c1', 'p_old', '返信'),
        _P('2026/08/22 10:20:00', '和光', '子2（遅）', 'c2', 'p_old', '返信'),
        _P('2026/08/22 10:30:00', '朝霞', '孫（c1への返信）', 'g1', 'c1', '返信'),
    ]
    t = app_logic.build_allboard_tree(rows)
    ids = [n['投稿ID'] for n in t]
    check('大元は新しい順＝新しい p_new が先頭', ids[0] == 'p_new')
    # p_old のツリーは 深さ優先・子は昇順： p_old → c1 →（c1の孫）g1 → c2
    check('p_old のツリーは子が昇順のまま（p_old→c1→g1→c2）',
          ids[1:] == ['p_old', 'c1', 'g1', 'c2'])
    depth = {n['投稿ID']: n['深さ'] for n in t}
    check('孫は深さ2（多段でも子は昇順で並ぶ）', depth['g1'] == 2)


def t_tree_roots_same_time_stable():
    print('■ tree：同時刻の大元でも並びが安定（入力順を変えても毎回同じ）')
    same = '2026/08/22 11:00:00'
    base = [
        _P(same, '和光', 'A', 'aaa'),
        _P(same, '徳丸', 'B', 'bbb'),
        _P(same, '志木', 'C', 'ccc'),
    ]
    rows_a = base
    rows_b = [base[2], base[0], base[1]]   # 入力の並びだけ入れ替える
    ids_a = [n['投稿ID'] for n in app_logic.build_allboard_tree(rows_a)]
    ids_b = [n['投稿ID'] for n in app_logic.build_allboard_tree(rows_b)]
    check('入力順を変えても並びは同じ（安定）', ids_a == ids_b)
    # roots は reverse=True＝第2キー（投稿ID）も降順。実質ユニークなので一意に定まる。
    check('同時刻は投稿IDで一意に並ぶ（降順 ccc→bbb→aaa）',
          ids_a == ['ccc', 'bbb', 'aaa'])


def t_tree_resolved_keeps_position():
    print('■ tree：解決済みにしても投稿の位置は変わらない（日付順のまま・持ち上げない）')
    base = [
        _P('2026/08/22 10:00:00', '和光', '古', 'p1'),
        _P('2026/08/22 11:00:00', '徳丸', '中', 'p2'),
        _P('2026/08/22 12:00:00', '志木', '新', 'p3'),
    ]
    before = [n['投稿ID'] for n in app_logic.build_allboard_tree(base)]
    check('解決前は新しい順（p3→p2→p1）', before == ['p3', 'p2', 'p1'])
    # 真ん中の p2 を「持ち主（徳丸）」が解決済みにする（状態行を1行足すだけ）。
    rows = base + [_P('2026/08/22 12:30:00', '徳丸', '解決済み', 's1', 'p2', '状態')]
    t = app_logic.build_allboard_tree(rows)
    after = [n['投稿ID'] for n in t]
    check('解決済みにしても並びは変わらない（p2 は持ち上がらない）',
          after == ['p3', 'p2', 'p1'])
    resolved = {n['投稿ID']: n['ルート解決済み'] for n in t}
    check('p2 は解決済みフラグが立つ', resolved['p2'] is True)
    check('p1・p3 は未解決のまま', resolved['p1'] is False and resolved['p3'] is False)


# ============================================================================
# 3) mailer：allboard_reply の文面／宛先1店・自返信抑止
# ============================================================================
def t_mail_build_allboard_reply():
    print('■ mail：build_notification(allboard_reply) の件名・本文')
    subject, body = mailer.build_notification(
        'allboard_reply', '和光', '徳丸', [], 'これは返信の本文です', 'http://example/app?store=徳丸')
    check('件名に「全店板の返信があります」', '全店板の返信があります' in subject)
    check('件名に投稿者店名（和光）', '和光' in subject)
    check('本文に返信の冒頭が入る', 'これは返信の本文です' in body)
    check('本文にアプリのリンクが入る', 'http://example/app?store=徳丸' in body)


def t_mail_notify_reply_targets_one():
    print('■ mail：notify_allboard_reply は「直近の親の投稿店1店だけ」を宛先にする')
    captured = {}
    orig = mailer.send_notifications

    def fake_send(secrets, kind, actor, targets, timeout=None):
        captured['kind'] = kind
        captured['actor'] = actor
        captured['targets'] = targets
        return {'messages': [], 'sent': []}

    mailer.send_notifications = fake_send
    try:
        mailer.notify_allboard_reply({}, '和光', '徳丸', '本文です')
        check("kind は 'allboard_reply'", captured.get('kind') == 'allboard_reply')
        check('宛先はちょうど1店', len(captured.get('targets', [])) == 1)
        check('宛先は親の投稿店（徳丸）', captured['targets'][0]['store'] == '徳丸')
        # 自分の投稿への自返信（親＝自店）は send を呼ばない
        captured.clear()
        res = mailer.notify_allboard_reply({}, '和光', '和光', '本文です')
        check('自返信は送信を呼ばない（captured 空）', 'kind' not in captured)
        check('自返信は空の結果を返す', res == {'messages': [], 'sent': []})
    finally:
        mailer.send_notifications = orig


# ============================================================================
def main():
    print('=== 全店板ツリー返信・解決済み 検証テスト（2026-08-22） ===\n')
    for fn in [
        # gsheet_store
        t_headers_are_six,
        t_append_new_post_six_cols,
        t_header_migration_3to6,
        t_six_col_header_not_overwritten,
        t_read_keeps_reply_and_state_rows,
        t_has_post_id_helper,
        t_concurrent_write_rewrites_once,
        t_concurrent_no_rewrite_when_present,
        # app_logic
        t_tree_basic,
        t_tree_multilevel,
        t_tree_sibling_order,
        t_tree_orphan_reply_becomes_root,
        t_tree_cycle_no_infinite_loop,
        t_tree_legacy_id_stable_and_linkable,
        t_tree_resolved_owner_only,
        t_tree_resolved_other_store_ignored,
        t_tree_resolved_latest_wins,
        t_unread_counts_reply_not_state,
        t_new_id_unique,
        # 新着順（2026-08-22）：大元は降順・子は昇順・解決済みは位置を変えない
        t_tree_roots_desc_children_asc,
        t_tree_roots_same_time_stable,
        t_tree_resolved_keeps_position,
        # mailer
        t_mail_build_allboard_reply,
        t_mail_notify_reply_targets_one,
    ]:
        fn()
        print('')
    print('=== 結果: %d 件成功 / %d 件失敗 ===' % (_passed, _failed))
    sys.exit(1 if _failed else 0)


if __name__ == '__main__':
    main()
