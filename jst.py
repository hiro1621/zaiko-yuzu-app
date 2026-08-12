# -*- coding: utf-8 -*-
"""
日本時間（JST）で「いま」と「今日」を返す小さな部品（jst.py）

【何をするファイルか】
Streamlit Cloud のサーバーは UTC（協定世界時＝日本より9時間おそい時計）で動いているため、
そのまま datetime.datetime.now() を使うと、記録される時刻が実際より9時間ずれます
（午前10時半の投稿が 01:24 で保存される、など）。締切を「日付」で判定するこのアプリでは、
このずれがあると「10日の15時以降はもう11日あつかい」になって月次スケジュールが成立しません。
そこで、このファイルの now()／today() を通して必ず日本時間で取り直します。

【なぜ独立したファイルにするか】
gsheet_store.py も app_logic.py も yuzu_core（計算エンジン）を import しているため、
日付関数をそれらの中に置くと import の依存が絡みます。この jst.py は標準の datetime しか
import しないので、どのファイルからでも安全に読み込めます。

【必要なライブラリのインストール】
　追加のインストールは不要です（Python 標準の datetime だけを使います）。

※ venv不要。Windows専用パス。コメント・メッセージはすべて日本語です。
"""

import datetime

# 日本標準時（UTC より9時間すすんでいる）
JST = datetime.timezone(datetime.timedelta(hours=9))


def now():
    """ 日本時間の「いま」を返す。
        ★tzinfo（時差の情報）は落として素の datetime で返す。
          既存コードが素の datetime を前提にしており、時差つきと素のものを比較すると
          TypeError（型が合わないエラー）になるため、ここで時差情報を外しておく。 """
    return datetime.datetime.now(JST).replace(tzinfo=None)


def today():
    """ 日本時間の「今日」（date）を返す。 """
    return now().date()
