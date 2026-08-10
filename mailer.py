# -*- coding: utf-8 -*-
"""
デッドストックリスト（mailer.py）＝店舗間のやり取り・予約のメール通知

【何をするスクリプトか】
デッドストックリスト（在庫融通アプリ）で
  1) 新着メッセージ  … 相手店へ
  2) 予約が入った    … 出し手店へ
  3) 予約が取り消された … 出し手店へ
の3種類の「お知らせメール」を送ります。掲示板・予約の保存が成功した直後に、
streamlit_app.py から1回だけ呼ばれます（メールはあくまで“おまけ”＝メールが送れなくても
掲示板・予約は必ず残ります）。

【送信のしくみ】
  ・カゴヤの送信サーバー（smtp.kagoya.net）を使います（Microsoft 365 ではありません）。
    ポート465はSSL、それ以外（587など）はSTARTTLS。実装は
    「技術料データ作成Webアプリ\backend\mailer.py」の書き方を写経しています
    （標準ライブラリ smtplib / email / ssl だけ・追加の pip install は不要）。
  ・設定は Streamlit の Secrets から読みます（★リポジトリがPublicなので、
    メールアドレス・パスワード・URLはコードに一切書きません。下は“書き方の見本”で、
    実際の値は Streamlit Cloud の Secrets 画面にだけ入れます）：
        [smtp]
        host = "smtp.kagoya.net"          # カゴヤの送信サーバー（公開サーバー名）
        port = 587
        user = "＜カゴヤのログインID＞"
        password = "＜パスワード＞"
        from = "＜差出人アドレス（noreply など）＞"

        app_url = "＜アプリのURL＞"          # ← [smtp] と同じ階層（トップレベル）

        [store_emails]
        "＜店名＞" = "＜その店のメールアドレス＞"
        …（14店分）
  ・[smtp] が無ければ送信を黙ってスキップし、画面に「メール通知は未設定です。」を1行だけ出します。
  ・宛先が [store_emails] に無い店はスキップし、店名を画面に出して知らせます（黙って捨てません）。
  ・送信に失敗しても投稿・予約は残し、警告（黄色）だけ出します（画面は止めません）。
  ・タイムアウト（既定15秒）を入れ、メール送信で画面が固まらないようにします。

【画面（Streamlit）との切り分け】
  ・このモジュールは Streamlit を import しません（品質管理部が単体テストしやすいように）。
  ・「画面に出す文言」は send_notifications() の戻り値 messages＝[(レベル, 文言), …] として返し、
    実際の表示は呼び出し側（streamlit_app.py）が行います。

【必要ライブラリ】標準ライブラリのみ（smtplib / email / ssl / urllib）。requirements.txt の変更は不要。

※ venv不要。Windows専用パス。コメント・メッセージはすべて日本語です。
"""
import ssl
import smtplib
from urllib.parse import quote
from email.message import EmailMessage
from email.utils import formataddr

# 差出人の表示名（例：デッドストックリスト <差出人アドレス>）
FROM_DISPLAY_NAME = 'デッドストックリスト'
# メール送信のタイムアウト（秒）。メールで画面を固めないための保険。
_DEFAULT_TIMEOUT = 15
# 本文に載せるメッセージ冒頭の文字数（★中身の全文は載せない＝在庫情報を社外へ撒かないため）
_EXCERPT_LEN = 100


# ============================================================================
# Secrets（秘密）を安全に読む … st.secrets でも普通の dict でも動くようにする
#   （st を import しないので、テストからは dict をそのまま渡せる）
# ============================================================================
def get_smtp_config(secrets):
    """ Secrets（マッピング）から [smtp] を取り出す。無ければ None（＝送信スキップ）。
        host と from が欠けている設定も None 扱いにする（送信元不明の誤送信を防ぐ）。 """
    try:
        smtp = secrets['smtp'] if ('smtp' in secrets) else None
    except Exception:
        smtp = None
    if not smtp:
        return None
    try:
        conf = {
            'host': str(smtp.get('host', 'smtp.kagoya.net') or '').strip(),
            'port': str(smtp.get('port', '587') or '587').strip(),
            'user': str(smtp.get('user', '') or '').strip(),
            'password': str(smtp.get('password', '') or '').strip(),
            'from': str(smtp.get('from', '') or '').strip(),
        }
    except Exception:
        return None
    if not conf['host'] or not conf['from']:
        return None
    return conf


def get_store_email(secrets, store):
    """ [store_emails] から店のメールアドレスを引く。無ければ ''（＝スキップ＋画面で知らせる）。 """
    try:
        emails = secrets['store_emails'] if ('store_emails' in secrets) else {}
    except Exception:
        emails = {}
    try:
        return str(emails.get(store, '') or '').strip()
    except Exception:
        return ''


def get_app_url(secrets):
    """ トップレベルの app_url を読む。無ければ ''（＝本文にリンクを付けない）。 """
    try:
        return str(secrets['app_url']).strip() if ('app_url' in secrets) else ''
    except Exception:
        return ''


# ============================================================================
# 文面の組み立て（純関数＝テストしやすい）
# ============================================================================
def _disp_store(name):
    """ 店名を件名・本文用の表示名にする。★アプリの画面と同じ呼び方をそのまま使う。

        いったん「末尾が店／薬局／堂でなければ『店』を足す」実装にしたが、
        本間部長の指示で『店』は付けないことにした（2026-08-10）。
        理由：アプリの画面はどこも「東立石」「さと和光」と呼んでいるので、
        メールだけ呼び方が変わると、メールと画面を突き合わせにくくなる。
        （「さと和光店」のような据わりの悪い表記も出てしまう。）
        前後の空白だけ落として、そのまま返す。 """
    return str(name or '').strip()


def _excerpt(text, n=_EXCERPT_LEN):
    """ メッセージ本文の冒頭 n 文字。長ければ末尾に … を付ける（★全文は載せない）。 """
    t = str(text or '').strip()
    return (t[:n] + '…') if len(t) > n else t


def build_app_link(app_url, store):
    """ ?store=<店名> を付けたアプリのURL（既存の _remember_store_in_url / _store_from_query と同じ仕組み）。
        app_url が空なら '' を返す（＝本文にリンク行を出さない）。日本語店名はURLエンコードする。 """
    if not app_url:
        return ''
    sep = '&' if ('?' in app_url) else '?'
    return '%s%sstore=%s' % (app_url, sep, quote(str(store or '')))


def _link_block(app_link):
    """ 本文末尾の共通ブロック（返信しないでほしい旨＋リンク）。
        ★文言は本間部長の指示（2026-08-10）。
          「返信できません」ではなく「返信しないでください」とお願いする言い方にする。
          差出人が noreply なので返信しても誰にも届かない＝返した本人が
          「送ったのに返事がない」と待ってしまうのを防ぐのが目的。 """
    lines = ['※ このメールには返信しないでください。',
             '　 お返事は、下のリンクからアプリの「⑤ やり取り」でお願いします。']
    if app_link:
        lines.append(app_link)
    else:
        lines.append('（アプリのURLが未設定のため、リンクは省略しています。'
                     'ふだんお使いのブックマークからお開きください。）')
    return '\n'.join(lines)


def build_notification(kind, actor_store, recipient_store, drugs, body_excerpt, app_link):
    """
    お知らせメールの件名と本文を組み立てる純関数。
      kind            … 'message'（新着メッセージ）/ 'reserved'（予約が入った）/ 'cancelled'（予約取消）
      actor_store     … 動作した店（メッセージの送信者／予約者／取消者）
      recipient_store … 宛先の店（メッセージの相手店／予約・取消なら出し手店）
      drugs           … 対象の薬品名のリスト（メッセージは0〜1件、予約・取消は0件以上）
      body_excerpt    … 新着メッセージのときだけ本文の冒頭（他は ''）
      app_link        … ?store=<宛先店> 付きのアプリURL（空ならリンク行なし）
    戻り値：(件名, 本文)
    """
    disp_actor = _disp_store(actor_store)
    disp_recipient = _disp_store(recipient_store)
    drugs = [str(d or '').strip() for d in (drugs or []) if str(d or '').strip()]

    if kind == 'message':
        subject = '【デッドストックリスト】%sから連絡があります' % disp_actor
        parts = ['%s から、デッドストックリストで連絡がありました。' % disp_actor, '']
        if drugs:
            parts.append('対象の薬：%s' % drugs[0])
            parts.append('')
        parts.append('メッセージの冒頭：')
        parts.append(body_excerpt or '（本文なし）')
        parts.append('')
        parts.append('（メッセージの全文はアプリでご確認ください。'
                     '在庫の詳しい情報はメールには載せていません。）')
        parts.append('')
        parts.append(_link_block(app_link))
        return subject, '\n'.join(parts)

    if kind == 'reserved':
        subject = '【デッドストックリスト】%sが予約しました' % disp_actor
        head = '%s が、%sのデッドストックを予約しました。' % (disp_actor, disp_recipient)
    elif kind == 'cancelled':
        subject = '【デッドストックリスト】%sが予約を取り消しました' % disp_actor
        head = '%s が、%sへのデッドストックの予約を取り消しました。' % (disp_actor, disp_recipient)
    else:
        # ★知らない種類は、黙って何かを送らずにここで止める。
        #   以前は else が『取り消しました』の文面になっていたため、
        #   呼び出し側の綴り違い1つで「予約されていないのに取消の通知が飛ぶ」
        #   という、いちばん困る間違いが起こり得た（2026-08-10 に塞いだ）。
        #   呼び出し元（send_notifications）は例外を握って画面に警告を出すだけなので、
        #   ここで止めても投稿・予約・取消そのものは必ず残る。
        raise ValueError('お知らせメールの種類が不明です：%r'
                         '（message／reserved／cancelled のどれかにしてください）' % (kind,))

    parts = [head, '']
    if drugs:
        parts.append('対象の薬：')
        parts.extend(['・%s' % d for d in drugs])
        parts.append('')
    parts.append('数量の相談・受け渡しは、アプリのやり取り（掲示板）か'
                 '電話・デスクネッツでお願いします。')
    parts.append('')
    parts.append(_link_block(app_link))
    return subject, '\n'.join(parts)


# ============================================================================
# 送信（カゴヤSMTP）… backend\mailer.py の send_report_mail を写経
# ============================================================================
def send_mail(smtp_conf, to_addr, subject, body, timeout=_DEFAULT_TIMEOUT):
    """ メールを1通送る。ポート465はSSL(SMTPS)、それ以外（587など）はSTARTTLS。
        件名・本文（日本語）のエンコードは EmailMessage が自動で行う。
        失敗時は日本語の RuntimeError を送出する（呼び出し側で握って画面に警告表示）。 """
    host = smtp_conf['host']
    try:
        port = int(str(smtp_conf.get('port', '587') or '587').strip())
    except ValueError:
        port = 587
    from_addr = smtp_conf['from']
    user = smtp_conf.get('user') or ''
    password = smtp_conf.get('password') or ''

    msg = EmailMessage()
    # 差出人の表示名を付ける（例：デッドストックリスト <差出人アドレス>）。
    msg['From'] = formataddr((FROM_DISPLAY_NAME, from_addr))
    msg['To'] = to_addr
    msg['Subject'] = subject
    msg.set_content(body)

    # SSLコンテキスト（既定＝厳密。証明書のCA署名＋ホスト名一致を検証）
    ctx = ssl.create_default_context()
    try:
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=timeout, context=ctx) as s:
                if user and password:
                    s.login(user, password)
                s.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=timeout) as s:
                s.ehlo()
                s.starttls(context=ctx)
                s.ehlo()
                if user and password:
                    s.login(user, password)
                s.send_message(msg)
    except Exception as e:
        # SMTP例外・ネットワーク不通・タイムアウトなどをまとめて日本語で返す
        raise RuntimeError('メール送信に失敗しました（%s:%d）：%s' % (host, port, e))


# ============================================================================
# 通知のまとめ役（画面に出す文言はここで作って“返す”だけ。表示は呼び出し側）
# ============================================================================
def send_notifications(secrets, kind, actor_store, targets, timeout=_DEFAULT_TIMEOUT):
    """
    複数の宛先へお知らせメールをまとめて送る。
      secrets … Secrets（st.secrets でも dict でも可）
      kind    … 'message' / 'reserved' / 'cancelled'
      actor_store … 動作した店
      targets … [{'store': 宛先店, 'drugs': [薬品名...], 'body_excerpt': 冒頭文字}, ...]
      timeout … 送信のタイムアウト（秒）
    戻り値：
      {'messages': [(レベル, 文言), ...],  # 呼び出し側が画面に出す（'info'/'warning'）
       'sent':     [送れた宛先アドレス, ...]}
    ※ [smtp] 未設定なら送信せず、'info' の案内1行だけを返す（掲示板・予約は呼び出し側で保存済み）。
    """
    smtp_conf = get_smtp_config(secrets)
    if smtp_conf is None:
        return {'messages': [('info', 'メール通知は未設定です。')], 'sent': []}

    app_url = get_app_url(secrets)
    messages = []
    sent = []
    for tg in (targets or []):
        store = str((tg or {}).get('store', '') or '').strip()
        if not store:
            continue
        to_addr = get_store_email(secrets, store)
        if not to_addr:
            messages.append((
                'warning',
                'メール通知先が未登録のため、%s には通知できませんでした'
                '（掲示板・予約は保存済みです）。' % store))
            continue
        link = build_app_link(app_url, store)
        subject, body = build_notification(
            kind, actor_store, store, tg.get('drugs') or [],
            tg.get('body_excerpt', ''), link)
        try:
            send_mail(smtp_conf, to_addr, subject, body, timeout=timeout)
            sent.append(to_addr)
        except Exception as e:
            messages.append((
                'warning',
                '%s へのメール通知に失敗しました（掲示板・予約は保存済みです）：%s' % (store, e)))
    return {'messages': messages, 'sent': sent}


def _group_by_store(rows):
    """ [{'出し手店' or '_出し手店', '薬品名'}] を 出し手店ごとに [(store, [薬品名...]), ...] へまとめる。
        並びは初出順、薬品名は店ごとに重複を除く（黙って捨てない・空店名は無視）。 """
    order = []
    by = {}
    seen = {}
    for r in (rows or []):
        store = str((r or {}).get('出し手店', '') or (r or {}).get('_出し手店', '') or '').strip()
        if not store:
            continue
        if store not in by:
            by[store] = []
            seen[store] = set()
            order.append(store)
        name = str((r or {}).get('薬品名', '') or '').strip()
        if name and name not in seen[store]:
            seen[store].add(name)
            by[store].append(name)
    return [(s, by[s]) for s in order]


def notify_new_message(secrets, actor_store, recipient_store, drug, body,
                       timeout=_DEFAULT_TIMEOUT):
    """ 新着メッセージを相手店へ通知する。drug は任意（空なら薬の指定なし）。 """
    drugs = [drug] if str(drug or '').strip() else []
    targets = [{'store': recipient_store, 'drugs': drugs, 'body_excerpt': _excerpt(body)}]
    return send_notifications(secrets, 'message', actor_store, targets, timeout=timeout)


def notify_reservation(secrets, actor_store, rows, timeout=_DEFAULT_TIMEOUT):
    """ 予約が入ったことを出し手店へ通知する。rows＝新しく予約できた行（出し手店・薬品名を持つ）。 """
    targets = [{'store': s, 'drugs': ds, 'body_excerpt': ''}
               for s, ds in _group_by_store(rows)]
    if not targets:
        return {'messages': [], 'sent': []}
    return send_notifications(secrets, 'reserved', actor_store, targets, timeout=timeout)


def notify_cancellation(secrets, actor_store, rows, timeout=_DEFAULT_TIMEOUT):
    """ 予約が取り消されたことを出し手店へ通知する。rows＝取り消した予約の行（出し手店・薬品名を持つ）。 """
    targets = [{'store': s, 'drugs': ds, 'body_excerpt': ''}
               for s, ds in _group_by_store(rows)]
    if not targets:
        return {'messages': [], 'sent': []}
    return send_notifications(secrets, 'cancelled', actor_store, targets, timeout=timeout)
