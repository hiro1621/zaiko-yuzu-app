# 在庫融通アプリ 起動・更新・ロールバック手順（本間部長用）

このフォルダ（`streamlit_app`）は、**店舗が自分でファイルをアップロードできる在庫融通アプリ**一式です。
このフォルダまるごとが「1つのGitHubリポジトリ」になります。プログラミングの知識がなくても、
下の手順どおりにコピー＆貼り付けすれば公開できます。

> **結論（先に全体像）**
> 1. Google側の準備（スプレッドシート＋鍵を作る） →
> 2. GitHubに新しい置き場を作ってこのフォルダを上げる →
> 3. Streamlit Cloud という無料サービスにつなぐ →
> 4. 秘密（鍵とパスワード）を Streamlit に貼る →
> 5. 「Deploy」を押すと、店舗が使えるURLが1本できあがります。
>
> 一度作れば、あとは**このフォルダを直して GitHub に上げ直すだけで自動で反映**されます。

---

## このフォルダに入っているもの

| ファイル | 役割 |
|---|---|
| `streamlit_app.py` | 画面本体（これが「起動ファイル」） |
| `yuzu_core.py` | 計算の頭脳（コマンド版 `create_yuzu_list.py` と**同じ計算を共有**） |
| `app_logic.py` | 画面に依存しない下ごしらえ処理 |
| `gsheet_store.py` | Googleシート（裏の保管庫）の読み書き |
| `stores_config.py` | **36店の店名リスト（ソユーズ9・内観堂5・飛鳥22＝3法人）**（増減はここだけ直す。2026-09-18 に3法人対応） |
| `requirements.txt` | 必要な部品の一覧（Streamlitが自動で入れる） |
| `.streamlit/secrets.toml.sample` | 秘密の「見本」（管理者7人の合言葉 `[admin_passwords]`・店ごとの合言葉 `[store_passwords]`・メール宛先36店の型）。本物はここには置きません |
| `mailer.py` | メール通知（カゴヤSMTP・1回の接続で複数通・全店板は全体60秒の予算つき） |
| `jst.py` | 日本時間の「いま」（Streamlit Cloud は UTC で動くため） |
| `test_*.py` | 本番シート・メール・ブラウザに触らない単体テスト（品質管理部の検証用） |
| `.streamlit/config.toml` | アップロード上限などの設定 |
| `.gitignore` | 秘密・在庫データを**絶対にGitHubへ上げない**ための除外設定 |

> ⚠️ **大事**：`yuzu_core.py` はコマンド版ツールも参照しています。
> このフォルダの名前や場所（`…\在庫融通\streamlit_app`）は**変えないで**ください。
> 変えるときは情報システム部にご相談を（コマンド版の参照先も直す必要があります）。

---

## STEP 1：Google側の準備（スプレッドシート＋サービスアカウント鍵）

「サービスアカウント」＝**アプリ専用のGoogleの作業ロボット**のことです。
このロボットにだけシートを預け、店舗は直接シートを触りません。

1. **スプレッドシートを新規作成**します（Googleドライブ ▸ 新規 ▸ Googleスプレッドシート）。
   名前は例えば「在庫融通_保管庫」。
   - URL `https://docs.google.com/spreadsheets/d/【この部分】/edit` の
     **【この部分】＝スプレッドシートID** を控えます（あとで使います）。
   - タブは空のままでOK（アプリが自動で `_index` や `raw_○○` を作ります）。
2. **Google Cloud** （https://console.cloud.google.com/）で
   - 上の「プロジェクトを選択」→「新しいプロジェクト」を作る（名前は「yuzu」等）。
   - 検索窓で「**Google Sheets API**」を探し「**有効にする**」を押す。
   - 左メニュー「APIとサービス ▸ 認証情報」→「**認証情報を作成 ▸ サービスアカウント**」。
     名前は「yuzu-bot」等。役割は付けなくてOK。作成する。
   - できたサービスアカウントを開き「**キー ▸ 鍵を追加 ▸ 新しい鍵 ▸ JSON**」。
     → **JSONファイルがダウンロード**されます（＝これが鍵。中身は絶対に人に見せない）。
   - JSONの中の `client_email`（例 `yuzu-bot@...iam.gserviceaccount.com`）を控えます。
3. **STEP 1-1で作ったスプレッドシートを、このロボットに共有**します。
   - スプレッドシート右上「共有」→ 先ほどの `client_email` を貼り付け、
     **権限は「編集者」**にして共有。
   - これで「アプリ専用の裏の保管庫（サービスアカウントにだけ編集者共有）」になります。
     「リンクを知っている全員＝編集者」の共有は**しません**（廃止）。

---

## STEP 2：GitHubに新しい置き場（リポジトリ）を作る

1. https://github.com/ にログイン → 右上「＋ ▸ New repository」。
2. 名前は例えば `zaiko-yuzu-app`。**Private（非公開）**を選ぶ。「Create repository」。
3. 「uploading an existing file」リンクから、
   **この `streamlit_app` フォルダの中身をすべてドラッグ＆ドロップ**して「Commit changes」。
   - `.gitignore` のおかげで、秘密（`secrets.toml`）や在庫データ（CSV/XLS）は
     間違って選んでも**上がりません**。`secrets.toml.sample`（見本）は上げてOKです。

---

## STEP 3：Streamlit Cloud につなぐ（無料）

1. https://share.streamlit.io/ に GitHub アカウントでサインイン。
2. 「**Create app ▸ Deploy a public app from GitHub**」を選び、
   - Repository：`（あなたの名前）/zaiko-yuzu-app`
   - Branch：`main`
   - **Main file path：`streamlit_app.py`**
3. 「**Advanced settings ▸ Secrets**」を開き、下の内容を貼り付けます
   （`secrets.toml.sample` と同じ形。**中身を本物に置き換え**）。

   ```toml
   spreadsheet_id = "STEP1で控えたスプレッドシートID"
   app_url = "アプリのURL"

   [admin_passwords]
   "本間" = "その人の合言葉"
   # …管理者7人ぶん（本間／森田／小島／ソユーズ担当者／内観堂担当者／加藤／飛鳥薬局担当者。どの名前でも36店を自由に切替できる。2026-09-19）

   [store_passwords]
   "東大泉" = "その店の合言葉"
   # …36店ぶん（キーは stores_config.py の店名そのまま。見本 secrets.toml.sample 参照）

   [gcp_service_account]
   # ダウンロードしたJSONの中身をこの下に丸ごと写す（見本 secrets.toml.sample 参照）
   type = "service_account"
   project_id = "..."
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "yuzu-bot@....iam.gserviceaccount.com"
   client_id = "..."
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "..."
   universe_domain = "googleapis.com"
   ```

   > `private_key` は、JSONに書いてある `\n` を**そのまま**残してコピーします（改行に直さない）。
   > 旧型の `app_password`（共有パスワード1つ）は、`[store_passwords]` が無いときだけ使われます（後方互換）。
   > 旧型の `admin_password`（本部用1つだけ）も後方互換で通ります（名前は「本部」）が、`[admin_passwords]` を推奨。同じ値を2か所に入れると、その値では入れません。
4. 「**Deploy**」を押す。1〜3分で **`https://○○.streamlit.app` というURL**ができます。
   これを店舗に配ります（社外には出さない）。

---

## STEP 4：ふだんの使い方（店舗）

1. URLを開く → **自店の合言葉を1回入れる**（合言葉で自店が決まります。店の選び直しはできません）。
   管理者の合言葉（`[admin_passwords]`・名前ごと7人）で入ったときだけ、左バーに「本部モード：名前」と出て、法人（ソユーズ／内観堂／飛鳥）で絞って店を選べます。
2. **薬VANの在庫ファイル（.xls / .csv / .xlsx）を選んでアップロード**（ファイル名は「店名_YYYYMM」）。
3. 画面に「現在 N/36店 アップ済み」と、②デッド品／③期限切迫品／④引き取れる薬／⑤やり取りが出ます。
   全店そろうほど精度が上がります。
4. ④で予約するときの受取時期は「今すぐ（随時便）」か「次のまとめ便（1・4・7・10月）」を選びます。

---

## 更新のしかた（あとで直したいとき）

- 直したいファイルを情報システム部が修正 → **GitHubに上げ直す**だけで、
  Streamlit Cloud が**自動で新しい版に切り替わります**（数分）。
- 例：店を1つ増減 → `stores_config.py` を直す → GitHubへ。「○/36」の分母も自動で変わります。
  ★店を足したら、Secrets の `[store_passwords]`（合言葉）と `[store_emails]`（メール宛先）にもその店を足します。
- **部品（`app_logic.py`・`yuzu_core.py`・`gsheet_store.py`・`mailer.py`・`stores_config.py`）を変えたときは、
  Streamlit Cloud の「Reboot app」を押してください**（押さないと古い部品のまま動くことがあります）。
- 秘密（パスワード・鍵）を変えたいときは、**Streamlit の Secrets 画面**で直します
  （GitHub側は触りません）。

## ロールバック（元に戻したいとき）

- GitHub の対象ファイルの「History（履歴）」から前の版を開き「Revert（戻す）」。
  → 自動で前の版に戻ります。
- または Streamlit Cloud の管理画面で、そのアプリを一度停止（Delete/Reboot）して
  安定版のブランチで作り直すこともできます。困ったら情報システム部へ。

---

## 安全のための約束（必ず守る）

- **鍵JSON・共有パスワードは GitHub に置かない**（`.gitignore` で自動除外済み。`secrets.toml.sample` の見本だけ置く）。
- スプレッドシートは**サービスアカウントにだけ「編集者」共有**。店舗には直接共有しない。
- 配ったアプリURLは**社外に出さない**（店舗の在庫情報が入っています）。
- このアプリは**まったく新しい別リポジトリ**です。技術料Webアプリ・LINE受付・自動印刷など、
  既存の本番システムには一切関係しません（触れていません）。

---

## 2つ目のアプリ（検証用）と、3法人36店への切替（2026-10）

3法人対応（飛鳥22店の追加・店ごとの合言葉・500円ルール・便の月・引取依頼書の記名欄）は、
本番を止めずに **改修用ブランチ `feature/3houjin-36stores` ＋ 2つ目の Streamlit アプリ（検証用）** で試運転してから本番へ合流します。
手順書は次の2枚です（本間部長用・値は書いてありません）：

- `..\docs\3法人対応改修_202610\検証用アプリ登録手順_本間さん用.md` … 検証用アプリの登録（ブランチ→シートのコピー→New app→Secrets→メール実測）
- `..\docs\3法人対応改修_202610\切替日の手順_本間さん用.md` … 10月末の本番合流（バックアップ→merge→push→Secrets→Reboot→確認→案内）と戻し方
