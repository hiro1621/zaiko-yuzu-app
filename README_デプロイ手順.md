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
| `stores_config.py` | **15店の店名リスト**（増減はここだけ直す） |
| `requirements.txt` | 必要な部品の一覧（Streamlitが自動で入れる） |
| `.streamlit/secrets.toml.sample` | 秘密の「見本」。本物はここには置きません |
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
   app_password = "店舗に配る共有パスワード"
   spreadsheet_id = "STEP1で控えたスプレッドシートID"

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
4. 「**Deploy**」を押す。1〜3分で **`https://○○.streamlit.app` というURL**ができます。
   これを店舗に配ります（社外には出さない）。

---

## STEP 4：ふだんの使い方（店舗）

1. URLを開く → **共有パスワードを1回入れる**。
2. **自分の店を選ぶ**（ドロップダウン。自由入力はできません）。
3. **薬VANの在庫ファイル（.xls / .csv / .xlsx）を選んでアップロード**。
4. 画面に「現在 N/15店 アップ済み」「自店視点の2ビュー」「全店一覧」が出ます。
   全店そろうほど精度が上がります。

---

## 更新のしかた（あとで直したいとき）

- 直したいファイルを情報システム部が修正 → **GitHubに上げ直す**だけで、
  Streamlit Cloud が**自動で新しい版に切り替わります**（数分）。
- 例：店を1つ増減 → `stores_config.py` を直す → GitHubへ。「○/15」の分母も自動で変わります。
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
