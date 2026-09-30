# 業務日誌 入力補助（包括）（Outlook アドイン）

Outlook の予定の作成画面に「業務日誌 入力補助（包括）」ボタンを出し、件名と本文を地域包括の業務日誌の自動記入ルール
（件名の先頭に【種別】、参加者数・外部参加者は本文）どおりに入力するアドイン。

- 使える: Outlook on the web、新しい Outlook（Windows・Mac）、従来の Outlook（Windows・Mac）
- 使えない: スマホの Outlook（予定作成画面のアドインに非対応。スマホでは手入力）
- 費用: なし（画面は GitHub Pages に置く。配布は Microsoft 365 管理センター）

## 中身

| ファイル | 内容 |
|---|---|
| `docs/taskpane.html` `taskpane.js` `taskpane.css` | アドインの画面（GitHub Pages で公開する） |
| `docs/rules.js` | 件名・本文の組み立てルール。**Office スクリプト「業務日誌作成_包括」の TAGS と合わせること** |
| `docs/assets/` | アイコン |
| `manifest.template.xml` → `manifest.xml` | 管理センターに登録する設定ファイル。`make_manifest.py` で作る |
| `addin_id.txt` | アドインの ID（変えない。変えると別のアドイン扱いになる） |

## 1. GitHub Pages に置く（初回だけ）

法人の GitHub アカウントで、組織の下に **公開（Public）** リポジトリを作る（無料プランで Pages を使うには公開が必要。
中身は画面のコードだけで、利用者情報やパスワードは含まない）。

```sh
cd outlook-addin
gh auth login -h github.com -p https -w          # 法人アカウントでサインイン（初回だけ）
git init -b main && git add docs README.md manifest.template.xml make_manifest.py addin_id.txt
git commit -m "業務日誌 入力補助アドイン"
gh repo create <組織名>/houkatsu-nisshi-addin --public --source . --push
gh api -X POST repos/<組織名>/houkatsu-nisshi-addin/pages -f "source[branch]=main" -f "source[path]=/docs"
```

数分後に `https://<組織名>.github.io/houkatsu-nisshi-addin/taskpane.html` が開ければ公開できている。

## 2. manifest.xml を作る

```sh
python make_manifest.py https://<組織名>.github.io/houkatsu-nisshi-addin
```

## 3. 包括のメンバーに配布する（Microsoft 365 管理センター、画面操作だけ）

1. https://admin.microsoft.com →「設定」→「統合アプリ」→「カスタム アプリをアップロード」
2. アプリの種類「Office アドイン」→「デバイスからマニフェスト ファイルをアップロード」で `manifest.xml` を選ぶ
3. 「ユーザー」で「特定のユーザー/グループ」→ **地域包括支援センター**（チーム）を選ぶ
4. 内容を確認して「展開」

反映まで最大24時間。各自の Outlook の予定の作成画面（リボンや「…」の中）に「業務日誌 入力補助（包括）」が出る。
先に1人で試すときは、手順3で自分だけを選ぶか、Outlook on the web の「アドインを取得」→「個人用アドイン」→「カスタム アドインの追加」→「ファイルから追加」で manifest.xml を読み込む。

## 画面やルールを直すとき

1. `docs/` を直して `git commit` → `git push`（数分で反映。管理センターの作業は不要）
2. ボタン名・アイコン・URL などマニフェストを変えたときだけ、`make_manifest.py <URL> 1.0.1.0` のようにバージョンを上げて作り直し、管理センターの「統合アプリ」→ このアプリ →「更新」で manifest.xml を差し替える
3. 【】の種類を増やすなどルールを変えたときは、Office スクリプト（`../tools/業務日誌作成_包括.ts` の TAGS）も同じように直す

## 自動化用アカウントに切り替えるとき

GitHub の組織・リポジトリはそのままでよい（個人アカウントに依存しない）。管理センターの配布もそのまま。
