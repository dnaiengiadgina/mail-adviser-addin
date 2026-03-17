# メール送信確認アドイン for Microsoft 365 Outlook

メール送信前に宛先・件名・添付ファイルを確認し、誤送信を防止するOutlookアドインです。  
[m-FILTER MailAdviser](https://www.daj.jp/bs/ma/) のようなポップアップ型誤送信対策を自社で実装します。

## ✅ 主な機能

| 機能 | 説明 |
|------|------|
| 宛先確認 | To/CC/BCC の全宛先を一覧表示 |
| 社外ドメイン警告 | `@fujilogi.co.jp` 以外の宛先を赤色でハイライト |
| 件名チェック | 件名が未入力の場合に警告 |
| 添付忘れチェック | 本文に「添付」「別紙」等のキーワードがあるのにファイルがない場合に警告 |
| 送信前チェックリスト | 4項目のチェックボックスで確認を促す |
| Smart Alerts | 送信ボタン押下時に自動でポップアップダイアログを表示 |

## 🗂 ファイル構成

```
mail-adviser-addin/
├── manifest.xml                      # M365管理センターにアップロードするXMLマニフェスト
├── taskpane.html                     # タスクパネルUI（リボンボタンから開く確認パネル）
├── src/
│   └── launchevent/
│       ├── launchevent.html          # event-based activation用HTML（新Outlook/OWA用）
│       └── launchevent.js            # Smart Alertsイベントハンドラ
├── assets/
│   ├── icon-16.png                   # アイコン 16x16
│   ├── icon-32.png                   # アイコン 32x32
│   ├── icon-64.png                   # アイコン 64x64
│   ├── icon-80.png                   # アイコン 80x80
│   └── icon-128.png                  # アイコン 128x128
└── generate_icons.py                 # アイコン再生成スクリプト
```

## 🚀 デプロイ手順

### Step 1: GitHub Pages を有効化

1. GitHubで `mail-adviser-addin` というリポジトリを **Public** で作成
2. このフォルダの全ファイルをリポジトリにプッシュ
3. GitHubリポジトリ → **Settings** → **Pages**
4. Source: **Deploy from a branch** → Branch: `main` / `/(root)` → Save
5. 数分後に `https://dnaiengiadgina.github.io/mail-adviser-addin/` でアクセス可能になることを確認

> ⚠️ GitHub Pages は **HTTPS** で配信されるため、Outlookアドインの要件（SSL必須）を満たします。

### Step 2: M365管理センターからアドインを展開

1. [Microsoft 365 管理センター](https://admin.microsoft.com) に **グローバル管理者** でサインイン
2. 左メニュー: **設定** → **統合アプリ** (Integrated apps)
3. **アプリをアップロード** → **Office アドイン** → **マニフェストファイルをアップロード**
4. `manifest.xml` を選択してアップロード
5. デプロイ対象ユーザー/グループを選択（テスト時は特定ユーザーのみ推奨）
6. **デプロイ** をクリック

> 📝 展開後、Outlookへの反映には最大24時間かかる場合があります。

### Step 3: 動作確認

1. 対象ユーザーの **新しいOutlook** を開く
2. **新しいメール** を作成
3. リボンに「**送信確認**」ボタンが表示されることを確認
4. 送信ボタンを押すと Smart Alerts ダイアログが表示されることを確認

## ⚙️ カスタマイズ

### 社内ドメインの変更

`src/launchevent/launchevent.js` の以下の行を変更:

```javascript
const INTERNAL_DOMAIN = 'fujilogi.co.jp';  // ← 自社ドメインに変更
```

### 警告レベルの変更

`manifest.xml` の `SendMode` を変更:

```xml
<LaunchEvent Type="OnMessageSend"
             FunctionName="onMessageSendHandler"
             SendMode="PromptUser" />   <!-- PromptUser / SoftBlock / Block -->
```

| オプション | 動作 |
|-----------|------|
| `PromptUser` | 警告ダイアログを表示、ユーザーが「送信する」「戻る」を選択可能 |
| `SoftBlock` | ユーザーは警告を無視して送信可能 |
| `Block` | 条件未達成の場合は送信不可（厳格モード） |

### キーワードチェックの追加

`src/launchevent/launchevent.js` の `keywords` 配列に追加:

```javascript
const keywords = ['添付', '別紙', 'ファイルを', '資料を', 'attached', 'attachment'];
```

## 📋 動作要件

- Microsoft 365 Business / Enterprise サブスクリプション
- 新しいOutlook for Windows または Outlook on the web (OWA)
- Mailbox requirement set 1.12 以上（Smart Alerts）
- M365管理者によるアドイン展開（ユーザー自身によるインストールは不可）

## 🔒 セキュリティ・プライバシー

- このアドインはメールの内容をサーバーに送信しません
- すべての処理はクライアント側（ブラウザ内）で実行されます
- Office.js API経由でのみメール情報にアクセスします

## 📖 参考ドキュメント

- [Office Add-ins manifest (XML)](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview)
- [Smart Alerts walkthrough](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Handle OnMessageSend event](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events)
- [Event-based activation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation)
