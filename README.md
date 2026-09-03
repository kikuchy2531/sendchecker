# 送信前 宛先チェック アドイン（新しい Outlook 用）

新しい Outlook（バージョン 1.2026.x 系）は Web ベースのため、VBA / VSTO / COM アドインは動作しません。
本ツールは Office.js の **イベントベース アドイン（Smart Alerts / OnMessageSend）** として実装しています。

## 機能
- 送信ボタン押下時に To / Cc / Bcc の宛先一覧をダイアログ表示し、確認を促す
- 社外アドレスが含まれる場合は件数付きの警告を先頭に表示し、社外宛先を分けて列挙
- 「このまま送信」／「宛先を修正する」を選択可能（`SendMode="PromptUser"`）

## ファイル
| ファイル | 役割 |
|---|---|
| `manifest.xml` | アドイン定義。`YOUR-HOST` と `<Id>` を置き換える |
| `launchevent.js` | 宛先取得・社外判定・ダイアログ生成 |
| `commands.html` | ランタイム用の空ページ（office.js と launchevent.js を読み込む） |

## セットアップ
1. `launchevent.js` の `INTERNAL_DOMAINS` に自社ドメインを設定
   （空にすると Outlook の `recipientType` 判定にフォールバック）
2. 3 ファイルとアイコン（16/32/64/80/128px PNG）を **HTTPS** の Web サーバーへ配置
   - 検証用: `npx http-server -S -C cert.pem -K key.pem` などでローカル HTTPS 可
3. `manifest.xml` の `https://YOUR-HOST` を実際の URL に、`<Id>` を新規 GUID に置き換え
4. 配布
   - **検証**: 新しい Outlook → 「アプリ」→「アドインを追加」→「カスタム アドインを追加」→「ファイルから」で manifest.xml を指定
   - **本番**: Microsoft 365 管理センター → 統合アプリ（一元展開）で組織配布（推奨。イベントベース アドインは管理者展開が安定）

## 動作の調整
- **社外宛は送信不可にしたい**: `manifest.xml` の `SendMode` を `SoftBlock` または `Block` に変更
  （この場合は社内のみのメールでもダイアログが出るため、`launchevent.js` で社外なし時に `allowEvent: true` を返すよう変更してください）
- **一覧表示件数**: `MAX_LIST`
- **エラー時の挙動**: 現在は例外発生時に送信を通します（`catch` 内を変更で厳格化可）

## 注意
- `errorMessageMarkdown` は Mailbox 要件セット 1.14 以上で有効。旧クライアントでは `errorMessage`（プレーンテキスト）が使われます
- OnMessageSend ハンドラー内では独自の UI（ダイアログ / 作業ウィンドウ）は開けないため、表示内容は Smart Alerts のダイアログ文面に限定されます
- 送信時に Outlook 側で「アドインの処理中」の表示が数秒出ることがあります
