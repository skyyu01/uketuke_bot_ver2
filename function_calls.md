# プロジェクト関数呼び出し関係

このドキュメントは、プロジェクト内の各スクリプトファイルで定義されている関数と、それらの間の呼び出し関係をまとめたものです。

---

## `01_webアプリ.js`

Google Site埋め込み用のWebアプリケーションとして動作し、ユーザーをChatのDMにリダイレクトする役割を担います。

### 定義されている関数

- `doGet(e)`: WebアプリにGETリクエストがあったときのエントリーポイント。

### 呼び出し関係

- `Chat.Spaces.setup`: ユーザーとBotのDMスペースを作成または取得します。
- `HtmlService`: リダイレクト用のHTMLページを生成します。

---

## `02_メッセージ処理.js`

Google Chatからのユーザーメッセージを受け取り、質問としてスプレッドシートに記録する役割を担います。

### 定義されている関数

- `onMessage(e)`: Chatでメッセージが送信されたときのエントリーポイント。
- `isDriveFolderAttachment_(att)`: 添付がフォルダか判定するヘルパー。
- `extractSenderEmailStrict_(e)`: 送信者メールアドレスを厳密に抽出するヘルパー。
- `handleConfirmationAction(e)`: 「はい/いいえ」の確認カードのアクションを処理します。
- `loadWelcomeUrls_()`: ウェルカムメッセージの画像URLを読み込みます。
- `buildImageGridCard_(urls, perRow)`: ウェルカムカードを生成します。
- `onAddedToSpace(e)`: Botがスペースに追加されたときのエントリーポイント。
- `getHeaderWithAppCredentials()`: Chat API用の認証ヘッダーを取得します。
- `disableButtonsAndPatch_(messageName, statusText)`: 操作後のカードを無効化します。
- `extractMessageName_(e)`: イベントオブジェクトからメッセージ名を取得します。
- `fetchSpaceAndThreadByMessageName_(messageName)`: メッセージ名からスペースとスレッドの情報を取得します。

### 呼び出し関係

- `onMessage(e)`
  - `handleReviewAction(e)` (`05_chatスペース転記.js` 内) を呼び出すことがあります。
  - `handleConfirmationAction` をトリガーするカードを返します。
- `handleConfirmationAction(e)`
  - `saveAttachmentWithSubject_(...)` (`03_google drive転記.js` 内) を呼び出します。
  - `appendRowToSheetWithSA_(...)` (`03_google drive転記.js` 内) を呼び出してスプレッドシートに質問を記録します。
  - `fetchSpaceAndThreadByMessageName_(...)` を呼び出します。
  - `disableButtonsAndPatch_(...)` を呼び出します。
- `onAddedToSpace(e)`
  - `buildImageGridCard_(...)` を呼び出します。

---

## `03_google drive転記.js`

Google Driveへのファイル保存や、Google Sheets APIの直接操作など、サービスアカウントを利用したバックエンド処理を担当します。

### 定義されている関数

- `saveAttachmentWithSubject_(att, subjectEmail)`: 添付ファイルをDriveに保存します。
- `getSaService_()`: サービスアカウントの認証サービスを取得します。
- `downloadFromChatWithSA_(resourceName)`: Chat添付をダウンロードします。
- `uploadToDriveWithSA_(bytes, filename, mimeType)`: Driveにファイルをアップロードします。
- `appendRowToSheetWithSA_(row, sheetName, ...)`: シートに行を追記します。
- `embedText_(text)`: テキストをベクトル化します（RAG用）。
- `rebuildPastCasesIndex_()`: 過去案件のRAGインデックスを再構築します。
- `ragSearch_v3(question, k)`: RAGインデックスを検索します。
- `downloadDriveFileAsUser_(fileId, ...)`: Driveファイルをダウンロードします。
- `saveDriveAttachmentTwoToken_(...)`: ユーザー権限とサービスアカウント権限を組み合わせてファイルを転送します。
- `saveAnswerLogJson_(row, payload)`: 回答ログをJSONファイルとしてDriveに保存します。

### 呼び出し関係

- このファイルの関数群は、主に `02_メッセージ処理.js` と `04_検索エージェント群.js` から呼び出されます。
- `UrlFetchApp`, `DriveApp`, `SpreadsheetApp`, `OAuth2` ライブラリなどを多用します。

---

## `04_検索エージェント群.js`

質問に対する回答を生成するコアロジック。RAGとWeb検索を組み合わせています。

### 定義されている関数

- `generateAnswers()`: メインのトリガー関数。
- `generateSearchKeywords_(question)`: RAG用のキーワードを生成。
- `chunkText_(text)`: テキストをチャンクに分割。
- `findRelevantChunks_(chunks, keywords, topK)`: 関連チャンクを検索。
- `getTextFromGoogleSlide_(slideUrl)`
- `getTextFromGoogleSheet_(sheetUrl)`
- `getTextFromGoogleDoc_(docUrl)`
- `researchAgent(initialQuestion)`: RAGとWeb検索を実行するエージェント。
- `generateQuery(state, config)`: Web検索用のクエリを生成。
- `webResearch(searchQuery, allowedUrls)`: Web検索を実行。
- `reflection(state)`: 検索結果を評価。
- `finalizeAnswer(state)`: 最終回答を生成。
- `callGeminiApi(apiUrl, payload)`: Gemini APIのラッパー。
- `analyzeQuestion(question)`: 質問を分析。

### 呼び出し関係

- `generateAnswers()`
  - `analyzeQuestion(...)` を呼び出します。
  - `researchAgent(...)` を呼び出します。
  - `saveAnswerLogJson_(...)` (`03_google drive転記.js` 内) を呼び出します。
  - `finalizeAndNotify_(...)` (`05_chatスペース転記.js` 内) を呼び出します。
  - `postReviewCard_(...)` (`05_chatスペース転記.js` 内) を呼び出します。
  - `sendNotificationToChat(...)` (`05_chatスペース転記.js` 内) を呼び出します。
- `researchAgent(initialQuestion)`
  - `getTextFrom...` 関数群を呼び出します。
  - `generateSearchKeywords_(...)` を呼び出します。
  - `chunkText_(...)` を呼び出します。
  - `findRelevantChunks_(...)` を呼び出します。
  - `generateQuery(...)` を呼び出します。
  - `webResearch(...)` を呼び出します。
  - `reflection(...)` を呼び出します。
  - `finalizeAnswer(...)` を呼び出します。
  - `ragSearch_v3(...)` または `ragSearch_v2(...)` (`03_google drive転記.js` 内) を呼び出すことがあります。
- `finalizeAnswer(state)`, `generateSearchKeywords_`, `generateQuery`, `webResearch`, `reflection`, `analyzeQuestion`
  - `callGeminiApi(...)` を呼び出します。

---

## `05_chatスペース転記.js`

Google Chatへの通知やメッセージ送信、レビューフローを担当します。

### 定義されている関数

- `sendNotificationToChat(userName, question, filelink, answer)`: Webhookで生成完了通知カードを送信。
- `finalizeAndNotify_(row, finalAnswer, userName, dmSpaceName)`: Chat APIで最終回答をDM送信。
- `postReviewCard_(row, userName, question, draftAnswer, ...)`: Webhookでレビュー依頼カードを送信。
- `handleReviewAction(e)`: レビュー操作（承認/差戻など）を処理。
- `setupReviewActionToken()`: レビュー機能用の認証トークンをセットアップ。

### 呼び出し関係

- `sendNotificationToChat(...)`
  - `UrlFetchApp.fetch(...)` を使用してWebhookにPOSTします。
- `finalizeAndNotify_(...)`
  - `Chat.Spaces.Messages.create(...)` (Chat API Advanced Service) を使用します。
- `postReviewCard_(...)`
  - `UrlFetchApp.fetch(...)` を使用してWebhookにPOSTします。
- `handleReviewAction(e)`
  - `finalizeAndNotify_(...)` を呼び出すことがあります。
