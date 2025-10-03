/**
 * @fileoverview
 * Google Chatからの質問をスプレッドシートに記録し、
 * 定期的にGoogle検索連携のGemini API（リサーチエージェント）で回答案を生成するスクリプト。
 *
 * @version 2.0.0
 */

// スクリプトプロパティから設定値を取得
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = SCRIPT_PROPERTIES.getProperty('SPREADSHEET_ID');
const GEMINI_API_KEY = SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
const SHEET_NAME = '問い合わせ一覧'; // スプレッドシートのシート名

// ==== テスト／本番 切り替え ====
const test_mode = true;
const CHAT_WEBHOOK_URL = test_mode
  ? SCRIPT_PROPERTIES.getProperty('CHAT_WEBHOOK_URL_TEST')  // テスト用
  : SCRIPT_PROPERTIES.getProperty('CHAT_WEBHOOK_URL');      // 本番用

if (!CHAT_WEBHOOK_URL) {
  console.warn('CHAT_WEBHOOK_URL が未設定です（mode=' + (test_mode ? 'TEST' : 'PROD') + '）');
}

const DRIVE_FOLDER_ID = SCRIPT_PROPERTIES.getProperty('DRIVE_FOLDER_ID'); // 添付保存先DriveフォルダID

// --- APIエンドポイント（必要に応じて変更） ---
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=' + GEMINI_API_KEY;
const GEMINI_API_URL_FLASH = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + GEMINI_API_KEY;

// --- 列の定義（A列=1, B列=2...） ---
const TIMESTAMP_COL = 1; // メッセージ受信日時
const USER_NAME_COL = 2;  // メッセージ送信者情報
const TOOL_SERVICE_COL = 3; // ツール・サービス分類列
const INQUIRY_TYPE_COL = 4; // 問い合わせ内容分類列
const QUESTION_COL = 6; // 質問本文
const SUMMARY_COL = 7;  // 質問要約列
const FILE_LINK_COL = 8; // 添付ファイルリンク列
const FINAL_ANSWER_COL = 9;   // ← 要確認：実シートの「最終回答」列の列番号
const DONE_CHECK_COL   = 11;   // ← 要確認：実シートの「対応完了（チェックボックス）」列
const FAQ_CHECK_COL    = 12;   // ← 要確認：実シートの「FAQ公開（チェックボックス）」列
const GEMINI_STATUS_COL = 13; // geminiでの回答作成状況
const GEMINI_ANSWER_COL = 14; // geminiでの回答案
const LOG_LINK_COL     = 17; // 回答ログURL（空いている列を使用）
const DM_WEBHOOK_URL_COL = 18; // 対象ユーザーとのDM URL,使っていない列を使用
// DM（1:1）の space 名（例: "spaces/AAAA..."）を保存する列
const DM_SPACE_NAME_COL   = 19;   // 未使用列に合わせて調整OK
// （任意）返信スレッドを保持したい場合
const CHAT_THREAD_NAME_COL = 20;  // 不要なら定義しなくてOK
const CHAT_MESSAGE_NAME_COL = 21; // メッセージIDを格納する列（Z列）




// --- 定数 ---
const STATUS_NEW_QUESTION = '新規質問';
const STATUS_ANSWER_GENERATED = '回答案生成済み';
const STATUS_TEMP_ERROR = '一時エラー';
const STATUS_ERROR = 'エラー';
const STATUS_PROCESSING = '処理中';
const AUTO_ANSWER_THRESHOLD = 0.7;
const NG_CATEGORIES = ['ルール・セキュリティ', '社内機密', '個人情報'];


// カテゴリの選択肢 (2軸)
const CATEGORY_TOOL_SERVICE = ["Gemini", "NotebookLM", "Google Workspace", "スマホ", "社内システム", "ツール選定", "全般・その他"];
const CATEGORY_INQUIRY_TYPE = ["使い方・基本操作", "プロンプト", "機能・仕様", "エラー・不具合", "活用アイデア相談", "ルール・セキュリティ", "その他"];

/* ==================================================================
 * 1. Google Chat 連携機能 (onMessage & handleConfirmationAction)
 * ================================================================== */

function onMessage(e) {
  console.log(e.chat.messagePayload.message.name); // メッセージのメタ情報
  const user = e.chat ? e.chat.user : e.user;
  const userName = user.displayName;
  const message = (e.chat && e.chat.messagePayload && e.chat.messagePayload.message) ? e.chat.messagePayload.message : e.message;
  const senderEmail = extractSenderEmailStrict_(message); // 空なら throw

  // レビュー操作のディスパッチ
  if (e.common && e.common.invokedFunction === 'handleReviewAction') {
    return handleReviewAction(e);
  }


  // ボット自身の発言には反応しない
  if (user.type === 'BOT') return;

  // フォルダ添付ブロック
  const hasFolder = Array.isArray(message?.attachment) && message.attachment.some(isDriveFolderAttachment_);
  if (hasFolder) {
    return {
      hostAppDataAction: {
        chatDataAction: {
          createMessageAction: {
            message: {
              text: '📁 フォルダは受け付けておりません。ファイルを添付してください🙇'
            }
          }
        }
      }
    };
  }

  if (!message.text) {
    return {
      hostAppDataAction: {
        chatDataAction: {
          createMessageAction: {
            message: { text: "ファイルのみの質問は受け付けておりません。質問文も合わせて入力をお願いします🙇" }
          }
        }
      }
    };
  }

  const questionText = message.text.trim();

  // ★追加: DMのspace名/スレッド名を取得（必ずDMで受ける前提）
  const dmSpaceName =
    (e?.chat?.messagePayload?.message?.space?.name) ||
    (e?.space?.name) || '';
  const dmThreadName =
    (e?.chat?.messagePayload?.message?.thread?.name) ||
    (message?.thread?.name) || '';


  // 添付ファイルの情報を取得（最初の添付のみ）
  let attachmentData = null;
  if (message.attachment && message.attachment.length > 0) {
    const att = message.attachment[0];
    if (att.driveDataRef && att.driveDataRef.driveFileId) {
      attachmentData = {
        type: 'DRIVE',
        fileId: att.driveDataRef.driveFileId,
        contentName: att.contentName || '',
        mimeType: att.contentType || 'application/octet-stream'
      };
    } else if (att.attachmentDataRef && att.attachmentDataRef.resourceName) {
      attachmentData = {
        type: 'CHAT',
        resourceName: att.attachmentDataRef.resourceName,
        contentName: att.contentName || '',
        mimeType: att.contentType || 'application/octet-stream'
      };
    }
    console.log('attachmentData(normalized)=', attachmentData);
  }

  // 確認カード
  return {
    hostAppDataAction: {
      chatDataAction: {
        createMessageAction: {
          message: {
            cardsV2: [{
              cardId: 'question-confirmation-card',
              card: {
                header: {
                  title: '内容の確認',
                  subtitle: '以下の内容で送信します。よろしければ「はい」をクリックしてください！ \n※追加の質問・相談も「はい」をクリックしてください！'
                },
                sections: [{
                  widgets: [
                    { textParagraph: { text: '<b>内容:</b><br>' + questionText } },
                    (attachmentData
                      ? { textParagraph: { text: '<b>ファイル名:</b><br>' + attachmentData.contentName } }
                      : {}),
                    {
                      buttonList: {
                        buttons: [
                          {
                            text: 'はい',
                            onClick: {
                              action: {
                                function: 'handleConfirmationAction',
                                parameters: [
                                  { key: 'isConfirmed', value: 'true' },
                                  { key: 'questionText', value: questionText },
                                  { key: 'user', value: userName },
                                  { key: 'attachment', value: JSON.stringify(attachmentData) },
                                  { key: 'senderEmail', value: senderEmail }
                                ]
                              }
                            }
                          },
                          {
                            text: 'いいえ',
                            onClick: {
                              action: {
                                function: 'handleConfirmationAction',
                                parameters: [{ key: 'isConfirmed', value: 'false' }]
                              }
                            }
                          }
                        ]
                      }
                    }
                  ]
                }]
              }
            }]
          }
        }
      }
    }
  };
}

// フォルダ添付かどうか
function isDriveFolderAttachment_(att) {
  const t = String(att?.contentType || att?.internalContentType || '').toLowerCase();
  return t === 'application/vnd.google-apps.folder';
}

// 空や許可外なら throw
function extractSenderEmailStrict_(e) {
  const email = (e?.sender?.email || '').trim().toLowerCase();
  if (!email) throw new Error('sender.email が取得できません');
  return email;
}

/**
 * 確認カードのボタン操作
 */
function handleConfirmationAction(e) {
  console.log("情報", e);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    const name = extractMessageName_(e);
    if (name) disableButtonsAndPatch_(name, '⚠️ この操作はすでに処理中/処理済みです。');
    return {};
  }

  try {
    const p = e?.commonEventObject?.parameters || {};
    const isConfirmed = p.isConfirmed === 'true';
    const messageName = extractMessageName_(e);

    if (!isConfirmed) {
      if (messageName) disableButtonsAndPatch_(messageName, '🚫 操作をキャンセルしました。\n再度メッセージをご入力ください。');
      return {};
    }

    const questionText = p.questionText || '';
    const userName = p.user || (e?.sender?.displayName || 'ユーザー');
    const attachmentParam = p.attachment || null;

    // 送信者メール（必須）
    let senderEmail = (p.senderEmail || '').trim().toLowerCase();
    if (!senderEmail) {
      try { senderEmail = extractSenderEmailStrict_(e); } catch (_) { senderEmail = ''; }
    }
    if (!senderEmail) {
      if (messageName) disableButtonsAndPatch_(messageName,
        '⚠️ エラーが発生しました。\n' +
        'お手数ですが、質問文を下記アドレス宛にメールにて送付お願いいたします。\n' +
        'genai@skylark.co.jp'
      );
      return {};
    }

    // 添付保存
    let fileUrl = '';
    try {
      if (attachmentParam && attachmentParam !== 'null') {
        const att = JSON.parse(attachmentParam);
        fileUrl = saveAttachmentWithSubject_(att, senderEmail);
      }
    } catch (err) {
      if (String(err?.code) === 'UNSUPPORTED_EXPORT' || String(err?.message).includes('UNSUPPORTED_EXPORT')) {
        if (messageName) disableButtonsAndPatch_(messageName,
          '⚠️ 申し訳ありません。現在非対応のファイルとなっております。\n' +
          'お手数ですが、記載いただいた質問文と添付ファイルを\n' +
          '下記アドレス宛にメールにて送付お願いいたします。\n' +
          'genai@skylark.co.jp'
        );
        return {};
      }
      if (String(err).includes('File not found')) {
        if (messageName) disableButtonsAndPatch_(messageName,
          '⚠️ エラーが発生しました。\n' +
          'お手数ですが、記載いただいた質問文と添付ファイルを\n' +
          '下記アドレス宛にメールにて送付お願いいたします。\n' +
          'genai@skylark.co.jp'
        );
        return {};
      }
      console.error('saveAttachment error:', err);
      if (messageName) disableButtonsAndPatch_(messageName,
        '⚠️ エラーが発生しました。\n' +
        'お手数ですが、記載いただいた質問文と添付ファイルを\n' +
        '下記アドレス宛にメールにて送付お願いいたします。\n' +
        'genai@skylark.co.jp'
      );
      return {};
    }

    // スプレッドシートへ追記
    const senderDisplay = senderEmail ? (userName + ' <' + senderEmail + '>') : userName;
    // messageName から Chat API でメッセージ詳細を取得し、space / thread を特定
    const msgName = messageName;
    const meta = fetchSpaceAndThreadByMessageName_(msgName); // { spaceName, threadName }

    const rowData = [
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'), //1
      senderDisplay, //2
      '', //3
      '', //4
      '', //5
      questionText, //6
      '', //7
      fileUrl, //7
      '', //8
      '', //9
      '', //10
      '', //11
      STATUS_NEW_QUESTION, //12
      '', //13
      '', //14
      '', //15
      '', //16
      '', //17
      meta.spaceName, //18
      meta.threadName //19
    ];

    appendRowToSheetWithSA_(rowData, SHEET_NAME, 'USER_ENTERED');

    if (messageName) {
      disableButtonsAndPatch_(messageName, 'お問い合わせありがとうございます❗️ \n内容を確認し、担当者よりメールにてご連絡いたします。');
    }
    return {};
  } finally {
    lock.releaseLock();
  }
}

/* ==================================================================
 * 2. Bot追加時（onAddedToSpace）— 改行説明＋画像グリッドを左側表示
 * ================================================================== */

/** 画像URLの読み込み（WELCOME_GCS_LIST: 改行 / カンマ / セミコロン / JSON配列 どれでもOK） */
function loadWelcomeUrls_() {
  const raw = (PropertiesService.getScriptProperties().getProperty('WELCOME_GCS_LIST') || '').trim();
  if (!raw) return [];
  if (raw.startsWith('[')) {
    try { return JSON.parse(raw).map(function (s) { return String(s).trim(); }).filter(Boolean); }
    catch (e) {}
  }
  return raw.split(/\r?\n|[,;]\s*/).map(function (s) { return s.trim(); }).filter(Boolean);
}

/** 画像グリッドカードを作る（画像クリックで openLink ＝ 原寸を開く） */
/** 画像グリッドカードを作る（画像クリックで openLink ＝ 原寸を開く）
 *  ウェルカムカードのみセクション間の線と内部の divider で強調
 */
function buildImageGridCard_(urls, perRow) {
  var n = Math.max(1, perRow || 3);
  var rows = [];
  for (var i = 0; i < urls.length; i += n) rows.push(urls.slice(i, i + n));

  // --- セクション構成 ---
  // 1) ご案内テキスト（最後に divider を入れて“二重の区切り”効果）
  var sections = [{
    widgets: [{
      textParagraph: {
        // \n ではなく <br> を使う（既存仕様どおり）
        text: 'このチャットボットでは生成AIに関する相談・質問を受け付けています！<br>' +
              '社内業務への活用・AIツールのエラー解決など気軽にご相談ください！<br>' +
              'メッセージ送信の数秒後に内容の最終確認が届きます(画像参照)。  回答はメールにて返信いたします。<br>' +
              'ご不明な点がございましたら、[genai@skylark.co.jp] までお問い合わせください。'
      }
    }]
  }];

  // 2) 画像ギャラリー（各行を1セクションに）
  rows.forEach(function (row, idx) {
    var section = {
      widgets: [{
        columns: {
          columnItems: row.map(function (u) {
            return {
              widgets: [{
                image: {
                  imageUrl: u,
                  altText: 'preview',
                  onClick: { openLink: { url: u } } // 画像クリックで原寸
                }
              }]
            };
          }),
          wrapStyle: 'WRAP',
          horizontalAlignment: 'CENTER'
        }
      }]
    };
    // 先頭の画像セクションにだけ見出しを付ける（線＋見出しで切れ目を強調）
    if (idx === 0) section.header = '使い方イメージ';
    sections.push(section);
  });

  return {
    cardsV2: [{
      cardId: 'welcome-gallery',
      card: {
        header: {
          title: '追加ありがとうございます！'
        },
        sections: sections
      }
    }]
  };
}

/** Botがスペースに追加されたとき（左側＝Bot名義でカードを返す） */
function onAddedToSpace(e) {
  var urls = loadWelcomeUrls_();
  var message = urls.length
    ? buildImageGridCard_(urls, 3)
    : { text: 'ようこそ！画像は現在未設定です。' };

  // 返した内容がそのまま “Botのメッセージ（左側）” として表示される
  return {
    hostAppDataAction: {
      chatDataAction: {
        createMessageAction: { message: message }
      }
    }
  };
}

/* ==================================================================
 * 3. ヘルパー群（既存のまま）
 * ================================================================== */

/** Apps Script → Chat API を Bot名義で呼ぶヘッダ */
function getHeaderWithAppCredentials() {
  const sa = getSaService_(); // 既存のSA(JWT)取得関数を使用
  return { Authorization: 'Bearer ' + sa.getAccessToken() };
}

function disableButtonsAndPatch_(messageName, statusText) {
  const current = Chat.Spaces.Messages.get(messageName, {}, getHeaderWithAppCredentials());
  const cards = JSON.parse(JSON.stringify(current.cardsV2 || [])); // deep copy

  cards.forEach(function (wrapped) {
    const card = wrapped.card || {};
    if (card.fixedFooter) delete card.fixedFooter;
    (card.sections || []).forEach(function (sec) {
      const widgets = sec.widgets || [];
      sec.widgets = widgets
        .filter(function (w) { return !w.buttonList; })
        .map(function (w) {
          if (w.decoratedText && w.decoratedText.button) delete w.decoratedText.button;
          if (w.decoratedText && w.decoratedText.buttons) delete w.decoratedText.buttons;
          return w;
        });
    });
  });

  const statusSection = { widgets: [{ textParagraph: { text: statusText } }] };
  if (cards.length === 0) {
    cards.push({ card: { sections: [statusSection] } });
  } else {
    const last = cards[cards.length - 1];
    last.card = last.card || {};
    last.card.sections = last.card.sections || [];
    last.card.sections.push(statusSection);
  }

  Chat.Spaces.Messages.patch(
    { cardsV2: cards },
    messageName,
    { updateMask: 'cardsV2' },
    getHeaderWithAppCredentials()
  );
}

function getSaEmail_() {
  const p = PropertiesService.getScriptProperties();
  return p.getProperty('SA_CLIENT_EMAIL') ||
         (function(){ try { return JSON.parse(p.getProperty('SA_JSON')||'{}').client_email; } catch(e){ return ''; } })() ||
         '(SAメール未設定)';
}

function extractMessageName_(e) {
  return (
    e?.chat?.buttonClickedPayload?.message?.name ||
    e?.chat?.messagePayload?.message?.name ||
    e?.commonEventObject?.message?.name ||
    e?.message?.name || null
  );
}

/**
 * Chat メッセージ名（messages/xxxx）から space 名と thread 名を取得
 * Advanced Google services の「Google Chat」を有効化している前提
 */
function fetchSpaceAndThreadByMessageName_(messageName) {
  try {
    if (!messageName) return { spaceName: '', threadName: '' };
    const msg = Chat.Spaces.Messages.get(messageName, {}, getHeaderWithAppCredentials());
    const spaceName  = (msg && msg.space && msg.space.name) ? String(msg.space.name) : '';
    const threadName = (msg && msg.thread && msg.thread.name) ? String(msg.thread.name) : '';
    return { spaceName, threadName };
  } catch (e) {
    console.warn('fetchSpaceAndThreadByMessageName_ failed:', e);
    return { spaceName: '', threadName: '' };
  }
}

