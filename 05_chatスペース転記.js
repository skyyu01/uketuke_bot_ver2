// ==================================================================
// 既存 doGet をルーター化：?app=review ならレビューUIへ
// ==================================================================

// ==== globals (load once per runtime) ====
const REVIEW_ACTION_TOKEN = (SCRIPT_PROPERTIES.getProperty('REVIEW_ACTION_TOKEN') || '').trim();

// Webhook 選択（URLパラメータ > DEFAULT）
function chooseWebhookUrl_(urlFromParam /*, urlFromSheetOpt */) {
  const def = (SCRIPT_PROPERTIES.getProperty('USER_DM_WEBHOOK_URL_DEFAULT') || '').trim();
  return String(urlFromParam || /*urlFromSheetOpt ||*/ def || '').trim();
}

// Webhook バリデーション
function isValidWebhookUrl(u) {
  try {
    const s = String(u || '').trim();
    if (!s) return false;
    const url = new URL(s);
    if (url.protocol !== 'https:') return false;
    const h = url.hostname.toLowerCase();
    if (h.endsWith('slack.com')) return true;
    if (h.endsWith('chat.googleapis.com')) return true;
    return false;
  } catch (_) { return false; }
}

// ==================================================================
// ルーター：既存の doGet を拡張
// ==================================================================
/**
 * google site記載用
 * インストールおよび遷移リンク
*/
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  // ★ ここでレビューUIへ振り分け
  if (params.app === 'review') {
    return reviewAppEntry_(e); // ← 下で定義（旧「レビュー用 doGet」）
  }

  // ---- ここから既存の処理（元コードをそのまま） ----
  // 1) 呼び出しユーザー ↔ このChatアプリ のDMを用意（既存があればそれを返す）
  const dm = Chat.Spaces.setup({
    space: { spaceType: 'DIRECT_MESSAGE', singleUserBotDm: true }
  });

  const spaceName = dm.name;            // 例: "spaces/1234567890"
  const spaceId   = spaceName.split('/')[1];

  // 2) 多重ログイン対策：/u/{index} をクエリで選択（デフォルト0）
  const idx = Number(params.u ? params.u : 0);
  const chatUrl = `https://mail.google.com/chat/u/${idx}/#chat/dm/${spaceId}`;

  // 3) 即リダイレクト（iframe=Sites埋め込み時にブロックされる可能性があるのでフォールバック付）
  const html = HtmlService.createHtmlOutput(
  `<!doctype html><meta charset="utf-8">
  <base target="_top">
  <style>
    body{font:14px system-ui;padding:16px;line-height:1.6}
    a.btn{display:inline-block;padding:10px 14px;border:1px solid #ccc;border-radius:8px;text-decoration:none}
    .note{margin-top:16px;padding:12px;border:1px solid #eee;border-radius:8px;background:#fffbe6}
    .note h2{margin:0 0 8px;font-size:16px}
    .note ol{margin:0;padding-left:1.4em}
  </style>
  <script>
  (function(){
    var url = ${JSON.stringify(chatUrl)};
    try {
      if (window.top === window.self) {
        window.location.replace(url); // トップで開いているとき
        return;
      }
      window.top.location.href = url;
      setTimeout(function(){ document.getElementById('fb').style.display='inline-block'; }, 400);
    } catch(e) {
      document.addEventListener('DOMContentLoaded', function(){
        document.getElementById('fb').style.display='inline-block';
      });
    }
  })();
  </script>

  <div class="note">
    <li>「Google Chat を開く」ボタンをクリックしてください。</li>
  </div>

  <p><a id="fb" class="btn" href="${chatUrl}" style="display:none">Google Chat を開く</a></p>`
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return html;
}

// ==================================================================
// レビューUI側のエントリ（旧 doGet を関数化）
// ==================================================================
function reviewAppEntry_(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const token  = params.t || '';
    const ok = token && token === REVIEW_ACTION_TOKEN;
    if (!ok) return HtmlService.createHtmlOutput('invalid token');

    const row = Number(params.row || 0);
    const action = String(params.action || 'approve');
    const userName = params.userName ? decodeURIComponent(params.userName) : '';
    const userDmWebhookUrlParam = params.userDmWebhookUrl ? decodeURIComponent(params.userDmWebhookUrl) : '';

    if (!SPREADSHEET_ID || !SHEET_NAME) {
      return HtmlService.createHtmlOutput('config error: spreadsheet id / sheet name not set');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!row || !sheet) return HtmlService.createHtmlOutput('invalid row');

    // ★ ステータスをチェックして重複実行を防止
    if (GEMINI_STATUS_COL) {
      const status = sheet.getRange(row, GEMINI_STATUS_COL).getValue();
      if (status === '承認・送付済' || status === '差戻') {
        return HtmlService.createHtmlOutput(`行${row} はすでに「${status}」です。このウィンドウを閉じてください。`);
      }
    }

    const draft = sheet.getRange(row, GEMINI_ANSWER_COL).getValue();

    // 差戻し
    if (action === 'reject') {
      if (GEMINI_STATUS_COL) sheet.getRange(row, GEMINI_STATUS_COL).setValue('差戻');
      updateCardStatus_(row, '差戻し済み');
      return HtmlService.createHtmlOutput(`行${row} を差戻しました。`);
    }

    // 修正して送付：エディタ画面を表示
    if (action === 'fix_and_approve') {
      const comment = params.comment ? decodeURIComponent(params.comment) : '';
      const initialText = String(draft || '') + (comment ? `\n\n【レビュー修正メモ】\n${comment}` : '');
      const dmSpaceNameParam  = params.dmSpaceName ? decodeURIComponent(params.dmSpaceName) : '';
      const dmThreadNameParam = params.dmThreadName ? decodeURIComponent(params.dmThreadName) : '';
      return renderEditorPage_(
        row, initialText, userName,
        { space: dmSpaceNameParam, thread: dmThreadNameParam },
        REVIEW_ACTION_TOKEN
      );
    }

    // 「通常の承認 → 即送付」
    const dmSpaceNameParam = params.dmSpaceName ? decodeURIComponent(params.dmSpaceName) : '';
    const dmThreadNameParam = params.dmThreadName ? decodeURIComponent(params.dmThreadName) : '';
    let dm = { space: dmSpaceNameParam, thread: dmThreadNameParam };
    if (!dm.space) dm = getDmFromRow_(row);  // パラメータ無ければシートから

    finalizeAndNotify_(row, String(draft || ''), dm.space, dm.thread);
    if (GEMINI_STATUS_COL) sheet.getRange(row, GEMINI_STATUS_COL).setValue('承認・送付済');
    updateCardStatus_(row, '承認・送付済み');
    return HtmlService.createHtmlOutput(`行${row} を送付しました。`);

  } catch (err) {
    console.error('review action error:', err);
    return HtmlService.createHtmlOutput('error: ' + (err && err.message));
  }
}

// ★ 新規追加：Chatメッセージを取得する内部関数
function getMessage_(messageName) {
  const url = `https://chat.googleapis.com/v1/${messageName}`;
  const token = getSaAccessToken_();
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Get message failed: ${code} ${res.getContentText()}`);
  }
  return JSON.parse(res.getContentText());
}

// ★ 変更：カードのステータス表示を更新する内部関数
function updateCardStatus_(row, statusText) {
  try {
    if (typeof CHAT_MESSAGE_NAME_COL !== 'undefined') {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
      const messageName = sheet.getRange(row, CHAT_MESSAGE_NAME_COL).getValue();
      if (messageName) {
        // Get the existing message/card object
        const message = getMessage_(messageName);
        if (!message.cardsV2 || message.cardsV2.length === 0) {
          console.error(`updateCardStatus_ failed for row ${row}: No card found in message ${messageName}`);
          return;
        }

        // Use the existing card structure
        const newCardPayload = { cardsV2: message.cardsV2 };
        const card = newCardPayload.cardsV2[0].card;

        // 1. Update the header title
        card.header.title = `レビュー依頼（${statusText}）`;

        // 2. Find and replace the '操作' section's widgets
        let operationSectionFound = false;
        for (let i = 0; i < card.sections.length; i++) {
          if (card.sections[i].header === '操作') {
            card.sections[i].widgets = [{
              textParagraph: { text: `このレビューは <b>${statusText}</b> として処理されました。` }
            }];
            operationSectionFound = true;
            break;
          }
        }

        // Fallback if '操作' section is not found for some reason
        if (!operationSectionFound) {
          card.sections.push({
            header: '状態',
            widgets: [{
              textParagraph: { text: `このレビューは <b>${statusText}</b> として処理されました。` }
            }]
          });
        }

        updateReviewCard_(messageName, newCardPayload);
      }
    }
  } catch (e) {
    console.error(`updateCardStatus_ failed for row ${row}:`, e);
  }
}

// ==================================================================
// 編集UI（Textarea版）
// ==================================================================
function renderEditorPage_(row, initialText, userName, dm, token) {
  const data = { row, userName, dm: (dm||{space:'',thread:''}), token, initialText };
  const dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>回答の修正と送付</title>
<style>
  body { font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Hiragino Sans", "Noto Sans JP", sans-serif; margin: 24px; }
  .wrap { max-width: 920px; margin: 0 auto; }
  h1 { font-size: 20px; margin: 0 0 12px; }
  textarea { width: 100%; min-height: 420px; line-height: 1.6; padding: 12px; box-sizing: border-box; }
  .row { display: flex; gap: 12px; margin-top: 12px; flex-wrap: wrap; }
  button { padding: 10px 14px; border-radius: 8px; border: 1px solid #ddd; background: #f7f7f7; cursor: pointer; }
  button.primary { background: #2563eb; color: #fff; border-color: #1d4ed8; }
  .muted { color: #666; font-size: 12px; margin-top: 6px; }
  .msg { margin-top: 14px; padding: 10px 12px; border-radius: 8px; background: #f0fdf4; border: 1px solid #86efac; display: none; }
  .err { background: #fef2f2; border-color: #fecaca; }
  .topbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
  .topbar small { color: #666; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; background: #eee; padding: 2px 6px; border-radius: 6px; border: 1px solid #ddd; }
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <h1>回答の修正と送付（行 <span id="row"></span>）</h1>
    <small>ショートカット: <span class="kbd">⌘/Ctrl + S</span> 送付、<span class="kbd">⌘/Ctrl + D</span> 下書き保存</small>
  </div>
  <div class="muted" id="meta"></div>
  <textarea id="editor" spellcheck="true"></textarea>
  <div class="row">
    <button id="send" class="primary">⏩ 修正して送付</button>
    <button id="save">💾 下書き保存（送付しない）</button>
    <button id="cancel" onclick="window.close()">閉じる</button>
  </div>
  <div id="msg" class="msg"></div>
</div>

<script>
  console.log('BOOT');

  google.script.run
    .withSuccessHandler(list => console.log('CATALOG', list))
    .withFailureHandler(e => console.log('CATALOG_FAIL', e && e.message))
    .__catalog();


  // ★ ここから通常処理
  const data = ${dataJson};
  const $ = (id)=>document.getElementById(id);

  $('row').textContent = data.row;
  $('editor').value = data.initialText || '';
  $('meta').textContent =
    (data.userName ? ('宛先: ' + data.userName + ' ／ ') : '') +
    (data.dm.space ? ('DM: ' + data.dm.space + (data.dm.thread ? ' (' + data.dm.thread + ')' : '')) : 'DM: 未設定');

  function showMsg(text, isErr){
    const el = $('msg'); el.textContent = text; el.style.display = 'block';
    if (isErr) el.classList.add('err'); else el.classList.remove('err');
  }
  function disableUI(disabled){
    $('send').disabled = disabled; $('save').disabled = disabled; $('editor').disabled = disabled;
  }

  function submit(type){
    const content = $('editor').value;
    disableUI(true);
    if (type === 'send') {
      google.script.run
      .withSuccessHandler((r)=>{ showMsg(r || '送付しました。'); setTimeout(()=>{ google.script.host.close(); }, 800); })
      .withFailureHandler((e)=>{ console.log('APPROVE_FAIL', e); showMsg('送付に失敗: ' + (e && e.message || e), true); disableUI(false); })
      .approveWithEditedAnswer( // ← アンダースコア無し
          data.row, content, data.userName, data.dm, data.token);
    } else {
      google.script.run
      .withSuccessHandler((r)=>{ showMsg(r || '下書きを保存しました。'); disableUI(false); })
      .withFailureHandler((e)=>{ showMsg('保存に失敗: ' + (e && e.message || e), true); disableUI(false); })
      .saveEditedDraft( // ← アンダースコア無し
          data.row, content, data.token);
    }
  }

  // ハンドラ登録
  document.getElementById('send').addEventListener('click', ()=>submit('send'));
  document.getElementById('save').addEventListener('click', ()=>submit('save'));
  document.addEventListener('keydown', (e)=>{
    const ctrl = e.metaKey || e.ctrlKey;
    if (!ctrl) return;
    if (e.key === 's' || e.key === 'S') { e.preventDefault(); submit('send'); }
    if (e.key === 'd' || e.key === 'D') { e.preventDefault(); submit('save'); }
  });

  // デバッグログ
  console.log('HANDLERS_READY');
  document.getElementById('send').addEventListener('click', ()=>console.log('CLICK_SEND'));
</script>

</body>
</html>
  `;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================================================================
// サーバ処理（送付／下書き保存／最終反映）
// ==================================================================
function approveWithEditedAnswer_(row, editedText, userName, dm, token) {
  if (!token || token !== REVIEW_ACTION_TOKEN) throw new Error('invalid token');
  row = Number(row);
  if (!row) throw new Error('invalid row');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('sheet not found');

  // ★ ステータスをチェックして重複実行を防止
  if (GEMINI_STATUS_COL) {
    const status = sheet.getRange(row, GEMINI_STATUS_COL).getValue();
    if (status === '承認・送付済') {
      throw new Error(`行${row} はすでに送付済みです。`);
    }
  }

  // ★ dm は {space, thread} を想定。無ければ行から取得
  let target = { space: '', thread: '' };
  if (dm && typeof dm === 'object') {
    target.space  = String(dm.space  || '').trim();
    target.thread = String(dm.thread || '').trim();
  }
  if (!target.space) {
    const fromRow = getDmFromRow_(row);
    target.space  = target.space  || fromRow.space;
    target.thread = target.thread || fromRow.thread;
  }

  // ★ シート反映 → DM（ベストエフォート）の順
  const note = finalizeAndNotify_(row, String(editedText || ''), target.space, target.thread);

  if (GEMINI_STATUS_COL) sheet.getRange(row, GEMINI_STATUS_COL).setValue('承認・送付済');

  // ★ カードを更新
  updateCardStatus_(row, '承認・送付済み');

  // UI には成功扱いで返す（DM結果は注記）
  return `行${row} を修正して送付しました。${note ? '（' + note + '）' : ''}`;
}


function saveEditedDraft_(row, editedText, token) {
  if (!token || token !== REVIEW_ACTION_TOKEN) throw new Error('invalid token');
  row = Number(row);
  if (!row) throw new Error('invalid row');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('sheet not found');

  sheet.getRange(row, GEMINI_ANSWER_COL).setValue(String(editedText || ''));
  if (GEMINI_STATUS_COL) sheet.getRange(row, GEMINI_STATUS_COL).setValue('レビュー中（修正保存）');

  return `行${row} の下書きを保存しました。`;
}

// 旧 finalizeAndNotify_ を全削除して ↓ で置き換え
function finalizeAndNotify_(row, finalAnswer, dmSpaceName, dmThreadName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  // --- 最終反映（シート）---
  if (FINAL_ANSWER_COL) sheet.getRange(row, FINAL_ANSWER_COL).setValue(finalAnswer);
  if (FAQ_CHECK_COL) sheet.getRange(row, FAQ_CHECK_COL).setValue(true); // FAQ公開はチェック
  if (GEMINI_STATUS_COL) sheet.getRange(row, GEMINI_STATUS_COL).setValue('承認・送付済');
  // ※ 「対応完了」チェックはユーザーの「はい」フィードバックを待つため、ここではセットしない

  // --- DM宛先の決定 ---
  const space = String(dmSpaceName || '').trim()
    || (DM_SPACE_NAME_COL ? String(sheet.getRange(row, DM_SPACE_NAME_COL).getValue() || '').trim() : '');
  const thread = String(dmThreadName || '').trim()
    || (CHAT_THREAD_NAME_COL ? String(sheet.getRange(row, CHAT_THREAD_NAME_COL).getValue() || '').trim() : '');

  if (!/^spaces\//.test(space)) {
    console.warn('DM skipped: space not set for row', row);
    return 'DM宛先未設定のため通知スキップ';
  }

  // --- フィードバックカードの作成 ---
  const cardBody = '【回答のご連絡】\n\n' + String(finalAnswer || '');
  const cardPayload = {
    cardsV2: [{
      cardId: 'feedback-card-' + row,
      card: {
        header: { title: '回答のご連絡' },
        sections: [
          { widgets: [{ textParagraph: { text: cardBody } }] },
          {
            header: 'この回答で問題は解決しましたか？',
            widgets: [{
              buttonList: {
                buttons: [
                  {
                    text: 'はい',
                    onClick: {
                      action: {
                        function: 'handleFeedbackAction',
                        parameters: [
                          { key: 'isResolved', value: 'true' },
                          { key: 'row', value: String(row) }
                        ]
                      }
                    }
                  },
                  {
                    text: 'いいえ',
                    onClick: {
                      action: {
                        function: 'handleFeedbackAction',
                        parameters: [
                          { key: 'isResolved', value: 'false' },
                          { key: 'row', value: String(row) }
                        ]
                      }
                    }
                  }
                ]
              }
            }]
          }
        ]
      }
    }]
  };

  // --- DM送信 ---
  try {
    sendChatAsBot_(space, null, thread, cardPayload); // テキストをnullにし、カードペイロードを渡す
    console.log('Feedback card sent to:', space, thread || '(no thread)');
    return ''; // 成功
  } catch (e) {
    console.error('DM failed but sheet committed:', e);
    return 'DM送付に失敗（ログ参照）';
  }
}




// ==================================================================
// 診断ユーティリティ
// ==================================================================
function __authTest() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}

function __debugValidateWebhook() {
  const raw = SCRIPT_PROPERTIES.getProperty('USER_DM_WEBHOOK_URL_DEFAULT');
  const trimmed = String(raw || '').trim();
  let host = '';
  try { host = new URL(trimmed).hostname; } catch (e) {}
  console.log({ raw, trimmed, host, valid: isValidWebhookUrl(raw) });
}

function sendNotificationToChat(userName, question, filelink, answer) {
  if (!CHAT_WEBHOOK_URL) {
    console.warn('Webhook URLが設定されていないため、Google Chatへの通知をスキップしました。');
    return;
  }
}

function postReviewCard_(row, userName, question, draftAnswer, spreadsheetUrl) {
  // ===== 事前チェック =====
  if (!CHAT_WEBHOOK_URL) {
    console.warn('postReviewCard_: CHAT_WEBHOOK_URL が未設定のため送信をスキップします');
    return;
  }
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.error('postReviewCard_: sheet not found');
    return;
  }

  // ===== WebアプリURL（review ルーター） =====
  const WEBAPP_BASE = (SCRIPT_PROPERTIES.getProperty('WEBAPP_URL') || '').trim();

  function buildUrl_(base, params) {
    var q = Object.keys(params)
      .filter(function(k){ return params[k] !== '' && params[k] != null; })
      .map(function(k){ return k + '=' + encodeURIComponent(String(params[k])); })
      .join('&');
    return base + (base.indexOf('?') >= 0 ? '&' : '?') + q;
  }

  // ===== 行から DM 情報（参考） =====
  const dmSpaceName  = (typeof DM_SPACE_NAME_COL === 'number' && DM_SPACE_NAME_COL > 0)
    ? String(sheet.getRange(row, DM_SPACE_NAME_COL).getValue() || '')
    : '';
  const dmThreadName = (typeof CHAT_THREAD_NAME_COL === 'number' && CHAT_THREAD_NAME_COL > 0)
    ? String(sheet.getRange(row, CHAT_THREAD_NAME_COL).getValue() || '')
    : '';

  // ===== 文字列の整形（プレーンテキスト、長文は切り詰め） =====
  function toPlain(s) {
    s = String(s || '');
    // 念のためHTMLタグを除去
    s = s.replace(/<[^>]*>/g, '');
    // 連続空白を整える（任意）
    s = s.replace(/\r\n/g, '\n').replace(/\u00a0/g, ' ');
    return s;
  }
  function truncate(s, max) {
    s = String(s || '');
    if (s.length > max) return s.slice(0, max - 1) + '…';
    return s;
  }

  const uname = toPlain(userName);
  const qText = truncate(toPlain(question), 4000);
  const aText = truncate(toPlain(draftAnswer), 7000); // カード全体サイズ対策

  // ===== レビュー用リンクを生成（WEBAPP_URL があれば） =====
  function buildActionUrl(action) {
    if (!WEBAPP_BASE) return '';
    return buildUrl_(WEBAPP_BASE, {
      app: 'review',
      t: REVIEW_ACTION_TOKEN || '',
      row: String(row),
      action: action,
      userName: uname || '',
      dmSpaceName: dmSpaceName || '',
      dmThreadName: dmThreadName || ''
    });
  }

  const urlApprove   = buildActionUrl('approve');
  const urlReject    = buildActionUrl('reject');
  const urlFixAndUI  = buildActionUrl('fix_and_approve');

  // ===== ボタン（null は混ぜない） =====
  const opButtons = [];
  if (urlApprove) opButtons.push({ text: '✅ 承認して送付', onClick: { openLink: { url: urlApprove } } });
  if (urlReject)  opButtons.push({ text: '↩︎ 差戻',         onClick: { openLink: { url: urlReject } } });
  if (urlFixAndUI)opButtons.push({ text: '✏️ 修正して送付', onClick: { openLink: { url: urlFixAndUI } } });

  const infoButtons = [];
  if (spreadsheetUrl) infoButtons.push({ text: 'スプレッドシートを開く', onClick: { openLink: { url: spreadsheetUrl } } });

  // ===== カード本体（すべてプレーンテキスト） =====
  const card = {
    text: 'レビュー依頼', // ← Fallback（必須ではないが安全）
    cardsV2: [{
      cardId: 'review-card-' + row,
      card: {
        header: {
          title: 'レビュー依頼（承認 / 差戻 / 修正）',
          subtitle: '行番号: ' + row
        },
        sections: [
          {
            header: '内容',
            widgets: [
              { decoratedText: { topLabel: '質問者', text: uname || '（不明）' } },
              { textParagraph: { text: '【質問】\n' + (qText || '（未入力）') } },
              { textParagraph: { text: '【下書き回答】\n' + (aText || '（未作成）') } }
            ].concat(infoButtons.length ? [{ buttonList: { buttons: infoButtons } }] : [])
          },
          {
            header: '操作',
            widgets: opButtons.length ? [{ buttonList: { buttons: opButtons } }] : [
              { textParagraph: { text: 'レビュー用リンクが未設定です。管理者にWEBAPP_URLの設定をご確認ください。' } }
            ]
          }
        ]
      }
    }]
  };

  // ===== 送信（Webhook → API direct call に変更）=====
  try {
    const match = String(CHAT_WEBHOOK_URL).match(/(spaces\/[^\/]+)/);
    if (!match) {
      throw new Error('postReviewCard_: Could not extract space name from CHAT_WEBHOOK_URL.');
    }
    const spaceName = match[1];
    const url = `https://chat.googleapis.com/v1/${spaceName}/messages`;
    const token = getSaAccessToken_();

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify(card),
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code < 200 || code >= 300) {
      console.error('postReviewCard_ payload(head)=', JSON.stringify({
        text: card.text,
        header: card.cardsV2[0].card.header,
      }));
      throw new Error('review card post failed: ' + code + ' ' + body);
    }

    const message = JSON.parse(body);
    const messageName = message.name;

    if (messageName && typeof CHAT_MESSAGE_NAME_COL !== 'undefined') {
        sheet.getRange(row, CHAT_MESSAGE_NAME_COL).setValue(messageName);
    }

    console.log('postReviewCard_: posted. row=', row, 'status=', code, 'messageName=', messageName);
  } catch (e) {
    console.error('postReviewCard_ error:', e);
    throw e;
  }
}

// ★ 新規追加：カード更新用関数
function updateReviewCard_(messageName, cardPayload) {
  if (!messageName) {
    console.warn('updateReviewCard_ skipped: messageName is empty');
    return;
  }

  const url = `https://chat.googleapis.com/v1/${messageName}?updateMask=cardsV2`;
  const token = getSaAccessToken_();

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'PATCH',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify(cardPayload),
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error(`Card update failed: ${code} ${res.getContentText()}`);
    }
    console.log(`Card ${messageName} updated successfully.`);
  } catch (e) {
    // メインの処理は成功しているので、カード更新の失敗はログ出力に留める
    console.error(`updateReviewCard_ failed for ${messageName}:`, e);
  }
}


// 新規追加（任意）：行から DM 宛先を読む
function getDmFromRow_(row) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const space = DM_SPACE_NAME_COL ? String(sheet.getRange(row, DM_SPACE_NAME_COL).getValue() || '').trim() : '';
  const thread = CHAT_THREAD_NAME_COL ? String(sheet.getRange(row, CHAT_THREAD_NAME_COL).getValue() || '').trim() : '';
  return { space, thread };
}


function getSaAccessToken_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sa_chat_token');
  if (cached) return cached;

  var raw = SCRIPT_PROPERTIES.getProperty('SA_JSON');
  if (!raw) throw new Error('SA_JSON is not set');
  var sa = JSON.parse(raw);

  var clientEmail = sa.client_email;
  var privateKey  = sa.private_key;
  if (!clientEmail || !privateKey) throw new Error('SA_JSON missing client_email/private_key');

  var now = Math.floor(Date.now() / 1000);
  var header = { alg: 'RS256', typ: 'JWT' };
  var claim  = {
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/chat.bot',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };
  var sub = (SCRIPT_PROPERTIES.getProperty('SA_SUBJECT') || '').trim();
  if (sub) claim.sub = sub;

  function b64url(s) {
    var bytes = Utilities.newBlob(typeof s === 'string' ? s : JSON.stringify(s)).getBytes();
    return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/,'');
  }
  var unsigned = b64url(header) + '.' + b64url(claim);
  var sig = Utilities.computeRsaSha256Signature(unsigned, privateKey);
  var jwt = unsigned + '.' + Utilities.base64EncodeWebSafe(sig).replace(/=+$/,'');

  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('SA token exchange failed: ' + code + ' ' + res.getContentText());
  var json = JSON.parse(res.getContentText());
  var token = json.access_token;
  var ttl = Math.max(60, Math.min(300, Number(json.expires_in || 300)));
  cache.put('sa_chat_token', token, Math.floor(ttl * 0.8));
  return token;
}

function sendChatAsBot_(spaceName, text, threadName, cardPayload) {
  if (!/^spaces\//.test(spaceName)) throw new Error('invalid spaceName: ' + spaceName);
  var url = 'https://chat.googleapis.com/v1/' + spaceName + '/messages';
  
  var body;
  if (cardPayload) {
    body = cardPayload;
  } else {
    body = { text: text };
  }

  if (threadName) {
    body.thread = { name: threadName };
  }

  var token = getSaAccessToken_();
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(body),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  var sc = res.getResponseCode();
  if (sc < 200 || sc >= 300) {
    console.error('sendChatAsBot_ failed with code ' + sc, res.getContentText());
    throw new Error('Chat REST failed: ' + sc + ' ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

// 公開用ラッパー（クライアントから呼ぶのはこちら）
function approveWithEditedAnswer() {
  return approveWithEditedAnswer_.apply(this, arguments);
}
function saveEditedDraft() {
  return saveEditedDraft_.apply(this, arguments);
}


function ping_(){ console.log('ping_ called'); return 'ok'; }
function __ping(){ return 'ok'; }


function __catalog() {
  try {
    return Object.getOwnPropertyNames(this)
      .filter(n => /^(approveWithEditedAnswer_?|saveEditedDraft_?|finalizeAndNotify_?|ping_?)$/.test(n))
      .sort();
  } catch (e) {
    return 'ERR:' + (e && e.message);
  }
}

/**
 * 「いいえ」が押されたときに、質問者とサポートメンバーのDMスペースを作成する
 * @param {string} questionerId - 質問者のユーザーID (例: 'users/12345')
 * @param {string} questionerName - 質問者の表示名
 * @param {number} row - スプレッドシートの行番号
 */
function createEscalationSpace(questionerId, questionerName, row) {
  // ★要設定: スクリプトプロパティにカンマ区切りでユーザーIDを保存 (例: users/123,users/456)
  const SUPPORT_MEMBER_USER_IDS = (PropertiesService.getScriptProperties().getProperty('SUPPORT_MEMBER_USER_IDS') || '')
    .split(',')
    .map(id => id.trim())
    .filter(Boolean);

  if (SUPPORT_MEMBER_USER_IDS.length === 0) {
    console.error('エスカレーション先のサポートメンバーが設定されていません (Script Property: SUPPORT_MEMBER_USER_IDS)');
    return;
  }

  try {
    // Create a space with the questioner and support members
    const members = [questionerId, ...SUPPORT_MEMBER_USER_IDS].map(id => ({ member: { name: id } }));
    
    // Using the Chat advanced service
    const space = Chat.Spaces.create({ members: members });

    // Post an initial message to the new space
    const initialMessage = `【要対応・詳細ヒアリング】\n行番号: ${row}\n質問者: ${questionerName} さん\n\nこの回答では問題が解決しなかったため、詳細なヒアリングをお願いします。`;
    sendChatAsBot_(space.name, initialMessage);

    return space.name;
  } catch (e) {
    console.error('エスカレーションスペースの作成に失敗しました。', e);
    
    // Fallback notification to a general support channel
    // ★要設定: スクリプトプロパティにフォールバック先のスペース名を保存 (例: spaces/ABCDEFG)
    const fallbackSpace = PropertiesService.getScriptProperties().getProperty('FALLBACK_SUPPORT_SPACE_NAME');
    if (fallbackSpace) {
      const fallbackMessage = `【要手動対応】\n行番号: ${row} の質問者 ${questionerName} さんとのスペース作成に失敗しました。手動でスペースを作成し、ヒアリングをお願いします。`;
      try {
        sendChatAsBot_(fallbackSpace, fallbackMessage);
      } catch (e2) {
        console.error('フォールバック通知の送信にも失敗しました。', e2);
      }
    }
    // Re-throw the original error to be caught by the calling function
    throw e;
  }
}
