/**
 * 添付保存（subjectEmail を必須指定）
 * DRIVE: DWD（ユーザーとして読み）→ SA で共有ドライブに書き込み
 * CHAT : 従来どおり media.download → SA でアップロード
 */
function saveAttachmentWithSubject_(att, subjectEmail) {
  if (!att || !att.type) return '';

  if (att.type === 'DRIVE') {
    if (!subjectEmail) throw new Error('subjectEmail is required for DRIVE attachment');
    console.log('[saveAttachmentWithSubject_] DRIVE two-token path, subject=', subjectEmail);
    return saveDriveAttachmentTwoToken_(att.fileId, att.contentName, att.mimeType, subjectEmail);
  }

  if (att.type === 'CHAT') {
    console.log('[saveAttachmentWithSubject_] CHAT media path');
    const { bytes, mimeType } = downloadFromChatWithSA_(att.resourceName);
    const saved = uploadToDriveWithSA_(bytes, att.contentName || ('attachment_' + Date.now()), mimeType);
    return saved.webViewLink;
  }

  return '';
}


/** ------------ SAトークン取得（JWT） ------------- */
function getSaService_() {
  const p = PropertiesService.getScriptProperties();

  // 1) どちらかの方式で保存：A) 個別に保存（推奨） / B) JSON丸ごと保存
  const clientEmail =
    p.getProperty('SA_CLIENT_EMAIL') ||
    (function() {
      try { return JSON.parse(p.getProperty('SA_JSON') || '{}').client_email; } catch(e) { return null; }
    })();

  let rawKey =
    p.getProperty('SA_PRIVATE_KEY') ||
    (function() {
      try { return JSON.parse(p.getProperty('SA_JSON') || '{}').private_key; } catch(e) { return null; }
    })();

  if (!clientEmail) throw new Error('SA_CLIENT_EMAIL が未設定です（または SA_JSON に client_email がありません）');
  if (!rawKey)      throw new Error('SA_PRIVATE_KEY が未設定です（または SA_JSON に private_key がありません）');

  // 2) キーの正規化
  let privateKey = String(rawKey)
    .replace(/\r\n/g, '\n')         // CRLF→LF
    .replace(/^\s*"+|"+\s*$/g, '')  // 先頭/末尾のダブルクォート除去
    .trim();

  // プロパティに \n を そのまま貼っている場合の復元
  if (privateKey.indexOf('\\n') !== -1) {
    privateKey = privateKey.replace(/\\n/g, '\n');
  }

  // 3) フォーマット検証（PKCS#8 の BEGIN/END）
  if (!/^-----BEGIN PRIVATE KEY-----\n[\s\S]+?\n-----END PRIVATE KEY-----\n?$/.test(privateKey)) {
    throw new Error('SA_PRIVATE_KEY が PEM(PKCS#8) 形式ではありません。BEGIN/END 行を含め、改行を正しく保存してください。');
  }

  const service = OAuth2.createService('sa-jwt')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setIssuer(clientEmail)
    .setPrivateKey(privateKey)
    .setScope([
      'https://www.googleapis.com/auth/chat.bot', 
      'https://www.googleapis.com/auth/chat.messages.readonly',
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/spreadsheets',
      "https://www.googleapis.com/auth/script.external_request"
    ].join(' '))
    .setPropertyStore(p);

  if (!service.hasAccess()) {
    // ライブラリ側の詳細エラーを吐く
    throw new Error(service.getLastError());
  }
  return service;
}

/** ------ Chat添付をSAでダウンロード（media.download） ------ */
function downloadFromChatWithSA_(resourceName) {
  const sa = getSaService_();
  const url = 'https://chat.googleapis.com/v1/media/' + resourceName + '?alt=media';
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + sa.getAccessToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Chat media.download failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  return {
    bytes: res.getContent(),
    mimeType: res.getHeaders()['Content-Type'] || 'application/octet-stream'
  };
}

/** ------ DriveへSAでmultipartアップロード（files.create） ------ */
function uploadToDriveWithSA_(bytes, filename, mimeType) {
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  const sa = getSaService_();

  const boundary = 'xxxxxxxxxx' + Date.now();
  const delimiter = '\r\n--' + boundary + '\r\n';
  const closeDelim = '\r\n--' + boundary + '--';

  const metadata = { name: filename, parents: [folderId] };
  const body =
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: ' + mimeType + '\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    Utilities.base64Encode(bytes) +
    closeDelim;

  const res = UrlFetchApp.fetch(
   'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink&supportsAllDrives=true',
    {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + sa.getAccessToken(),
        'Content-Type': 'multipart/related; boundary=' + boundary
      },
      payload: body,
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error('Drive upload failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  return JSON.parse(res.getContentText()); // {id, webViewLink}
}

/** SAで1行追記（A列基準） */
function appendRowToSheetWithSA_(row, sheetName, valueInputOption = 'USER_ENTERED') {
  const p = PropertiesService.getScriptProperties();
  const sheetId = p.getProperty('SPREADSHEET_ID');
  const sa = getSaService_();

  // 1) A列の最終非空行を取得（A:A を縦→横に返させて長さで判定）
  const colRange = `${quoteSheetName_(sheetName)}!A:A`;
  const getUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(colRange)}?majorDimension=COLUMNS`;
  const getRes = UrlFetchApp.fetch(getUrl, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + sa.getAccessToken() },
    muteHttpExceptions: true
  });
  if (getRes.getResponseCode() !== 200) {
    throw new Error('Sheets read failed: ' + getRes.getResponseCode() + ' ' + getRes.getContentText());
  }
  const getData = JSON.parse(getRes.getContentText());
  const colA = (getData.values && getData.values[0]) || [];
  // colA  = A列の最後に値がある行番号
  const lastRowOnA = colA.length;        // 例: ヘッダのみなら 1
  const nextRow    = lastRowOnA + 1;     // 次に書くべき行

  // 2) A{nextRow} から横に1行分を書き込む（update）
  const a1 = `${quoteSheetName_(sheetName)}!A${nextRow}`;
  const putUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(a1)}?valueInputOption=${encodeURIComponent(valueInputOption)}`;
  const payload = { range: a1, values: [row] };

  const putRes = UrlFetchApp.fetch(putUrl, {
    method: 'put',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + sa.getAccessToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = putRes.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Sheets update failed: ' + code + ' ' + putRes.getContentText());
  }

  return { updatedRange: a1, rowNumber: nextRow };
}


/** シート名をA1表記で安全にクォート */
function quoteSheetName_(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

/** ===== Embeddings (text-embedding-004) を用いた RAG インデクス ===== */
function embedText_(text) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/text-embedding-004:embedContent?key=' + GEMINI_API_KEY;
  const payload = { content: { parts: [{ text: String(text||'') }] } };
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText() || '{}');
  return (data.embedding && data.embedding.values) || [];
}

function rebuildPastCasesIndex_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow(); if (lastRow < 2) return null;
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const index = [];

  for (let i = 0; i < values.length; i++) {
    const row = i + 2;
    const user = values[i][USER_NAME_COL - 1] || '';
    const q = values[i][QUESTION_COL - 1] || '';
    const a = values[i][GEMINI_ANSWER_COL - 1] || '';
    const doc = `質問者:${user}\n質問:${q}\n回答:${a}`;
    const vec = embedText_(doc);
    index.push({ row, url: SpreadsheetApp.openById(SPREADSHEET_ID).getUrl() + '#gid=' + sheet.getSheetId() + '&range=A' + row, embedding: vec });
    Utilities.sleep(200);
  }
  const name = (PropertiesService.getScriptProperties().getProperty('RAG_INDEX_FILE_NAME') || 'past_cases_index.json');
  const it = DriveApp.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true);
  DriveApp.createFile(name, JSON.stringify(index), MimeType.JSON);
  return name;
}

function ragSearch_v3(question, k){
  const name = (PropertiesService.getScriptProperties().getProperty('RAG_INDEX_FILE_NAME') || 'past_cases_index.json');
  let it = DriveApp.getFilesByName(name); if (!it.hasNext()) return [];
  const list = JSON.parse(it.next().getBlob().getDataAsString() || '[]');
  const q = embedText_(question);

  function cosine(a,b){ let s=0,na=0,nb=0; const n=Math.min(a.length,b.length);
    for (let i=0;i<n;i++){ s+=a[i]*b[i]; na+=a[i]*a[i]; nb+=b[i]*b[i]; }
    return s/Math.sqrt((na||1)*(nb||1));
  }
  return list.map(d => ({ row:d.row, url:d.url, score: cosine(q, d.embedding||[]) }))
             .sort((x,y)=>y.score-x.score)
             .slice(0, Number(PropertiesService.getScriptProperties().getProperty('RAG_TOP_K')||k||5));
}


/** Googleネイティブの“安全な”エクスポート先を解決する */
function resolveExportForGApps_(srcMime) {
  switch (srcMime) {
    case 'application/vnd.google-apps.document':      // Google ドキュメント
      return { 
        mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
        ext: '.docx' 
        };
    case 'application/vnd.google-apps.spreadsheet':   // スプレッドシート
      return { 
        mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
        ext: '.xlsx' 
        };
    case 'application/vnd.google-apps.presentation':  // スライド
      return { 
        mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 
        ext: '.pptx' 
        };

    // ---- ここから “Drive APIでは export 非対応” ----
    case 'application/vnd.google-apps.form':          // Google フォーム
    case 'application/vnd.google-apps.site':          // Google サイト
    case 'application/vnd.google-apps.map':           // マイマップ
    case 'application/vnd.google-apps.shortcut':      // ショートカット
    case 'application/vnd.google-apps.folder':        // フォルダ
    case 'application/vnd.google-apps.drive-sdk':     // 外部アプリ
    case 'application/vnd.google-apps.script':
    case 'application/vnd.google-apps.jam':           // Jamboard
    case 'application/vnd.google-apps.drawing':       // 図形描画
      return { unsupported: true };
    default:
      return { unsupported: true };
  }
}


/** XLSX の bytes を “Google スプレッドシート”として新規作成して保存
 * @param {Byte[]} xlsxBytes
 * @param {string} nameWithoutExt  例: '受付bot_テスト用'
 * @return {{id:string, webViewLink:string}}
 */
function uploadAsGoogleSheetWithSA_(xlsxBytes, nameWithoutExt) {
  const sa = getSaService_();
  const folderId = DRIVE_FOLDER_ID;

  const boundary = 'xxxxxxxxxx' + Date.now();
  const delimiter = '\r\n--' + boundary + '\r\n';
  const closeDelim = '\r\n--' + boundary + '--';

  const metadata = {
    name: nameWithoutExt,
    mimeType: 'application/vnd.google-apps.spreadsheet', // ← ここがポイント
    parents: [folderId]
  };

  const body =
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    // メディア部は元の XLSX の MIME を付ける
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    Utilities.base64Encode(xlsxBytes) +
    closeDelim;

  const res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink',
    {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + sa.getAccessToken(),
        'Content-Type': 'multipart/related; boundary=' + boundary
      },
      payload: body,
      muteHttpExceptions: true
    }
  );

  const code = res.getResponseCode();
  if (code >= 200 && code < 300) return JSON.parse(res.getContentText());
  throw new Error('upload(convert) failed: ' + code + ' ' + res.getContentText());
}



/** デプロイユーザとして Drive ファイルを取得（メタ → bytes）
 *  DWD（なりすまし）は使いません。subjectEmail は無視されます。
 *  ※ デプロイユーザが元ファイルを閲覧できることが前提です。
 */
function downloadDriveFileAsUser_(fileId, subjectEmail /* unused */) {
  if (!fileId) throw new Error('downloadDriveFileAsUser_: fileId が空です');

  // Apps Script の実行主体（＝デプロイユーザ）の OAuth トークン
  const token = ScriptApp.getOAuthToken();
  const auth = { Authorization: 'Bearer ' + token };

  console.log(auth)

  // 1) メタ取得
  const metaUrl =
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
    '?fields=id,name,mimeType&supportsAllDrives=true';
  const m = UrlFetchApp.fetch(metaUrl, {
    method: 'get',
    headers: auth,
    muteHttpExceptions: true
  });

  if (m.getResponseCode() !== 200) {
    throw new Error('files.get(meta) failed: ' + m.getResponseCode() + ' ' + m.getContentText());
  }
  const meta = JSON.parse(m.getContentText());
  const name = meta.name || ('file_' + fileId);
  const mime = meta.mimeType || 'application/octet-stream';

  console.log('[meta]', { id: fileId, name, srcMime: mime });


  // 2) 本体
  if (String(mime).startsWith('application/vnd.google-apps')) {
    // ネイティブ → エクスポート
    const exp = resolveExportForGApps_(mime); 
    console.log('[export-plan]', { srcMime: mime, toMime: exp.mime, ext: exp.ext, unsupported: !!exp.unsupported });


    if (exp.unsupported) {
      // 400を出さずに、上位でハンドリングできるよう専用コード付きで throw
      const err = new Error('UNSUPPORTED_EXPORT: ' + mime);
      err.name = 'UnsupportedExportError';
      err.code = 'UNSUPPORTED_EXPORT';
      throw err;
    }

    const exUrl =
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
      '/export?mimeType=' + encodeURIComponent(exp.mime);
    const ex = UrlFetchApp.fetch(exUrl, {
      method: 'get',
      headers: auth,
      muteHttpExceptions: true
    });
    if (ex.getResponseCode() !== 200) {
      throw new Error('files.export failed: ' + ex.getResponseCode() + ' ' + ex.getContentText());
    }
    return {
      bytes: ex.getContent(),
      mimeType: exp.mime,
      suggestedName: name.endsWith(exp.ext) ? name : (name + exp.ext)
    };
  } else {
    // バイナリ → alt=media
    const dlUrl =
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
      '?alt=media&supportsAllDrives=true';
    const dl = UrlFetchApp.fetch(dlUrl, {
      method: 'get',
      headers: auth,
      muteHttpExceptions: true
    });
    if (dl.getResponseCode() !== 200) {
      throw new Error('files.get(media) failed: ' + dl.getResponseCode() + ' ' + dl.getContentText());
    }
    return { bytes: dl.getContent(), mimeType: mime, suggestedName: name };
  }
}



/** 二段トークン転送：ユーザーで読んで、SA で宛先に保存 */
function saveDriveAttachmentTwoToken_(fileId, fileName, srcMime, subjectEmail) {
  // 1) （本実装ではデプロイユーザ権限で）元ファイルの bytes を取得
  const { bytes, mimeType, suggestedName } = downloadDriveFileAsUser_(fileId, subjectEmail);
  const nameForSave = fileName || suggestedName || ('attachment_' + Date.now());

  // 2) ★ Excel なら Google スプレッドシートとして作成（方法B）
  if (isExcelFile_(mimeType, nameForSave)) {
    const created = uploadAsGoogleSheetWithSA_(bytes, removeExt_(nameForSave));
    return created.webViewLink;
  }

  // 3) それ以外は従来どおりそのまま保存
  const saved = uploadToDriveWithSA_(bytes, nameForSave, mimeType);
  return saved.webViewLink;
}

// --- Excel 判定＆拡張子処理 ---
function isExcelFile_(mime, name) {
  const m = String(mime || '').toLowerCase();
  const n = String(name || '').toLowerCase();
  return (
    m === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || // .xlsx
    m === 'application/vnd.ms-excel' ||                                         // .xls
    /\.xlsx?$/.test(n)                                                          // 拡張子での保険
  );
}
function removeExt_(name) {
  const s = String(name || '');
  return s.replace(/\.[^.]+$/, ''); // 末尾の拡張子を1個だけ削除
}




// 送信者ID・メールをイベントから頑健に取り出す（複数パス対応）
function extractSenderId_(e) {
  // 例: "users/115424474243460216959"
  const namePath =
    e?.sender?.name ||
    e?.commonEventObject?.user?.name ||
    e?.chat?.user?.name ||
    e?.user?.name || '';
  const m = String(namePath).match(/users\/(\d+)/);
  if (m) return m[1];
  // fallback: message から取得
  try {
    const msgName =
      e?.chat?.buttonClickedPayload?.message?.name ||
      e?.chat?.messagePayload?.message?.name ||
      e?.commonEventObject?.message?.name || null;
    if (msgName) {
      const cur = Chat.Spaces.Messages.get(msgName, {}, getHeaderWithAppCredentials());
      if (cur?.messageMetadata?.sender) return String(cur.messageMetadata.sender);
    }
  } catch (_) {}
  return '';
}

function extractSenderEmail_(e) {
  // あなたのログにある形（最優先）
  if (e?.sender?.email) return String(e.sender.email).trim().toLowerCase();
  // そのほかの Chat Apps 形式で入る可能性
  if (e?.commonEventObject?.user?.email) return String(e.commonEventObject.user.email).trim().toLowerCase();
  if (e?.chat?.user?.email) return String(e.chat.user.email).trim().toLowerCase();
  return '';
}

function normalizeEmail_(s){ return String(s||'').trim().toLowerCase(); }
function isAllowedDomain_(email){
  const allow = (PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAIL_DOMAIN')||'').toLowerCase();
  return !allow || email.endsWith('@'+allow);
}

/** SAトークンをリセット（スコープ変更後に1回実行） */
function resetSaToken() {
  const svc = getSaService_();
  svc.reset();           // OAuth2 ライブラリのトークン・キャッシュを破棄
  Logger.log('SA token has been reset.');
}



// ===== RAG（過去案件 Top-K）ユーティリティ =====
function __p(k){ return PropertiesService.getScriptProperties().getProperty(k); }
function __folder(){ var id=__p('DRIVE_FOLDER_ID'); return id ? DriveApp.getFolderById(id) : DriveApp.getRootFolder(); }
function __getSheetByGid(ssId, gid){
  var ss=SpreadsheetApp.openById(ssId);
  var sh=ss.getSheets().find(function(s){return s.getSheetId()===Number(gid)});
  return sh || ss.getActiveSheet();
}
function embedText_v2_(text){
  var arr=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text));
  return Array.from(arr).map(function(x){return (x&0xff)/255;}).slice(0,128);
}
function rebuildPastCasesIndex_v2(){
  var sheetId=__p('PAST_CASES_SHEET_ID'); var gid=Number(__p('PAST_CASES_SHEET_GID'));
  if (!sheetId || !gid) throw new Error('PAST_CASES_SHEET_ID/GID 未設定');
  var sh=__getSheetByGid(sheetId, gid);
  var lastRow=sh.getLastRow(); if (lastRow<2) return null;
  var lastCol=sh.getLastColumn();
  var headers=sh.getRange(1,1,1,lastCol).getValues()[0];
  var values =sh.getRange(2,1,lastRow-1,lastCol).getValues();

  var base=SpreadsheetApp.openById(sheetId).getUrl()+'#gid='+gid+'&range=';
  var docs=values.map(function(row,i){
    var obj={}; headers.forEach(function(h,c){ obj[String(h)]=row[c]; });
    var text=headers.map(function(h){ return h + ': ' + (obj[h]||''); }).join('\n');
    return { row:i+2, url: base+'A'+(i+2), text:text };
  });

  var index=docs.map(function(d){ return { row:d.row, url:d.url, embedding: embedText_v2_(d.text) }; });
  var name=__p('RAG_INDEX_FILE_NAME')||'past_cases_index.json';

  var it=__folder().getFilesByName(name); while(it.hasNext()) it.next().setTrashed(true);
  __folder().createFile(name, JSON.stringify(index), MimeType.PLAIN_TEXT);
  return name;
}
function ragSearch_v2(question, k){
  var name=__p('RAG_INDEX_FILE_NAME')||'past_cases_index.json';
  var it=__folder().getFilesByName(name); if(!it.hasNext()){ it=DriveApp.getFilesByName(name); if(!it.hasNext()) return []; }
  var list=JSON.parse(it.next().getBlob().getDataAsString()||'[]');
  var q=embedText_v2_(question);
  function cos(a,b){ var s=0,na=0,nb=0; for (var i=0;i<a.length;i++){ s+=a[i]*b[i]; na+=a[i]*a[i]; nb+=b[i]*b[i]; } return s/Math.sqrt(na*nb); }
  return list.map(function(d){ return { row:d.row, url:d.url, score: cos(d.embedding,q) }; })
             .sort(function(a,b){ return b.score-a.score; })
             .slice(0, Number(__p('RAG_TOP_K')||k||5));
}



function saveAnswerLogJson_(row, payload) {
  const name = `answer_log_row_${row}_${Date.now()}.json`;
  const blob = Utilities.newBlob(JSON.stringify(payload, null, 2), 'application/json', name);
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

