// ==================================================================
// 2. 回答案自動生成機能 (トリガーで実行)
// ==================================================================

/**
 * スプレッドシートをチェックし、回答案が空の質問に対してリサーチエージェントで回答を生成する関数。
 * 時間ベースのトリガーで定期的に実行します。
 */
function generateAnswers() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var lastRow = sheet.getRange('A:A').getValues().filter(String).length;
  if (lastRow < 2) {
    console.log('処理対象のデータがありません。');
    return;
  }
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var values = dataRange.getValues();
  var newQuestionFound = false;

  for (var i = 0; i < values.length; i++) {
    var userName = values[i][USER_NAME_COL - 1];
    var question = values[i][QUESTION_COL - 1];
    var answer = values[i][GEMINI_ANSWER_COL - 1];
    var status = values[i][GEMINI_STATUS_COL - 1];
    var filelink = values[i][FILE_LINK_COL - 1];
    var currentRow = i + 2;
    var isNewAndEmpty = (status === STATUS_NEW_QUESTION && !answer);
    var isTempRetry = (status === STATUS_TEMP_ERROR);

    if (question && (isNewAndEmpty || isTempRetry)) {
      newQuestionFound = true;
      console.log(currentRow + '行目の質問「' + question + '」の回答案を生成します。');
      console.log(currentRow + '行目の質問内容を分析中...');
      sheet.getRange(currentRow, GEMINI_STATUS_COL).setValue(STATUS_PROCESSING);
      SpreadsheetApp.flush();
      var analysisResult = analyzeQuestion(question);

      var confidence = Number(analysisResult.confidence || 0);
      var category = String(analysisResult.inquiry_type || '');
      var route = 'review';
      if (confidence >= AUTO_ANSWER_THRESHOLD && !NG_CATEGORIES.some(function(ng) { return category.includes(ng); })) {
        route = 'auto';
      }

      sheet.getRange(currentRow, TOOL_SERVICE_COL).setValue(analysisResult.tool_service);
      sheet.getRange(currentRow, INQUIRY_TYPE_COL).setValue(analysisResult.inquiry_type);
      sheet.getRange(currentRow, SUMMARY_COL).setValue(analysisResult.summary);
      console.log(currentRow + '行目に分析結果を書き込みました。');

      // 修正:
      var finalResult = researchAgent(question);
      // 回答セルには 600字だけ（レコードには3部構成でもOKなら出し分け）
      var finalAnswer = finalResult.final_600;

      // Google Chat 送信用も過剰な装飾を除去
      var chatMessage = String(finalAnswer || '').replace(/---/g, '');


      sheet.getRange(currentRow, GEMINI_ANSWER_COL).setValue(finalAnswer);
      sheet.getRange(currentRow, GEMINI_STATUS_COL).setValue(STATUS_ANSWER_GENERATED);

      try {
        const logPayload = {
          // 元の項目（後方互換）
          row: currentRow,
          question: question,
          analysis: analysisResult,
          queries: finalResult ? finalResult.search_query : [],
          sources: finalResult ? finalResult.web_research_result : [],
          draftAnswer: finalAnswer,
          createdAt: new Date().toISOString(),

          // ▼ ここからデバッグに効く追加入力 ▼
          // RAG内部
          used_internal_chunks: finalResult ? finalResult.used_internal_chunks : [],
          reranked_internal_top3: finalResult ? finalResult.reranked_internal_top3 : [],
          internal_refs: finalResult ? finalResult.internal_refs : [],  // 例: ["スライド 7","シート: 案件一覧"]
          citations: finalResult ? finalResult.web_refs : [],           // 外部URL 1–2件

          // 生成された要約と最終体裁（後から再現しやすい）
          final_internal_300: finalResult ? finalResult.final_internal_300 : '',
          final_web_300: finalResult ? finalResult.final_web_300 : '',
          final_600: finalResult ? finalResult.final_600 : finalAnswer,

          // 任意（調査ループのメタ）
          research_loop_count: finalResult ? finalResult.research_loop_count : undefined
        };

        const logUrl = saveAnswerLogJson_(currentRow, logPayload);
        if (LOG_LINK_COL) sheet.getRange(currentRow, LOG_LINK_COL).setValue(logUrl);
      } catch (e) {
        console.warn('回答ログ保存に失敗: row=' + currentRow, e);
      }

      var spreadsheetUrl = SpreadsheetApp.openById(SPREADSHEET_ID).getUrl();
      var dmSpaceName = values[i][DM_SPACE_NAME_COL - 1];

      if (route === 'auto') {
        const dm = getDmFromRow_(currentRow);
        finalizeAndNotify_(currentRow, finalAnswer, dm.space, dm.thread);
        sheet.getRange(currentRow, GEMINI_STATUS_COL).setValue('自動送付済');
        console.log("自動回答を行いました");
      } else {
        postReviewCard_(currentRow, userName, question, chatMessage, spreadsheetUrl, dmSpaceName);
      }

      console.log(currentRow + '行目に回答案を書き込みました。');
      sendNotificationToChat(userName, question, filelink, chatMessage);

    }
  }
  if (!newQuestionFound) {
    console.log('処理対象の新規質問はありませんでした。');
  }
}

// ==================================================================
// 3. リサーチエージェント (RAG対応)
// ==================================================================

// --- RAG実装のためのヘルパー関数群 ---

/**
 * 質問から検索キーワードを抽出する
 */
function generateSearchKeywords_(question) {
  try {
    var prompt = '以下の質問内容から、社内ドキュメントを検索するための最も重要なキーワードを5つ、JSON配列形式で抽出してください。\n' +
                 'キーワードは、具体的で検索に適した名詞や専門用語を選んでください。\n\n' +
                 '質問: "' + question + '"\n\n' +
                 '例: {"keywords": ["キーワード1", "キーワード2"]}';
    var payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: "application/json", temperature: 0 }
    };
    var response = callGeminiApi(GEMINI_API_URL_FLASH, payload);
    var result = JSON.parse(response.candidates[0].content.parts[0].text);
    return result.keywords || [];
  } catch (e) {
    console.error('キーワード抽出に失敗:', e);
    return [];
  }
}

/**
 * テキストを意味のあるチャンクに分割する
 */
function chunkText_(text) {
  if (!text) return [];
  // (^|\n) を追加して先頭の見出しも拾う
  var splitter = /(?:^|\n)--- (?:スライド \d+|シート: [^\n]+) ---\n/;
  var parts = text.split(splitter);
  // 見出し自体も抽出（先頭も拾う）
  var headers = text.match(/(?:^|\n)--- (?:スライド \d+|シート: [^\n]+) ---\n/g) || [];

  var chunks = [];
  // parts[0] は splitter の前置（多くは空）。以降を headers とペアにする
  for (var i = 1; i < parts.length; i++) {
    var header = headers[i - 1].replace(/^\n/, '').trim(); // 先頭改行を除去
    var body = parts[i].trim();
    if (body) chunks.push(header + '\n' + body);
  }
  return chunks;
}

/**
 * キーワードに基づいて関連チャンクを検索・スコアリングする
 */
function findRelevantChunks_(chunks, keywords, topK) {
  topK = topK || 3;
  if (!chunks.length) return [];
  var qToks = tokenize_(((keywords||[]).join(' ')));
  var idf = computeIdf_(chunks);
  var lengths = chunks.map(function(c){ return tokenize_(c).length; });
  var avgLen = lengths.reduce(function(a,b){return a+b;},0)/lengths.length;

  var scored = chunks.map(function(c, idx){
    return { idx: idx, chunk: c, score: bm25Score_(c, qToks, idf, 1.5, 0.75, avgLen) };
  }).filter(function(o){ return o.score>0; });

  scored.sort(function(a,b){ return b.score - a.score; });
  // まずTop20までに拡げて後段のLLM再ランクへ
  return scored.slice(0, Math.max(topK, 20));
}


// --- Googleドキュメントからテキストを抽出するヘルパー関数群 ---

function getTextFromGoogleSlide_(slideUrl) {
  try {
    var m = slideUrl.match(/presentation\/d\/([a-zA-Z0-9-_]+)/);
    if (!m) return '';
    var p = SlidesApp.openById(m[1]);
    var out = [];
    p.getSlides().forEach(function(slide, idx){
      var title = '';
      var body  = [];
      slide.getShapes().forEach(function(sh){
        if (!sh.getText) return;
        var t = sh.getText().asString().trim();
        if (!t) return;
        if (!title && (sh.getShapeType && sh.getShapeType() === SlidesApp.ShapeType.TITLE)) title = t;
        else body.push(t);
      });
      var chunk = [
        '--- スライド ' + (idx+1) + ' ---',
        '【タイトル】' + (title || '(無題)'),
        body.join('\n')
      ].join('\n');
      out.push(chunk);
    });
    return out.join('\n\n');
  } catch(e){ 
    console.error('Slide抽出失敗:', slideUrl, e); 
    return '（スライド取得失敗）';
  }
}

function getTextFromGoogleSheet_(sheetUrl) {
  try {
    var m = sheetUrl.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!m) return '';
    var ss = SpreadsheetApp.openById(m[1]);

    // ★ スプレッドシート（ファイル）の名前で分岐
    var isSpecialFile = (typeof ss.getName === 'function') && (ss.getName() === '案件別要件整理シート');

    // 走査対象シートを決定
    var targetSheets;
    if (isSpecialFile) {
      var only = ss.getSheetByName('対応案件表');
      if (!only) {
        // 期待するシートが無い場合はわかるように返す（運用に応じて空文字にしてもOK）
        return '（対象スプレッドシートは「案件別要件整理シート」ですが、シート「対応案件表」が見つかりません）';
      }
      targetSheets = [only];
    } else {
      targetSheets = ss.getSheets(); // 従来動作：全シート
    }

    var out = [];
    targetSheets.forEach(function(sh){
      var values = sh.getDataRange().getValues();
      if (!values.length) return;
      var header = values[0].join(' | ');
      var rows = values.slice(1).map(function(r,i){
        return (i+2)+': '+ r.join(' | ');
      }).join('\n');
      out.push('--- シート: ' + sh.getName() + ' ---\n【ヘッダ】' + header + '\n' + rows);
    });

    return out.join('\n\n');
  } catch(e){
    console.error('Sheet抽出失敗:', sheetUrl, e);
    return '（シート取得失敗）';
  }
}

function getTextFromGoogleDoc_(docUrl) {
  try {
    var documentIdMatch = docUrl.match(/document\/d\/([a-zA-Z0-9-_]+)/);
    if (!documentIdMatch || !documentIdMatch[1]) {
      console.log('無効なGoogle Document URLです: ' + docUrl);
      return '';
    }
    var documentId = documentIdMatch[1];
    var doc = DocumentApp.openById(documentId);
    return doc.getBody().getText();
  } catch (e) {
    console.error('Google Documentからのテキスト抽出に失敗しました: ' + docUrl, e);
    return '（ドキュメント「' + docUrl + '」の読み込みに失敗しました）';
  }
}

/**
 * 質問に基づき、Web検索と複数回の思考を経て回答を生成するエージェント
 */
/**
 * 質問に基づき、社内資料（Slides/Sheets/Docs）と外部Webの要点を集約し、
 * 300字×2（社内/外部）→最終600字回答を合成して返す。
 * - 社内資料は BM25風の予選（Top20相当）→ LLM再ランク（Top3）で精度を上げる
 * - Webは複数クエリ・リフレクションで必要十分性を判断
 * - 出典（内部: スライド番号/シート名、外部: URL）を 1–2件だけ末尾に併記
 * - 返却値 state にデバッグ用フィールド（used_internal_chunks など）を含める
 */
function researchAgent(initialQuestion) {
  // ========= 基本設定 =========
  var config = {
    initial_search_query_count: 3, // 生成する初期クエリ数
    max_research_loops: 2,         // Webリサーチの最大ループ
    max_total_queries: 10          // Web検索の総クエリ上限
  };

  // ========= ステート =========
  var state = {
    messages: [{ role: "user", content: initialQuestion }],
    search_query: [],
    web_research_result: [],   // 後方互換: 文字列 or オブジェクトの text を格納
    web_citations: [],         // {url, title?} の配列
    sources_gathered: [],
    research_loop_count: 0,
    number_of_ran_queries: 0,
    used_internal_chunks: [],  // 予選に使った内部チャンク（スコア付き）
    reranked_internal_top3: [] // 最終的に採用したTop3チャンク
  };

  // ========= 社内資料（RAG） =========
  var INTERNAL_GUIDE_URLS = (PropertiesService.getScriptProperties().getProperty('INTERNAL_GUIDE_URLS') || '')
    .split(',')
    .map(function (s) { return s.trim(); })
    .filter(Boolean);

  var internalSummary = '';      // 社内300字
  var internalRefs = [];         // “スライド7”や“シート:案件一覧”など
  var prelim = [];               // BM25風の予選（Top20相当のスコア付き想定）

  if (INTERNAL_GUIDE_URLS.length > 0) {
    var allGuideText = '';
    try {
      var slideUrls = INTERNAL_GUIDE_URLS.filter(function (url) { return url.includes('docs.google.com/presentation/d/'); });
      var sheetUrls = INTERNAL_GUIDE_URLS.filter(function (url) { return url.includes('docs.google.com/spreadsheets/d/'); });
      console.log("csvファイルがあるか",sheetUrls)
      var docUrls   = INTERNAL_GUIDE_URLS.filter(function (url) { return url.includes('docs.google.com/document/d/'); });

      if (slideUrls.length) allGuideText += slideUrls.map(getTextFromGoogleSlide_).join('\n\n');
      if (sheetUrls.length) allGuideText += (allGuideText ? '\n\n' : '') + sheetUrls.map(getTextFromGoogleSheet_).join('\n\n');
      console.log("csvみたい", sheetUrls.map(getTextFromGoogleSheet_).join('\n\n'))
      if (docUrls.length)   allGuideText += (allGuideText ? '\n\n' : '') + docUrls.map(getTextFromGoogleDoc_).join('\n\n');
    } catch (e) {
      console.warn('社内資料の取得中にエラー:', e);
    }

    if (allGuideText && allGuideText.trim()) {
      // 1) キーワード生成
      var keywords = [];
      try {
        keywords = generateSearchKeywords_(initialQuestion) || [];
      } catch (e) {
        console.warn('キーワード抽出失敗: フォールバックで無視', e);
        keywords = [];
      }
      console.log('抽出された検索キーワード: ' + (keywords.join(', ') || '(なし)'));

      // 2) チャンク化（スライド/シート単位のチャンクを想定）
      var chunks = [];
      try {
        chunks = chunkText_(allGuideText) || [];
      } catch (e) {
        console.warn('チャンク化失敗: ', e);
        chunks = [];
      }

      // 3) BM25風の予選（Top20程度まで拡張して返すよう findRelevantChunks_ を利用）
      try {
        var prelimCandidates = findRelevantChunks_(chunks, keywords, 3 /* topK引数は内部でTop20へ拡張を推奨 */) || [];
        // 後方互換（古い findRelevantChunks_ が「文字列配列」を返す場合に対応）
        prelim = prelimCandidates.map(function (c) {
          if (typeof c === 'string') {
            return { idx: -1, chunk: c, score: 1.0 };
          } else {
            // {idx, chunk, score} を想定
            return {
              idx: (typeof c.idx === 'number' ? c.idx : -1),
              chunk: c.chunk || c.text || '',
              score: (typeof c.score === 'number' ? c.score : 1.0)
            };
          }
        });
        // ログ・デバッグ用に保持
        state.used_internal_chunks = prelim.slice(0, 20);
      } catch (e) {
        console.warn('BM25風予選（findRelevantChunks_）でエラー:', e);
        prelim = [];
      }

      // 4) LLM再ランク（Top20 → Top3）
      var top3 = [];
      console.log("20個の候補",prelim)
      try {
        top3 = rerankWithLLM_(initialQuestion, prelim) || [];
        console.log("最終候補",top3)
      } catch (e) {
        console.warn('再ランク失敗→予選上位を流用:', e);
        top3 = prelim.slice(0, 3).map(function (o) { return o.chunk; });
      }
      state.reranked_internal_top3 = top3.slice();

      // 5) Top3から社内300字要約を作成（出典を末尾に併記）
      if (top3.length) {
        // ① 社内300字要約（既存）
        internalSummary = summarizeToLength_(
          top3.join('\n\n'),
          300,
          '社内資料の記述のみ。数字・条件は原文通り'
        );

        // ② ▼ ここから“出典の最低限の表示”を追加 ▼
        //   各チャンク先頭のヘッダから「スライド番号 or シート名」を抽出
        //   （chunkText_ が '--- スライド n ---' / '--- シート: name ---' を付けている前提）
        var internalRefs = top3.map(function (c) {
          var m = c.match(/^--- (スライド [\d]+|シート: [^\n]+) ---/m);
          return m ? m[1] : null;
        }).filter(Boolean);

        // 末尾に 1〜2件だけ併記
        if (internalRefs.length) {
          internalSummary += '\n(参考: ' + internalRefs.slice(0, 2).join(', ') + ')';
        }
        // 呼び出し元でも使えるよう state に保存（返却用）
        state.internal_refs = internalRefs.slice(0, 2);
        // ② ▲ ここまで追加 ▲

        console.log('社内ガイド Top3 → 300字要約を作成');
      } else {
        console.log('社内ガイド内に関連チャンクは見つかりませんでした。');
      }
    } else {
      console.log('社内ガイドの本文テキストが空でした。');
    }
  } else {
    console.log('INTERNAL_GUIDE_URLS が未設定または空です。');
  }

  // ========= Webリサーチ =========
  var seenQueries = [];
  var seenDomains = [];
  function domainOf(u) { try { return (new URL(u)).hostname.replace(/^www\./, ''); } catch (e) { return ''; } }

  // 初期クエリ生成
  try {
    var initialQueries = (generateQuery(state, config) || [])
      .filter(function (q) { return q && seenQueries.indexOf(q) === -1; })
      .slice(0, config.initial_search_query_count);

    for (var k = 0; k < initialQueries.length; k++) seenQueries.push(initialQueries[k]);
    state.search_query = state.search_query.concat(initialQueries);
    state.number_of_ran_queries += initialQueries.length;
    console.log('生成された検索クエリ: ' + initialQueries.join(', '));
  } catch (e) {
    console.warn('検索クエリ生成で失敗（generateQuery）:', e);
  }

  // ループ調査
  for (var loop = 0; loop < config.max_research_loops; loop++) {
    state.research_loop_count = loop + 1;
    if (seenQueries.length >= config.max_total_queries) break;

    var searchResults = [];
    var currentQueries = state.search_query.slice(-state.number_of_ran_queries);

    for (var l = 0; l < currentQueries.length; l++) {
      var query = currentQueries[l];
      console.log('ウェブ調査中 (' + (loop + 1) + '/' + config.max_research_loops + '): "' + query + '"');
      try {
        var researchResult = webResearch(query);

        // 後方互換: webResearch が文字列を返す場合に対応
        if (typeof researchResult === 'string') {
          searchResults.push(researchResult);
          state.web_research_result.push(researchResult);
        } else if (researchResult && typeof researchResult === 'object') {
          var txt = researchResult.text || '';
          var cits = Array.isArray(researchResult.citations) ? researchResult.citations : [];
          searchResults.push(txt);
          state.web_research_result.push(txt);
          cits.forEach(function (c) {
            if (!c || !c.url) return;
            var d = domainOf(c.url);
            if (d && seenDomains.indexOf(d) === -1) seenDomains.push(d);
            state.web_citations.push(c);
          });
        } else {
          // 形式不明
          var fallback = String(researchResult || '（検索結果なし）');
          searchResults.push(fallback);
          state.web_research_result.push(fallback);
        }
      } catch (e) {
        console.warn('webResearch 実行中にエラー:', e);
      }
    }

    console.log("調査結果を評価中...");
    var reflectionResult = null;
    try {
      reflectionResult = reflection(state);
    } catch (e) {
      console.warn('reflection の実行に失敗。十分と見なして打ち切り:', e);
      reflectionResult = { is_sufficient: true, knowledge_gap: "", follow_up_queries: [] };
    }

    // 追加クエリ生成
    var nextQs = ((reflectionResult && reflectionResult.follow_up_queries) || [])
      .filter(function (q) { return q && seenQueries.indexOf(q) === -1; })
      .slice(0, 3);

    // 同一ドメインに偏らないよう、簡易ペナルティを文字列レベルで付与
    var penalize = seenDomains.slice(0, 2).map(function (d) { return '-site:' + d; }).join(' ');
    var diversified = nextQs.map(function (q) { return penalize ? (q + ' ' + penalize) : q; });

    for (var m = 0; m < diversified.length; m++) {
      if (seenQueries.length < config.max_total_queries) {
        state.search_query.push(diversified[m]);
        seenQueries.push(diversified[m]);
      }
    }

    if (reflectionResult && reflectionResult.is_sufficient) {
      console.log("情報が十分であると判断。最終回答を生成します。");
      break;
    } else {
      var gap = reflectionResult ? reflectionResult.knowledge_gap : '(不明)';
      var fuq = reflectionResult ? reflectionResult.follow_up_queries : [];
      console.log('知識ギャップを検出: ' + gap);
      console.log('追加のクエリを生成: ' + fuq.join(', '));
      state.search_query = state.search_query.concat(fuq);
      state.number_of_ran_queries = fuq.length;
    }

    if (loop === config.max_research_loops - 1) {
      console.log("最大ループ回数に達しました。");
    }
  }

  // ========= Web300字 要約（参考URLを1件だけ付与） =========
  var webSummary = '';
  try {
    var webTextJoined = '';
    if (state.web_research_result && state.web_research_result.length) {
      // 文字列配列を想定。オブジェクト対応は上で text を push 済み。
      webTextJoined = state.web_research_result.join('\n\n');
    }
    // 代表URLを1件だけ選ぶ（先頭の citation を優先）
    var refUrl = '';
    if (state.web_citations && state.web_citations.length) {
      refUrl = state.web_citations[0].url || '';
    } else {
      // citationsが空の時は、テキスト内からURLらしきものを拾う簡易フォールバック
      var m = webTextJoined.match(/https?:\/\/[^\s\)\]]+/);
      refUrl = m ? m[0] : '';
    }

    webSummary = summarizeToLength_(
      webTextJoined + (refUrl ? ('\n(参考:' + refUrl + ')') : ''),
      300,
      '本文以外の一般論は書かない'
    );
  } catch (e) {
    console.warn('Web300字要約でエラー:', e);
    webSummary = '';
  }

  // ========= 参考：過去案件（ragSearch_v3/v2 があれば添付） =========
  try {
    var ragHits = (typeof ragSearch_v3 === 'function')
      ? ragSearch_v3(state.messages[0].content, Number(PropertiesService.getScriptProperties().getProperty('RAG_TOP_K') || 5))
      : ((typeof ragSearch_v2 === 'function')
        ? ragSearch_v2(state.messages[0].content, Number(PropertiesService.getScriptProperties().getProperty('RAG_TOP_K') || 5))
        : []);
    if (Array.isArray(ragHits) && ragHits.length) {
      var appendix = '\n\n---\n【類似事例（過去案件）】\n' + ragHits.map(function (r, i) { return (i + 1) + '. ' + r.url; }).join('\n');
      state.messages.push({ role: "model", content: appendix });
    }
  } catch (e) {
    console.warn('RAG付与に失敗しました（処理は継続）:', e);
  }

  // ========= 最終600字の合成 =========
  console.log("最終回答を生成中...");
  var final600 = '';
  try {
    final600 = finalizeAnswerWithSections_(state.messages[0].content, internalSummary, webSummary);
  } catch (e) {
    console.warn('最終合成でエラー。空回答にフォールバック:', e);
    final600 = '・現時点の社内/外部素材からは根拠が不足しています。\n・質問の前提(対象ツール/期間/部署など)を補足してください。';
  }

  // ========= メッセージ配列にも残す（互換） =========
  state.messages.push({
    role: "model",
    content:
      (internalSummary ? '【社内300字】\n' + internalSummary + '\n\n' : '') +
      (webSummary ? '【外部300字】\n' + webSummary + '\n\n' : '') +
      '【最終回答(600字以内)】\n' + final600
  });

  // ========= 呼び出し元で直接使えるフィールド =========
  state.final_internal_300 = internalSummary;
  state.final_web_300 = webSummary;
  state.final_600 = final600;

  // 代表出典も返すと便利（オプション）
  state.internal_refs = internalRefs.slice(0, 2);
  state.web_refs = (state.web_citations.slice(0, 2).map(function (c) { return c.url; }));

  return state;
}



// --- リサーチエージェントのヘルパー関数群 ---

function braveWebSearch_(query, count) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('BRAVE_API_KEY');
  if (!apiKey) throw new Error('BRAVE_API_KEY not set');
  var url = 'https://api.search.brave.com/res/v1/web/search'
          + '?q=' + encodeURIComponent(query)
          + '&country=JP&search_lang=ja&count=' + (count || 5);
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Accept': 'application/json', 'X-Subscription-Token': apiKey },
    muteHttpExceptions: true
  });
  var json = JSON.parse(res.getContentText() || '{}');
  var results = (json.web && json.web.results) ? json.web.results : [];
  var seen = {};
  return results.filter(function(r){
    var u=(r.url||'').replace(/[?#].*$/,'');
    if(seen[u]) return false; seen[u]=1; return true;
  }).map(function(r){ return { title:r.title, url:r.url, snippet:r.description||'' }; });
}

function callGeminiApiWithUrlContext_(modelUrl, prompt, urls) {
  var text = prompt;
  if (urls && urls.length) {
    text += '\n\n参照URL:\n' + urls.slice(0, 20).join('\n');
  }
  var payload = {
    contents: [{ parts: [{ text: text }] }],
    generationConfig: { temperature: 0.0 },
    tools: [{ "url_context": {} }]
  };
  return callGeminiApi(modelUrl, payload);
}

function generateQuery(state, config) {
  var prompt = 'あなたは高度なリサーチアシスタントです。ユーザーの質問に基づいて、ウェブ検索に最適な、具体的で多様な検索クエリを生成してください。現在の日は' + new Date().toLocaleDateString('ja-JP') + 'です。\n\n質問: "' + state.messages[0].content + '"\n\nJSON形式で、"query"というキーに' + config.initial_search_query_count + '個の検索クエリのリストを持たせてください。例: {"query": ["クエリ1", "クエリ2"]}';
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" }
  };
  var response = callGeminiApi(GEMINI_API_URL_FLASH, payload);
  try {
    return JSON.parse(response.candidates[0].content.parts[0].text).query;
  } catch (e) {
    console.error("generateQuery JSONパースエラー:", response.candidates[0].content.parts[0].text);
    throw new Error("検索クエリ生成失敗");
  }
}

function webResearch(searchQuery, allowedUrls) {
  if (allowedUrls && allowedUrls.length) {
    var prompt = '次のトピックについて、一次ソース本文（URL-Context）を直接参照し、' +
                 '最新・信頼できる内容だけを要点化してください。根拠URLも併記してください。\n\n' +
                 'トピック: ' + searchQuery;
    var resp = callGeminiApiWithUrlContext_(GEMINI_API_URL_FLASH, prompt, allowedUrls.slice(0,8));
    return resp.candidates[0].content.parts[0].text || '（検索結果なし）';
  }
  var prompts = '「' + searchQuery + '」についてGoogle検索で最新かつ信頼できる情報を収集し、要点をまとめてください。';
  var payload = {
    contents: [{ parts: [{ text: prompts }] }],
    tools: [{ "google_search": {} }],
    generationConfig: { temperature: 0.0 }
  };
  var response = callGeminiApi(GEMINI_API_URL_FLASH, payload);
  return response.candidates[0].content.parts[0].text || "（検索結果なし）";
}

function reflection(state) {
  var prompt = 'あなたはリサーチ内容の評価者です。以下のユーザーの質問と、それに対して収集した情報の要約を読んで、回答するのに情報が十分か、それとも不足しているかを判断してください。\n\n質問: "' + state.messages[0].content + '"\n\n収集した情報の要約:\n---\n' + state.web_research_result.join("\n\n") + '\n---\n\n判断結果をJSON形式で、"is_sufficient"(boolean)、"knowledge_gap"(string)、"follow_up_queries"(string[])のキーで返してください。情報が十分な場合、knowledge_gapとfollow_up_queriesは空にしてください。';
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" }
  };
  var response = callGeminiApi(GEMINI_API_URL, payload);
  try {
    return JSON.parse(response.candidates[0].content.parts[0].text);
  } catch (e) {
    console.error("reflection JSONパースエラー:", response.candidates[0].content.parts[0].text);
    return { is_sufficient: true, knowledge_gap: "", follow_up_queries: [] };
  }
}


function finalizeAnswerWithSections_(question, internalSummary, webSummary) {
  var sys =
    'あなたは優秀なアシスタントです。\n' +
    '以下の二つの素材（社内要約300字/外部要約300字）だけを根拠に、' +
    '重複を統合して日本語で最終回答を600字以内で作成してください。\n' +
    '・見出しは付けず、箇条書き主体\n' +
    '・社内情報を優先、その後に外部情報\n' +
    '・根拠URLはあれば1〜2個だけ末尾に(参考: …)で併記\n' +
    '・推測は禁止、素材に無い情報を足さない';

  var user =
    '# 質問\n' + question + '\n\n' +
    '# 社内300字\n' + (internalSummary || '') + '\n\n' +
    '# 外部300字\n' + (webSummary || '');

  var payload = {
    system_instruction: { parts: [{ text: sys }] },
    contents: [{ parts: [{ text: user }] }],
    generationConfig: { temperature: 0.0 }
  };
  var resp = callGeminiApi(GEMINI_API_URL, payload);
  var out = (resp.candidates && resp.candidates[0] && resp.candidates[0].content.parts[0].text) || '';
  return smartTrim_(out, 600);
}


function callGeminiApi(apiUrl, payload) {
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
  if (responseCode === 200) {
    return JSON.parse(responseBody);
  } else {
    console.error('API Error: ' + responseCode + ' ' + responseBody);
    throw new Error('Gemini APIリクエスト失敗 (Status: ' + responseCode + ')');
  }
}

// ==================================================================
// 5. 質問分析機能 (カテゴリ分類・要約)
// ==================================================================

function analyzeQuestion(question) {
  var prompt = '# 指示\n' +
    'あなたは「生成AI相談窓口」の担当者です。以下の質問内容を分析し、次の3つのタスクを実行してください。\n\n' +
    '1.  **ツール・サービスの分類**: 相談内容がどのAIツールに関するものか、以下のリストから最も近いものを1つ選択してください。\n' +
    '# ツール・サービス リスト\n' +
    '- ' + CATEGORY_TOOL_SERVICE.join('\n- ') + '\n\n' +
    '2.  **問い合わせ内容の分類**: 相談者が何を知りたいのか、以下のリストから最も近いものを1つ選択してください。\n' +
    '# 問い合わせ内容 リスト\n' +
    '- ' + CATEGORY_INQUIRY_TYPE.join('\n- ') + '\n\n' +
    '3.  **要約**: 相談内容の要点を、目的が明確にわかるように100文字程度で要約してください。要約はFAQとして公開される可能性があるため、個人名や部署名などの個人・組織情報は含めないでください。\n\n' +
    '# 出力形式\n' +
    '結果は、必ず指示されたスキーマのJSONオブジェクトで返してください。\n\n' +
    '---\n' +
    '# 質問内容:\n' +
    question +
    '\n---';

  var payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          tool_service: { "type": "STRING", "enum": CATEGORY_TOOL_SERVICE },
          inquiry_type: { "type": "STRING", "enum": CATEGORY_INQUIRY_TYPE },
          summary: { "type": "STRING" }
        },
        required: ["tool_service", "inquiry_type", "summary"]
      }
    }
  };

  try {
    var response = callGeminiApi(GEMINI_API_URL_FLASH, payload);
    var resultText = response.candidates[0].content.parts[0].text;
    return JSON.parse(resultText);
  } catch (e) {
    console.error('質問内容の分析中にエラーが発生しました: ' + e.message);
    throw new Error('質問分析APIの呼び出しまたは結果の解析に失敗しました。(' + e.message + ')');
  }
}


// === 共通: 文字数制御ユーティリティ ===
// 文末優先で「だいたい600字」に収める（途中で切らない）
// ・targetChars 付近〜+flex の範囲で、最初に現れる文末（。．！？!? または改行）で切る
// ・見つからなければ targetChars より手前の最後の文末で切る
// ・それも無ければ最大長（targetChars + flex）で切る
function smartTrim_(s, targetChars, opts) {
  s = String(s || '').trim();
  if (!s) return '';

  // デフォルト設定：目標600字、上ぶれ最大+200字、手前側は-80字まではOK
  var o = opts || {};
  var flex = typeof o.flex === 'number' ? o.flex : 200;    // どこまで上ぶれ許容するか
  var back = typeof o.back === 'number' ? o.back : 80;     // どこまで手前に戻ってよいか
  var min = typeof o.min === 'number' ? o.min : 520;       // 極端に短くならないよう下限を目安化

  if (s.length <= targetChars + flex) return s; // 目標+許容内ならそのまま返す

  var startAfter = Math.max(0, targetChars - back);
  var endBefore  = Math.min(s.length, targetChars + flex);

  // 文末候補の区切り（句点・感嘆・疑問・改行）
  var delimiter = /[。．！!？?\n]/g;

  // ① 目標の少し手前から+flexの範囲内で「先に現れる」文末を探す（途中切断を避けたいので優先）
  delimiter.lastIndex = startAfter;
  var cutIdx = -1, m;
  while ((m = delimiter.exec(s)) !== null) {
    var idx = m.index;
    if (idx >= startAfter && idx <= endBefore) { cutIdx = idx + 1; break; }
    if (idx > endBefore) break;
  }

  // ② 見つからなければ、目標直前までで「最後に現れた」文末を探す
  if (cutIdx === -1) {
    var lastBefore = -1;
    delimiter.lastIndex = 0;
    while ((m = delimiter.exec(s)) !== null) {
      if (m.index < targetChars) lastBefore = m.index + 1;
      else break;
    }
    if (lastBefore !== -1 && lastBefore >= Math.min(min, targetChars - back)) cutIdx = lastBefore;
  }

  // ③ まだ無ければ、強制的に target+flex で切る（途中切断の可能性はあるが最小限）
  if (cutIdx === -1) cutIdx = endBefore;

  // 末尾の中途記号や空白を整理（…等の付与はしない：文を切らない方針）
  var out = s.slice(0, cutIdx).replace(/\s+$/,'');
  return out;
}

// 既存互換ラッパー：厳密カットはやめ、文末優先のソフト制限に委譲
function hardTrim_(s, maxChars) {
  return smartTrim_(s, maxChars || 600, { flex: 200, back: 80, min: 520 });
}

// 厳密に600文字できりたいときはこっちを使用
function hardTrim_(s, maxChars) {
  s = String(s || '').trim();
  if (s.length <= maxChars) return s;
  return s.slice(0, Math.max(0, maxChars - 1)) + '…';
}

// maxChars 文字以内 / 箇条書き最大5点 / 余計な前置きや末尾注意書き禁止
function summarizeToLength_(text, maxChars, extraInstruction) {
  if (!text || !String(text).trim()) return '';
  var prompt =
    '以下の内容を日本語で' + maxChars + '文字以内に厳密に要約してください。' +
    '・重要ポイントのみ、箇条書き(最大5点)\n' +
    '・事実のみ。推測や一般論を入れない\n' +
    '・前置き/まとめ/注意書きは書かない\n' +
    (extraInstruction ? ('・' + extraInstruction + '\n') : '') +
    '\n---\n' + String(text).trim();

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.0 }
  };
  var resp = callGeminiApi(GEMINI_API_URL_FLASH, payload);
  var out = (resp.candidates && resp.candidates[0] && resp.candidates[0].content.parts[0].text) || '';
  // 念のためハードカット
  return hardTrim_(out.replace(/\n{3,}/g, '\n\n').trim(), maxChars);
}

function normalizeJa_(s) {
  if (!s) return '';
  // 全角→半角 / ひら→カタ / 記号・空白正規化 / 小文字化
  s = s.normalize('NFKC')
       .replace(/[‐-‒–—―ー]/g, '-')        // ダッシュ類統一
       .replace(/\s+/g, ' ')
       .trim()
       .toLowerCase();
  return s;
}
function tokenize_(s) {
  // 粗い日本語用トークナイザ（記号で分割＋2-4gramを追加）
  s = normalizeJa_(s).replace(/[^\p{L}\p{N}\s-]/gu, ' ');
  var toks = s.split(/\s+/).filter(Boolean);
  var grams = [];
  for (var n=2; n<=4; n++){
    for (var i=0; i+n<=toks.length; i++) grams.push(toks.slice(i,i+n).join(' '));
  }
  return toks.concat(grams);
}


var __idfCache = null;
function computeIdf_(chunks) {
  if (__idfCache) return __idfCache;
  var df = {};
  chunks.forEach(function(c){
    var seen = {};
    tokenize_(c).forEach(function(t){ seen[t]=1; });
    Object.keys(seen).forEach(function(t){ df[t]=(df[t]||0)+1; });
  });
  var N = chunks.length;
  var idf = {};
  Object.keys(df).forEach(function(t){
    idf[t] = Math.log( (N - df[t] + 0.5) / (df[t] + 0.5) + 1 );
  });
  __idfCache = idf;
  return idf;
}

function bm25Score_(chunk, queryTokens, idf, k1, b, avgLen) {
  k1 = k1 || 1.5; b = b || 0.75;
  var toks = tokenize_((chunk));
  var len = toks.length;
  var tf = {};
  toks.forEach(function(t){ tf[t]=(tf[t]||0)+1; });

  var score = 0;
  var seen = {};
  queryTokens.forEach(function(qt){
    if (seen[qt]) return; seen[qt]=1; // 同一語の重複カウント抑制
    var f = tf[qt] || 0;
    var id = idf[qt] || 0;
    var denom = f + k1*(1 - b + b*len/avgLen);
    score += id * (f * (k1+1)) / (denom || 1);
  });
  return score;
}

function rerankWithLLM_(question, scoredTop) {
  if (!scoredTop || !scoredTop.length) return [];
  var items = scoredTop.slice(0, 20).map(function(o, i){
    return (i+1) + '. ' + o.chunk.slice(0, 800); // LLM入力安全化
  }).join('\n\n');

  var prompt = '質問への関連度で次の抜粋をスコア0〜1で評価し、上位3件を返す。' +
    'JSON: {"ranking":[{"idx":番号,"score":数値}...]}。番号は与えた番号を用いる。\n\n' +
    '質問:\n' + question + '\n\n候補:\n' + items;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.0 }
  };
  var resp = callGeminiApi(GEMINI_API_URL_FLASH, payload);
  try {
    var js = JSON.parse(resp.candidates[0].content.parts[0].text);
    var topIdx = (js.ranking||[]).sort(function(a,b){ return b.score-a.score; }).slice(0,3).map(function(r){ return r.idx-1; });
    return topIdx.map(function(i){ return scoredTop[i].chunk; });
  } catch(e){
    console.warn('再ランク失敗→BM25上位を流用');
    return scoredTop.slice(0,3).map(function(o){ return o.chunk; });
  }
}





