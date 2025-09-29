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
        var logUrl = saveAnswerLogJson_(currentRow, {
          row: currentRow,
          question: question,
          analysis: analysisResult,
          queries: finalResult ? finalResult.search_query : [],
          sources: finalResult ? finalResult.web_research_result : [],
          draftAnswer: finalAnswer,
          createdAt: new Date().toISOString()
        });
        if (LOG_LINK_COL) sheet.getRange(currentRow, LOG_LINK_COL).setValue(logUrl);
      } catch (e) {
        console.warn('回答ログ保存に失敗: row=' + currentRow, e);
      }

      var spreadsheetUrl = SpreadsheetApp.openById(SPREADSHEET_ID).getUrl();
      var dmSpaceName = values[i][DM_SPACE_NAME_COL - 1];

      // Google Chat用にメッセージを整形（区切り線などを削除）
      var chatMessage = finalAnswer.replace(/---/g, '');

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
  var chunks = text.split(/\n---\sスライド\s\d+\s---\n|\n\n+/);
  return chunks.filter(function(c) { return c.trim() !== ''; });
}

/**
 * キーワードに基づいて関連チャンクを検索・スコアリングする
 */
function findRelevantChunks_(chunks, keywords, topK) {
  topK = topK || 3;
  if (!keywords || keywords.length === 0) {
    return chunks.slice(0, topK);
  }
  var scoredChunks = chunks.map(function(chunk) {
    var score = 0;
    keywords.forEach(function(keyword) {
      if (chunk.toLowerCase().indexOf(keyword.toLowerCase()) !== -1) {
        score++;
      }
    });
    return { chunk: chunk, score: score };
  });
  var relevantChunks = scoredChunks.filter(function(item) {
    return item.score > 0;
  });
  relevantChunks.sort(function(a, b) {
    return b.score - a.score;
  });
  return relevantChunks.slice(0, topK).map(function(item) {
    return item.chunk;
  });
}

// --- Googleドキュメントからテキストを抽出するヘルパー関数群 ---

function getTextFromGoogleSlide_(slideUrl) {
  try {
    var presentationIdMatch = slideUrl.match(/presentation\/d\/([a-zA-Z0-9-_]+)/);
    if (!presentationIdMatch || !presentationIdMatch[1]) {
      console.log('無効なGoogle Slide URLです: ' + slideUrl);
      return '';
    }
    var presentationId = presentationIdMatch[1];
    var presentation = SlidesApp.openById(presentationId);
    var slides = presentation.getSlides();
    var allText = '';
    for (var i = 0; i < slides.length; i++) {
      var slide = slides[i];
      allText += '\n--- スライド ' + (i + 1) + ' ---\n';
      var shapes = slide.getShapes();
      for (var j = 0; j < shapes.length; j++) {
        var shape = shapes[j];
        if (shape.getText) {
          var textRange = shape.getText();
          if (textRange) {
            allText += textRange.asString();
          }
        }
      }
    }
    return allText;
  } catch (e) {
    console.error('Google Slideからのテキスト抽出に失敗しました: ' + slideUrl, e);
    return '（スライド「' + slideUrl + '」の読み込みに失敗しました）';
  }
}

function getTextFromGoogleSheet_(sheetUrl) {
  try {
    var spreadsheetIdMatch = sheetUrl.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!spreadsheetIdMatch || !spreadsheetIdMatch[1]) {
      console.log('無効なGoogle Sheet URLです: ' + sheetUrl);
      return '';
    }
    var spreadsheetId = spreadsheetIdMatch[1];
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var allText = '';
    var sheets = spreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      allText += '\n--- シート: ' + sheet.getName() + ' ---\n';
      var data = sheet.getDataRange().getValues();
      for (var j = 0; j < data.length; j++) {
        allText += data[j].join('\t') + '\n';
      }
    }
    return allText;
  } catch (e) {
    console.error('Google Sheetからのテキスト抽出に失敗しました: ' + sheetUrl, e);
    return '（シート「' + sheetUrl + '」の読み込みに失敗しました）';
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
function researchAgent(initialQuestion) {
  var config = {
    initial_search_query_count: 3,
    max_research_loops: 2,
  };
  var INTERNAL_GUIDE_URLS = (PropertiesService.getScriptProperties().getProperty('INTERNAL_GUIDE_URLS') || '')
    .split(',').map(function(s){return s.trim()}).filter(Boolean);
  var internalGuideSection = '';

  // --- ▼ RAG実装 ここから ▼ ---
  var internalSummary = '';  // ← 300字サマリをここに入れる

  if (INTERNAL_GUIDE_URLS.length > 0) {
    var allGuideText = '';
    var slideUrls = INTERNAL_GUIDE_URLS.filter(function(url){ return url.includes('docs.google.com/presentation/d/'); });
    var sheetUrls = INTERNAL_GUIDE_URLS.filter(function(url){ return url.includes('docs.google.com/spreadsheets/d/'); });
    var docUrls   = INTERNAL_GUIDE_URLS.filter(function(url){ return url.includes('docs.google.com/document/d/'); });

    if (slideUrls.length) allGuideText += slideUrls.map(getTextFromGoogleSlide_).join('\n\n');
    if (sheetUrls.length) allGuideText += (allGuideText ? '\n\n' : '') + sheetUrls.map(getTextFromGoogleSheet_).join('\n\n');
    if (docUrls.length)   allGuideText += (allGuideText ? '\n\n' : '') + docUrls.map(getTextFromGoogleDoc_).join('\n\n');

    if (allGuideText.trim()) {
      var keywords = generateSearchKeywords_(initialQuestion);
      console.log('抽出された検索キーワード: ' + keywords.join(', '));
      var chunks = chunkText_(allGuideText);
      var relevantChunks = findRelevantChunks_(chunks, keywords, 3); // 上位3チャンク

      if (relevantChunks.length > 0) {
        // 300字で要約（関連が薄い場合は短くなるだけ。スキップ判定は下流で実施）
        internalSummary = summarizeToLength_(relevantChunks.join('\n\n'), 300, '社内資料に書かれていない推測は書かない');
        console.log('社内ガイドから ' + relevantChunks.length + ' 件抽出 → 300字要約を作成');
      } else {
        console.log('社内ガイド内に質問と関連する情報は見つかりませんでした。');
      }
    }
  }
  // --- ▲ RAG実装 ここまで ▲ ---

  config.max_total_queries = 10;
  var seenQueries = [];
  var seenDomains = [];
  function domainOf(u){ try{ return (new URL(u)).hostname.replace(/^www\./,''); }catch(e){ return ''; } }

  var state = {
    messages: [{ role: "user", content: initialQuestion }],
    search_query: [],
    web_research_result: [],
    sources_gathered: [],
    research_loop_count: 0,
    number_of_ran_queries: 0,
  };

  var initialQueries = generateQuery(state, config)
    .filter(function(q) { return q && seenQueries.indexOf(q) === -1; })
    .slice(0, config.initial_search_query_count);
  for(var k=0; k<initialQueries.length; k++) seenQueries.push(initialQueries[k]);
  state.search_query = state.search_query.concat(initialQueries);
  state.number_of_ran_queries += initialQueries.length;
  console.log('生成された検索クエリ: ' + initialQueries.join(', '));

  for (var loop = 0; loop < config.max_research_loops; loop++) {
    state.research_loop_count = loop + 1;
    var searchResults = [];
    var currentQueries = state.search_query.slice(-state.number_of_ran_queries);
    if (seenQueries.length >= config.max_total_queries) break;

    for(var l=0; l<currentQueries.length; l++){
      var query = currentQueries[l];
      console.log('ウェブ調査中 (' + (loop + 1) + '/' + config.max_research_loops + '): "' + query + '"');
      var researchResult = webResearch(query);
      if(researchResult.citations){
        researchResult.citations.forEach(function(c) { 
          var d = domainOf(c.url); 
          if (d && seenDomains.indexOf(d) === -1) seenDomains.push(d); 
        });
      }
      searchResults.push(researchResult);
    }
    state.web_research_result = state.web_research_result.concat(searchResults);

    console.log("調査結果を評価中...");
    var reflectionResult = reflection(state);

    var nextQs = (reflectionResult.follow_up_queries || [])
      .filter(function(q) { return q && seenQueries.indexOf(q) === -1; })
      .slice(0, 3);

    var penalize = seenDomains.slice(0,2).map(function(d){ return '-site:' + d; }).join(' ');
    var diversified = nextQs.map(function(q){ return penalize ? q + ' ' + penalize : q; });

    for(var m=0; m<diversified.length; m++){
      if (seenQueries.length < config.max_total_queries) {
        state.search_query.push(diversified[m]);
        seenQueries.push(diversified[m]);
      }
    }

    if (reflectionResult.is_sufficient) {
      console.log("情報が十分であると判断。最終回答を生成します。");
      break;
    } else {
      console.log('知識ギャップを検出: ' + reflectionResult.knowledge_gap);
      console.log('追加のクエリを生成: ' + reflectionResult.follow_up_queries.join(', '));
      state.search_query = state.search_query.concat(reflectionResult.follow_up_queries);
      state.number_of_ran_queries = reflectionResult.follow_up_queries.length;
    }

    if (loop === config.max_research_loops - 1) {
      console.log("最大ループ回数に達しました。");
    }
  }

  // --- Web要約(300字) ---
  var webSummary = '';
  if (state.web_research_result && state.web_research_result.length) {
    // 収集テキストを一本化して300字要約。必要ならURLが本文に含まれていれば1つだけ末尾に括弧で併記
    webSummary = summarizeToLength_(
      state.web_research_result.join('\n\n'),
      300,
      '本文にURLが含まれる場合は1つだけ末尾に(参考:URL)の形式で併記'
    );
  }


  try {
    var ragHits = (typeof ragSearch_v3 === 'function')
      ? ragSearch_v3(state.messages[0].content, Number(PropertiesService.getScriptProperties().getProperty('RAG_TOP_K')||5))
      : ((typeof ragSearch_v2 === 'function')
         ? ragSearch_v2(state.messages[0].content, Number(PropertiesService.getScriptProperties().getProperty('RAG_TOP_K')||5))
         : []);
    if (Array.isArray(ragHits) && ragHits.length) {
      var appendix = '\n\n---\n【類似事例（過去案件）】\n' + ragHits.map(function(r, i){ return (i+1)+'. ' + r.url; }).join('\n');
      state.messages.push({ role: "model", content: appendix });
    }
  } catch (e) {
    console.warn('RAG付与に失敗しました（処理は継続）:', e);
  }

  console.log("最終回答を生成中...");

  // 旧 finalizeAnswer を使わず、新しい合成器で 600 字に収める
  var final600 = finalizeAnswerWithSections_(state.messages[0].content, internalSummary, webSummary);

  // メッセージ配列にも残す（互換のため）
  state.messages.push({ role: "model", content:
    (internalSummary ? '【社内300字】\n' + internalSummary + '\n\n' : '') +
    (webSummary ? '【外部300字】\n' + webSummary + '\n\n' : '') +
    '【最終回答(600字以内)】\n' + final600
  });

  // 呼び出し元で直接使えるよう返却フィールドも付ける
  state.final_internal_300 = internalSummary;
  state.final_web_300      = webSummary;
  state.final_600          = final600;

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
  return hardTrim_(out, 600);
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
