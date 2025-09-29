// /**
//  * google site記載用
//  * インストールおよび遷移リンク
// */
// function doGet(e) {
//   // 1) 呼び出しユーザー ↔ このChatアプリ のDMを用意（既存があればそれを返す）
//   const dm = Chat.Spaces.setup({
//     space: { spaceType: 'DIRECT_MESSAGE', singleUserBotDm: true }
//   });

//   const spaceName = dm.name;            // 例: "spaces/1234567890"
//   const spaceId   = spaceName.split('/')[1];

//   // 2) 多重ログイン対策：/u/{index} をクエリで選択（デフォルト0）
//   const idx = Number(e && e.parameter && e.parameter.u ? e.parameter.u : 0);
//   const chatUrl = `https://mail.google.com/chat/u/${idx}/#chat/dm/${spaceId}`;

//   // 3) 即リダイレクト（iframe=Sites埋め込み時にブロックされる可能性があるのでフォールバック付）
//   const html = HtmlService.createHtmlOutput(
//   `<!doctype html><meta charset="utf-8">
//   <base target="_top">
//   <style>
//     body{font:14px system-ui;padding:16px;line-height:1.6}
//     a.btn{display:inline-block;padding:10px 14px;border:1px solid #ccc;border-radius:8px;text-decoration:none}
//     .note{margin-top:16px;padding:12px;border:1px solid #eee;border-radius:8px;background:#fffbe6}
//     .note h2{margin:0 0 8px;font-size:16px}
//     .note ol{margin:0;padding-left:1.4em}
//   </style>
//   <script>
//   (function(){
//     var url = ${JSON.stringify(chatUrl)};
//     try {
//       if (window.top === window.self) {
//         window.location.replace(url); // トップで開いているとき
//         return;
//       }
//       // Sitesのiframeなら親へ遷移を試みる
//       window.top.location.href = url;
//       setTimeout(function(){ document.getElementById('fb').style.display='inline-block'; }, 400);
//     } catch(e) {
//       document.addEventListener('DOMContentLoaded', function(){
//         document.getElementById('fb').style.display='inline-block';
//       });
//     }
//   })();
//   </script>

//   <div class="note">
//     <li>「Google Chat を開く」ボタンをクリックしてください。</li>
//   </div>

//   <p><a id="fb" class="btn" href="${chatUrl}" style="display:none">Google Chat を開く</a></p>`
//   ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

//   return html;
// }
