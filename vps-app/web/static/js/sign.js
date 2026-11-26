/* sign.js — GAS 版 src/sign.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var REQ_ID = __T.REQ_ID;
var APP_URL = __T.APP_URL;

var approverList = [];
var selectedApprover = null;
var hasStamp = false;

// ====== Canvas setup ======
var canvas = document.getElementById('signCanvas');
var ctx = canvas.getContext('2d');
var drawing = false;
var hasDrawn = false;

function resizeCanvas() {
  var rect = canvas.getBoundingClientRect();
  var dpr = window.devicePixelRatio || 1;
  canvas.width = rect.width * dpr;
  canvas.height = rect.height * dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.strokeStyle = '#0b2545';
  ctx.lineWidth = 2.5;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
}

canvas.addEventListener('pointerdown', function(e) {
  drawing = true; hasDrawn = true;
  var rect = canvas.getBoundingClientRect();
  ctx.beginPath();
  ctx.moveTo(e.clientX - rect.left, e.clientY - rect.top);
  e.preventDefault();
});
canvas.addEventListener('pointermove', function(e) {
  if (!drawing) return;
  var rect = canvas.getBoundingClientRect();
  ctx.lineTo(e.clientX - rect.left, e.clientY - rect.top);
  ctx.stroke();
  e.preventDefault();
});
canvas.addEventListener('pointerup', function(e) { drawing = false; e.preventDefault(); });
canvas.addEventListener('pointerleave', function() { drawing = false; });

function clearCanvas() {
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  hasDrawn = false;
}

// ====== 承認者一覧を読み込み ======
window.addEventListener('resize', function() { resizeCanvas(); clearCanvas(); });

window.addEventListener('DOMContentLoaded', function() {
  resizeCanvas();

  google.script.run
    .withSuccessHandler(function(list) {
      approverList = list || [];
      var sel = document.getElementById('approverSelect');
      sel.innerHTML = '<option value="">-- 承認者を選択 --</option>';
      // 手書きサインの選択肢を先頭に追加
      var handOpt = document.createElement('option');
      handOpt.value = '__handwrite__';
      handOpt.textContent = '✍ 手書きサイン（電子印なし）';
      sel.appendChild(handOpt);
      // 電子印登録者を追加
      for (var i = 0; i < approverList.length; i++) {
        var opt = document.createElement('option');
        opt.value = approverList[i].email;
        opt.textContent = approverList[i].name + '（' + approverList[i].email + '）';
        sel.appendChild(opt);
      }
    })
    .withFailureHandler(function(err) {
      document.getElementById('approverSelect').innerHTML =
        '<option value="">取得エラー: ' + __escText(err.message) + '</option>';
    })
    .api_getApproverList();
});

// ====== 承認者変更時 ======
function onApproverChange() {
  var sel = document.getElementById('approverSelect');
  var email = sel.value;
  var statusEl = document.getElementById('stampStatus');
  var signArea = document.getElementById('signArea');
  var btn = document.getElementById('approveBtn');

  if (!email) {
    statusEl.innerHTML = '';
    signArea.style.display = 'none';
    btn.disabled = true;
    selectedApprover = null;
    hasStamp = false;
    return;
  }

  if (email === '__handwrite__') {
    // 手書きサイン選択
    selectedApprover = '';
    hasStamp = false;
    statusEl.innerHTML = '<span class="stamp-badge no-stamp">手書きサイン</span>';
    signArea.style.display = 'block';
    document.querySelector('.sign-area-label').textContent = '下の枠内にサインを書いてください';
    btn.disabled = false;
    resizeCanvas();
    clearCanvas();
    return;
  }

  selectedApprover = email;
  // StampMapに登録されている = 電子印あり
  hasStamp = true;
  statusEl.innerHTML = '<span class="stamp-badge has-stamp">電子印あり</span>';
  signArea.style.display = 'none';
  btn.disabled = false;
}

// ====== 承認実行 ======
function goBack() {
  window.top.location.href = APP_URL + '?page=preview&reqId=' + REQ_ID;
}

function doApprove() {
  var selVal = document.getElementById('approverSelect').value;
  if (!selVal) {
    alert('承認者を選択してください。');
    return;
  }
  if (!hasStamp && !hasDrawn) {
    alert('サインを書いてください。');
    return;
  }
  if (!confirm('承認してよろしいですか？')) return;

  var btn = document.getElementById('approveBtn');
  btn.disabled = true;
  btn.textContent = '処理中...';
  var msg = document.getElementById('statusMsg');
  msg.textContent = '承認処理中... しばらくお待ちください。';
  msg.className = 'status-msg';

  // サインBase64（電子印がない場合のみ）
  var signBase64 = '';
  if (!hasStamp && hasDrawn) {
    signBase64 = canvas.toDataURL('image/png');
  }

  console.log('承認実行: reqId=' + REQ_ID + ', approverEmail=' + selectedApprover + ', hasStamp=' + hasStamp + ', signBase64.length=' + signBase64.length);

  google.script.run
    .withSuccessHandler(function(result) {
      console.log('approve result:', result);
      if (result && result.ok) {
        msg.className = 'status-msg success';
        btn.style.display = 'none';
        document.querySelector('.btn-back').style.display = 'none';
        var previewUrl = APP_URL + '?page=preview&reqId=' + REQ_ID + '&approved=1';
        var meta = document.createElement('meta');
        meta.httpEquiv = 'refresh';
        meta.content = '1;url=' + previewUrl;
        document.head.appendChild(meta);
        msg.innerHTML = '承認が完了しました。<br>承認済プレビューを表示します...<br>'
          + '<a href="' + previewUrl + '" target="_top" style="font-size:13px;color:#0ea5e9;">自動で切り替わらない場合はこちら</a>';
      } else {
        msg.textContent = '承認に失敗しました。';
        msg.className = 'status-msg error';
        btn.disabled = false;
        btn.textContent = '承認する';
      }
    })
    .withFailureHandler(function(err) {
      console.error('approve error:', err);
      msg.textContent = 'エラー: ' + err.message;
      msg.className = 'status-msg error';
      btn.disabled = false;
      btn.textContent = '承認する';
    })
    .api_approveLeaveRequest(REQ_ID, signBase64, selectedApprover);
}

/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/sign.html のまま）
 * =================================================================== */

// 「承認する」を押す前に、ログインが続いているかを確かめる（仕様 3-5。開いたときはシムが確かめる）
(function(){
  var orig = doApprove;
  doApprove = function(){
    if (window.__vpsCheckSession) window.__vpsCheckSession();
    return orig.apply(this, arguments);
  };
})();

// 操作中にログインが切れたときの書きかけ（承認者の選択と手書きサイン）の退避と復元。
// サインは canvas.toDataURL の PNG を sessionStorage に置き、ログインから戻ったら canvas に描き戻す
(function(){
  if (!window.__vpsDraft) return;
  var when = window.__vpsDraft.when;

  window.__vpsDraft.register({
    name: 'sign',
    collect: function(){
      var sel = document.getElementById('approverSelect');
      if (!sel || !sel.value) return null;
      return {
        reqId: REQ_ID,
        approver: sel.value,
        sign: (sel.value === '__handwrite__' && hasDrawn) ? canvas.toDataURL('image/png') : ''
      };
    },
    restore: function(d){
      if (!d || d.reqId !== REQ_ID || !d.approver) return;
      // 承認者の一覧が届いてから選び直す（手書きなら枠が出て大きさが決まる）
      when(function(){ return document.getElementById('approverSelect').options.length > 1; }, function(){
        var sel = document.getElementById('approverSelect');
        sel.value = d.approver;
        if (sel.value !== d.approver) return;
        onApproverChange();
        if (d.approver !== '__handwrite__' || !d.sign || d.sign.indexOf('data:image/png;base64,') !== 0) return;
        var img = new Image();
        img.onload = function(){
          var rect = canvas.getBoundingClientRect();
          ctx.drawImage(img, 0, 0, rect.width, rect.height);
          hasDrawn = true;
        };
        img.src = d.sign;
      });
    }
  });
})();
