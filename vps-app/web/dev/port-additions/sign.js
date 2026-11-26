
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
