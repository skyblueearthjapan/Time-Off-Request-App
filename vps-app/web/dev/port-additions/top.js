
/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/top.html のまま）
 * =================================================================== */

// 申請モーダルを開く前に、ログインが続いているかを確かめる（操作中のログイン切れ）
(function(){
  var orig = openModal;
  openModal = function(){
    if (window.__vpsCheckSession) window.__vpsCheckSession();
    return orig.apply(this, arguments);
  };
})();

// 操作中にログインが切れたときの書きかけ（申請モーダル・編集モーダル）の退避と復元
(function(){
  if (!window.__vpsDraft) return;
  var when = window.__vpsDraft.when;
  function byId(id){ return document.getElementById(id); }
  function isActive(id){ var el = byId(id); return !!el && el.classList.contains('active'); }
  function checkedValue(name){ var r = document.querySelector('input[name="' + name + '"]:checked'); return r ? r.value : ''; }
  function setChecked(name, value){
    document.querySelectorAll('input[name="' + name + '"]').forEach(function(r){ r.checked = (r.value === value); });
  }

  function restoreRequest(r){
    // 部署 → 作業員（部署一覧が読めてから。作業員を選ぶと申請ボタンが押せる）
    when(function(){ var s = byId('deptSelect'); return s.options.length > 1 && s.options[0].value === ''; }, function(){
      var sel = byId('deptSelect');
      sel.value = r.dept;
      if (sel.value !== r.dept) return;
      selectedDeptId = r.dept;
      selectedDeptName = sel.options[sel.selectedIndex].dataset.name || '';
      loadWorkers(r.dept, r.worker);
      // 作業員と休暇日の一覧が届いてからモーダルを開き直す
      when(function(){ return selectedWorkerId === r.worker && byId('leaveDate').options.length > 1; }, function(){
        openModal();
        setChecked('leaveKubun', r.kubun);
        setChecked('halfDayType', r.halfDay);
        if (r.retro) {
          var p = r.retro.split('-');
          retroCalYear = Number(p[0]);
          retroCalMonth = Number(p[1]);
          if (r.retroOpen) toggleRetroSection(); else renderRetroCalendar();
          selectRetroDate(r.retro);
        } else {
          byId('leaveDate').value = r.leaveDate || '';
          if (r.retroOpen) toggleRetroSection();
        }
        if (r.kubun) {
          buildLeaveTypeOptions(r.kubun);
          setChecked('leaveType', r.leaveType);
        }
        byId('specialReason').value = r.specialReason || '';
        byId('paidDetail').value = r.paidDetail || '';
        var step = Number(r.step) || 1;
        // 振替元は候補が届いてから選び直す
        when(function(){ return ALL_HOLIDAY_CANDIDATES.length > 0 || !r.substituteDate; }, function(){
          renderSubstitutePulldown();
          byId('substituteDate').value = r.substituteDate || '';
          validateSubstituteSelection();
          if (step === 9) buildConfirmation();
        });
        if (step === 9) buildConfirmation();
        showStep(step);
      });
    });
  }

  function restoreEdit(e){
    when(function(){
      return (window.LEAVE_ITEMS || []).some(function(it){ return it.reqId === e.reqId; });
    }, function(){
      openEditLeaveModal(e.reqId);
      byId('editLeaveDate').value = e.leaveDate || '';
      setChecked('editLeaveKubun', e.kubun);
      byId('editHalfDayWrap').style.display = e.kubun === '半日休' ? '' : 'none';
      setChecked('editHalfDayType', e.halfDay);
    });
  }

  window.__vpsDraft.register({
    name: 'top',
    collect: function(){
      var out = {};
      if (isActive('modalOverlay') && selectedWorkerId && byId('submitBtn')) {
        out.request = {
          dept: selectedDeptId, worker: selectedWorkerId, step: currentStep,
          kubun: checkedValue('leaveKubun'), halfDay: checkedValue('halfDayType'), leaveType: checkedValue('leaveType'),
          leaveDate: byId('leaveDate').value, retro: selectedRetroactiveLeaveDate,
          retroOpen: byId('retroSection').classList.contains('open'),
          substituteDate: byId('substituteDate').value,
          specialReason: byId('specialReason').value, paidDetail: byId('paidDetail').value
        };
      }
      if (isActive('editOverlay') && EDIT_LEAVE_CTX && EDIT_LEAVE_CTX.reqId) {
        out.edit = {
          reqId: EDIT_LEAVE_CTX.reqId, leaveDate: byId('editLeaveDate').value,
          kubun: checkedValue('editLeaveKubun'), halfDay: checkedValue('editHalfDayType')
        };
      }
      return (out.request || out.edit) ? out : null;
    },
    restore: function(d){
      if (d.request) restoreRequest(d.request);
      else if (d.edit) restoreEdit(d.edit);
    }
  });
})();
