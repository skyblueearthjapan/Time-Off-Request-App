/* leaveStatus.js — GAS 版 src/leaveStatus.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
  // ===== 定数 =====
  // タイムラインのノード横位置（年度内のおおよその日付位置：6/30,8/31,12/31,2/15,3/15）
  var NODE_POS = [29.0, 46.0, 79.5, 92.0, 100.0];

  // 3/16始まりの年度をクライアント側でも算出（サーバの computeFiscalYear_ と一致）
  function computeFyJs(d){
    d = d || new Date();
    var y = d.getFullYear(), m = d.getMonth()+1, day = d.getDate();
    return (m < 3 || (m === 3 && day < 16)) ? y-1 : y;
  }
  var CUR_FY = computeFyJs();
  var FY_MIN = CUR_FY - 5;               // 過去5年度まで遡れる
  var state = { fy: CUR_FY, dept: '', staff: '' };

  // 今日の位置（%）：当年度3/16からの経過日数 / 365。過去年度は100%扱い（呼び出し側で判定）
  function todayPct(fy){
    var start = new Date(fy, 2, 16); // 3/16（月index=2）
    var now = new Date();
    var pct = (now - start) / (365*24*3600*1000) * 100;
    return Math.max(0, Math.min(100, pct));
  }

  function statusMap(wl){
    if(wl === '危険') return {cls:'danger', label:'危険'};
    if(wl === '警告') return {cls:'warn',   label:'警告'};
    if(wl === '注意') return {cls:'info',   label:'注意'};
    return {cls:'ok', label:'順調'};
  }
  function stRank(wl){ return wl==='危険'?3:(wl==='警告'?2:(wl==='注意'?1:0)); }
  function kindClass(t){ t = t || ''; return t.indexOf('有給')===0 ? 'paid' : (t.indexOf('振替')===0 ? 'sub' : 'spec'); }
  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
  function dashIfEmpty(s){ return (s==null || s==='') ? '—' : esc(s); }

  // ===== セレクタ =====
  var deptSel = document.getElementById('deptSel');
  var staffSel = document.getElementById('staffSel');

  function fillDept(){
    google.script.run
      .withSuccessHandler(function(list){
        list = list || [];
        deptSel.innerHTML = '<option value="">全部署</option>' +
          list.map(function(d){ return '<option value="'+esc(d.deptId)+'">'+esc(d.deptName)+'</option>'; }).join('');
        deptSel.value = state.dept;
      })
      .withFailureHandler(function(){ deptSel.innerHTML = '<option value="">全部署</option>'; })
      .api_getDeptList();
  }

  function fillStaff(){
    if(!state.dept){ // 全部署のときはスタッフ指定不可（全員のみ）
      staffSel.innerHTML = '<option value="">全員</option>';
      staffSel.value = '';
      return;
    }
    google.script.run
      .withSuccessHandler(function(list){
        list = list || [];
        staffSel.innerHTML = '<option value="">全員</option>' +
          list.map(function(w){ return '<option value="'+esc(w.workerId)+'">'+esc(w.workerName)+'</option>'; }).join('');
        // 選択中スタッフが新しい部署に存在しなければ全員に戻す
        if(!list.some(function(w){ return w.workerId === state.staff; })) state.staff = '';
        staffSel.value = state.staff;
      })
      .withFailureHandler(function(){ staffSel.innerHTML = '<option value="">全員</option>'; })
      .api_getWorkersByDept(state.dept);
  }

  function setYearUI(){
    document.getElementById('yearM').textContent = state.fy + '年度';
    document.getElementById('yearP').textContent = '3/16 〜 翌3/15（年間）';
    document.getElementById('periodLabel').textContent = state.fy + '年度（3/16〜翌3/15）';
    document.getElementById('prevY').disabled = (state.fy <= FY_MIN);
    document.getElementById('nextY').disabled = (state.fy >= CUR_FY);
  }

  // ===== メインの検索・描画 =====
  function render(){
    setYearUI();
    var body = document.getElementById('body');
    body.classList.remove('fade'); void body.offsetWidth; body.classList.add('fade');
    body.innerHTML = '<div class="empty">読み込み中...</div>';
    google.script.run
      .withSuccessHandler(function(res){
        if(!res){ body.innerHTML = '<div class="empty">データがありません</div>'; return; }
        if(res.mode === 'personal') body.innerHTML = renderPersonal(res);
        else body.innerHTML = renderList(res);
      })
      .withFailureHandler(function(err){
        body.innerHTML = '<div class="empty">読み込みに失敗しました：' + esc(err && err.message || err) + '</div>';
      })
      .api_getLeaveStatusSearch(state.fy, state.dept, state.staff);
  }

  function renderPersonal(res){
    var w = res.worker;
    var isCur = (res.fyYear === CUR_FY);
    if(!w){
      return '<div class="empty">選択したスタッフの' + esc(res.fyYear) + '年度データが見つかりませんでした。</div>';
    }
    var sm = statusMap(w.warnLevel);
    var count = w.paidCount, rest = Math.max(0, 5 - count);
    var halfNote = (w.paidHalfCount > 0) ? ('うち半休 ' + esc(w.paidHalfCount) + '回') : '半休なし';

    var dots = '';
    for(var i=0;i<5;i++){ dots += '<span class="'+(i<Math.floor(count)?'on':'off')+'">●</span>'; }

    var ms = res.milestones || [];
    var asOf = isCur ? todayPct(res.fyYear) : 100;
    var nodes = ms.map(function(m, i){
      var tag = m.met ? '達成' : (m.state==='future' ? '残'+esc(m.remaining) : '不足'+esc(m.deficit));
      return '<div class="node '+(['done','risk','warn','note','future'].indexOf(m.state) >= 0 ? m.state : '')+'" style="left:'+NODE_POS[i]+'%">'+
             '<div class="circle">'+esc(m.required)+'</div>'+
             '<div class="date">'+esc(m.dateText)+'</div>'+
             '<div class="st">'+tag+'</div></div>';
    }).join('');
    var todayEl = isCur ? '<div class="today" style="left:'+asOf+'%"><span class="lab">今日</span></div>' : '';

    // 次アクション コールアウト
    var passedUnmet = ms.filter(function(m){ return m.passed && !m.met; });
    var nextFuture  = ms.filter(function(m){ return !m.passed; })[0];
    var callout='', cclass='ok', cic='✓';
    if(passedUnmet.length){
      var pw = passedUnmet[passedUnmet.length-1], def = pw.deficit;
      cclass = def>=3?'danger':(def>=2?'warn':'info'); cic = def>=2?'⚠':'！';
      callout = esc(pw.dateText)+' 期限の必要回数に <b>'+esc(def)+'回</b> 足りていません。早めに有給（私用）を取得してください。';
    } else if(nextFuture){
      var need = nextFuture.remaining;
      if(need===0){ cclass='ok'; cic='✓'; callout='次の期限 '+esc(nextFuture.dateText)+'（'+esc(nextFuture.required)+'回）は達成済み。年度末の 5回 に向けて計画的に取得しましょう。'; }
      else { cclass='info'; cic='🗓'; callout='次の期限 <b>'+esc(nextFuture.dateText)+'</b> までに <b>'+esc(nextFuture.required)+'回</b>（あと <b>'+esc(need)+'回</b>）。計画的に取得しましょう。'; }
    } else {
      cic = count>=5?'🎉':'✓'; cclass = count>=5?'ok':sm.cls;
      callout = count>=5 ? '年度の取得義務 5回を達成しています。' : '年度が終了しました。';
    }

    // 今後の休暇予定（当年度のみ・実データ）
    var up = res.upcoming || [];
    var upRows = isCur
      ? (up.length ? up.map(function(u){
          var kubun = esc(u.leaveKubun) + (u.halfDayType ? '（'+esc(u.halfDayType)+'）' : '');
          var stCell = (u.status==='承認済') ? '<span class="pill ok">承認済</span>' : '<span class="pill warn">'+esc(u.status||'申請中')+'</span>';
          return '<tr><td class="tabnum">'+esc(u.leaveDate)+'</td><td>'+kubun+'</td><td><span class="kindChip '+kindClass(u.leaveType)+'">'+esc(u.leaveType)+'</span></td><td>'+stCell+'</td></tr>';
        }).join('') : '<tr><td colspan="4" class="empty">今後の休暇予定はありません</td></tr>')
      : '<tr><td colspan="4" class="empty">過去年度のため予定はありません</td></tr>';

    // 当年度の取得内訳（実データの回数ベース。※個別日付の履歴は現状APIが返さないため回数で表示）
    var brk = [
      {type:'有給消化（私用）', n:w.paidCount, note:'最終取得日 '+dashIfEmpty(w.lastPaidDate)},
      {type:'振替休暇',        n:w.substituteCount, note:''},
      {type:'特別休暇',        n:w.specialCount, note:''}
    ];
    var brkRows = brk.map(function(b){
      return '<tr><td><span class="kindChip '+kindClass(b.type)+'">'+esc(b.type)+'</span></td><td class="num tabnum">'+esc(b.n)+'</td><td class="muted">'+b.note+'</td></tr>';
    }).join('');

    var total = w.totalCount, initial = (w.workerName||'？').trim().charAt(0);
    return ''+
    '<div class="meHead"><div class="avatar">'+esc(initial)+'</div><div><div class="nm">'+esc(w.workerName)+'</div><div class="dp">'+esc(w.deptName)+'　/　'+esc(res.fyYear)+'年度（3/16〜翌3/15）</div></div><div style="flex:1"></div><span class="pill '+sm.cls+'">有給義務：'+sm.label+'</span></div>'+
    '<div class="kpiGrid">'+
      '<div class="kpi paid"><span class="rail"></span><div class="k-label">有給消化（私用）</div><div class="k-val tabnum">'+esc(count)+'<span class="slash"> / 5回</span></div><div class="dots">'+dots+'</div><div class="k-sub">年度目標5回・残 '+esc(rest)+'回 ／ '+halfNote+'</div></div>'+
      '<div class="kpi sub"><span class="rail"></span><div class="k-label">振替休暇</div><div class="k-val tabnum">'+esc(w.substituteCount)+'<span class="slash"> 回</span></div><div class="k-sub">前月16〜当月15の期間内</div></div>'+
      '<div class="kpi spec"><span class="rail"></span><div class="k-label">特別休暇</div><div class="k-val tabnum">'+esc(w.specialCount)+'<span class="slash"> 回</span></div><div class="k-sub">結婚・忌引 等</div></div>'+
      '<div class="kpi tot"><span class="rail"></span><div class="k-label">合計取得</div><div class="k-val tabnum">'+esc(total)+'<span class="slash"> 回</span></div><div class="k-sub">最終有給日 '+dashIfEmpty(w.lastPaidDate)+'</div></div>'+
    '</div>'+
    '<div class="obl"><div class="o-top"><span class="ttl">有給取得義務の達成状況（年度内の期限）</span><span class="note">数字＝その期限までに必要な累計回数</span></div>'+
      '<div class="timeline"><div class="baseline"></div><div class="elapsed" style="width:'+asOf+'%"></div>'+nodes+todayEl+'</div>'+
      '<div class="o-legend"><span><i class="sw green"></i>達成</span><span><i class="sw grey"></i>これから（残N）</span><span><i class="sw amber"></i>期限超過・注意/警告</span><span><i class="sw red"></i>期限超過・危険</span></div>'+
    '</div>'+
    '<div class="callout '+cclass+'"><span class="ic">'+cic+'</span><span>'+callout+'</span></div>'+
    '<div class="twoCol" style="display:grid; gap:16px; grid-template-columns:1fr 1fr; margin-top:18px;">'+
      '<div><div class="sectionTitle" style="margin-bottom:6px"><div class="lead"><span>今後の休暇予定</span></div></div><div class="tableWrap"><table><thead><tr><th>日付</th><th>区分</th><th>種類</th><th>状態</th></tr></thead><tbody>'+upRows+'</tbody></table></div></div>'+
      '<div><div class="sectionTitle" style="margin-bottom:6px"><div class="lead"><span>'+esc(res.fyYear)+'年度の取得内訳</span></div></div><div class="tableWrap"><table><thead><tr><th>種類</th><th class="num">回数</th><th>備考</th></tr></thead><tbody>'+brkRows+'</tbody></table></div></div>'+
    '</div>';
  }

  function renderList(res){
    var workers = (res.workers || []).slice();
    // 状況が悪い順（severity降順）→ 有給回数昇順
    workers.sort(function(a, b){
      return (stRank(b.warnLevel) - stRank(a.warnLevel)) || (a.paidCount - b.paidCount);
    });

    var n = workers.length, warnCnt = 0, dangerCnt = 0, sum = 0;
    workers.forEach(function(o){
      if(o.warnLevel !== '') warnCnt++;
      if(o.warnLevel === '危険') dangerCnt++;
      sum += o.paidCount;
    });
    var avg = n ? (Math.round(sum/n*10)/10) : 0;

    var rows = workers.map(function(o){
      var sm = statusMap(o.warnLevel), dots = '';
      for(var i=0;i<5;i++){ dots += '<span class="'+(i<Math.floor(o.paidCount)?'on':'off')+'">●</span>'; }
      return '<tr class="rowLink" data-worker="'+esc(o.workerId)+'" data-dept="'+esc(o.deptName)+'" role="button" tabindex="0" title="クリックで '+esc(o.workerName)+' さんの詳細を表示">'+
        '<td><div class="rowname">'+esc(o.workerName)+'<span class="goArrow">›</span></div><div class="rowdept">'+esc(o.deptName)+'</div></td>'+
        '<td><span class="tabnum">'+esc(o.paidCount)+' / 5</span> <span class="minidots">'+dots+'</span></td>'+
        '<td class="num tabnum">'+esc(o.substituteCount)+'</td>'+
        '<td class="num tabnum">'+esc(o.specialCount)+'</td>'+
        '<td class="num tabnum">'+esc(o.totalCount)+'</td>'+
        '<td class="tabnum">'+dashIfEmpty(o.lastPaidDate)+'</td>'+
        '<td><span class="pill '+sm.cls+'">'+sm.label+'</span></td>'+
      '</tr>';
    }).join('');
    if(!rows) rows = '<tr><td colspan="7" class="empty">対象のスタッフがいません</td></tr>';

    var scope = res.deptFilter ? res.deptFilter : '全部署';
    return ''+
    '<div class="listKpi">'+
      '<div class="lk"><div class="v tabnum">'+n+'</div><div class="l">対象者</div></div>'+
      '<div class="lk"><div class="v warn tabnum">'+warnCnt+'</div><div class="l">有給 未達（注意以上）</div></div>'+
      '<div class="lk"><div class="v danger tabnum">'+dangerCnt+'</div><div class="l">危険レベル</div></div>'+
      '<div class="lk"><div class="v tabnum">'+avg+'</div><div class="l">平均取得回数</div></div>'+
    '</div>'+
    '<div class="sectionTitle" style="margin-bottom:8px"><div class="lead"><span>有給取得状況 一覧 — '+esc(scope)+'</span></div><div class="section-count">'+n+'名 ・ '+esc(res.fyYear)+'年度</div></div>'+
    '<div class="tableWrap"><table>'+
      '<thead><tr><th>氏名 / 部署</th><th>有給（私用）</th><th class="num">振替</th><th class="num">特別</th><th class="num">合計</th><th>最終有給日</th><th>状況</th></tr></thead>'+
      '<tbody>'+rows+'</tbody>'+
    '</table></div>'+
    '<div class="hint" style="margin-top:12px">💡 状況が悪い順に並べています。<b>各行をクリック</b>すると、その人の達成状況タイムライン（詳細）が開きます（上のプルダウンでも切替可）。</div>';
  }

  // 一覧の行クリック → その人の詳細ビューへ（部署/スタッフの絞り込みも自動で切替）
  function goToPerson(dept, worker){
    if(!worker) return;
    state.dept = dept || '';
    state.staff = worker;
    deptSel.value = state.dept;   // 部署プルダウンを同期（全部署→その人の部署に切替）
    fillStaff();                  // スタッフ候補を再取得し、その人を選択状態に
    render();                     // 詳細（本人）ビューを表示
    document.getElementById('body').scrollIntoView({ behavior:'smooth', block:'start' });
  }

  // ===== イベント配線 =====
  document.getElementById('prevY').addEventListener('click', function(){ if(state.fy > FY_MIN){ state.fy--; render(); } });
  document.getElementById('nextY').addEventListener('click', function(){ if(state.fy < CUR_FY){ state.fy++; render(); } });
  deptSel.addEventListener('change', function(){ state.dept = deptSel.value; state.staff = ''; fillStaff(); render(); });
  staffSel.addEventListener('change', function(){ state.staff = staffSel.value; render(); });

  // 一覧行のクリック／キーボード操作（body は再描画されるため委譲でハンドル）
  var bodyEl = document.getElementById('body');
  bodyEl.addEventListener('click', function(ev){
    var tr = ev.target && ev.target.closest ? ev.target.closest('tr.rowLink') : null;
    if(tr) goToPerson(tr.getAttribute('data-dept'), tr.getAttribute('data-worker'));
  });
  bodyEl.addEventListener('keydown', function(ev){
    if(ev.key !== 'Enter' && ev.key !== ' ') return;
    var tr = ev.target && ev.target.closest ? ev.target.closest('tr.rowLink') : null;
    if(tr){ ev.preventDefault(); goToPerson(tr.getAttribute('data-dept'), tr.getAttribute('data-worker')); }
  });

  // ===== 初期化：ログイン本人を初期選択（本人ビューで開く） =====
  function init(){
    setYearUI();
    google.script.run
      .withSuccessHandler(function(me){
        if(me && me.workerName){ document.getElementById('whoName').textContent = me.workerName + (me.deptName ? '（'+me.deptName+'）' : ''); }
        else { document.getElementById('whoName').textContent = '—'; }
        if(me && me.deptId){ state.dept = me.deptId; state.staff = me.workerId; }
        fillDept();
        fillStaff();
        render();
      })
      .withFailureHandler(function(){
        document.getElementById('whoName').textContent = '—';
        fillDept();
        fillStaff();
        render();
      })
      .api_getWorkerInfo();
  }
  init();
