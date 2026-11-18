// Browsing controls use bounded API pages; downloads always contain the full run.
function debounce(callback, delay=300) {
  let timer;
  return (...args) => {clearTimeout(timer); timer=setTimeout(()=>callback(...args),delay);};
}
function metric(label, value, detail='') {
  return `<div class="metric-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong><small>${escapeHtml(detail)}</small></div>`;
}
function elapsed(milliseconds) {
  if(milliseconds==null) return '—';
  return milliseconds<1000 ? `${Math.round(milliseconds)} ms` : `${(milliseconds/1000).toFixed(2)} 秒`;
}
function renderOverview() {
  const enabled=state.reports.filter(r=>r.enabled && r.schedule_type!=='manual');
  const next=enabled.filter(r=>r.next_run_at).sort((a,b)=>new Date(a.next_run_at)-new Date(b.next_run_at))[0];
  $('#workspace-overview').innerHTML=metric('报表定义',state.reports.length,'在左侧选择或新建')+metric('启用的自动计划',enabled.length,'人工运行始终可用')+metric('最近的计划',next?formatTime(next.next_run_at):'未安排',next?.name || '可在生成计划中设置');
}
function downloadLinks(run) {
  const names={xlsx:'Excel',docx:'Word',html:'HTML',csv:'CSV'};
  return (run.output_formats || []).map(format=>`<a href="/api/runs/${run.id}/file?format=${encodeURIComponent(format)}" ${format==='html'?'target="_blank" rel="noopener"':''}>${names[format] || escapeHtml(format)}</a>`).join('');
}
function updatePager(prefix, data, count) {
  const pages=Math.max(1,Math.ceil(count/data.page_size));
  const start=count?(data.page-1)*data.page_size+1:0;
  const end=Math.min(data.page*data.page_size,count);
  $(`#${prefix}-page-info`).textContent=`${start}–${end} / ${count} 条 · 第 ${data.page} / ${pages} 页`;
  $(`#${prefix}-prev`).disabled=data.page<=1;
  $(`#${prefix}-next`).disabled=data.page>=pages;
}
function historyQuery() {
  return new URLSearchParams({page:state.historyPage,page_size:$('#history-page-size').value,q:$('#run-search').value.trim(),status:$('#run-status-filter').value,trigger_type:$('#run-trigger-filter').value});
}
function renderRuns() {
  $('#runs-list').innerHTML=state.runs.map(run=>`<article class="run-row">
    <div><strong>${escapeHtml(run.report_name)}</strong><small>#${run.id} · ${escapeHtml(triggerLabel(run))} · ${escapeHtml(scheduleSummary(run.schedule_snapshot || {}))}</small><small>开始：${formatTime(run.started_at)}${escapeHtml(runParameterSummary(run))}</small>${run.scheduled_for?`<small>计划时间：${formatTime(run.scheduled_for)}</small>`:''}</div>
    <div><span class="run-status ${run.status}">${{success:'成功',failed:'失败',running:'运行中'}[run.status] || escapeHtml(run.status)}</span><small>${run.row_count || 0} 行 · 耗时 ${elapsed(run.metrics?.duration_ms)}</small>${run.error?`<details class="run-error"><summary>查看错误详情</summary><pre>${escapeHtml(run.error)}</pre></details>`:''}</div>
    <div><small>完成时间</small><small>${formatTime(run.finished_at)}</small></div>
    <div class="run-actions">${run.status==='success'?`<button class="ghost small view-run" data-id="${run.id}">查看</button>${downloadLinks(run)}`:''}<button class="ghost small reuse-run" data-id="${run.id}">复用参数</button></div>
  </article>`).join('') || '<div class="empty">没有符合当前条件的运行记录。可清除筛选或运行一份报表。</div>';
  $$('.view-run').forEach(el=>el.addEventListener('click',()=>viewRun(Number(el.dataset.id))));
  $$('.reuse-run').forEach(el=>el.addEventListener('click',()=>reuseRun(Number(el.dataset.id))));
}
async function loadRuns() {
  const request=++state.historyRequest;
  try {
    const data=await api(`/api/runs?${historyQuery()}`);
    if(request!==state.historyRequest) return;
    const changed=JSON.stringify(state.runs)!==JSON.stringify(data.items);
    state.historyPage=data.page; state.runs=data.items;
    if(changed || !$('#runs-list').children.length) renderRuns();
    updatePager('history',data,data.total);
    $('#history-metrics').innerHTML=metric('筛选匹配',data.total,'覆盖全部历史')+metric('成功',data.success)+metric('失败',data.failed)+metric('运行中',data.running);
    $('#history-message').textContent='';
  } catch(error) {if(request===state.historyRequest) $('#history-message').textContent=`历史加载失败：${error.message}`;}
}
async function reuseRun(id) {
  if(!canLeaveEditor()) return;
  const request=++state.editorRequest, snapshot=JSON.stringify(formPayload());
  try {
    const run=await api(`/api/runs/${id}`);
    const report=await api(`/api/reports/${run.report_id}`);
    if(request!==state.editorRequest) return;
    if(snapshot!==JSON.stringify(formPayload())) {toast('编辑内容已变化，请再次选择复用参数'); return;}
    state.selectedId=report.id; fillForm(report); renderReports();
    $$('#run-parameters [data-param]').forEach(el=>{if(Object.hasOwn(run.parameters || {},el.dataset.param)) el.value=run.parameters[el.dataset.param] ?? '';});
    setTab('definition'); $('#editor-run').scrollIntoView({behavior:'smooth',block:'center'});
    toast('已载入历史参数，请检查后运行；将使用当前报表定义');
  } catch(error) {toast(error.message);}
}
function resultQuery() {
  return new URLSearchParams({page:state.resultPage,page_size:$('#result-page-size').value,q:$('#result-search').value.trim(),sort:$('#result-sort').value,direction:$('#result-direction').value});
}
async function loadResultPage() {
  if(!state.selectedRunId) return;
  const request=++state.resultRequest;
  $('#result-page-message').textContent='正在查询明细…';
  try {
    const data=await api(`/api/runs/${state.selectedRunId}/data?${resultQuery()}`);
    if(request!==state.resultRequest) return;
    state.resultPage=data.page;
    const selection=$('#result-sort').value;
    $('#result-sort').innerHTML='<option value="">原始顺序</option>'+data.columns.map(c=>`<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
    $('#result-sort').value=selection;
    $('#result-table thead').innerHTML='<tr>'+data.columns.map(c=>`<th scope="col">${escapeHtml(c)}</th>`).join('')+'</tr>';
    $('#result-table tbody').innerHTML=data.items.length?data.items.map(row=>'<tr>'+data.columns.map(c=>`<td>${escapeHtml(row[c])}</td>`).join('')+'</tr>').join(''):`<tr><td colspan="${Math.max(1,data.columns.length)}" class="table-empty">没有匹配的数据。</td></tr>`;
    $('#result-page-message').textContent=`共 ${data.total} 行 · 匹配 ${data.filtered} 行 · ${data.columns.length} 列。搜索和排序覆盖全部明细，下载包含完整数据。`;
    updatePager('result',data,data.filtered);
    $('#result-content').classList.remove('hidden'); $('#result-empty').classList.add('hidden');
  } catch(error) {
    if(request!==state.resultRequest) return;
    $('#result-page-message').textContent=`明细加载失败：${error.message}`;
    $('#result-empty').textContent=`明细加载失败：${error.message}`;
    $('#result-empty').classList.remove('hidden');
  }
}
async function viewRun(id) {
  const request=++state.resultRequest;
  const changed=state.selectedRunId!==id;
  state.selectedRunId=id;
  setTab('results');
  $('#result-content').classList.add('hidden');
  $('#result-empty').textContent='正在加载运行结果…'; $('#result-empty').classList.remove('hidden');
  if(changed){state.resultPage=1; $('#result-search').value=''; $('#result-sort').innerHTML='<option value="">原始顺序</option>'; $('#result-direction').value='asc';}
  try {
    const [run,summaries]=await Promise.all([api(`/api/runs/${id}`),api(`/api/runs/${id}/summary`)]);
    if(request!==state.resultRequest) return;
    state.resultInfo=run;
    $('#result-meta').textContent=`${run.report_name} · 运行 #${id} · ${triggerLabel(run)} · 完成于 ${formatTime(run.finished_at)}${runParameterSummary(run)}`;
    const m=run.metrics || {};
    $('#result-metrics').innerHTML=metric('数据行数',run.row_count)+metric('总耗时',elapsed(m.duration_ms),`主查询 ${elapsed(m.query_ms)} · 生成 ${elapsed(m.output_ms)}`)+metric('独立 SQL 查询',m.source_queries ?? '—','同一次运行内相同查询复用')+metric('数据源复用',m.cache_hits ?? '—','跨单元格及输出格式');
    $('#result-downloads').innerHTML='<strong>下载完整报表</strong>'+downloadLinks(run);
    $('#result-summary').classList.toggle('hidden',!summaries.length);
    $('#result-summary').innerHTML='<div class="summary-title">统计值</div><div class="summary-grid">'+summaries.map(item=>`<div class="summary-card"><span>${escapeHtml(item.name)}</span><strong>${escapeHtml(item.value)}</strong><small title="${escapeHtml(item.formula)}">${escapeHtml(item.formula)}</small></div>`).join('')+'</div>';
    $('#result-preview').open=false; $('#result-chart').innerHTML='';
    $('#result-preview').classList.toggle('hidden',!(run.output_formats || []).includes('html'));
    await loadResultPage();
  } catch(error) {if(request===state.resultRequest) $('#result-empty').textContent=`结果加载失败：${error.message}`;}
}
document.addEventListener('DOMContentLoaded',()=>{
  $('#quick-save').addEventListener('click',saveReport);
  const refreshHistory=()=>{state.historyPage=1; loadRuns();};
  const refreshResult=()=>{state.resultPage=1; loadResultPage();};
  $('#run-search').addEventListener('input',debounce(refreshHistory));
  ['run-status-filter','run-trigger-filter','history-page-size'].forEach(id=>$(`#${id}`).addEventListener('change',refreshHistory));
  $('#reset-run-filters').addEventListener('click',()=>{['run-search','run-status-filter','run-trigger-filter'].forEach(id=>$(`#${id}`).value=''); refreshHistory();});
  $('#history-prev').addEventListener('click',()=>{state.historyPage=Math.max(1,state.historyPage-1);loadRuns();});
  $('#history-next').addEventListener('click',()=>{state.historyPage+=1;loadRuns();});
  $('#result-search').addEventListener('input',debounce(refreshResult));
  ['result-sort','result-direction','result-page-size'].forEach(id=>$(`#${id}`).addEventListener('change',refreshResult));
  $('#reset-result-filters').addEventListener('click',()=>{$('#result-search').value='';$('#result-sort').value='';$('#result-direction').value='asc';refreshResult();});
  $('#result-prev').addEventListener('click',()=>{state.resultPage=Math.max(1,state.resultPage-1);loadResultPage();});
  $('#result-next').addEventListener('click',()=>{state.resultPage+=1;loadResultPage();});
  $('#result-preview').addEventListener('toggle',()=>{if($('#result-preview').open && !$('#result-chart iframe') && state.selectedRunId) $('#result-chart').innerHTML=`<iframe title="HTML 报表结果" src="/api/runs/${state.selectedRunId}/file?format=html"></iframe>`;});
  $$('.editor-nav a').forEach(link=>link.addEventListener('click',event=>{
    event.preventDefault(); const target=$(link.getAttribute('href')); if(target.tagName==='DETAILS') target.open=true;
    target.scrollIntoView({behavior:'smooth',block:'start'});
  }));
  $('#expand-editor').addEventListener('click',()=>{
    const expand=$$('#definition-panel details.editor-section').some(el=>!el.open);
    $$('#definition-panel details.editor-section').forEach(el=>el.open=expand);
    $('#expand-editor').textContent=expand?'收起全部':'展开全部';
  });
});
