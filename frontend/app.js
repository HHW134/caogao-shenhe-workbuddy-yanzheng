/* 草案审核 WorkBuddy 验证 —— 前端交互 */

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

const LEVEL_CN = { error: '错误', warning: '警告', info: '提示' };

const state = {
  file: null,
  modules: new Set(['chapters']),
  result: null,
  ruleKey: null,
  ruleRows: [],
};

/* ---------------- Tab 切换 ---------------- */
$$('.tab').forEach((btn) => {
  btn.addEventListener('click', () => {
    $$('.tab').forEach((b) => b.classList.remove('active'));
    $$('.panel').forEach((p) => p.classList.remove('active'));
    btn.classList.add('active');
    $('#panel-' + btn.dataset.tab).classList.add('active');
    if (btn.dataset.tab === 'history') loadHistory();
  });
});

/* ---------------- 模块列表 ---------------- */
async function loadModules() {
  const res = await fetch('/api/modules');
  const data = await res.json();
  const box = $('#modules');
  box.innerHTML = '';
  data.modules.forEach((m) => {
    const div = document.createElement('div');
    div.className = 'mod' + (state.modules.has(m.key) ? ' on' : '');
    div.innerHTML = `
      <div class="mod-top">
        <input type="checkbox" ${state.modules.has(m.key) ? 'checked' : ''}>
        <span>${m.name}</span>
      </div>
      <div class="mod-desc">${m.desc}</div>`;
    div.addEventListener('click', (e) => {
      if (e.target.tagName !== 'INPUT') {
        const cb = div.querySelector('input');
        cb.checked = !cb.checked;
      }
      const on = div.querySelector('input').checked;
      on ? state.modules.add(m.key) : state.modules.delete(m.key);
      div.classList.toggle('on', on);
      refreshRunBtn();
    });
    box.appendChild(div);
  });
  renderRuleFiles(data.rules);
}

/* ---------------- 上传 ---------------- */
const drop = $('#drop');
const fileInput = $('#file');

drop.addEventListener('click', () => fileInput.click());
drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  drop.classList.remove('over');
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) setFile(fileInput.files[0]); });

function setFile(f) {
  if (!f.name.toLowerCase().endsWith('.docx')) {
    setStatus('仅支持 .docx 文件', true);
    return;
  }
  state.file = f;
  const info = $('#fileinfo');
  info.classList.remove('hidden');
  info.innerHTML = `<strong>${escapeHtml(f.name)}</strong><br>${(f.size / 1024).toFixed(1)} KB`;
  setStatus('');
  refreshRunBtn();
}

function refreshRunBtn() {
  $('#run').disabled = !(state.file && state.modules.size);
}

function setStatus(text, isErr = false) {
  const el = $('#status');
  el.innerHTML = text;
  el.classList.toggle('err', isErr);
}

/* ---------------- 执行审核 ---------------- */
$('#run').addEventListener('click', async () => {
  if (!state.file) return;
  const fd = new FormData();
  fd.append('file', state.file);
  fd.append('modules', Array.from(state.modules).join(','));

  $('#run').disabled = true;
  setStatus('<span class="spinner"></span>正在审核，请稍候…');

  try {
    const res = await fetch('/api/review', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '审核失败');
    state.result = data;
    renderResult(data);
    setStatus(`审核完成，共 ${data.问题总数} 条问题`);
  } catch (err) {
    setStatus('审核失败：' + err.message, true);
  } finally {
    $('#run').disabled = false;
  }
});

/* ---------------- 结果渲染 ---------------- */
function renderResult(data) {
  $('#empty').classList.add('hidden');
  $('#result').classList.remove('hidden');

  const s = data.级别统计 || {};
  const modStats = Object.entries(data.模块统计 || {})
    .map(([k, v]) => `${k} ${v}`).join('　');
  $('#stats').innerHTML = `
    <div class="stat"><div class="k">问题总数</div><div class="v">${data.问题总数}</div>
      <div class="sub">${data.审核结论}</div></div>
    <div class="stat error"><div class="k">错误 error</div><div class="v">${s.error || 0}</div>
      <div class="sub">须修改</div></div>
    <div class="stat warning"><div class="k">警告 warning</div><div class="v">${s.warning || 0}</div>
      <div class="sub">建议核实</div></div>
    <div class="stat info"><div class="k">提示 info</div><div class="v">${s.info || 0}</div>
      <div class="sub">参考</div></div>
    <div class="stat"><div class="k">审核文件</div>
      <div class="v" style="font-size:14px;line-height:1.5;word-break:break-all">${escapeHtml(data.文件名)}</div>
      <div class="sub">${data.审核时间}</div></div>
    <div class="stat"><div class="k">执行模块</div>
      <div class="v" style="font-size:13px;line-height:1.6">${(data.执行模块 || []).join('<br>')}</div>
      <div class="sub">${modStats}</div></div>`;

  const sel = $('#f-module');
  const mods = Array.from(new Set((data.问题列表 || []).map((r) => r.模块名称)));
  sel.innerHTML = '<option value="">全部模块</option>' +
    mods.map((m) => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');

  renderIssues();

  $('#downloads').innerHTML = (data.产物 || []).map((d) => `
    <a class="dl" href="${d.url}">
      <span class="ico">${d.文件名.endsWith('.xlsx') ? '📊' : '📝'}</span>
      <span>
        <span class="t">${d.类型}</span>
        <span class="s">${escapeHtml(d.文件名)} · ${(d.大小 / 1024).toFixed(0)} KB</span>
      </span>
    </a>`).join('') || '<p class="tablefoot">本次未生成产物文件</p>';

  if (data.模块异常) {
    const msg = Object.entries(data.模块异常).map(([k, v]) => `${k}: ${v}`).join('；');
    setStatus('部分模块异常 → ' + msg, true);
  }
}

function renderIssues() {
  const data = state.result;
  if (!data) return;
  const lv = $('#f-level').value;
  const mod = $('#f-module').value;
  const kw = $('#f-kw').value.trim().toLowerCase();

  const rows = (data.问题列表 || []).filter((r) => {
    if (lv && r.级别 !== lv) return false;
    if (mod && r.模块名称 !== mod) return false;
    if (kw) {
      const hay = `${r.审核要点}${r.问题}${r.位置}${r.修改建议}${r.原文 || ''}`.toLowerCase();
      if (!hay.includes(kw)) return false;
    }
    return true;
  });

  $('#issues tbody').innerHTML = rows.map((r) => `
    <tr class="${r.级别}">
      <td>${r.序号}</td>
      <td>${escapeHtml(r.审核要点)}<div class="mono">${escapeHtml(r.模块名称 || '')}</div></td>
      <td><span class="badge ${r.级别}">${LEVEL_CN[r.级别] || r.级别}</span></td>
      <td>${escapeHtml(r.问题)}${r.原文 ? `<div class="mono">原文：${escapeHtml(r.原文)}</div>` : ''}</td>
      <td>${escapeHtml(r.位置)}</td>
      <td>${escapeHtml(r.修改建议)}</td>
    </tr>`).join('') ||
    '<tr><td colspan="6" style="text-align:center;color:#64748b;padding:26px">没有符合条件的问题</td></tr>';

  $('#tablefoot').textContent =
    `显示 ${rows.length} / ${(data.问题列表 || []).length} 条　·　每条问题已作为 Word 批注写入「批注文档」对应段落`;
}

['#f-level', '#f-module', '#f-kw'].forEach((sel) => {
  $(sel).addEventListener('input', renderIssues);
});

/* ---------------- 规则中心 ---------------- */
function renderRuleFiles(rules) {
  const box = $('#rulefiles');
  box.innerHTML = '';
  rules.forEach((r, i) => {
    const div = document.createElement('div');
    div.className = 'rf' + (i === 0 ? ' on' : '');
    div.innerHTML = `<div class="n">${r.名称}</div>
      <div class="c">${r.规则数} 条 · 启用 ${r.启用数} · ${r.文件}</div>`;
    div.addEventListener('click', () => {
      $$('.rf').forEach((x) => x.classList.remove('on'));
      div.classList.add('on');
      loadRules(r.key);
    });
    box.appendChild(div);
  });
  if (rules.length) loadRules(rules[0].key);
}

async function loadRules(key) {
  state.ruleKey = key;
  const res = await fetch('/api/rules/' + key);
  const data = await res.json();
  state.ruleRows = data.rows || [];
  renderRules();
}

function renderRules() {
  const kw = $('#r-kw').value.trim().toLowerCase();
  const rows = state.ruleRows.filter((r) => {
    if (!kw) return true;
    return Object.values(r).join(' ').toLowerCase().includes(kw);
  });
  $('#rules-table tbody').innerHTML = rows.map((r) => `
    <tr>
      <td class="mono">${escapeHtml(r.编号)}</td>
      <td>${escapeHtml(r.分类)}</td>
      <td>${escapeHtml(r.审核要点)}</td>
      <td>${escapeHtml(r.问题)}</td>
      <td>${escapeHtml(r.修改建议)}</td>
      <td>${r.级别 ? `<span class="badge gray">${escapeHtml(r.级别)}</span>` : ''}</td>
      <td><span class="badge ${r.启用 === '启用' ? 'info' : 'gray'}">${r.启用}</span></td>
    </tr>`).join('') ||
    '<tr><td colspan="7" style="text-align:center;color:#64748b;padding:26px">无匹配规则</td></tr>';
  $('#rulefoot').textContent =
    `显示 ${rows.length} / ${state.ruleRows.length} 条　·　规则存放在仓库 rules/ 目录，改 JSON 后重跑审核即可生效，无需改代码`;
}

$('#r-kw').addEventListener('input', renderRules);

/* ---------------- 历史记录 ---------------- */
async function loadHistory() {
  const res = await fetch('/api/runs');
  const data = await res.json();
  const box = $('#history');
  if (!data.runs.length) {
    box.innerHTML = '<p class="tablefoot">暂无历史记录</p>';
    return;
  }
  box.innerHTML = data.runs.map((r) => {
    const s = r.级别统计 || {};
    return `<div class="run" data-run="${r.run_id}">
      <div>
        <div class="name">${escapeHtml(r.文件名)}</div>
        <div class="meta">${r.审核时间}　·　${(r.执行模块 || []).join('、')}</div>
      </div>
      <div class="nums">
        <span class="badge error">错误 ${s.error || 0}</span>
        <span class="badge warning">警告 ${s.warning || 0}</span>
        <span class="badge info">提示 ${s.info || 0}</span>
        <span class="badge gray">共 ${r.问题总数}</span>
      </div>
    </div>`;
  }).join('');

  $$('.run').forEach((el) => {
    el.addEventListener('click', async () => {
      const res2 = await fetch('/api/result/' + el.dataset.run);
      const data2 = await res2.json();
      state.result = data2;
      renderResult(data2);
      $$('.tab').forEach((b) => b.classList.remove('active'));
      $$('.panel').forEach((p) => p.classList.remove('active'));
      document.querySelector('.tab[data-tab="review"]').classList.add('active');
      $('#panel-review').classList.add('active');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });
}

$('#refresh-history').addEventListener('click', loadHistory);

/* ---------------- 工具 ---------------- */
function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

loadModules();
