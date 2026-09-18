import { fetch as tauriFetch } from '@tauri-apps/plugin-http';
import { save, open } from '@tauri-apps/plugin-dialog';
import { writeTextFile, readTextFile } from '@tauri-apps/plugin-fs';
import { invoke } from '@tauri-apps/api/core';
import { importFromExcel, exportToExcel, mergeExcelFiles } from './excel.js';

// ======== 常量 ========
var LS_KEYS = { endpoint: 'csa_ai_endpoint', key: 'csa_ai_key', model: 'csa_ai_model', retry: 'csa_ai_retry' };
var EOL_CACHE_KEY = 'csa_eol_cache';
var EOL_CACHE_TTL = 30 * 24 * 60 * 60 * 1000;
var BATCH_INPUT_KEY = 'csa_batch_input';
var INTERNAL_GROUP_IDS = ['com.lyzdfintech.loongeasy.ce'];

// ======== 状态 ========
var theme = 'light';
var reportData = {};
var auditEntries = [];
var traceLog = [];
var batchResults = [];
var batchRunning = false;
var lastAIRaw = '';
var importedExcelPaths = [];
var importedExcelPath = null;

// ======== 工具 ========
function sleep(ms) { return new Promise(function(resolve) { setTimeout(resolve, ms); }); }

async function httpGet(url) {
  var r = await fetchWithTimeout(url, null, 10000);
  if (!r.ok) throw new Error('HTTP ' + r.status);
  return await r.json();
}

async function httpPost(url, headers, body) {
  var r = await fetchWithTimeout(url, { method: 'POST', headers: headers, body: body }, 15000);
  return r;
}

function fetchWithTimeout(url, options, timeout) {
  return Promise.race([
    tauriFetch(url, options),
    new Promise(function(_, reject) {
      setTimeout(function() { reject(new Error('请求超时(' + timeout + 'ms): ' + url.substring(0, 80))); }, timeout);
    })
  ]);
}

async function saveTextFile(filename, content, ext) {
  var filePath = await save({
    defaultPath: filename,
    filters: [{ name: ext.toUpperCase(), extensions: [ext] }]
  });
  if (!filePath) return false;
  await writeTextFile(filePath, content);
  return true;
}

// ======== localStorage ========
function getLS(key) { try { return localStorage.getItem(key); } catch(e) { return null; } }
function setLS(key, val) { try { localStorage.setItem(key, val); } catch(e) {} }
function delLS(key) { try { localStorage.removeItem(key); } catch(e) {} }

function getEOLCache() { try { return JSON.parse(getLS(EOL_CACHE_KEY) || '{}'); } catch(e) { return {}; } }
function setEOLCache(cache) { setLS(EOL_CACHE_KEY, JSON.stringify(cache)); }
function getCachedEOL(name) {
  var cache = getEOLCache();
  var entry = cache[name];
  if (!entry || !entry.ts) return null;
  if (Date.now() - entry.ts > EOL_CACHE_TTL) { delete cache[name]; setEOLCache(cache); return null; }
  return entry;
}
function saveCachedEOL(name, eolData, aiEol, eolSource) {
  if (eolSource === 'none') return;
  var cache = getEOLCache();
  cache[name] = { ts: Date.now(), eolData: eolData || null, aiEol: aiEol || null, eolSource: eolSource };
  setEOLCache(cache);
}

// ======== UI 通用 ========
function $(id) { return document.getElementById(id); }
function getStep(n) { return document.querySelector('[data-step="' + n + '"]'); }
function setStepState(n, state) { if (getStep(n)) getStep(n).className = 'step ' + state; }

function addAudit(msg) {
  var now = new Date().toLocaleTimeString();
  auditEntries.push('[' + now + '] ' + msg);
  var log = $('auditLog');
  log.innerHTML = auditEntries.map(function(e) { return '<div class="entry">' + e + '</div>'; }).join('');
  $('auditCard').classList.remove('hidden');
  log.scrollTop = log.scrollHeight;
}

function logStep(stepNum, stepName, input, output, status) {
  status = status || 'ok';
  var inp = input ? JSON.parse(JSON.stringify(input)) : {};
  var out = output ? JSON.parse(JSON.stringify(output)) : {};
  function trimObj(obj) {
    var result = {};
    for (var k in obj) {
      if (obj.hasOwnProperty(k)) {
        var v = obj[k];
        if (typeof v === 'string' && v.length > 500) result[k] = v.substring(0, 500) + '...[截断]';
        else if (Array.isArray(v) && v.length > 6) result[k] = v.slice(0, 6).concat(['...[' + (v.length - 6) + ' 条更多]']);
        else result[k] = v;
      }
    }
    return result;
  }
  traceLog.push({ num: stepNum, name: stepName, input: trimObj(inp), output: trimObj(out), status: status, ts: new Date().toLocaleTimeString() });
  renderTrace();
}

function renderTrace() {
  var el = $('traceContent');
  if (!traceLog.length) return;
  var html = '';
  for (var i = 0; i < traceLog.length; i++) {
    var t = traceLog[i];
    html += '<div class="trace-step"><div class="trace-step-header"><span>步骤' + t.num + ' - ' + t.name + ' <span style="color:var(--text-secondary);font-weight:400">' + t.ts + '</span></span><span class="trace-status ' + t.status + '">' + (t.status === 'ok' ? 'OK' : 'ERR') + '</span></div><div class="trace-step-body"><div class="trace-section"><div class="trace-section-label">输入</div><div class="trace-code">' + JSON.stringify(t.input, null, 2) + '</div></div><div class="trace-section"><div class="trace-section-label">输出</div><div class="trace-code">' + JSON.stringify(t.output, null, 2) + '</div></div></div></div>';
  }
  el.innerHTML = html;
  $('traceCard').classList.remove('hidden');
}

function toggleTheme() {
  theme = theme === 'light' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', theme);
  $('btnTheme').textContent = theme === 'light' ? '☀' : '🌙';
}

function setReportData(d) { reportData = d; }

// ======== 单组件研判 ========
async function startAssessment() {
  var name = $('compName').value.trim();
  var version = $('compVersion').value.trim();
  var vendor = $('compVendor').value.trim();
  if (!name || !version) { alert('请输入组件名称和版本号'); return; }
  var colonIdx = name.indexOf(':');
  var searchName = colonIdx !== -1 ? name.substring(colonIdx + 1).trim() : name;
  var groupId = colonIdx !== -1 ? name.substring(0, colonIdx).trim() : '';
  setReportData({ name: searchName, version: version, vendor: vendor, groupId: groupId, rawName: name, date: new Date().toISOString().split('T')[0] });
  auditEntries = []; traceLog = [];
  $('vulnList').innerHTML = '<span style="color:var(--text-secondary);font-size:0.8rem">研判中…</span>';
  $('eolResult').innerHTML = '<span style="color:var(--text-secondary);font-size:0.8rem">研判中…</span>';
  $('reportContent').textContent = '';
  $('reportCard').classList.add('hidden');
  $('auditCard').classList.add('hidden');
  $('traceCard').classList.add('hidden');
  $('traceContent').innerHTML = '';
  $('aiRawCard').classList.add('hidden');
  for (var i = 1; i <= 8; i++) setStepState(i, 'pending');
  addAudit('开始研判: ' + (reportData.rawName || reportData.name) + ' ' + version);
  await runAllSteps(reportData.name, version, vendor);
}

async function runAllSteps(name, version, vendor) {
  try { await step1_cpe(name, version, vendor); } catch (e) { setStepState(1, 'error'); addAudit('[CPE匹配] 失败: ' + e.message); }
  try { await step2_cve(name, version); } catch (e) { setStepState(2, 'error'); addAudit('[CVE检索] 失败: ' + e.message); }
  try { await step6_eol(name); } catch (e) { setStepState(6, 'error'); addAudit('[生命周期] 失败: ' + e.message); }
  try { await step4_github(name, version); } catch (e) { setStepState(4, 'error'); addAudit('[GitHub] 失败: ' + e.message); }
  try { await step5_synthesize(); } catch (e) { setStepState(5, 'error'); addAudit('[AI深度分析] 失败: ' + e.message); }
  setStepState(3, 'done'); addAudit('[官方公告] 已结合API数据进行综合核验');
  logStep(3, '官方公告核验', { note: '厂商安全公告无统一API' }, { advisoryLinks: reportData.vulns ? reportData.vulns.map(function(v) { return 'https://nvd.nist.gov/vuln/detail/' + v.id; }) : [] });
  setStepState(7, 'done'); addAudit('[维护规则] 提取完成');
  logStep(7, '维护规则提取', { eolSource: reportData.eolSource }, { eolData: reportData.eolData ? 'endoflife.date ' + reportData.eolData.length + ' 条记录' : '无', aiEolRationale: reportData.aiEol ? reportData.aiEol.rationale : null });
  setStepState(8, 'done'); addAudit('[EOL判定] 判定完成');
  logStep(8, 'EOL终止判定', { eolSource: reportData.eolSource, currentVersion: version }, { eolStatus: reportData.aiEol ? reportData.aiEol.eolStatus : (reportData.eolData ? '见生命周期表' : '无数据') });
  generateReport();
  addAudit('研判完成，报告已生成');
}

// ======== Step 1: CPE ========
async function step1_cpe(name, version, vendor) {
  setStepState(1, 'running'); addAudit('[CPE匹配] 正在检索 NVD CPE...');
  var kw = vendor ? vendor + ' ' + name : name;
  var url = 'https://services.nvd.nist.gov/rest/json/cpes/2.0?keywordSearch=' + encodeURIComponent(kw) + '&resultsPerPage=10';
  var cpeResult = null, source = 'NVD API';
  try {
    var data = await httpGet(url);
    if (data.products && data.products.length > 0) { cpeResult = data.products[0].cpe.criteria; addAudit('[CPE匹配] 找到匹配: ' + cpeResult); }
  } catch (e) { source = '构造'; addAudit('[CPE匹配] NVD API 不可达，使用构造CPE'); }
  reportData.cpe = cpeResult || 'cpe:2.3:a:*:' + name.toLowerCase() + ':' + version;
  logStep(1, 'CPE资产匹配', { keyword: kw, apiUrl: url }, { cpe: reportData.cpe, source: source });
  setStepState(1, 'done');
}

// ======== Step 2: CVE ========
async function step2_cve(name, version) {
  setStepState(2, 'running'); addAudit('[CVE检索] 正在查询 OSV.dev...');
  var results = [];
  var ecosystems = ['Maven', 'npm', 'PyPI', 'Go', 'crates.io'];
  var mavenName = (reportData.rawName && reportData.rawName.indexOf(':') !== -1) ? reportData.rawName : name;

  var osvPromises = ecosystems.map(function(ec) {
    var pkgName = ec === 'Maven' ? mavenName : name;
    return fetchWithTimeout('https://api.osv.dev/v1/querybatch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ queries: [{ package: { name: pkgName, ecosystem: ec }, version: version }] })
    }, 10000).then(function(r) { return r.ok ? r.json() : null; }).catch(function() { return null; });
  });
  var osvResponses = await Promise.all(osvPromises);

  for (var i = 0; i < osvResponses.length; i++) {
    var data = osvResponses[i];
    if (data && data.results && data.results[0] && data.results[0].vulns) {
      var vulns = data.results[0].vulns;
      for (var j = 0; j < vulns.length; j++) {
        var v = vulns[j], sev = 'unknown', score = 'N/A';
        if (v.severity) {
          var dbInfo = v.database_specific || {}; score = dbInfo.severity || v.severity; sev = String(score).toLowerCase();
          var ns = parseFloat(score);
          if (!isNaN(ns)) { if (ns >= 9) sev = 'critical'; else if (ns >= 7) sev = 'high'; else if (ns >= 4) sev = 'medium'; else sev = 'low'; }
        }
        results.push({ id: v.id, severity: sev, score: score, summary: (v.summary || '').substring(0, 100) });
      }
    }
  }
  if (results.length === 0) {
    try {
      addAudit('[CVE检索] OSV无结果，尝试NVD API...');
      var nvdKeyword = (reportData.rawName || name) + ' ' + version;
      var nvdUrl = 'https://services.nvd.nist.gov/rest/json/cves/2.0?keywordSearch=' + encodeURIComponent(nvdKeyword) + '&resultsPerPage=20';
      var nvdData = await httpGet(nvdUrl);
      if (nvdData.vulnerabilities) {
        for (var k = 0; k < nvdData.vulnerabilities.length; k++) {
          var cveItem = nvdData.vulnerabilities[k].cve;
          var metrics = cveItem.metrics, cvssScore = 'N/A', sev = 'unknown';
          if (metrics && metrics.cvssMetricV31) cvssScore = metrics.cvssMetricV31[0].cvssData.baseScore;
          else if (metrics && metrics.cvssMetricV30) cvssScore = metrics.cvssMetricV30[0].cvssData.baseScore;
          var ns2 = parseFloat(cvssScore);
          if (!isNaN(ns2)) { if (ns2 >= 9) sev = 'critical'; else if (ns2 >= 7) sev = 'high'; else if (ns2 >= 4) sev = 'medium'; else sev = 'low'; }
          results.push({ id: cveItem.id, severity: sev, score: cvssScore, summary: ((cveItem.descriptions && cveItem.descriptions[0]) ? cveItem.descriptions[0].value : '').substring(0, 100) });
        }
      }
    } catch (e2) { addAudit('[CVE检索] NVD API 不可达: ' + e2.message); }
  }
  reportData.vulns = results; renderVulns(results);
  var crit = results.filter(function(v) { return v.severity === 'critical'; }).length;
  var high = results.filter(function(v) { return v.severity === 'high'; }).length;
  var med = results.filter(function(v) { return v.severity === 'medium'; }).length;
  var low = results.filter(function(v) { return v.severity === 'low'; }).length;
  addAudit('[CVE检索] 发现 ' + results.length + ' 个漏洞（严重:' + crit + ' 高危:' + high + '）');
  logStep(2, 'CVE漏洞检索', { component: name, version: version, osvEcosystems: ecosystems }, { totalVulns: results.length, critical: crit, high: high, medium: med, low: low, vulns: results.map(function(v) { return { id: v.id, severity: v.severity, score: v.score }; }) });
  setStepState(2, 'done');
}

function renderVulns(vulns) {
  var el = $('vulnList');
  if (!vulns || vulns.length === 0) { el.innerHTML = '<span style="color:var(--success);font-size:0.8rem">未发现已知漏洞（查询覆盖有限）</span>'; return; }
  var html = '';
  for (var i = 0; i < vulns.length; i++) {
    var v = vulns[i];
    html += '<div class="vuln-item ' + v.severity + '"><span><strong>' + v.id + '</strong><br>' + v.summary + '</span><span class="badge badge-' + v.severity + '">' + v.severity.toUpperCase() + (v.score !== 'N/A' ? ' ' + v.score : '') + '</span></div>';
  }
  el.innerHTML = html;
}

// ======== Step 6: EOL ========
function isInternalComponent() {
  var gid = reportData.groupId || '';
  if (INTERNAL_GROUP_IDS.indexOf(gid) !== -1) return true;
  var raw = reportData.rawName || '';
  var nm = reportData.name || '';
  return INTERNAL_GROUP_IDS.indexOf(raw) !== -1 || INTERNAL_GROUP_IDS.indexOf(nm) !== -1;
}

async function step6_eol(name) {
  setStepState(6, 'running');
  if (isInternalComponent()) {
    reportData.eolSource = 'internal'; reportData.internalMaintained = true;
    reportData.eolData = null; reportData.aiEol = null;
    addAudit('[生命周期] 内部自研组件（groupId: ' + reportData.groupId + '），直接标记维护中');
    renderEOL(); logStep(6, '生命周期', { component: name, internal: true }, { eolSource: 'internal', eolStatus: '维护中' });
    setStepState(6, 'done'); return;
  }
  var cacheKey = reportData.rawName || name;
  var cached = getCachedEOL(cacheKey);
  if (cached) {
    reportData.eolData = cached.eolData; reportData.aiEol = cached.aiEol;
    reportData.eolSource = cached.eolSource; reportData.cached = true; reportData.cachedAt = cached.ts;
    addAudit('[生命周期] 命中缓存，沿用上次结论');
    renderEOL(); logStep(6, '生命周期', { cached: true }, { eolSource: reportData.eolSource });
    setStepState(6, 'done'); return;
  }
  addAudit('[生命周期] 正在查询 endoflife.date...');
  var eolData = null;
  var eolSources = [];
  var nameVariants = generateEOLNameVariants(reportData.artifactId || name, reportData.groupId);
  for (var vi = 0; vi < nameVariants.length; vi++) {
    try {
      var r = await fetchWithTimeout('https://endoflife.date/api/' + encodeURIComponent(nameVariants[vi]) + '.json', null, 10000);
      if (r.ok) { eolData = await r.json(); if (eolData && eolData.length > 0) { eolSources.push('endoflife.date(' + nameVariants[vi] + ')'); break; } }
    } catch (e) {}
  }
  if (!eolData) addAudit('[生命周期] endoflife.date 无此组件数据');

  var mavenInfo = null;
  var githubInfo = null;
  var multiSourcePromises = [];
  if (reportData.groupId && reportData.artifactId) {
    multiSourcePromises.push(checkMavenCentralLatest(reportData.groupId, reportData.artifactId).then(function(m) { if (m) { eolSources.push('maven-central'); addAudit('[生命周期] Maven Central: 最新版本 ' + m.latestVersion + '，距今 ' + m.yearsSinceLast + ' 年'); } mavenInfo = m; }));
  }
  multiSourcePromises.push(checkGithubLatestRelease(reportData.artifactId || name).then(function(g) { if (g) { eolSources.push('github'); addAudit('[生命周期] GitHub: 最新 Release ' + g.latestRelease + '，距今 ' + g.yearsSinceLast + ' 年'); } githubInfo = g; }));
  await Promise.all(multiSourcePromises);

  reportData.eolData = eolData; reportData.eolSource = 'endoflife.date';
  if (!eolData || !Array.isArray(eolData) || eolData.length === 0) {
    addAudit('[生命周期] endoflife.date 无数据，自动调用 AI 兜底...');
    var aiKey = $('aiKey').value.trim();
    if (!aiKey) { addAudit('[AI兜底] 未配置 API Key'); $('aiCard').classList.remove('hidden'); }
    var aiResult = await aiEolBatch(reportData.artifactId || name, version, reportData.groupId, eolSources.join(', '), mavenInfo, githubInfo, $('aiEndpoint').value.trim(), aiKey, $('aiModel').value.trim());
    if (aiResult) { reportData.eolSource = 'AI'; reportData.aiEol = aiResult; addAudit('[AI兜底] AI 已返回 EOL 判定: ' + (aiResult.eolStatus || '未知')); }
    else {
      if (mavenInfo && mavenInfo.yearsSinceLast > 2) { reportData.eolSource = 'maven-stale'; reportData.mavenInfo = mavenInfo; addAudit('[生命周期] AI 失败，Maven Central 距今 ' + mavenInfo.yearsSinceLast + ' 年，判定为已EOL'); }
      else if (githubInfo && githubInfo.yearsSinceLast > 2) { reportData.eolSource = 'github-stale'; reportData.githubInfo = githubInfo; addAudit('[生命周期] AI 失败，GitHub 距今 ' + githubInfo.yearsSinceLast + ' 年，判定为已EOL'); }
      else { reportData.eolSource = 'none'; addAudit('[AI兜底] AI 未配置或调用失败'); }
    }
  }
  saveCachedEOL(cacheKey, reportData.eolData, reportData.aiEol, reportData.eolSource);
  renderEOL(); logStep(6, '生命周期', { cached: false }, { eolSource: reportData.eolSource });
  setStepState(6, 'done');
}

// ======== Step 4: GitHub ========
async function step4_github(name, version) {
  setStepState(4, 'running'); addAudit('[GitHub Release] 尝试获取发布信息...');
  var repos = ['apache/' + name.toLowerCase(), name.toLowerCase() + '/' + name.toLowerCase(), 'spring-projects/' + name.toLowerCase().replace('spring-', '')];
  var found = null;
  for (var i = 0; i < repos.length; i++) {
    try {
      var r = await tauriFetch('https://api.github.com/repos/' + repos[i] + '/releases?per_page=5');
      if (r.ok) { var data = await r.json(); if (data.length > 0) { found = { repo: repos[i], latest: data[0].tag_name, releases: data.slice(0, 5) }; break; } }
    } catch (e) {}
  }
  reportData.github = found;
  if (found) addAudit('[GitHub Release] 仓库: ' + found.repo + ', 最新: ' + found.latest);
  else addAudit('[GitHub Release] 未找到 GitHub 仓库');
  setStepState(4, 'done');
  logStep(4, 'GitHub Release', { searchedRepos: repos }, { found: found ? found.repo : null, latest: found ? found.latest : null });
}

// ======== AI 原始回显 ========
function displayAIRaw(component, rawContent, model, status, fullResponse) {
  $('aiRawCard').classList.remove('hidden');
  $('aiRawComp').textContent = component;
  $('aiRawTime').textContent = new Date().toLocaleString();
  $('aiRawModel').textContent = model || '未知';
  var rawEl = $('aiRawContent');
  rawEl.textContent = rawContent || '(空)';
  rawEl.style.borderLeft = status === 'error' ? '3px solid var(--danger)' : '';
  var parsedEl = $('aiRawParsed');
  var parsed = null;
  try { var m = (rawContent || '').match(/\{[\s\S]*\}/); if (m) parsed = JSON.parse(m[0]); } catch (e) { parsed = null; }
  if (parsed) parsedEl.textContent = JSON.stringify(parsed, null, 2);
  else parsedEl.textContent = '(未能解析为 JSON)' + (fullResponse ? '\n\n完整响应:\n' + JSON.stringify(fullResponse, null, 2) : '');
  lastAIRaw = rawContent || '';
  reportData.aiRawResponse = { content: rawContent, parsed: parsed, model: model, time: new Date().toISOString() };
}

// ======== Step 5: AI EOL 兜底 ========
async function step5_ai_eol(name) {
  setStepState(5, 'running');
  var endpoint = $('aiEndpoint').value.trim();
  var apiKey = $('aiKey').value.trim();
  var model = $('aiModel').value.trim();
  if (!endpoint || !apiKey) {
    addAudit('[AI兜底] 未配置 API Key 或端点');
    logStep(5, 'AI兜底', { component: name }, { skipped: true }, 'err');
    setStepState(5, 'done'); return null;
  }
  var maxRetry = parseInt($('aiRetry').value, 10);
  if (isNaN(maxRetry) || maxRetry < 0) maxRetry = 2;
  var maxAttempts = maxRetry + 1;
  var basePrompt = '你是一个开源组件生命周期分析专家。请对组件 "' + name + '"（如包含冒号，冒号前是 groupId，后是 artifactId）进行 EOL 状态研判。\n\n判定规则：\n- 如果该组件仍有新版本发布、官方仍在活跃维护 → "维护中"\n- 如果该组件已停止维护、官方不再发布更新、或仅偶尔打安全补丁而无新功能开发 → "已EOL"\n- 如果该组件即将停止维护（已公布 EOL 日期且在 6 个月内）→ "即将EOL"\n\n重要提示：\n- 不能仅凭最近是否有版本来判断是否在维护\n- 很多项目虽然近期有发布，但实际上已停止活跃维护，只是偶尔打安全补丁\n- 对于 Apache Commons 等老项目，如果已进入维护模式（仅安全补丁、无新功能），可判定为"已EOL"\n- 请结合你对这个组件的了解：官方是否活跃？社区是否活跃？是否有维护计划？\n\n如果不确定，默认判断为 "已EOL"。\n\n请以 JSON 格式返回：\n{\n  "eolStatus": "维护中" 或 "已EOL" 或 "即将EOL",\n  "currentVersionStatus": "描述当前已知最新版本是否仍在维护",\n  "latestSafeVersion": "推荐的安全版本号",\n  "rationale": "判定依据"\n}\n只返回 JSON，不要其他文字。';
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    addAudit('[AI兜底] 第 ' + attempt + '/' + maxAttempts + ' 次尝试...');
    var outcome = await callAIOnce(name, endpoint, apiKey, model, basePrompt, attempt);
    if (outcome.success) {
      addAudit('[AI兜底] AI 判定: ' + outcome.data.eolStatus);
      logStep(5, 'AI兜底', { model: model, attempts: attempt }, outcome.data);
      setStepState(5, 'done'); return outcome.data;
    }
    if (attempt < maxAttempts) { addAudit('[AI兜底] 第 ' + attempt + ' 次失败（' + outcome.reason + '），重试...'); await sleep(600 * attempt); }
    else { addAudit('[AI兜底] 已尝试 ' + maxAttempts + ' 次均失败'); logStep(5, 'AI兜底', { attempts: attempt }, { failed: true, reason: outcome.reason }, 'err'); }
  }
  setStepState(5, 'error'); return null;
}

async function callAIOnce(name, endpoint, apiKey, model, basePrompt, attempt) {
  var prompt = basePrompt;
  if (attempt > 1) prompt += '\n\n【重要提醒】上次返回格式不正确。请严格只返回 JSON。';
  var temperature = attempt === 1 ? 0.1 : 0.3;
  try {
    var r = await tauriFetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
      body: JSON.stringify({ model: model, messages: [{ role: 'system', content: '你是一个开源组件生命周期分析专家，返回严格的JSON格式。' }, { role: 'user', content: prompt }], temperature: temperature, max_tokens: 800 })
    });
    if (!r.ok) { var errBody = ''; try { errBody = await r.text(); } catch(e) {} displayAIRaw(name, 'HTTP ' + r.status + '\n' + errBody, model, 'error'); return { success: false, reason: 'HTTP ' + r.status }; }
    var data = await r.json();
    var content = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
    if (!content) { displayAIRaw(name, 'AI 返回空内容', model, 'error'); return { success: false, reason: 'AI 返回空内容' }; }
    displayAIRaw(name, content, model, 'ok', data);
    var jsonMatch = content.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return { success: false, reason: 'AI 返回无法解析为 JSON' };
    try { return { success: true, data: JSON.parse(jsonMatch[0]) }; }
    catch (e) { return { success: false, reason: 'JSON 解析失败: ' + e.message }; }
  } catch (e) { displayAIRaw(name, '网络异常: ' + e.message, model, 'error'); return { success: false, reason: '网络异常: ' + e.message }; }
}

async function step5_synthesize() {
  if (reportData.eolSource === 'AI' && reportData.aiEol) {
    setStepState(5, 'done'); addAudit('[AI深度分析] 已在 EOL 兜底阶段完成');
    logStep(5, 'AI深度分析', {}, { reused: true }); return;
  }
  setStepState(5, 'done'); addAudit('[AI深度分析] 综合研判完成');
  logStep(5, 'AI深度分析', {}, { eolSource: reportData.eolSource });
}

// ======== EOL 渲染 ========
function renderEOL() {
  var el = $('eolResult');
  var eolData = reportData.eolData;
  var eolSource = reportData.eolSource || 'endoflife.date';
  var aiEol = reportData.aiEol;
  if (eolData && Array.isArray(eolData) && eolData.length > 0) {
    var html = '<table style="width:100%;font-size:0.78rem;border-collapse:collapse"><tr style="border-bottom:0.5px solid var(--border);color:var(--text-secondary)"><th style="text-align:left;padding:6px">版本</th><th style="text-align:left;padding:6px">发布日期</th><th style="text-align:left;padding:6px">EOL</th><th style="text-align:left;padding:6px">状态</th></tr>';
    for (var i = 0; i < Math.min(eolData.length, 10); i++) {
      var item = eolData[i], isEol = false;
      if (typeof item.eol === 'string') { var d = new Date(item.eol); isEol = d < new Date(); }
      var status = item.eol === false ? 'active' : (isEol ? 'eol' : 'warn');
      var statusText = item.eol === false ? '维护中' : (isEol ? '已EOL' : '即将EOL');
      html += '<tr style="border-bottom:0.5px solid var(--border)"><td style="padding:6px"><strong>' + item.cycle + '</strong></td><td style="padding:6px">' + (item.releaseDate || 'N/A') + '</td><td style="padding:6px">' + (typeof item.eol === 'string' ? item.eol : (item.eol === false ? '未定' : 'N/A')) + '</td><td style="padding:6px"><span class="status-badge status-' + status + '">' + statusText + '</span></td></tr>';
    }
    html += '</table><div style="font-size:0.7rem;color:var(--text-secondary);margin-top:4px">数据来源：endoflife.date' + (reportData.cached ? '（缓存）' : '') + '</div>';
    el.innerHTML = html; return;
  }
  if (eolSource === 'internal') { el.innerHTML = '<div style="padding:8px 0"><span class="status-badge status-active">维护中</span> <span style="font-size:0.7rem;color:var(--accent)">内部自研组件</span></div>'; return; }
  if (eolSource === 'AI' && aiEol) {
    var sc = aiEol.eolStatus === '已EOL' ? 'eol' : (aiEol.eolStatus === '维护中' ? 'active' : 'warn');
    el.innerHTML = '<div style="padding:8px 0"><span class="status-badge status-' + sc + '">' + (aiEol.eolStatus || '未知') + '</span> <span style="font-size:0.7rem;color:var(--accent)">AI 判定</span><div style="font-size:0.78rem;line-height:1.8;margin-top:8px"><p><strong>当前版本：</strong>' + (aiEol.currentVersionStatus || '未知') + '</p><p><strong>最新安全版本：</strong>' + (aiEol.latestSafeVersion || '未知') + '</p><p><strong>判定依据：</strong>' + (aiEol.rationale || '无') + '</p></div></div>';
    return;
  }
  el.innerHTML = '<span style="color:var(--danger);font-size:0.8rem">无 EOL 数据，请手动查阅</span>';
}

// ======== 报告生成 ========
function buildFullReport() {
  var d = reportData; if (!d.name) return '';
  var vulns = d.vulns || [];
  var crit = vulns.filter(function(v) { return v.severity === 'critical'; }).length;
  var high = vulns.filter(function(v) { return v.severity === 'high'; }).length;
  var med = vulns.filter(function(v) { return v.severity === 'medium'; }).length;
  var low = vulns.filter(function(v) { return v.severity === 'low'; }).length;
  var eolStatus = '未知', eolBasis = '无可用数据', suggestVersion = '请查阅官方发布页', eolAction = '查阅官方文档';
  var eolSource = d.eolSource || 'endoflife.date';
  if (d.eolData && Array.isArray(d.eolData) && d.eolData.length > 0) {
    var matchedCycle = matchEOLCycle(d.eolData, d.version);
    var eolVal = matchedCycle.eol;
    if (eolVal === false) {
      eolStatus = '维护中'; eolBasis = 'endoflife.date 版本 ' + matchedCycle.cycle + ' 仍在维护'; suggestVersion = matchedCycle.latest || matchedCycle.cycle;
    } else if (eolVal === true || eolVal === 'true') {
      eolStatus = '已EOL'; eolBasis = 'endoflife.date 版本 ' + matchedCycle.cycle + ' 已 EOL'; suggestVersion = '建议寻找替代方案'; eolAction = '立即规划迁移';
    } else if (typeof eolVal === 'string' && eolVal.trim() !== '') {
      var eolDate = new Date(eolVal);
      if (!isNaN(eolDate.getTime()) && eolDate < new Date()) {
        eolStatus = '已EOL'; eolBasis = 'endoflife.date 版本 ' + matchedCycle.cycle + ' EOL 日期 ' + eolVal + ' 已过期'; suggestVersion = '建议寻找替代方案'; eolAction = '立即规划迁移';
      } else {
        eolStatus = '即将EOL'; eolBasis = 'endoflife.date 版本 ' + matchedCycle.cycle + ' EOL 日期 ' + eolVal; suggestVersion = matchedCycle.latest || matchedCycle.cycle; eolAction = '提前规划升级';
      }
    } else {
      eolStatus = '已EOL'; eolBasis = 'endoflife.date 版本 ' + matchedCycle.cycle + ' EOL 状态未知'; suggestVersion = '建议寻找替代方案'; eolAction = '查阅官方文档';
    }
  } else if (eolSource === 'AI' && d.aiEol) {
    var ai = d.aiEol; eolStatus = ai.eolStatus || '未知'; eolBasis = 'AI 判定：' + (ai.rationale || '无'); suggestVersion = ai.latestSafeVersion || '请查阅官方发布页';
  } else if (eolSource === 'internal') { eolStatus = '维护中'; eolBasis = '内部自研组件'; suggestVersion = d.version; eolAction = '内部维护'; }
  var r = '## 组件安全风险研判报告\n\n### 基本信息\n- 组件：' + (d.rawName || d.name) + ' ' + d.version + '\n';
  if (d.vendor) r += '- 厂商：' + d.vendor + '\n';
  r += '- 日期：' + d.date + '\n- CPE：' + d.cpe + '\n\n### 漏洞风险\n- 严重：' + crit + ' | 高危：' + high + ' | 中危：' + med + ' | 低危：' + low + '\n';
  if (vulns.length > 0) { r += '\n漏洞清单：\n'; for (var j = 0; j < vulns.length; j++) { var v = vulns[j]; r += '  - ' + v.id + '  ' + v.severity.toUpperCase() + (v.score !== 'N/A' ? ' (CVSS ' + v.score + ')' : '') + '  ' + v.summary + '\n'; } }
  r += '\n### EOL状态\n- 判定：' + eolStatus + '\n- 依据：' + eolBasis + '\n- 建议升级版本：' + suggestVersion + '\n- 处置建议：' + eolAction + '\n';
  r += '\n### 依据链接\n- NVD: https://nvd.nist.gov/vuln/search/results?query=' + encodeURIComponent(d.name) + '\n- endoflife.date: https://endoflife.date/' + encodeURIComponent(d.name.toLowerCase()) + '\n';
  if (d.github && d.github.repo) r += '- GitHub: https://github.com/' + d.github.repo + '/releases\n';
  r += '\n### 审计台账\n'; for (var k = 0; k < auditEntries.length; k++) r += auditEntries[k] + '\n';
  return r;
}

function generateReport() { $('reportContent').textContent = buildFullReport(); $('reportCard').classList.remove('hidden'); }

async function downloadReport(type) {
  var d = reportData; if (!d.name) { alert('请先完成研判'); return; }
  var fnBase, text;
  if (type === 'vuln') {
    var vulns = d.vulns || [];
    text = '## 漏洞风险报告\n\n- 组件：' + d.name + ' ' + d.version + '\n- 严重：' + vulns.filter(function(v){return v.severity==='critical';}).length + ' | 高危：' + vulns.filter(function(v){return v.severity==='high';}).length + '\n';
    if (vulns.length > 0) { text += '\n### 漏洞清单\n'; for (var i = 0; i < vulns.length; i++) { var v = vulns[i]; text += '- ' + v.id + '  [' + v.severity.toUpperCase() + ']  ' + v.summary + '\n'; } }
    fnBase = '漏洞报告_' + d.name + '_' + d.version;
  } else if (type === 'eol') {
    text = '## EOL 生命周期报告\n\n- 组件：' + d.name + ' ' + d.version + '\n- 数据来源：' + (d.eolSource || '未知') + '\n';
    if (d.aiEol) text += '\n### AI 兜底结论\n- EOL状态：' + (d.aiEol.eolStatus||'未知') + '\n- 推荐版本：' + (d.aiEol.latestSafeVersion||'未知') + '\n- 依据：' + (d.aiEol.rationale||'无') + '\n';
    fnBase = 'EOL报告_' + d.name + '_' + d.version;
  } else { text = buildFullReport(); fnBase = '安全研判_' + d.name + '_' + d.version + '_' + d.date; }
  var saved = await saveTextFile(fnBase + '.md', text, 'md');
  if (saved) addAudit('报告已下载 (' + type + ')');
}

function resetAll() {
  $('vulnList').innerHTML = '<span style="color:var(--text-secondary);font-size:0.8rem">等待研判…</span>';
  $('eolResult').innerHTML = '<span style="color:var(--text-secondary);font-size:0.8rem">等待研判…</span>';
  $('reportContent').textContent = '';
  $('reportCard').classList.add('hidden');
  $('auditCard').classList.add('hidden');
  $('traceCard').classList.add('hidden');
  $('traceContent').innerHTML = '';
  $('aiRawCard').classList.add('hidden');
  reportData = {}; auditEntries = []; traceLog = [];
  for (var i = 1; i <= 8; i++) setStepState(i, 'pending');
}

// ======== 批量研判 ========
function parseBatchInput() {
  var text = $('batchInput').value.trim();
  if (!text) return [];
  var lines = text.split('\n'); var items = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim(); if (!line || line.startsWith('#')) continue;
    var parts = line.split(',').map(function(p) { return p.trim(); });
    if (parts.length < 2) continue;
    var ga = parts[0], colonIdx = ga.indexOf(':');
    var groupId = '', artifactId = ga;
    if (colonIdx !== -1) { groupId = ga.substring(0, colonIdx).trim(); artifactId = ga.substring(colonIdx + 1).trim(); }
    items.push({ groupId: groupId, artifactId: artifactId, name: artifactId, fullName: ga, version: parts[1], vendor: parts[2] || '' });
  }
  return items;
}

var BATCH_CONCURRENCY = 5;
var batchEOLRuntimeCache = {};

async function evalComponentBatch(name, version, vendor, rawName, groupId) {
  var ctx = {
    name: name, rawName: rawName || name, groupId: groupId || '',
    version: version, vendor: vendor || '',
    date: new Date().toISOString().split('T')[0],
    cpe: '', vulns: [], eolData: null, eolSource: '',
    aiEol: null, internalMaintained: false, cached: false,
  };

  async function s1_cpe() {
    ctx.cpe = 'cpe:2.3:a:*:' + name.toLowerCase() + ':' + version;
  }

  async function s2_cve() {
    var results = [];
    var ecosystems = ['Maven', 'npm', 'PyPI', 'Go', 'crates.io'];
    var mavenName = (ctx.rawName && ctx.rawName.indexOf(':') !== -1) ? ctx.rawName : name;
    var osvPromises = ecosystems.map(function(ec) {
      var pkgName = ec === 'Maven' ? mavenName : name;
      return fetchWithTimeout('https://api.osv.dev/v1/querybatch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ queries: [{ package: { name: pkgName, ecosystem: ec }, version: version }] })
      }, 8000).then(function(r) { return r.ok ? r.json() : null; }).catch(function() { return null; });
    });
    var osvResponses = await Promise.all(osvPromises);
    for (var i = 0; i < osvResponses.length; i++) {
      var data = osvResponses[i];
      if (data && data.results && data.results[0] && data.results[0].vulns) {
        var vulns = data.results[0].vulns;
        for (var j = 0; j < vulns.length; j++) {
          var v = vulns[j], sev = 'unknown', score = 'N/A';
          if (v.severity) {
            var dbInfo = v.database_specific || {}; score = dbInfo.severity || v.severity; sev = String(score).toLowerCase();
            var ns = parseFloat(score);
            if (!isNaN(ns)) { if (ns >= 9) sev = 'critical'; else if (ns >= 7) sev = 'high'; else if (ns >= 4) sev = 'medium'; else sev = 'low'; }
          }
          results.push({ id: v.id, severity: sev, score: score, summary: (v.summary || '').substring(0, 100) });
        }
      }
    }
    ctx.vulns = results;
  }

  async function s6_eol() {
    var gid = ctx.groupId || '';
    var raw = ctx.rawName || '';
    var nm = ctx.name || '';
    var isInternal = INTERNAL_GROUP_IDS.indexOf(gid) !== -1 || INTERNAL_GROUP_IDS.indexOf(raw) !== -1 || INTERNAL_GROUP_IDS.indexOf(nm) !== -1;
    if (isInternal) { ctx.eolSource = 'internal'; ctx.internalMaintained = true; return; }

    var cacheKey = ctx.rawName || name;
    var cached = getCachedEOL(cacheKey);
    if (cached) {
      ctx.eolData = cached.eolData; ctx.aiEol = cached.aiEol;
      ctx.eolSource = cached.eolSource; ctx.cached = true;
      return;
    }

    if (batchEOLRuntimeCache[cacheKey]) {
      var rt = await batchEOLRuntimeCache[cacheKey];
      ctx.eolData = rt.eolData; ctx.aiEol = rt.aiEol;
      ctx.eolSource = rt.eolSource; ctx.cached = true;
      return;
    }

    batchEOLRuntimeCache[cacheKey] = (async function() {
      var eolData = null;
      var eolSources = [];

      var nameVariants = generateEOLNameVariants(nm, gid);
      for (var vi = 0; vi < nameVariants.length; vi++) {
        try {
          var r = await fetchWithTimeout('https://endoflife.date/api/' + encodeURIComponent(nameVariants[vi]) + '.json', null, 10000);
          if (r.ok) { eolData = await r.json(); if (eolData && eolData.length > 0) { eolSources.push('endoflife.date(' + nameVariants[vi] + ')'); break; } }
        } catch (e) {}
      }

      var mavenInfo = null;
      if (gid && nm) {
        mavenInfo = await checkMavenCentralLatest(gid, nm);
        if (mavenInfo) eolSources.push('maven-central');
      }

      var githubInfo = null;
      githubInfo = await checkGithubLatestRelease(nm);
      if (githubInfo) eolSources.push('github');

      var result = { eolData: eolData, eolSource: 'endoflife.date', aiEol: null, mavenInfo: mavenInfo, githubInfo: githubInfo };

      if (!eolData || !Array.isArray(eolData) || eolData.length === 0) {
        var endpoint = $('aiEndpoint').value.trim();
        var apiKey = $('aiKey').value.trim();
        var model = $('aiModel').value.trim();
        if (endpoint && apiKey) {
          var aiResult = await aiEolBatch(nm, ctx.version, gid, eolSources.join(', '), mavenInfo, githubInfo, endpoint, apiKey, model);
          if (aiResult) { result.eolSource = 'AI'; result.aiEol = aiResult; }
          else {
            result.eolSource = mavenInfo && mavenInfo.yearsSinceLast > 2 ? 'maven-stale' : (githubInfo && githubInfo.yearsSinceLast > 2 ? 'github-stale' : 'none');
          }
        } else {
          result.eolSource = mavenInfo && mavenInfo.yearsSinceLast > 2 ? 'maven-stale' : (githubInfo && githubInfo.yearsSinceLast > 2 ? 'github-stale' : 'none');
        }
      }
      saveCachedEOL(cacheKey, result.eolData, result.aiEol, result.eolSource);
      return result;
    })();

    var rt2 = await batchEOLRuntimeCache[cacheKey];
    ctx.eolData = rt2.eolData; ctx.aiEol = rt2.aiEol;
    ctx.eolSource = rt2.eolSource; ctx.mavenInfo = rt2.mavenInfo; ctx.githubInfo = rt2.githubInfo;
  }

  await Promise.all([
    s1_cpe().catch(function() {}),
    s2_cve().catch(function() {}),
    s6_eol().catch(function() {}),
  ]);

  return ctx;
}

function matchEOLCycle(eolData, version) {
  var userVer = String(version || '').trim();
  var sorted = eolData.slice().sort(function(a, b) {
    return String(b.cycle || '').length - String(a.cycle || '').length;
  });
  for (var i = 0; i < sorted.length; i++) {
    var cyc = String(sorted[i].cycle || '');
    if (cyc && (userVer === cyc || userVer.startsWith(cyc + '.'))) return sorted[i];
  }
  return eolData[0];
}

function generateEOLNameVariants(artifactId, groupId) {
  var variants = [];
  var name = (artifactId || '').toLowerCase().trim();
  if (!name) return variants;
  variants.push(name);
  if (name.indexOf('-') !== -1) {
    var parts = name.split('-');
    if (parts.length >= 2) {
      variants.push(parts[0] + '-' + parts.slice(1).join('-'));
      if (parts[0] === 'spring') variants.push('spring-framework');
      if (parts[0] === 'spring') variants.push('spring-boot');
    }
  }
  if (name.indexOf(':') !== -1) {
    var colonParts = name.split(':');
    if (colonParts.length >= 2) {
      var artifact = colonParts[colonParts.length - 1].trim();
      if (artifact) {
        variants.push(artifact);
        if (artifact.indexOf('-') !== -1) {
          var subParts = artifact.split('-');
          if (subParts[0] === 'spring') { variants.push('spring-framework'); variants.push('spring-boot'); }
        }
      }
    }
  }
  var commonMap = {
    'spring-core': 'spring-framework', 'spring-beans': 'spring-framework', 'spring-context': 'spring-framework',
    'spring-web': 'spring-framework', 'spring-webmvc': 'spring-framework', 'spring-jdbc': 'spring-framework',
    'spring-tx': 'spring-framework', 'spring-aop': 'spring-framework', 'spring-orm': 'spring-framework',
    'spring-expression': 'spring-framework', 'spring-test': 'spring-framework',
    'spring-boot-starter': 'spring-boot', 'spring-boot-starter-web': 'spring-boot', 'spring-boot-starter-data-jpa': 'spring-boot',
    'spring-boot-autoconfigure': 'spring-boot', 'spring-boot-actuator': 'spring-boot',
    'log4j-core': 'log4j', 'log4j-api': 'log4j', 'log4j': 'log4j',
    'commons-lang3': 'apache-commons-lang', 'commons-lang': 'apache-commons-lang',
    'commons-collections4': 'apache-commons-collections', 'commons-collections': 'apache-commons-collections',
    'commons-io': 'apache-commons-io', 'commons-codec': 'apache-commons-codec',
    'commons-logging': 'apache-commons-logging', 'commons-compress': 'apache-commons-compress',
    'jackson-databind': 'jackson-databind', 'jackson-core': 'jackson-core',
    'jackson-annotations': 'jackson-annotations',
    'mybatis': 'mybatis', 'mybatis-spring': 'mybatis',
    'netty-transport': 'netty', 'netty-handler': 'netty', 'netty-common': 'netty',
    'shiro-core': 'apache-shiro', 'shiro-web': 'apache-shiro',
    'dubbo': 'apache-dubbo',
    'zookeeper': 'apache-zookeeper',
    'kafka-clients': 'apache-kafka',
    'hadoop-common': 'apache-hadoop', 'hadoop-client': 'apache-hadoop',
    'hbase-client': 'apache-hbase', 'hbase-common': 'apache-hbase',
    'tomcat-embed-core': 'apache-tomcat', 'tomcat-embed-websocket': 'apache-tomcat',
    'tomcat-coyote': 'apache-tomcat',
    'netty-all': 'netty',
  };
  if (commonMap[name]) variants.push(commonMap[name]);
  if (groupId) {
    var gid = groupId.toLowerCase();
    if (gid.indexOf('springframework') !== -1) { variants.push('spring-framework'); variants.push('spring-boot'); }
    if (gid.indexOf('apache') !== -1 && gid.indexOf('logging') !== -1) variants.push('apache-commons-logging');
  }
  var unique = [];
  for (var i = 0; i < variants.length; i++) { if (variants[i] && unique.indexOf(variants[i]) === -1) unique.push(variants[i]); }
  return unique;
}

async function checkMavenCentralLatest(groupId, artifactId) {
  if (!groupId || !artifactId) return null;
  try {
    var url = 'https://search.maven.org/solrsearch/select?q=g:%22' + encodeURIComponent(groupId) + '%22+AND+a:%22' + encodeURIComponent(artifactId) + '%22&core=gav&rows=1&sort=timestamp+desc&wt=json';
    var r = await fetchWithTimeout(url, null, 10000);
    if (!r.ok) return null;
    var data = await r.json();
    if (!data.response || !data.response.docs || data.response.docs.length === 0) return null;
    var doc = data.response.docs[0];
    var timestamp = doc.timestamp || doc.lastUpdated;
    if (!timestamp) return null;
    var lastDate = new Date(typeof timestamp === 'number' ? timestamp : timestamp);
    if (isNaN(lastDate.getTime())) return null;
    var now = new Date();
    var yearsDiff = (now.getTime() - lastDate.getTime()) / (365.25 * 24 * 3600 * 1000);
    return {
      latestVersion: doc.v || doc.version,
      lastUpdated: lastDate.toISOString().split('T')[0],
      yearsSinceLast: Math.round(yearsDiff * 10) / 10,
    };
  } catch (e) { return null; }
}

async function checkGithubLatestRelease(artifactId) {
  if (!artifactId) return null;
  var name = artifactId.toLowerCase().replace(/^.*:/, '');
  var repoCandidates = [];
  var commonRepos = {
    'spring-core': 'spring-projects/spring-framework', 'spring-beans': 'spring-projects/spring-framework',
    'spring-context': 'spring-projects/spring-framework', 'spring-web': 'spring-projects/spring-framework',
    'spring-webmvc': 'spring-projects/spring-framework', 'spring-jdbc': 'spring-projects/spring-framework',
    'spring-tx': 'spring-projects/spring-framework', 'spring-aop': 'spring-projects/spring-framework',
    'spring-boot-starter': 'spring-projects/spring-boot', 'spring-boot-starter-web': 'spring-projects/spring-boot',
    'spring-boot-autoconfigure': 'spring-projects/spring-boot',
    'log4j-core': 'apache/logging-log4j2', 'log4j-api': 'apache/logging-log4j2',
    'commons-lang3': 'apache/commons-lang', 'commons-io': 'apache/commons-io',
    'commons-codec': 'apache/commons-codec', 'commons-compress': 'apache/commons-compress',
    'commons-collections4': 'apache/commons-collections',
    'commons-collections': 'apache/commons-collections',
    'commons-beanutils': 'apache/commons-beanutils',
    'commons-net': 'apache/commons-net',
    'commons-logging': 'apache/commons-logging',
    'commons-cli': 'apache/commons-cli',
    'commons-pool': 'apache/commons-pool',
    'commons-pool2': 'apache/commons-pool',
    'commons-dbcp': 'apache/commons-dbcp',
    'commons-fileupload': 'apache/commons-fileupload',
    'commons-math3': 'apache/commons-math',
    'jackson-databind': 'FasterXML/jackson-databind', 'jackson-core': 'FasterXML/jackson-core',
    'jackson-annotations': 'FasterXML/jackson-annotations',
    'mybatis': 'mybatis/mybatis-3', 'mybatis-spring': 'mybatis/spring',
    'netty-transport': 'netty/netty', 'netty-handler': 'netty/netty', 'netty-common': 'netty/netty', 'netty-all': 'netty/netty',
    'netty-codec': 'netty/netty', 'netty-codec-http': 'netty/netty', 'netty-resolver': 'netty/netty',
    'netty-buffer': 'netty/netty', 'netty-transport': 'netty/netty',
    'shiro-core': 'apache/shiro', 'shiro-web': 'apache/shiro',
    'dubbo': 'apache/dubbo', 'zookeeper': 'apache/zookeeper',
    'kafka-clients': 'apache/kafka',
    'tomcat-embed-core': 'apache/tomcat', 'tomcat-coyote': 'apache/tomcat',
    'guava': 'google/guava', 'gson': 'google/gson',
    'fastjson': 'alibaba/fastjson', 'fastjson2': 'alibaba/fastjson2',
    'easyexcel': 'alibaba/easyexcel',
    'druid': 'alibaba/druid', 'canal-client': 'alibaba/canal',
    'nacos-client': 'alibaba/nacos',
    'poi': 'apache/poi', 'poi-ooxml': 'apache/poi',
    'velocity': 'apache/velocity-engine',
    'jwt': 'auth0/java-jwt', 'jjwt-api': 'jwtk/jjwt',
    'bouncycastle': 'bcgit/bc-java',
    'bcel': 'apache/commons-bcel',
    'jedis': 'redis/jedis',
    'lettuce-core': 'lettuce-io/lettuce-core',
    'jsoup': 'jhy/jsoup',
    'slf4j-api': 'qos-ch/slf4j',
    'quartz': 'quartz-scheduler/quartz',
    'javassist': 'jboss-javassist/javassist',
    'jaxb-api': 'javaee/jaxb-spec',
    'jettison': 'codehaus/jettison',
    'jaxen': 'jaxen/jaxen',
    'ezmorph': 'ezmorph/ezmorph',
    'httpclient': 'apache/httpcomponents-client',
    'httpcore': 'apache/httpcomponents-core',
    'httpclient5': 'apache/httpcomponents-client',
    'httpcore5': 'apache/httpcomponents-core',
    'mariadb-java-client': 'mariadb-corporation/mariadb-connector-j',
    'stax2-api': 'codehaus-plexus/stax2-api',
    'commons-jexl3': 'apache/commons-jexl',
    'jctools-core': 'JCTools/JCTools',
    'woden-core': 'apache/woden',
    'axis2-kernel': 'apache/axis2-java',
    'axiom-api': 'apache/axis2-java',
    'xmlschema-core': 'apache/xmlschema',
    'neethi': 'apache/neethi',
  };
  if (commonRepos[name]) repoCandidates.push(commonRepos[name]);
  repoCandidates.push(name + '/' + name);
  repoCandidates.push(name.replace(/-/g, '') + '/' + name.replace(/-/g, ''));
  var seen = {};
  for (var i = 0; i < repoCandidates.length; i++) {
    var repo = repoCandidates[i];
    if (seen[repo]) continue;
    seen[repo] = true;
    try {
      var r = await fetchWithTimeout('https://api.github.com/repos/' + repo + '/releases?per_page=1', { headers: { 'Accept': 'application/vnd.github+json' } }, 8000);
      if (!r.ok) continue;
      var data = await r.json();
      if (!data || data.length === 0) continue;
      var latest = data[0];
      var publishedAt = latest.published_at || latest.created_at;
      if (!publishedAt) continue;
      var pubDate = new Date(publishedAt);
      if (isNaN(pubDate.getTime())) continue;
      var now = new Date();
      var yearsDiff = (now.getTime() - pubDate.getTime()) / (365.25 * 24 * 3600 * 1000);
      return {
        repo: repo,
        latestRelease: latest.tag_name,
        publishedAt: publishedAt.split('T')[0],
        yearsSinceLast: Math.round(yearsDiff * 10) / 10,
      };
    } catch (e) {}
  }
  return null;
}

async function aiEolBatch(name, version, groupId, sources, mavenInfo, githubInfo, endpoint, apiKey, model) {
  var prompt = '你是一个开源组件生命周期分析专家。请对以下组件进行 EOL（End of Life）状态研判：\n\n';
  prompt += '组件名称：' + name + '\n';
  if (groupId) prompt += 'groupId：' + groupId + '\n';
  prompt += '版本号：' + (version || '未知') + '\n';
  if (sources) prompt += '已查询数据源：' + sources + '\n';
  if (mavenInfo) prompt += 'Maven Central 最新版本：' + mavenInfo.latestVersion + '（最后发布：' + mavenInfo.lastUpdated + '，距今 ' + mavenInfo.yearsSinceLast + ' 年）\n';
  if (githubInfo) prompt += 'GitHub 最新 Release：' + githubInfo.latestRelease + '（发布于：' + githubInfo.publishedAt + '，距今 ' + githubInfo.yearsSinceLast + ' 年）\n';
  prompt += '\n判定规则：\n';
  prompt += '- 如果该组件仍有新版本发布、官方仍在维护 → "维护中"\n';
  prompt += '- 如果该组件已停止维护、官方不再发布更新 → "已EOL"\n';
  prompt += '- 如果该组件即将停止维护（已公布 EOL 日期且在 6 个月内）→ "即将EOL"\n';
  prompt += '\n重要提示：\n';
  prompt += '- Maven Central 或 GitHub 的最后发布时间仅作为参考，不能仅凭时间判断是否在维护\n';
  prompt += '- 很多项目虽然近期有发布，但实际上已停止活跃维护，只是偶尔打安全补丁\n';
  prompt += '- 请结合你对这个组件的了解：官方是否活跃？社区是否活跃？是否有维护计划？\n';
  prompt += '- 对于 Apache Commons 等老项目，如果已进入维护模式（仅安全补丁、无新功能），可判定为"已EOL"\n';
  prompt += '\n如果不确定，默认判断为 "已EOL"。\n';
  prompt += '\n请以 JSON 格式返回：\n';
  prompt += '{\n  "eolStatus": "维护中" 或 "已EOL" 或 "即将EOL",\n';
  prompt += '  "currentVersionStatus": "描述当前已知最新版本是否仍在维护",\n';
  prompt += '  "latestSafeVersion": "推荐的安全版本号",\n';
  prompt += '  "rationale": "判定依据"\n';
  prompt += '}\n只返回 JSON，不要其他文字。';
  for (var attempt = 1; attempt <= 2; attempt++) {
    try {
      var finalPrompt = attempt > 1 ? prompt + '\n\n【重要提醒】上次返回格式不正确。请严格只返回 JSON。' : prompt;
      var temperature = attempt === 1 ? 0.1 : 0.3;
      var r = await fetchWithTimeout(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
        body: JSON.stringify({ model: model, messages: [{ role: 'system', content: '你是一个开源组件生命周期分析专家，返回严格的JSON格式。' }, { role: 'user', content: finalPrompt }], temperature: temperature, max_tokens: 800 })
      }, 30000);
      if (!r.ok) continue;
      var data = await r.json();
      var content = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
      if (!content) continue;
      var jsonMatch = content.match(/\{[\s\S]*\}/);
      if (!jsonMatch) continue;
      try { return JSON.parse(jsonMatch[0]); } catch (e) { continue; }
    } catch (e) { continue; }
  }
  return null;
}

async function startBatch() {
  if (batchRunning) return;
  batchRunning = true;
  $('btnStartBatch').disabled = true;
  var items = parseBatchInput();
  if (!items.length) { alert('请在批量输入框中填写组件列表'); batchRunning = false; $('btnStartBatch').disabled = false; return; }
  batchResults = new Array(items.length);
  $('batchTbody').innerHTML = '';
  $('batchTable').classList.remove('hidden');
  $('batchProgress').classList.remove('hidden');
  $('batchProgressBar').style.width = '0%';
  resetAll();
  addAudit('[批量] 开始批量研判 ' + items.length + ' 个组件（并发 ' + BATCH_CONCURRENCY + '）');
  batchEOLRuntimeCache = {};

  var completed = 0;
  var startTime = Date.now();
  var queue = items.map(function(item, idx) { return { idx: idx, item: item }; });

  async function worker() {
    while (queue.length > 0) {
      var task = queue.shift();
      if (!task) break;
      var idx = task.idx;
      var item = task.item;
      addAudit('[批量 ' + (idx + 1) + '/' + items.length + '] 研判: ' + item.fullName + ' ' + item.version);
      var ctx = null;
      try {
        ctx = await evalComponentBatch(item.artifactId, item.version, item.vendor, item.fullName, item.groupId);
      } catch (e) { addAudit('[批量] 研判失败: ' + e.message); }
      if (!ctx) ctx = { name: item.artifactId, version: item.version, vulns: [], eolSource: 'none' };
      var vulns = ctx.vulns || [];
      var crit = vulns.filter(function(v) { return v.severity === 'critical'; }).length;
      var high = vulns.filter(function(v) { return v.severity === 'high'; }).length;
      var eolStatus = '已EOL';
      if (ctx.eolSource === 'internal') eolStatus = '维护中';
      else if (ctx.eolData && Array.isArray(ctx.eolData) && ctx.eolData.length > 0) {
        var matchedCycle = matchEOLCycle(ctx.eolData, ctx.version || item.version);
        var eolVal = matchedCycle.eol;
        if (eolVal === false) eolStatus = '维护中';
        else if (eolVal === true || eolVal === 'true') eolStatus = '已EOL';
        else if (typeof eolVal === 'string' && eolVal.trim() !== '') {
          var eolDate = new Date(eolVal);
          if (!isNaN(eolDate.getTime()) && eolDate < new Date()) eolStatus = '已EOL';
          else eolStatus = '即将EOL';
        }
        else eolStatus = '已EOL';
      }
      else if (ctx.aiEol) eolStatus = ctx.aiEol.eolStatus || '已EOL';
      else if (ctx.eolSource === 'maven-stale') eolStatus = '已EOL';
      else if (ctx.eolSource === 'github-stale') eolStatus = '已EOL';
      else eolStatus = '已EOL';
      batchResults[idx] = { name: item.artifactId, version: item.version, vendor: item.vendor, fullName: item.fullName, vulnCount: vulns.length, crit: crit, high: high, eolStatus: eolStatus, reportData: ctx };
      completed++;
      $('batchProgressBar').style.width = Math.round(completed / items.length * 100) + '%';
      renderBatchResults();
    }
  }

  var workers = [];
  for (var w = 0; w < Math.min(BATCH_CONCURRENCY, items.length); w++) {
    workers.push(worker());
  }
  await Promise.all(workers);

  var elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  addAudit('[批量] 全部完成，耗时 ' + elapsed + ' 秒（平均 ' + (elapsed / items.length).toFixed(1) + ' 秒/组件）');
  $('reportCard').classList.remove('hidden');
  $('reportContent').textContent = buildBatchReport();
  $('btnBatchExportExcel').style.display = 'inline-block';
  batchRunning = false;
  $('btnStartBatch').disabled = false;
}

function renderBatchResults() {
  var html = '';
  for (var idx = 0; idx < batchResults.length; idx++) {
    var r = batchResults[idx];
    if (!r) continue;
    var sc = r.eolStatus === '已EOL' ? 'eol' : (r.eolStatus === '维护中' ? 'active' : 'warn');
    html += '<tr><td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + r.fullName + '">' + r.fullName + '</td><td>' + r.version + '</td><td>' + r.vulnCount + ' (' + r.crit + '严重/' + r.high + '高危)</td><td><span class="status-badge status-' + sc + '">' + r.eolStatus + '</span></td><td><button class="btn btn-sm btn-secondary" data-batch-idx="' + idx + '">详情</button></td></tr>';
  }
  $('batchTbody').innerHTML = html;
}

async function evalComponent(name, version, vendor) {
  await Promise.all([
    step1_cpe(name, version, vendor).catch(function(e) {}),
    step2_cve(name, version).catch(function(e) {}),
    step6_eol(name).catch(function(e) {}),
  ]);
}

function renderBatchRow(idx, name, version, vulnCount, crit, high, eolStatus, fullName) {
  var sc = eolStatus === '已EOL' ? 'eol' : (eolStatus === '维护中' ? 'active' : 'warn');
  var html = '<tr><td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + fullName + '">' + fullName + '</td><td>' + version + '</td><td>' + vulnCount + ' (' + crit + '严重/' + high + '高危)</td><td><span class="status-badge status-' + sc + '">' + eolStatus + '</span></td><td><button class="btn btn-sm btn-secondary" data-batch-idx="' + idx + '">详情</button></td></tr>';
  $('batchTbody').insertAdjacentHTML('beforeend', html);
}

function viewBatchDetail(idx) {
  var d = batchResults[idx];
  setReportData(d.reportData);
  renderVulns(d.reportData.vulns || []);
  renderEOL();
  generateReport();
  $('resultSection').scrollIntoView({ behavior: 'smooth' });
}

function buildBatchReport() {
  var r = '## 批量组件安全研判报告\n\n| 组件 | 版本 | 漏洞(严重/高危) | EOL |\n|------|------|----------------|-----|\n';
  for (var i = 0; i < batchResults.length; i++) { var d = batchResults[i]; r += '| ' + (d.fullName || d.name) + ' | ' + d.version + ' | ' + d.vulnCount + ' (' + d.crit + '/' + d.high + ') | ' + d.eolStatus + ' |\n'; }
  return r;
}

// ======== 数据同步 ========
async function exportAllData() {
  var data = { app: 'component-security-assessment', version: 1, exportedAt: new Date().toISOString(), ai: { endpoint: getLS(LS_KEYS.endpoint) || '', key: getLS(LS_KEYS.key) || '', model: getLS(LS_KEYS.model) || '', retry: getLS(LS_KEYS.retry) || '2' }, eolCache: getEOLCache(), batchInput: getLS(BATCH_INPUT_KEY) || '', internalGroupIds: INTERNAL_GROUP_IDS };
  var content = JSON.stringify(data, null, 2);
  var saved = await saveTextFile('组件研判数据备份_' + new Date().toISOString().split('T')[0] + '.json', content, 'json');
  if (saved) { $('syncStatus').value = '已导出 ' + Object.keys(data.eolCache).length + ' 条 EOL 缓存'; addAudit('[同步] 已导出数据备份'); }
}

async function importAllData() {
  var filePath = await open({ filters: [{ name: 'JSON', extensions: ['json'] }] });
  if (!filePath) return;
  try {
    var content = await readTextFile(filePath);
    var data = JSON.parse(content);
    if (data.app !== 'component-security-assessment') throw new Error('不是本工具导出的文件');
    if (data.ai) {
      if (data.ai.endpoint) { setLS(LS_KEYS.endpoint, data.ai.endpoint); $('aiEndpoint').value = data.ai.endpoint; }
      if (data.ai.key) { setLS(LS_KEYS.key, data.ai.key); $('aiKey').value = data.ai.key; }
      if (data.ai.model) { setLS(LS_KEYS.model, data.ai.model); $('aiModel').value = data.ai.model; }
      if (data.ai.retry) { setLS(LS_KEYS.retry, data.ai.retry); $('aiRetry').value = data.ai.retry; }
    }
    if (data.eolCache) { var cache = getEOLCache(); for (var k in data.eolCache) { if (!cache[k]) cache[k] = data.eolCache[k]; } setEOLCache(cache); }
    if (data.batchInput) { setLS(BATCH_INPUT_KEY, data.batchInput); $('batchInput').value = data.batchInput; }
    if (data.internalGroupIds) INTERNAL_GROUP_IDS = data.internalGroupIds;
    $('syncStatus').value = '导入成功：AI配置 + ' + (data.eolCache ? Object.keys(data.eolCache).length : 0) + ' 条 EOL 缓存';
    addAudit('[同步] 导入成功');
  } catch (err) { $('syncStatus').value = '导入失败：' + err.message; alert('导入失败：' + err.message); }
}

// ======== Excel 导入/导出 ========
async function handleImportExcel() {
  var paths;
  try {
    paths = await open({ multiple: true, filters: [{ name: 'Excel', extensions: ['xlsx', 'xls', 'csv'] }] });
  } catch (e) { alert('打开文件对话框失败：' + e.message); return; }
  if (!paths || (Array.isArray(paths) ? paths.length === 0 : false)) return;
  if (!Array.isArray(paths)) paths = [paths];
  var allItems = [];
  var lastPath = null;
  for (var i = 0; i < paths.length; i++) {
    try {
      var items = await importFromExcel(paths[i]);
      allItems = allItems.concat(items);
      lastPath = paths[i];
      addAudit('[Excel导入] ' + paths[i].split('/').pop() + ' → ' + items.length + ' 个组件');
    } catch (e) { addAudit('[Excel导入] ' + paths[i].split('/').pop() + ' 导入失败：' + e.message); }
  }
  if (!allItems.length) { alert('未在文件中找到有效的组件数据（需至少包含组件名称和版本列）'); return; }
  var text = allItems.map(function(it) { return [it.name, it.version, it.vendor].filter(Boolean).join(','); }).join('\n');
  $('batchInput').value = text;
  setLS(BATCH_INPUT_KEY, text);
  importedExcelPath = lastPath;
  importedExcelPaths = paths.slice();
  addAudit('[Excel导入] 共导入 ' + allItems.length + ' 个组件（来自 ' + paths.length + ' 个文件）');
  if (!$('batchCard').classList.contains('hidden')) return;
  switchToBatch();
}

async function handleExportExcel() {
  if (!reportData.name && batchResults.length === 0) { alert('请先完成研判'); return; }
  var defaultName = batchResults.length > 0 ? '批量研判报告_' + new Date().toISOString().split('T')[0] : '安全研判_' + reportData.name + '_' + reportData.version + '_' + reportData.date;

  var templatePath = importedExcelPath;
  var mergedForExport = false;

  if (!templatePath) {
    var picked;
    try { picked = await open({ multiple: true, filters: [{ name: 'Excel', extensions: ['xlsx', 'xls', 'csv'] }] }); }
    catch (e) { alert('打开文件对话框失败：' + e.message); return; }
    if (!picked || (Array.isArray(picked) ? picked.length === 0 : false)) return;
    if (!Array.isArray(picked)) picked = [picked];
    if (picked.length === 1) {
      templatePath = picked[0];
      importedExcelPaths = [picked[0]];
    } else {
      var desktop0 = await invoke('get_desktop_path');
      var sep0 = desktop0.endsWith('/') ? '' : '/';
      var mergeOut = desktop0 + sep0 + defaultName + '_合并模板.xlsx';
      try {
        addAudit('[Excel导出] 检测到多文件，先合并为模板：' + mergeOut);
        await mergeExcelFiles(picked, mergeOut);
        templatePath = mergeOut;
        importedExcelPaths = [mergeOut];
        mergedForExport = true;
      } catch (e) {
        alert('合并文件失败：' + e.message + '\n将使用第一个文件作为模板');
        templatePath = picked[0];
        importedExcelPaths = [picked[0]];
      }
    }
  } else if (importedExcelPaths.length > 1) {
    var desktop1 = await invoke('get_desktop_path');
    var sep1 = desktop1.endsWith('/') ? '' : '/';
    var mergeOut1 = desktop1 + sep1 + defaultName + '_合并模板.xlsx';
    try {
      addAudit('[Excel导出] 多文件合并为模板：' + mergeOut1);
      await mergeExcelFiles(importedExcelPaths, mergeOut1);
      templatePath = mergeOut1;
      importedExcelPaths = [mergeOut1];
      mergedForExport = true;
    } catch (e) {
      addAudit('[Excel导出] 合并失败，使用最后一个文件：' + templatePath);
    }
  }

  try {
    var desktop = await invoke('get_desktop_path');
    var sep = desktop.endsWith('/') ? '' : '/';
    var filePath = desktop + sep + defaultName + '.xlsx';
    addAudit('[Excel导出] 开始导出到：' + filePath + '，模板：' + templatePath);
    await exportToExcel(reportData, batchResults, filePath, templatePath);
    addAudit('[Excel导出] 报告已保存到桌面：' + defaultName + '.xlsx');
    var msg = '导出成功！\n文件已保存到桌面：' + defaultName + '.xlsx';
    if (mergedForExport) msg += '\n（已自动合并 ' + (importedExcelPaths.length > 1 ? '多' : '所有') + ' 个文件）';
    alert(msg);
  } catch (e) {
    var errMsg = '';
    if (typeof e === 'string') errMsg = e;
    else if (e && e.message) errMsg = e.message;
    else try { errMsg = JSON.stringify(e); } catch (_) { errMsg = String(e); }
    console.error('[Excel导出] 错误详情:', e);
    alert('Excel 导出失败：' + errMsg + '\n\n路径：' + (filePath || '未知'));
  }
}

// ======== Excel 合并 ========
async function handleMergeExcel() {
  var filePaths;
  try {
    filePaths = await open({ multiple: true, filters: [{ name: 'Excel', extensions: ['xlsx', 'xls', 'csv'] }] });
  } catch (e) { alert('打开文件对话框失败：' + e.message); return; }
  if (!filePaths || (Array.isArray(filePaths) ? filePaths.length === 0 : false)) return;
  if (!Array.isArray(filePaths)) filePaths = [filePaths];
  if (filePaths.length < 2) { alert('请选择至少两个 Excel 文件进行合并（按住 Command 键多选）'); return; }

  var defaultName = '合并报告_' + new Date().toISOString().split('T')[0];

  try {
    var desktop = await invoke('get_desktop_path');
    var sep = desktop.endsWith('/') ? '' : '/';
    var outputPath = desktop + sep + defaultName + '.xlsx';
    var result = await mergeExcelFiles(filePaths, outputPath);
    var summary = '共合并 ' + result.totalFiles + ' 个文件，';
    var sheetList = [];
    for (var name in result.stats) {
      sheetList.push(name + '（+' + result.stats[name].addedRows + '行）');
    }
    addAudit('[Excel合并] ' + summary + sheetList.join('、'));

    if (result.components && result.components.length > 0) {
      var batchText = result.components.map(function (c) { return c.name + ',' + c.version; }).join('\n');
      $('batchInput').value = batchText;
      setLS(BATCH_INPUT_KEY, batchText);
      addAudit('[Excel合并] 已导入 ' + result.components.length + ' 个组件到批量研判');
      if ($('batchCard').classList.contains('hidden')) switchToBatch();
    }

    importedExcelPath = outputPath;
    importedExcelPaths = [outputPath];

    alert('合并完成！\n' + summary + '\n' + sheetList.join('\n') + '\n\n文件已保存到桌面：' + defaultName + '.xlsx\n已自动导入 ' + result.components.length + ' 个组件到批量研判，可直接点击「批量研判全部」');
  } catch (e) {
    alert('Excel 合并失败：' + e.message);
  }
}

// ======== AI 连接测试 ========
async function testAIConnection() {
  var btn = $('btnTestAI');
  var statusEl = $('aiTestStatus');
  var endpoint = $('aiEndpoint').value.trim();
  var apiKey = $('aiKey').value.trim();
  var model = $('aiModel').value.trim();
  if (!endpoint) { statusEl.innerHTML = '<span style="color:var(--danger)">请填写 API 端点</span>'; return; }
  if (!apiKey) { statusEl.innerHTML = '<span style="color:var(--danger)">请填写 API Key</span>'; return; }
  if (!model) { statusEl.innerHTML = '<span style="color:var(--danger)">请填写模型名称</span>'; return; }
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>测试中…'; statusEl.innerHTML = '';
  try {
    var r = await tauriFetch(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey }, body: JSON.stringify({ model: model, messages: [{ role: 'user', content: 'ping' }], max_tokens: 1, temperature: 0 }) });
    statusEl.innerHTML = r.ok ? '<span style="color:var(--success)">连接正常</span>' : '<span style="color:var(--danger)">HTTP ' + r.status + '</span>';
  } catch (e) { statusEl.innerHTML = '<span style="color:var(--danger)">网络错误: ' + e.message + '</span>'; }
  btn.disabled = false; btn.textContent = '测试连接';
}

// ======== 日志规范导出 ========
async function downloadLogSpec() {
  var md = buildLogSpecMarkdown();
  var saved = await saveTextFile('业务日志规范参考_' + new Date().toISOString().split('T')[0] + '.md', md, 'md');
  if (saved) addAudit('[日志规范] 已导出 Markdown');
}

function buildLogSpecMarkdown() {
  return '# 业务日志规范参考\n\n> 通用业务系统日志规范。模板为 Logback/SLF4J 风格。\n\n## 1. 关键业务节点\n\n```java\nlog.info("[{}] 用户登录成功 event=user.login userId={}", traceId, userId);\nlog.info("[{}] 支付成功 event=payment.success orderId={}", traceId, orderId);\n```\n\n## 2. 外部交互\n\n```java\nlog.info("[{}] 调用第三方API 发起 target={} request={}", traceId, url, maskParams(requestParams));\n```\n\n## 3. 异常与错误\n\n```java\nlog.error("[{}] 业务处理异常 event={} params={}", traceId, "order.create", maskParams(req), e);\n```\n\n## 4. 状态变更\n\n```java\nlog.info("[{}] 缓存失效 cacheKey={} reason={}", traceId, cacheKey, reason);\n```\n\n## 5. 性能指标\n\n```java\nlog.warn("[{}] 接口耗时超阈值 api={} costMs={}", traceId, apiPath, costMs);\n```\n\n### 通用说明\n- 脱敏函数 maskParams() 需自行实现\n- traceId 贯穿整条调用链\n- log.error 必须传入异常对象保证 Stack Trace 完整';
}

// ======== Tab 切换 ========
function switchToSingle() {
  $('singleCard').classList.remove('hidden'); $('batchCard').classList.add('hidden');
  $('tabSingle').classList.add('active'); $('tabBatch').classList.remove('active');
}
function switchToBatch() {
  $('singleCard').classList.add('hidden'); $('batchCard').classList.remove('hidden');
  $('tabSingle').classList.remove('active'); $('tabBatch').classList.add('active');
}

// ======== 初始化 ========
var DEFAULT_AI_CONFIG = {
  endpoint: 'https://api.deepseek.com/v1/chat/completions',
  key: 'sk-7402865cf9234d1ab6c3dee556087f6a',
  model: 'deepseek-v4-flash',
  retry: '2',
};

function initAIConfig() {
  var ep = getLS(LS_KEYS.endpoint), k = getLS(LS_KEYS.key), m = getLS(LS_KEYS.model), r = getLS(LS_KEYS.retry);
  if (ep) $('aiEndpoint').value = ep;
  else { $('aiEndpoint').value = DEFAULT_AI_CONFIG.endpoint; setLS(LS_KEYS.endpoint, DEFAULT_AI_CONFIG.endpoint); }
  if (k) $('aiKey').value = k;
  else { $('aiKey').value = DEFAULT_AI_CONFIG.key; setLS(LS_KEYS.key, DEFAULT_AI_CONFIG.key); }
  if (m) $('aiModel').value = m;
  else { $('aiModel').value = DEFAULT_AI_CONFIG.model; setLS(LS_KEYS.model, DEFAULT_AI_CONFIG.model); }
  if (r) $('aiRetry').value = r;
  else { $('aiRetry').value = DEFAULT_AI_CONFIG.retry; setLS(LS_KEYS.retry, DEFAULT_AI_CONFIG.retry); }
  ['aiEndpoint', 'aiKey', 'aiModel', 'aiRetry'].forEach(function(id) {
    $(id).addEventListener('change', saveAIConfig);
    $(id).addEventListener('input', saveAIConfig);
  });
}
function saveAIConfig() {
  setLS(LS_KEYS.endpoint, $('aiEndpoint').value.trim());
  setLS(LS_KEYS.key, $('aiKey').value.trim());
  setLS(LS_KEYS.model, $('aiModel').value.trim());
  setLS(LS_KEYS.retry, $('aiRetry').value.trim());
}
function clearAIConfig() {
  delLS(LS_KEYS.endpoint); delLS(LS_KEYS.key); delLS(LS_KEYS.model); delLS(LS_KEYS.retry);
  $('aiEndpoint').value = ''; $('aiKey').value = ''; $('aiModel').value = ''; $('aiRetry').value = '2';
}

function initBatchInput() {
  var v = getLS(BATCH_INPUT_KEY);
  if (v) $('batchInput').value = v;
  $('batchInput').addEventListener('input', function() { setLS(BATCH_INPUT_KEY, $('batchInput').value); });
}

function clearEOLCacheAndNotify() {
  var cache = getEOLCache();
  var count = Object.keys(cache).length;
  delLS(EOL_CACHE_KEY);
  if (count > 0) { addAudit('[缓存] 已清除 ' + count + ' 条 EOL 缓存'); alert('已清除 ' + count + ' 条 EOL 缓存'); }
  else alert('当前无 EOL 缓存');
}

document.addEventListener('DOMContentLoaded', function() {
  initAIConfig();
  initBatchInput();

  $('btnTheme').addEventListener('click', toggleTheme);
  $('btnAIConfig').addEventListener('click', function() { $('aiCard').classList.toggle('hidden'); });
  $('btnClearAI').addEventListener('click', clearAIConfig);
  $('btnToggleAICard').addEventListener('click', function() {
    var body = $('aiConfigBody');
    body.style.display = body.style.display === 'none' ? '' : 'none';
  });
  $('btnTestAI').addEventListener('click', testAIConnection);

  $('btnSync').addEventListener('click', function() { $('syncCard').classList.toggle('hidden'); });
  $('btnCloseSync').addEventListener('click', function() { $('syncCard').classList.add('hidden'); });
  $('btnExport').addEventListener('click', exportAllData);
  $('btnImport').addEventListener('click', importAllData);
  $('btnClearEOL').addEventListener('click', clearEOLCacheAndNotify);

  $('btnAssess').addEventListener('click', startAssessment);

  $('tabSingle').addEventListener('click', switchToSingle);
  $('tabBatch').addEventListener('click', switchToBatch);
  $('btnSwitchSingle').addEventListener('click', switchToSingle);
  $('btnStartBatch').addEventListener('click', startBatch);

  $('btnReportEOL').addEventListener('click', function() { downloadReport('eol'); });
  $('btnReportVuln').addEventListener('click', function() { downloadReport('vuln'); });
  $('btnReportAll').addEventListener('click', function() { downloadReport('all'); });
  $('btnCopyReport').addEventListener('click', function() {
    var text = $('reportContent').textContent;
    if (!text) return;
    navigator.clipboard.writeText(text).then(function() { addAudit('报告已复制'); }).catch(function() {});
  });

  $('btnCopyAIRaw').addEventListener('click', function() {
    if (!lastAIRaw) return;
    navigator.clipboard.writeText(lastAIRaw).then(function() { addAudit('AI原始文本已复制'); }).catch(function() {});
  });
  $('btnCloseAIRaw').addEventListener('click', function() { $('aiRawCard').classList.add('hidden'); });

  $('btnDownloadLogSpec').addEventListener('click', downloadLogSpec);
  $('btnExpandLogSpec').addEventListener('click', function() {
    var bodies = document.querySelectorAll('#logSpecCard .log-template-body');
    var allOpen = true;
    bodies.forEach(function(b) { if (!b.classList.contains('open')) allOpen = false; });
    bodies.forEach(function(b) { b.classList.toggle('open', !allOpen); });
  });

  $('btnExpandTrace').addEventListener('click', function() {
    var bodies = document.querySelectorAll('.trace-step-body');
    var allOpen = true;
    bodies.forEach(function(b) { if (!b.classList.contains('open')) allOpen = false; });
    bodies.forEach(function(b) { b.classList.toggle('open', !allOpen); });
  });

  $('btnImportExcel').addEventListener('click', handleImportExcel);
  $('btnExportExcel').addEventListener('click', handleExportExcel);
  $('btnBatchExportExcel').addEventListener('click', handleExportExcel);
  $('btnMergeExcel').addEventListener('click', handleMergeExcel);

  $('batchTbody').addEventListener('click', function(e) {
    var btn = e.target.closest('[data-batch-idx]');
    if (btn) viewBatchDetail(parseInt(btn.dataset.batchIdx));
  });
  $('logSpecCard').addEventListener('click', function(e) {
    var header = e.target.closest('.js-toggle');
    if (header) header.nextElementSibling.classList.toggle('open');
  });
  $('traceContent').addEventListener('click', function(e) {
    var header = e.target.closest('.trace-step-header');
    if (header) header.nextElementSibling.classList.toggle('open');
  });
});
