import { computeTargetAddress } from './range-utils.js';

const STORAGE_KEY = 'excel-translator-settings-v1';

const state = {
  engine: 'bing',
  from: 'auto',
  to: 'zh-CN',
  fileBase64: '',
  fileName: '',
  inspectedSheets: [],
  settings: {
    bingApiKey: '',
    bingRegion: '',
    googleApiKey: '',
    googleEndpoint: ''
  }
};

const $ = (id) => document.getElementById(id);

function showToast(msg, type = 'info') {
  const el = $('toastBanner');
  el.textContent = msg;
  el.className = `toast ${type}`;
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (parsed.engine === 'bing' || parsed.engine === 'google') state.engine = parsed.engine;
    if (parsed.from) state.from = parsed.from;
    if (parsed.to) state.to = parsed.to;
    if (parsed.settings) state.settings = { ...state.settings, ...parsed.settings };
  } catch {
    // ignore
  }
}

function saveSettings() {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        engine: state.engine,
        from: state.from,
        to: state.to,
        settings: state.settings
      })
    );
  } catch {
    // ignore
  }
}

function updateEngineUI() {
  $('btnBing').classList.toggle('active', state.engine === 'bing');
  $('btnGoogle').classList.toggle('active', state.engine === 'google');

  const hasKey =
    state.engine === 'bing'
      ? Boolean(state.settings.bingApiKey.trim())
      : Boolean(state.settings.googleApiKey.trim());
  $('modeTag').textContent = hasKey ? '官方 API Key 模式' : '免 Key 模式';
}

function updateXlsxMappingPreview() {
  const src = $('xlsxSource').value.trim() || 'A1';
  const mode = $('xlsxMode').value;
  const custom = $('xlsxCustomTarget').value.trim();
  $('xlsxCustomWrap').classList.toggle('hidden', mode !== 'custom');
  $('xlsxMapSrc').textContent = src;

  try {
    $('xlsxMapDst').textContent = computeTargetAddress(src, mode, custom);
  } catch {
    $('xlsxMapDst').textContent = mode === 'custom' && !custom ? '请输入目标单元格' : '格式无效';
  }
}

async function checkHealth() {
  try {
    const res = await fetch('/api/health');
    if (res.ok) {
      $('serverStatus').textContent = '本地翻译服务在线';
      $('serverStatus').className = 'server-pill online';
      return;
    }
  } catch {
    // ignore
  }
  $('serverStatus').textContent = '服务未连接';
  $('serverStatus').className = 'server-pill offline';
}

async function loadLanguages() {
  try {
    const res = await fetch('/api/languages');
    const data = await res.json();
    const langs = data.languages || [];
    $('fromLang').innerHTML = '';
    $('toLang').innerHTML = '';

    for (const lang of langs) {
      const o1 = document.createElement('option');
      o1.value = lang.code;
      o1.textContent = lang.label;
      $('fromLang').appendChild(o1);

      if (lang.code !== 'auto') {
        const o2 = document.createElement('option');
        o2.value = lang.code;
        o2.textContent = lang.label;
        $('toLang').appendChild(o2);
      }
    }

    $('fromLang').value = state.from;
    $('toLang').value = state.to;
  } catch {
    // ignore
  }
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function base64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type: mimeType });
}

function renderXlsxPreview(sheetName) {
  const sheet = state.inspectedSheets.find((s) => s.name === sheetName) || state.inspectedSheets[0];
  const wrap = $('xlsxPreviewWrap');
  const table = $('xlsxPreviewTable');
  if (!sheet || !Array.isArray(sheet.previewRows) || sheet.previewRows.length === 0) {
    wrap.classList.add('hidden');
    return;
  }

  wrap.classList.remove('hidden');
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.textContent = '#';
  headRow.appendChild(corner);

  (sheet.columnRanges || []).forEach((colInfo) => {
    const th = document.createElement('th');
    th.className = 'col-head';
    th.textContent = colInfo.range ? `${colInfo.col} (${colInfo.range})` : colInfo.col;
    th.title = colInfo.range
      ? `点击选中整列非空范围 ${colInfo.range}`
      : `${colInfo.col} 列为空`;
    th.addEventListener('click', () => {
      if (!colInfo.range) {
        showToast(`${colInfo.col} 列没有可翻译的非空内容`, 'info');
        return;
      }
      $('xlsxSource').value = colInfo.range;
      updateXlsxMappingPreview();
      showToast(`已选中整列范围 ${colInfo.range}`, 'info');
    });
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  sheet.previewRows.forEach((rowVals, rIdx) => {
    const tr = document.createElement('tr');
    const rowTh = document.createElement('th');
    rowTh.textContent = String(rIdx + 1);
    tr.appendChild(rowTh);

    rowVals.forEach((val, cIdx) => {
      const td = document.createElement('td');
      const colLetter = sheet.columnRanges?.[cIdx]?.col || String.fromCharCode(65 + cIdx);
      const addr = `${colLetter}${rIdx + 1}`;
      td.className = val ? 'cell-item' : 'cell-item cell-empty';
      td.textContent = val || '(空)';
      td.title = val ? `${addr}: ${val}` : `${addr}: (空单元格)`;
      td.addEventListener('click', () => {
        $('xlsxSource').value = addr;
        updateXlsxMappingPreview();
      });
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

async function handleSelectedFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    showToast('请选择 .xlsx 格式的 Excel 文件', 'error');
    return;
  }

  const buf = await file.arrayBuffer();
  state.fileBase64 = arrayBufferToBase64(buf);
  state.fileName = file.name;
  $('fileSummary').textContent = `已选择: ${file.name} (${Math.round(file.size / 1024)} KB)`;

  try {
    const res = await fetch('/api/xlsx/inspect', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileBase64: state.fileBase64 })
    });
    const data = await res.json();
    if (res.ok && data.ok && Array.isArray(data.sheets)) {
      state.inspectedSheets = data.sheets;
      $('xlsxSheet').innerHTML = '';
      data.sheets.forEach((s) => {
        const opt = document.createElement('option');
        opt.value = s.name;
        opt.textContent = `${s.name} (${s.rowCount}行 × ${s.columnCount}列)`;
        $('xlsxSheet').appendChild(opt);
      });
      renderXlsxPreview($('xlsxSheet').value);
      showToast(`已读取工作簿，共 ${data.sheets.length} 个工作表`, 'info');
    }
  } catch (err) {
    showToast(`解析工作表失败: ${err.message}`, 'error');
  }
}

async function handleTranslateXlsx() {
  if (!state.fileBase64) {
    showToast('请先选择要翻译的 .xlsx 文件', 'error');
    return;
  }

  const btn = $('btnTranslateXlsx');
  btn.disabled = true;
  btn.textContent = '翻译处理中...';

  try {
    const res = await fetch('/api/xlsx/translate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        fileBase64: state.fileBase64,
        sheetName: $('xlsxSheet').value,
        sourceAddress: $('xlsxSource').value.trim() || 'A1',
        targetMode: $('xlsxMode').value,
        customTargetAddress: $('xlsxCustomTarget').value.trim(),
        engine: state.engine,
        from: state.from,
        to: state.to,
        config: state.settings
      })
    });
    const data = await res.json();
    if (!res.ok || !data.ok) {
      throw new Error(data.error || `HTTP ${res.status}`);
    }

    const blob = base64ToBlob(
      data.fileBase64,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const outName = state.fileName.replace(/\.xlsx$/i, '') + `_translated_${state.engine}.xlsx`;
    a.href = url;
    a.download = outName;
    a.click();
    URL.revokeObjectURL(url);

    showToast(
      `已通过 ${state.engine === 'bing' ? 'Bing' : 'Google'} 完成 ${data.sourceRange} → ${data.targetRange}（共 ${data.translatedCount} 格），已下载 ${outName}`,
      'success'
    );
  } catch (err) {
    showToast(`文件翻译失败: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '翻译 .xlsx 并下载结果';
  }
}

async function handleLiveRead() {
  try {
    const addr = encodeURIComponent($('liveSource').value.trim());
    const res = await fetch(`/api/excel/selection?address=${addr}`);
    const data = await res.json();
    if (!res.ok || !data.ok) {
      throw new Error(data.error || `HTTP ${res.status}`);
    }
    const preview = $('livePreview');
    preview.classList.remove('hidden');
    preview.textContent = `工作簿: ${data.workbook} | 工作表: ${data.sheet} | 选区: ${data.address}\n内容: ${JSON.stringify(data.values)}`;
    showToast(`已读取 Excel 选区 ${data.address}`, 'success');
  } catch (err) {
    showToast(`无法读取 Excel 窗口: ${err.message}`, 'error');
  }
}

async function handleLiveFill() {
  const btn = $('btnLiveFill');
  btn.disabled = true;
  btn.textContent = '翻译写入中...';

  try {
    const res = await fetch('/api/excel/translate-fill', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sourceAddress: $('liveSource').value.trim(),
        targetMode: $('liveMode').value,
        customTargetAddress: $('liveCustomTarget').value.trim(),
        engine: state.engine,
        from: state.from,
        to: state.to,
        config: state.settings
      })
    });
    const data = await res.json();
    if (!res.ok || !data.ok) {
      throw new Error(data.error || `HTTP ${res.status}`);
    }
    const preview = $('livePreview');
    preview.classList.remove('hidden');
    preview.textContent = `已写入 ${data.workbook} [${data.sheet}] ${data.sourceAddress} → ${data.targetAddress}\n首格译文: ${data.preview}`;
    showToast(
      `已通过 ${state.engine === 'bing' ? 'Bing' : 'Google'} 写入 ${data.targetAddress}`,
      'success'
    );
  } catch (err) {
    showToast(`写入 Excel 失败: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '翻译并写入 Excel 目标格';
  }
}

async function handleTranslatePaste() {
  const raw = $('pasteInput').value;
  if (!raw.trim()) {
    showToast('请先输入或粘贴需要翻译的文本', 'info');
    return;
  }

  const rows = raw.split(/\r?\n/).map((line) => line.split('\t'));
  const flat = rows.flat();
  const nonEmptyIdx = [];
  flat.forEach((cell, i) => {
    if (cell.trim() !== '') nonEmptyIdx.push(i);
  });

  if (nonEmptyIdx.length === 0) {
    showToast('没有可翻译的文本', 'info');
    return;
  }

  const btn = $('btnTranslatePaste');
  btn.disabled = true;
  btn.textContent = '翻译中...';

  try {
    const res = await fetch('/api/translate/batch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        texts: nonEmptyIdx.map((i) => flat[i]),
        engine: state.engine,
        from: state.from,
        to: state.to,
        config: state.settings
      })
    });
    const data = await res.json();
    if (!res.ok || !data.ok) {
      throw new Error(data.error || `HTTP ${res.status}`);
    }

    const outFlat = [...flat];
    const errors = [];
    data.results.forEach((item, idx) => {
      const targetIdx = nonEmptyIdx[idx];
      if (item.ok) outFlat[targetIdx] = item.translatedText;
      else errors.push(item.error);
    });

    if (errors.length === data.results.length) {
      throw new Error(errors[0] || '翻译失败');
    }

    let cursor = 0;
    const outText = rows
      .map((row) => {
        const lineCells = outFlat.slice(cursor, cursor + row.length);
        cursor += row.length;
        return lineCells.join('\t');
      })
      .join('\n');

    $('pasteOutput').value = outText;
    showToast(
      `已通过 ${state.engine === 'bing' ? 'Bing' : 'Google'} 完成 ${nonEmptyIdx.length - errors.length} 格翻译`,
      'success'
    );
  } catch (err) {
    showToast(`翻译失败: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '立即翻译';
  }
}

function bindEvents() {
  [$('btnBing'), $('btnGoogle')].forEach((btn) => {
    btn.addEventListener('click', () => {
      state.engine = btn.dataset.engine;
      updateEngineUI();
      saveSettings();
    });
  });

  $('fromLang').addEventListener('change', () => {
    state.from = $('fromLang').value;
    saveSettings();
  });

  $('toLang').addEventListener('change', () => {
    state.to = $('toLang').value;
    saveSettings();
  });

  $('btnSwap').addEventListener('click', () => {
    if (state.from === 'auto') {
      showToast('源语言为“自动检测”时无法互换，请先选择具体源语言', 'info');
      return;
    }
    const prev = state.from;
    state.from = state.to;
    state.to = prev;
    $('fromLang').value = state.from;
    $('toLang').value = state.to;
    saveSettings();
  });

  $('btnSaveConfig').addEventListener('click', () => {
    state.settings = {
      bingApiKey: $('bingApiKey').value.trim(),
      bingRegion: $('bingRegion').value.trim(),
      googleApiKey: $('googleApiKey').value.trim(),
      googleEndpoint: $('googleEndpoint').value.trim()
    };
    saveSettings();
    updateEngineUI();
    showToast('设置已保存', 'success');
  });

  document.querySelectorAll('.tab-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach((b) => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach((p) => p.classList.remove('active'));
      btn.classList.add('active');
      $(`tab-${btn.dataset.tab}`).classList.add('active');
    });
  });

  const dz = $('dropzone');
  dz.addEventListener('dragover', (e) => {
    e.preventDefault();
    dz.classList.add('dragover');
  });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('dragover');
    handleSelectedFile(e.dataTransfer.files[0]);
  });
  $('xlsxFileInput').addEventListener('change', (e) => handleSelectedFile(e.target.files[0]));

  $('xlsxSheet').addEventListener('change', () => renderXlsxPreview($('xlsxSheet').value));
  $('xlsxSource').addEventListener('input', updateXlsxMappingPreview);
  $('xlsxMode').addEventListener('change', updateXlsxMappingPreview);
  $('xlsxCustomTarget').addEventListener('input', updateXlsxMappingPreview);
  $('btnTranslateXlsx').addEventListener('click', handleTranslateXlsx);

  $('liveMode').addEventListener('change', () => {
    $('liveCustomWrap').classList.toggle('hidden', $('liveMode').value !== 'custom');
  });
  $('btnLiveRead').addEventListener('click', handleLiveRead);
  $('btnLiveFill').addEventListener('click', handleLiveFill);

  $('btnTranslatePaste').addEventListener('click', handleTranslatePaste);
  $('btnCopyOutput').addEventListener('click', async () => {
    const out = $('pasteOutput').value;
    if (!out) {
      showToast('暂无译文可复制', 'info');
      return;
    }
    await navigator.clipboard.writeText(out);
    showToast('译文已复制到剪贴板，可在 Excel 目标单元格直接 Ctrl+V 粘贴', 'success');
  });
}

async function init() {
  loadSettings();
  $('bingApiKey').value = state.settings.bingApiKey;
  $('bingRegion').value = state.settings.bingRegion;
  $('googleApiKey').value = state.settings.googleApiKey;
  $('googleEndpoint').value = state.settings.googleEndpoint;

  bindEvents();
  updateEngineUI();
  updateXlsxMappingPreview();
  await Promise.all([checkHealth(), loadLanguages()]);
}

init();
