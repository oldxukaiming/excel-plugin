import {
  apiFetch,
  loadTranslatorState,
  saveTranslatorState
} from './api-client.js';
import {
  isExcelHostReady,
  readCurrentSelection,
  readRangeValues,
  subscribeSelectionChanged,
  writeRangeValues
} from './excel-bridge.js';
import { computeTargetAddress } from './range-utils.js';

const loaded = loadTranslatorState();
const state = {
  ...loaded,
  sourceAddress: 'A1',
  customTargetAddress: '',
  previewText: ''
};

const els = {
  envBadge: document.getElementById('envBadge'),
  serverStatusDot: document.getElementById('serverStatusDot'),
  engineModeBadge: document.getElementById('engineModeBadge'),
  btnEngineBing: document.getElementById('btnEngineBing'),
  btnEngineGoogle: document.getElementById('btnEngineGoogle'),
  sourceLang: document.getElementById('sourceLang'),
  targetLang: document.getElementById('targetLang'),
  btnSwapLang: document.getElementById('btnSwapLang'),
  btnReadSelection: document.getElementById('btnReadSelection'),
  sourceAddress: document.getElementById('sourceAddress'),
  targetMode: document.getElementById('targetMode'),
  customTargetGroup: document.getElementById('customTargetGroup'),
  customTargetAddress: document.getElementById('customTargetAddress'),
  previewSource: document.getElementById('previewSource'),
  previewTarget: document.getElementById('previewTarget'),
  sourceTextPreview: document.getElementById('sourceTextPreview'),
  btnTranslateFill: document.getElementById('btnTranslateFill'),
  btnPreviewOnly: document.getElementById('btnPreviewOnly'),
  autoSyncSelection: document.getElementById('autoSyncSelection'),
  resultPreviewBox: document.getElementById('resultPreviewBox'),
  resultPreviewLabel: document.getElementById('resultPreviewLabel'),
  resultTextPreview: document.getElementById('resultTextPreview'),
  statusBanner: document.getElementById('statusBanner'),
  btnToggleSettings: document.getElementById('btnToggleSettings'),
  settingsChevron: document.getElementById('settingsChevron'),
  settingsBody: document.getElementById('settingsBody'),
  apiBaseUrl: document.getElementById('apiBaseUrl'),
  bingApiKey: document.getElementById('bingApiKey'),
  bingRegion: document.getElementById('bingRegion'),
  googleApiKey: document.getElementById('googleApiKey'),
  googleEndpoint: document.getElementById('googleEndpoint'),
  btnSaveSettings: document.getElementById('btnSaveSettings'),
  btnTestServer: document.getElementById('btnTestServer')
};

function persistSettings() {
  saveTranslatorState(state);
}

function showStatus(message, type = 'info') {
  els.statusBanner.textContent = message;
  els.statusBanner.className = `status-banner ${type}`;
}

function clearStatus() {
  els.statusBanner.className = 'status-banner hidden';
  els.statusBanner.textContent = '';
}

function updateEngineUI() {
  els.btnEngineBing.classList.toggle('active', state.engine === 'bing');
  els.btnEngineGoogle.classList.toggle('active', state.engine === 'google');

  const usingKey =
    state.engine === 'bing'
      ? Boolean(state.settings.bingApiKey.trim())
      : Boolean(state.settings.googleApiKey.trim());

  els.engineModeBadge.textContent = usingKey ? '官方 API Key 模式' : '免 Key 模式';
}

function updateMappingPreview() {
  const src = els.sourceAddress.value.trim() || 'A1';
  const mode = els.targetMode.value;
  const custom = els.customTargetAddress.value.trim();

  els.customTargetGroup.classList.toggle('hidden', mode !== 'custom');
  els.previewSource.textContent = src;

  try {
    const target = computeTargetAddress(src, mode, custom);
    els.previewTarget.textContent = target;
  } catch (err) {
    els.previewTarget.textContent = mode === 'custom' && !custom ? '请输入目标单元格' : '格式无效';
  }
}

function formatValuesPreview(values2D) {
  if (!Array.isArray(values2D) || values2D.length === 0) {
    return '(空)';
  }
  const flat = values2D.flat().map((v) => (v === null || v === undefined ? '' : String(v)));
  const nonEmpty = flat.filter((v) => v.trim() !== '');
  if (nonEmpty.length === 0) {
    return '(空单元格)';
  }
  if (flat.length === 1) {
    return flat[0];
  }
  return `共 ${flat.length} 个单元格（首项: ${nonEmpty[0].slice(0, 60)}）`;
}

async function syncFromExcelSelection() {
  if (!isExcelHostReady()) return;
  try {
    const selection = await readCurrentSelection();
    els.sourceAddress.value = selection.address;
    state.sourceAddress = selection.address;
    state.previewText = formatValuesPreview(selection.values);
    els.sourceTextPreview.textContent = state.previewText;
    updateMappingPreview();
  } catch (err) {
    showStatus(`读取选区失败: ${err.message}`, 'error');
  }
}

async function loadLanguages() {
  try {
    const res = await apiFetch('/api/languages', {}, state.settings);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    const langs = data.languages || [];

    els.sourceLang.innerHTML = '';
    els.targetLang.innerHTML = '';

    for (const lang of langs) {
      const optSrc = document.createElement('option');
      optSrc.value = lang.code;
      optSrc.textContent = lang.label;
      els.sourceLang.appendChild(optSrc);

      if (lang.code !== 'auto') {
        const optDst = document.createElement('option');
        optDst.value = lang.code;
        optDst.textContent = lang.label;
        els.targetLang.appendChild(optDst);
      }
    }

    els.sourceLang.value = state.from;
    els.targetLang.value = state.to;
  } catch {
    // Fallback list if API is unreachable
    els.sourceLang.innerHTML = `
      <option value="auto">自动检测 (Auto)</option>
      <option value="en">英语 (English)</option>
      <option value="zh-CN">简体中文 (zh-CN)</option>
      <option value="ja">日语 (日本語)</option>
    `;
    els.targetLang.innerHTML = `
      <option value="zh-CN">简体中文 (zh-CN)</option>
      <option value="en">英语 (English)</option>
      <option value="ja">日语 (日本語)</option>
    `;
  }
}

async function checkServerHealth() {
  try {
    const res = await apiFetch('/api/health', {}, state.settings);
    if (res.ok) {
      els.serverStatusDot.className = 'status-dot online';
      els.serverStatusDot.title = '翻译服务在线';
      return true;
    }
  } catch {
    // ignore
  }
  els.serverStatusDot.className = 'status-dot offline';
  els.serverStatusDot.title = '翻译服务未连接';
  return false;
}

async function callBatchTranslate(texts) {
  const res = await apiFetch(
    '/api/translate/batch',
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        texts,
        engine: state.engine,
        from: state.from,
        to: state.to,
        config: state.settings
      })
    },
    state.settings
  );

  const data = await res.json();
  if (!res.ok || !data.ok) {
    throw new Error(data.error || `HTTP ${res.status}`);
  }
  return data.results;
}

async function handleTranslate({ fillExcel }) {
  clearStatus();
  const sourceAddr = els.sourceAddress.value.trim();
  if (!sourceAddr) {
    showStatus('请先输入或选中源单元格地址（例如 A1）', 'error');
    return;
  }

  let targetAddr = '';
  try {
    targetAddr = computeTargetAddress(
      sourceAddr,
      els.targetMode.value,
      els.customTargetAddress.value.trim()
    );
  } catch (err) {
    showStatus(`目标位置无效: ${err.message}`, 'error');
    return;
  }

  els.btnTranslateFill.disabled = true;
  els.btnPreviewOnly.disabled = true;
  els.btnTranslateFill.textContent = '翻译中...';

  try {
    let values2D;
    if (isExcelHostReady()) {
      const rangeData = await readRangeValues(sourceAddr);
      values2D = rangeData.values;
      els.sourceTextPreview.textContent = formatValuesPreview(values2D);
    } else {
      // Browser standalone preview fallback
      const fallbackText = window.prompt('当前处于浏览器预览模式，请输入要测试翻译的文本：', 'Hello Excel');
      if (fallbackText === null) {
        return;
      }
      values2D = [[fallbackText]];
      els.sourceTextPreview.textContent = fallbackText;
    }

    const rowCount = values2D.length;
    const colCount = values2D[0]?.length || 0;
    const flatTexts = [];
    const nonEmptyIndices = [];

    for (let r = 0; r < rowCount; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        const val = values2D[r][c];
        const str = val === null || val === undefined ? '' : String(val);
        if (str.trim() !== '') {
          nonEmptyIndices.push(flatTexts.length);
          flatTexts.push(str);
        } else {
          flatTexts.push('');
        }
      }
    }

    if (nonEmptyIndices.length === 0) {
      showStatus(`源区域 ${sourceAddr} 中没有可翻译的文本`, 'info');
      return;
    }

    const textsToTranslate = nonEmptyIndices.map((idx) => flatTexts[idx]);
    const results = await callBatchTranslate(textsToTranslate);

    const translatedFlat = [...flatTexts];
    const errors = [];
    results.forEach((item, i) => {
      const targetFlatIdx = nonEmptyIndices[i];
      if (item.ok) {
        translatedFlat[targetFlatIdx] = item.translatedText;
      } else {
        errors.push(item.error);
      }
    });

    if (errors.length === results.length) {
      throw new Error(errors[0] || '所有单元格翻译均失败');
    }

    const output2D = [];
    for (let r = 0; r < rowCount; r += 1) {
      const row = [];
      for (let c = 0; c < colCount; c += 1) {
        row.push(translatedFlat[r * colCount + c]);
      }
      output2D.push(row);
    }

    const firstTranslated = output2D.flat().find((v) => String(v).trim() !== '') || '';
    els.resultPreviewBox.classList.remove('hidden');
    els.resultPreviewLabel.textContent = `译文结果（引擎: ${state.engine === 'bing' ? 'Bing' : 'Google'}）`;
    els.resultTextPreview.textContent =
      output2D.flat().length === 1
        ? firstTranslated
        : `已完成 ${nonEmptyIndices.length} 个单元格（首项: ${firstTranslated}）`;

    if (fillExcel && isExcelHostReady()) {
      await writeRangeValues(targetAddr, output2D);
      const engineName = state.engine === 'bing' ? 'Bing' : 'Google';
      if (errors.length > 0) {
        showStatus(
          `已通过 ${engineName} 写入 ${targetAddr}（部分失败 ${errors.length} 项: ${errors[0]}）`,
          'error'
        );
      } else {
        showStatus(`已通过 ${engineName} 将 ${sourceAddr} 翻译填充至 ${targetAddr}`, 'success');
      }
    } else if (fillExcel) {
      showStatus(`浏览器预览完成（未连接 Excel，目标地址为 ${targetAddr}）`, 'info');
    } else {
      showStatus(`预览完成（当前引擎: ${state.engine === 'bing' ? 'Bing' : 'Google'}）`, 'success');
    }
  } catch (err) {
    showStatus(`翻译失败: ${err.message}`, 'error');
  } finally {
    els.btnTranslateFill.disabled = false;
    els.btnPreviewOnly.disabled = false;
    els.btnTranslateFill.textContent = '翻译并填充到目标单元格';
  }
}

function bindEvents() {
  [els.btnEngineBing, els.btnEngineGoogle].forEach((btn) => {
    btn.addEventListener('click', () => {
      state.engine = btn.dataset.engine;
      updateEngineUI();
      persistSettings();
    });
  });

  els.sourceLang.addEventListener('change', () => {
    state.from = els.sourceLang.value;
    persistSettings();
  });

  els.targetLang.addEventListener('change', () => {
    state.to = els.targetLang.value;
    persistSettings();
  });

  els.btnSwapLang.addEventListener('click', () => {
    if (state.from === 'auto') {
      showStatus('源语言为“自动检测”时无法直接互换，请先选择具体源语言', 'info');
      return;
    }
    const prevFrom = state.from;
    state.from = state.to;
    state.to = prevFrom;
    els.sourceLang.value = state.from;
    els.targetLang.value = state.to;
    persistSettings();
  });

  els.sourceAddress.addEventListener('input', updateMappingPreview);
  els.customTargetAddress.addEventListener('input', updateMappingPreview);
  els.targetMode.addEventListener('change', () => {
    state.targetMode = els.targetMode.value;
    updateMappingPreview();
    persistSettings();
  });

  els.btnReadSelection.addEventListener('click', syncFromExcelSelection);
  els.btnTranslateFill.addEventListener('click', () => handleTranslate({ fillExcel: true }));
  els.btnPreviewOnly.addEventListener('click', () => handleTranslate({ fillExcel: false }));

  els.btnToggleSettings.addEventListener('click', () => {
    const isHidden = els.settingsBody.classList.toggle('hidden');
    els.btnToggleSettings.setAttribute('aria-expanded', String(!isHidden));
    els.settingsChevron.textContent = isHidden ? '▸' : '▾';
  });

  els.btnSaveSettings.addEventListener('click', () => {
    state.settings = {
      apiBaseUrl: els.apiBaseUrl.value.trim(),
      bingApiKey: els.bingApiKey.value.trim(),
      bingRegion: els.bingRegion.value.trim(),
      googleApiKey: els.googleApiKey.value.trim(),
      googleEndpoint: els.googleEndpoint.value.trim()
    };
    persistSettings();
    updateEngineUI();
    checkServerHealth();
    showStatus('高级设置已保存', 'success');
  });

  els.btnTestServer.addEventListener('click', async () => {
    const ok = await checkServerHealth();
    showStatus(ok ? '翻译服务连接正常' : '无法连接到翻译服务', ok ? 'success' : 'error');
  });
}

async function init() {
  // Sync settings inputs
  els.targetMode.value = state.targetMode;
  els.apiBaseUrl.value = state.settings.apiBaseUrl;
  els.bingApiKey.value = state.settings.bingApiKey;
  els.bingRegion.value = state.settings.bingRegion;
  els.googleApiKey.value = state.settings.googleApiKey;
  els.googleEndpoint.value = state.settings.googleEndpoint;

  bindEvents();
  updateEngineUI();
  updateMappingPreview();

  await Promise.all([loadLanguages(), checkServerHealth()]);

  if (typeof globalThis.Office !== 'undefined' && typeof globalThis.Office.onReady === 'function') {
    globalThis.Office.onReady(async (info) => {
      if (info.host === globalThis.Office.HostType.Excel && isExcelHostReady()) {
        els.envBadge.textContent = '已连接到 Microsoft Excel';
        await syncFromExcelSelection();
        await subscribeSelectionChanged((selection) => {
          if (!els.autoSyncSelection.checked) return;
          els.sourceAddress.value = selection.address;
          state.sourceAddress = selection.address;
          els.sourceTextPreview.textContent = formatValuesPreview(selection.values);
          updateMappingPreview();
        });
      } else {
        els.envBadge.textContent = '浏览器预览模式（未在 Excel 内打开）';
      }
    });
  } else {
    els.envBadge.textContent = '浏览器预览模式（Office.js 未加载）';
  }
}

init();
