/* AI 设置弹窗（工作流配置器用；与公式助手共用同一份设置：服务商、模型、每日上限存本浏览器，Key 只存本标签页） */
(function (EFA) {
  'use strict';
  const L = EFA.LLM, U = EFA.U;
  const S = {};
  const CUSTOM = '__custom__';
  let dlg = null, onSaved = null;

  function build() {
    dlg = document.createElement('dialog');
    dlg.className = 'dialog';
    dlg.id = 'aisDlg';
    dlg.innerHTML =
      '<form method="dialog" id="aisForm">' +
      '<h2>AI 设置 <span class="muted">（可选）</span></h2>' +
      '<p class="dlg-intro">不用 AI 也能手动添加步骤。接入国内大模型后，可以用一句话描述流程，由 AI 生成步骤供你确认（需要你自己的 API Key，按用量计费）。</p>' +
      '<label class="field"><span>服务商</span><select id="aisProvider"></select></label>' +
      '<label class="field"><span id="aisKeyLabel">API Key</span><input type="password" id="aisKey" autocomplete="off" placeholder="sk-…" spellcheck="false"><small id="aisKeyWhere"></small></label>' +
      '<label class="field"><span>模型</span><select id="aisModel"></select></label>' +
      '<label class="field" id="aisCustomField" hidden><span>模型名称</span><input type="text" id="aisCustom" placeholder="填写服务商文档中的模型名" spellcheck="false"></label>' +
      '<details class="adv"><summary>高级</summary>' +
      '<label class="field" id="aisWsField"><span>工作空间 ID（可选）</span><input type="text" id="aisWs" placeholder="百炼要求使用工作空间专属地址时填写" spellcheck="false"></label>' +
      '<label class="field"><span>接口地址（留空使用官方地址）</span><input type="url" id="aisBase" spellcheck="false"></label></details>' +
      '<label class="field"><span>每天最多调用 AI 次数（0 表示不限）</span><input type="number" id="aisDaily" min="0" max="10000" step="1" inputmode="numeric"><small id="aisUsage"></small></label>' +
      '<label class="check"><input type="checkbox" id="aisSamples"><span>把每列的少量示例值一起发给 AI（理解更准确；关闭后只发送列名和类型）</span></label>' +
      '<p class="dlg-privacy">使用 AI 时，会把<b>文件的列名、数据类型、行数</b>、<b>已有步骤</b>和你写的<b>描述</b>（以及勾选时的少量示例值）发送给所选服务商；文件本身不会上传。身份证号、手机号、姓名这类列始终不发送示例值。</p>' +
      '<div class="dlg-actions"><button type="button" class="btn btn-ghost" id="aisClear">清除 Key</button><span class="spacer"></span>' +
      '<button type="button" class="btn btn-ghost" id="aisCancel">取消</button><button type="submit" class="btn btn-primary">保存</button></div>' +
      '</form>';
    document.body.append(dlg);
    const $ = id => dlg.querySelector('#' + id);
    $('aisProvider').addEventListener('change', e => fill(e.target.value));
    $('aisModel').addEventListener('change', e => { $('aisCustomField').hidden = e.target.value !== CUSTOM; });
    $('aisCancel').addEventListener('click', () => dlg.close());
    $('aisClear').addEventListener('click', () => { L.clearKey($('aisProvider').value); $('aisKey').value = ''; });
    $('aisForm').addEventListener('submit', () => save());
  }

  function fill(pid) {
    const $ = id => dlg.querySelector('#' + id);
    const s = L.getSettings(pid);
    const p = L.PROVIDERS[s.provider];
    const esc = U.escapeHtml;
    $('aisProvider').innerHTML = L.PROVIDER_LIST.map(x => '<option value="' + x.id + '"' + (x.id === s.provider ? ' selected' : '') + '>' + esc(x.name) + '</option>').join('');
    $('aisKeyLabel').textContent = p.keyLabel;
    $('aisKeyWhere').textContent = '在' + p.keyWhere + ' 创建。只保存在当前浏览器标签页，关闭页面即清除。';
    $('aisKey').value = s.apiKey;
    const preset = p.models.some(m => m.id === s.model);
    $('aisModel').innerHTML = p.models.map(m => '<option value="' + m.id + '"' + (m.id === s.model ? ' selected' : '') + '>' + esc(m.name) + '</option>').join('') +
      '<option value="' + CUSTOM + '"' + (preset ? '' : ' selected') + '>自定义模型名…</option>';
    $('aisCustom').value = preset ? '' : s.model;
    $('aisCustomField').hidden = preset;
    $('aisWsField').hidden = !p.workspaceBase;
    $('aisWs').value = s.workspace;
    $('aisBase').value = s.baseURL;
    $('aisBase').placeholder = p.base;
    $('aisSamples').checked = s.sendSamples;
    $('aisDaily').value = s.dailyCalls;
    const u = L.getUsage();
    $('aisUsage').textContent = '今天已调用 ' + u.calls + ' 次，消耗约 ' + (u.inTok + u.outTok).toLocaleString() + ' tokens（统计只在本浏览器）。';
  }

  function save() {
    const $ = id => dlg.querySelector('#' + id);
    const pid = $('aisProvider').value;
    let model = $('aisModel').value;
    if (model === CUSTOM) model = $('aisCustom').value.trim() || L.PROVIDERS[pid].models[0].id;
    const daily = Math.max(0, Math.floor(Number($('aisDaily').value)));
    L.saveSettings({ provider: pid, apiKey: $('aisKey').value.trim(), model, baseURL: $('aisBase').value, workspace: $('aisWs').value,
      sendSamples: $('aisSamples').checked, dailyCalls: Number.isFinite(daily) ? daily : undefined });
    if (onSaved) onSaved();
  }

  S.open = function (cb) {
    if (!dlg) build();
    onSaved = cb || null;
    fill();
    dlg.showModal();
    setTimeout(() => dlg.querySelector('#aisKey').focus(), 50);
  };

  EFA.AISettings = S;
})(window.EFA = window.EFA || {});
