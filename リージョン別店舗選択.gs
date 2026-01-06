/**
 * リージョン別店舗選択フォーム - 既存フォーム更新 & プロパティ自動設定
 *  - 従業員選択肢に部署名を追加表示（例：「樋口真二 Region A」）
 *  - 店舗選択は全リージョンで **最後のセクション**（=「報告内容の作成」があればそこ）へ遷移
 *  - セクション下部の「このセクションの最後に移動」も **全リージョンで最後のセクション**に固定
 *  - セクション内の LIST はすべて MC に置換し、全選択肢へ遷移を付与
 *  - 足りない Region セクション＆店舗選択を自動作成
 *  - 店舗0件はプレースホルダー1件で例外回避
 *  - 「報告内容の作成」セクションは末尾へ移動（moveItem）
 */

// ==================== 【設定】 ====================

// フォームの編集URL か ID（/viewform は不可）
const KIZON_FORM_ID_OR_URL = 'https://docs.google.com/forms/d/xxxxxxxxxxxxxxxxxxxxxxxx/edit';

// フォーム設定（null/undefined は未変更）
const FORM_SETTINGS = {
  title: null,
  description: null,
  confirmationMessage: '回答ありがとうございました。',
  acceptingResponses: true,
  allowResponseEdits: false,
  showLinkToRespondAgain: false,
  requireLogin: null,
  limitOneResponsePerUser: null,
  progressBar: false,
  shuffleQuestions: false,
  collectEmail: null,
  isQuiz: null
};

// 実行結果を Properties に保存するか
const LOG_TO_PROPERTIES = true;

// 報告内容セクションのタイトル
const REPORT_SECTION_TITLE = '報告内容の作成';

// 店舗ラベル合成ルール
const STORE_LABEL_ORDER = ['name','code','area','address','phone'];
const STORE_LABEL_SEPARATOR = '｜';

// 0件時のプレースホルダー
const PLACEHOLDER_LABEL = '（店舗データなし）';


// ==================== 確認/診断 ====================

function setteiKakunin() {
  Logger.log('=== 現在の設定確認 ===');
  try {
    const { form, id, editUrl, pubUrl } = resolveForm_(KIZON_FORM_ID_OR_URL);
    Logger.log('✓ フォームOK');
    Logger.log('  ID: ' + id);
    Logger.log('  名称: ' + form.getTitle());
    Logger.log('  公開URL: ' + pubUrl);
    Logger.log('  編集URL: ' + editUrl);
    return true;
  } catch (e) {
    Logger.log('✗ フォームを開けません: ' + e);
    Logger.log('ヒント: /edit の編集URL か ID を指定（/viewform は不可）/ 共同編集者権限');
    return false;
  }
}

function debugFormAccess() {
  Logger.log('=== アクセス診断 ===');
  const eff = (function(){ try { return Session.getEffectiveUser().getEmail(); } catch(_) { return '(不明)'; } })();
  Logger.log('実行ユーザー: ' + eff);

  const cand = buildCandidates_(KIZON_FORM_ID_OR_URL, true);
  Logger.log('候補: ' + JSON.stringify(cand));

  for (const c of cand) {
    Logger.log('--- 候補検証: ' + c + ' ---');
    const isUrl = /^https?:\/\//i.test(c);
    const id = isUrl ? extractFormIdFromUrl_(c) : c;

    if (id && !isUrl) {
      try { const f = DriveApp.getFileById(id); Logger.log('Drive ✓ タイトル: ' + f.getName() + ' / ゴミ箱? ' + f.isTrashed()); }
      catch (e) { Logger.log('Drive ✗ ' + e); }
    }

    try {
      const f = isUrl ? FormApp.openByUrl(c) : FormApp.openById(c);
      Logger.log('FormApp ✓ ' + f.getTitle());
      Logger.log('編集URL: ' + f.getEditUrl());
      return;
    } catch (e) {
      Logger.log('FormApp ✗ ' + e);
      if (isUrl && /viewform/i.test(c)) Logger.log('ヒント: /viewform では開けません。/edit の編集URLを使用してください。');
    }
  }
  Logger.log('=== 診断終了: 全候補不可 ===');
}


// ==================== メイン更新 ====================

function formKoushin() {
  const runInfo = { ok: false, message: '', counts: {} };
  try {
    Logger.log('=== 既存フォーム更新開始 ===');
    const { form } = resolveForm_(KIZON_FORM_ID_OR_URL);
    Logger.log('✓ フォーム: ' + form.getTitle());

    // プロパティ適用
    applyFormSettings_(form, FORM_SETTINGS);

    const { jugyoin, tenpo, regions } = dataYomikomi();
    runInfo.counts = { jugyoin: jugyoin.length, tenpo: tenpo.length, regions: regions.length };

    // 現在のアイテム
    let allItems = form.getItems();

    // 従業員設問を特定（MC優先、無ければLIST）
    let empItem = null, empType = null, empIndex = -1;
    for (let i = 0; i < allItems.length; i++) {
      const it = allItems[i];
      if (it.getType() === FormApp.ItemType.MULTIPLE_CHOICE) { empItem = it.asMultipleChoiceItem(); empType = 'MC'; empIndex = i; break; }
      if (it.getType() === FormApp.ItemType.LIST)            { empItem = it.asListItem();          empType = 'LIST'; empIndex = i; break; }
    }
    if (!empItem) { Logger.log('✗ 従業員設問が見つかりません'); runInfo.message = '従業員設問なし'; saveRunProps_(runInfo); return; }
    Logger.log(`従業員設問: [${empIndex}] タイプ=${empType}`);

    // 既存の Region セクションを収集
    let regionSections = collectRegionSections_(allItems);
    Logger.log('既存リージョン: ' + (regionSections.map(r => r.region).join(', ') || '(なし)'));

    // 報告内容セクションを検出（無ければ後で最終セクションにフォールバック）
    let reportSection = findReportSection_(allItems);

    // 足りない Region セクションと店舗選択(MC)を自動作成（暫定ターゲットは reportSection または最後）
    const prelimNavTarget = getNavigationTarget_(form, reportSection);
    const created = ensureRegionSectionsAndStores_(form, regionSections, regions, tenpo, prelimNavTarget);
    if (created > 0) Logger.log(`✓ 追加作成: Region セクション/店舗選択 ${created} セット`);

    // ── 先に「報告内容の作成」を末尾へ移動（UI上の"最後のセクション"を確定）
    if (reportSection) moveReportSectionToEnd_(form);

    // 改めて nav 先（= 最後のセクション or 報告内容の作成）を取得
    allItems = form.getItems();
    regionSections = collectRegionSections_(allItems);
    reportSection = findReportSection_(allItems);
    const navTarget = getNavigationTarget_(form, reportSection);
    const navLabel  = navTarget ? navTarget.getTitle() : '遷移なし';
    if (reportSection) Logger.log(`✓ ナビ先: ${navLabel}（報告内容の作成）`);
    else Logger.log(`（注意）報告内容セクションなし。ナビ先: ${navLabel}`);

    // 従業員設問の更新（リージョン→セクション自動割当、部署名を表示名に追加）
    if (empType === 'LIST') {
      empItem.setChoiceValues(jugyoin.map(j => j.hyojiMei));
      Logger.log(`✓ 従業員(LIST)更新: ${jugyoin.length}名`);
    } else {
      const mc = empItem;
      const choices = jugyoin.map(j => {
        const label = j.hyojiMei;
        const target = findSectionForRegion_(regionSections, j.region);
        return target ? mc.createChoice(label, target) : mc.createChoice(label);
      });
      mc.setChoices(choices);
      Logger.log(`✓ 従業員(MC)更新: ${choices.length}名 / 自動割当 byRegion`);
    }

    // 各リージョンの店舗選択を **セクション内の全問** 対象に MC 化＆全選択肢へナビ付与
    updateRegionStoresForceMC_All_(form, tenpo, navTarget);

    // ★ 各リージョンのセクション下部「このセクションの最後に移動」を最後のセクションに固定
    setRegionSectionsGoToLast_(form, navTarget);

    Logger.log('=== 更新完了 ===');
    Logger.log('公開URL: ' + form.getPublishedUrl());
    runInfo.ok = true; runInfo.message = 'OK'; saveRunProps_(runInfo);
  } catch (e) {
    Logger.log('✗ エラー: ' + e);
    Logger.log('stack: ' + e.stack);
    runInfo.ok = false; runInfo.message = e && e.message || String(e); saveRunProps_(runInfo);
    throw e;
  }
}


// ==================== データ読み込み（部署名を表示名に追加） ====================

function dataYomikomi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ユーザー
  const userSheet = ss.getSheetByName('ユーザーリスト');
  if (!userSheet) throw new Error('「ユーザーリスト」シートが見つかりません');
  if (userSheet.getLastRow() < 2) throw new Error('ユーザーリストにデータがありません');
  const userLastCol = Math.max(userSheet.getLastColumn(), 21);
  const userValues  = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, userLastCol).getValues();

  const jugyoin = [];
  userValues.forEach(row => {
    const mei   = (row[1] || '').toString().trim();
    const sei   = (row[2] || '').toString().trim();
    const busho = (row[20]|| '').toString().trim();
    
    const namePart = (sei + ' ' + mei).trim();
    if (!namePart) return;
    
    const hyojiMei = busho ? namePart + ' ' + busho : namePart;
    
    const region = regionChuushutsu(busho);
    jugyoin.push({ hyojiMei, busho, region });
  });

  // 店舗
  const tenpoSheet = ss.getSheetByName('店舗リスト');
  if (!tenpoSheet) throw new Error('「店舗リスト」シートが見つかりません');
  if (tenpoSheet.getLastRow() < 2) throw new Error('店舗リストにデータがありません');

  const lastCol = Math.max(tenpoSheet.getLastColumn(), 3);
  const values = tenpoSheet.getRange(1, 1, tenpoSheet.getLastRow(), lastCol).getValues();
  const headers = (values[0] || []).map(h => (h || '').toString().trim());
  const normMap = buildStoreHeaderMap_(headers);

  const tenpo = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name   = (row[1] || '').toString().trim();
    const region = (row[2] || '').toString().trim();
    if (!name || !region) continue;

    const obj = { name, region };
    Object.keys(normMap).forEach(key => {
      const colIndex = normMap[key];
      if (colIndex == null) return;
      const v = (row[colIndex] || '').toString().trim();
      if (v) obj[key] = v;
    });
    obj.label = composeStoreLabel_(obj);
    tenpo.push(obj);
  }

  const regions = Array.from(new Set(tenpo.map(t => t.region))).sort();
  Logger.log(`読み込み: 従業員=${jugyoin.length}, 店舗=${tenpo.length}, リージョン=${regions.join(', ')}`);
  return { jugyoin, tenpo, regions };
}

function buildStoreHeaderMap_(headers) {
  const idxByLower = {};
  headers.forEach((h, i) => { idxByLower[(h || '').toString().trim().toLowerCase()] = i; });
  const aliases = {
    code:   ['店舗コード','店番','コード','code','store code','store_code'],
    area:   ['エリア','地区','エリア名','area','region name','zone'],
    address:['住所','所在地','address','addr'],
    phone:  ['電話','電話番号','tel','電話TEL','phone','phone number']
  };
  const map = {};
  Object.keys(aliases).forEach(key => {
    const cands = aliases[key];
    let found = null;
    for (const c of cands) {
      const j = idxByLower[c.toLowerCase()];
      if (j != null) { found = j; break; }
    }
    map[key] = found;
  });
  return map;
}

function composeStoreLabel_(store) {
  const parts = [];
  STORE_LABEL_ORDER.forEach(k => {
    if (k === 'name') { if (store.name) parts.push(store.name); }
    else if (store[k]) parts.push(store[k]);
  });
  return parts.join(STORE_LABEL_SEPARATOR);
}

function regionChuushutsu(busho) {
  if (!busho) return '';
  const patterns = [/Region[\s:-]*([A-Z])/i, /リージョン[\s:-]*([A-Z])/i];
  for (const p of patterns) { const m = busho.match(p); if (m) return m[1].toUpperCase(); }
  return '';
}


// ==================== 既存フォーム構造の操作 ====================

function collectRegionSections_(allItems) {
  const out = [];
  for (let i = 0; i < allItems.length; i++) {
    const it = allItems[i];
    if (it.getType() !== FormApp.ItemType.PAGE_BREAK) continue;
    const pb = it.asPageBreakItem();
    const title = pb.getTitle() || '';
    if (title === REPORT_SECTION_TITLE) continue;
    const m = title.match(/Region[\s:-]*([A-Z])(?:\b|[^A-Za-z]|$)/i);
    if (m) out.push({ region: m[1].toUpperCase(), section: pb, index: i, title });
  }
  out.sort((a,b) => a.region.localeCompare(b.region));
  return out;
}

function findSectionForRegion_(regionSections, regionLetter) {
  if (!regionLetter) return null;
  const hit = regionSections.find(r => r.region === regionLetter.toUpperCase());
  return hit ? hit.section : null;
}

function findReportSection_(allItems) {
  for (let i = 0; i < allItems.length; i++) {
    const it = allItems[i];
    if (it.getType() !== FormApp.ItemType.PAGE_BREAK) continue;
    const pb = it.asPageBreakItem();
    if (pb.getTitle() === REPORT_SECTION_TITLE) return pb;
  }
  return null;
}

function findLastSection_(form) {
  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i--) {
    if (items[i].getType() === FormApp.ItemType.PAGE_BREAK) {
      return items[i].asPageBreakItem();
    }
  }
  return null;
}

function getNavigationTarget_(form, reportSectionMaybe) {
  if (reportSectionMaybe) return reportSectionMaybe;
  const last = findLastSection_(form);
  if (last) return last;
  return null;
}


// ==================== セクション作成/更新（全リージョン→最後のセクション遷移） ====================

function ensureRegionSectionsAndStores_(form, existingSections, regionsLetters, tenpo, navTarget) {
  const have = new Set(existingSections.map(s => s.region));
  const need = Array.from(new Set(regionsLetters.map(r => r.toUpperCase())));
  let createdCount = 0;

  const tenpoByRegion = groupTenpo_(tenpo);

  existingSections.forEach(sec => {
    const hasChoice = hasStoreChoiceInSection_(form.getItems(), sec.index);
    if (!hasChoice) {
      const mc = form.addMultipleChoiceItem().setTitle(`店舗を選択 (Region ${sec.region})`);
      const stores = tenpoByRegion[sec.region] || [];
      const labels = stores.map(st => st.label || st.name);
      const choices = buildChoicesSafe_(mc, labels, navTarget);
      mc.setChoices(choices);
      createdCount++;
      Logger.log(`✓ 既存 Region ${sec.region} に店舗選択MCを追加`);
    }
  });

  need.forEach(letter => {
    if (have.has(letter)) return;
    form.addPageBreakItem().setTitle(`Region ${letter}`);
    const mc = form.addMultipleChoiceItem().setTitle(`店舗を選択 (Region ${letter})`);
    const stores = (tenpoByRegion[letter] || []);
    const labels = stores.map(st => st.label || st.name);
    const choices = buildChoicesSafe_(mc, labels, navTarget);
    mc.setChoices(choices);
    createdCount++;
    Logger.log(`✓ Region ${letter} セクション＋店舗選択MCを新規作成`);
  });

  return createdCount;
}

function hasStoreChoiceInSection_(allItems, sectionIndex) {
  for (let i = sectionIndex + 1; i < allItems.length; i++) {
    const it = allItems[i];
    if (it.getType() === FormApp.ItemType.PAGE_BREAK) break;
    if (it.getType() === FormApp.ItemType.LIST || it.getType() === FormApp.ItemType.MULTIPLE_CHOICE) return true;
  }
  return false;
}

function groupTenpo_(tenpo) {
  const m = {};
  tenpo.forEach(t => {
    const k = (t.region || '').toUpperCase();
    if (!m[k]) m[k] = [];
    m[k].push(t);
  });
  return m;
}

function updateRegionStoresForceMC_All_(form, tenpo, navTarget) {
  const navLabel = navTarget ? navTarget.getTitle() : '遷移なし';
  const tenpoByRegion = groupTenpo_(tenpo);

  let i = 0;
  while (i < form.getItems().length) {
    const items = form.getItems();
    const it = items[i];
    if (it.getType() !== FormApp.ItemType.PAGE_BREAK) { i++; continue; }

    const pb = it.asPageBreakItem();
    const title = pb.getTitle() || '';
    if (title === REPORT_SECTION_TITLE) { i++; continue; }

    const m = title.match(/Region[\s:-]*([A-Z])(?:\b|[^A-Za-z]|$)/i);
    if (!m) { i++; continue; }
    const region = m[1].toUpperCase();
    const stores = tenpoByRegion[region] || [];
    const labels = stores.map(st => st.label || st.name);

    let j = i + 1;
    while (j < form.getItems().length) {
      const items2 = form.getItems();
      const it2 = items2[j];
      if (it2.getType() === FormApp.ItemType.PAGE_BREAK) break;

      if (it2.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
        const mcItem = it2.asMultipleChoiceItem();
        const choices = buildChoicesSafe_(mcItem, labels, navTarget);
        mcItem.setChoices(choices);
        Logger.log(`✓ Region ${region} の店舗選択MC更新（→${navLabel}）: ${labels.length || 1}件`);
        j++;
        continue;
      }

      if (it2.getType() === FormApp.ItemType.LIST) {
        const listItem = it2.asListItem();
        const oldTitle = listItem.getTitle();
        const mc = form.addMultipleChoiceItem().setTitle(oldTitle || `店舗を選択 (Region ${region})`);
        const choices = buildChoicesSafe_(mc, labels, navTarget);
        mc.setChoices(choices);

        const afterAdd = form.getItems();
        const newMc = afterAdd[afterAdd.length - 1];
        form.moveItem(newMc, j);
        form.deleteItem(it2);
        Logger.log(`✓ Region ${region} の店舗LISTをMCへ置換（→${navLabel}）: ${labels.length || 1}件`);
        continue;
      }
      j++;
    }
    i = j;
  }
}

function buildChoicesSafe_(mcItem, labels, navTarget) {
  if (labels && labels.length > 0) {
    return labels.map(lbl => navTarget ? mcItem.createChoice(lbl, navTarget) : mcItem.createChoice(lbl));
  }
  return [mcItem.createChoice(PLACEHOLDER_LABEL)];
}

function setRegionSectionsGoToLast_(form, navTarget) {
  const navLabel = navTarget ? navTarget.getTitle() : '遷移なし';
  const items = form.getItems();
  let cnt = 0;

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.getType() !== FormApp.ItemType.PAGE_BREAK) continue;
    const pb = it.asPageBreakItem();
    const title = pb.getTitle() || '';

    if (title === REPORT_SECTION_TITLE) continue;

    const m = title.match(/Region[\s:-]*([A-Z])(?:\b|[^A-Za-z]|$)/i);
    if (!m) continue;

    if (navTarget) {
      pb.setGoToPage(navTarget);
      Logger.log(`✓ Region ${m[1].toUpperCase()} セクションの after-navigation を『${navLabel}』に設定`);
      cnt++;
    }
  }
  if (cnt === 0) Logger.log('（注意）Region セクションが見つからず after-navigation は未設定');
}

function moveReportSectionToEnd_(form) {
  const items = form.getItems();
  let start = -1, end = items.length;

  for (let k = 0; k < items.length; k++) {
    const it = items[k];
    if (it.getType() !== FormApp.ItemType.PAGE_BREAK) continue;
    if (it.asPageBreakItem().getTitle() === REPORT_SECTION_TITLE) {
      start = k;
      for (let j = k + 1; j < items.length; j++) {
        if (items[j].getType() === FormApp.ItemType.PAGE_BREAK) { end = j; break; }
      }
      break;
    }
  }
  if (start === -1) { Logger.log('（move）報告内容セクションが見つかりません'); return; }

  for (let pos = start; pos < end; pos++) {
    const current = form.getItems();
    form.moveItem(current[start], current.length - 1);
  }
  Logger.log('✓ 報告内容セクションを最後に配置しました');
}


// ==================== フォーム解決ユーティリティ ====================

function resolveForm_(idOrUrl) {
  const cand = buildCandidates_(idOrUrl, true);
  let lastErr = null;
  for (const c of cand) {
    try {
      const f = /^https?:\/\//i.test(c) ? FormApp.openByUrl(c) : FormApp.openById(c);
      return { form: f, id: f.getId(), editUrl: f.getEditUrl(), pubUrl: f.getPublishedUrl() };
    } catch (e) { lastErr = e; }
  }
  throw new Error('フォームを開けませんでした。詳細: ' + (lastErr || '(不明)'));
}

function buildCandidates_(idOrUrl, includeFromLinkedSheet) {
  const arr = [];
  if (idOrUrl && idOrUrl.trim()) {
    const s = idOrUrl.trim();
    if (/^https?:\/\//i.test(s)) {
      const id = extractFormIdFromUrl_(s);
      if (id) arr.push(id);
      arr.push(s);
    } else {
      arr.push(s);
    }
  }
  if (includeFromLinkedSheet) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const formUrl = ss.getFormUrl();
      if (formUrl) {
        const id = extractFormIdFromUrl_(formUrl);
        if (id) arr.push(id);
        arr.push(formUrl);
      }
    } catch (_) {}
  }
  return Array.from(new Set(arr));
}

function extractFormIdFromUrl_(url) {
  const m = url && url.match(/\/forms\/d\/([a-zA-Z0-9_-]+)(?:\/|$)/);
  return m ? m[1] : null;
}


// ==================== プロパティ適用 & 実行結果保存 ====================

function applyFormSettings_(form, s) {
  tryApply_(() => has(s,'title') && form.setTitle(s.title));
  tryApply_(() => has(s,'description') && form.setDescription(s.description));
  tryApply_(() => has(s,'confirmationMessage') && form.setConfirmationMessage(s.confirmationMessage));
  tryApply_(() => has(s,'acceptingResponses') && form.setAcceptingResponses(!!s.acceptingResponses));
  tryApply_(() => has(s,'allowResponseEdits') && form.setAllowResponseEdits(!!s.allowResponseEdits));
  tryApply_(() => has(s,'showLinkToRespondAgain') && form.setShowLinkToRespondAgain(!!s.showLinkToRespondAgain));
  tryApply_(() => has(s,'requireLogin') && form.setRequireLogin(!!s.requireLogin));
  tryApply_(() => has(s,'limitOneResponsePerUser') && form.setLimitOneResponsePerUser(!!s.limitOneResponsePerUser));
  tryApply_(() => has(s,'progressBar') && form.setProgressBar(!!s.progressBar));
  tryApply_(() => has(s,'shuffleQuestions') && form.setShuffleQuestions(!!s.shuffleQuestions));
  tryApply_(() => has(s,'collectEmail') && form.setCollectEmail(!!s.collectEmail));
  tryApply_(() => has(s,'isQuiz') && form.setIsQuiz(!!s.isQuiz));
  Logger.log('✓ フォーム設定を適用');
}

function saveRunProps_(info) {
  if (!LOG_TO_PROPERTIES) return;
  try {
    const t = new Date().toISOString();
    const base = {
      lastRunISO: t,
      lastOk: String(!!info.ok),
      lastMessage: info.message || '',
      jugyoinCount: info.counts && info.counts.jugyoin != null ? String(info.counts.jugyoin) : '',
      tenpoCount: info.counts && info.counts.tenpo != null ? String(info.counts.tenpo) : '',
      regionCount: info.counts && info.counts.regions != null ? String(info.counts.regions) : ''
    };
    PropertiesService.getScriptProperties().setProperties(base, true);
    PropertiesService.getUserProperties().setProperties(base, true);
    Logger.log('✓ 実行結果を Properties に保存');
  } catch(e) { Logger.log('(Properties保存スキップ) ' + e); }
}

function has(o,k) { return Object.prototype.hasOwnProperty.call(o,k) && o[k] != null; }
function tryApply_(fn) { try { fn(); } catch(e) { Logger.log('(設定スキップ) ' + e); } }