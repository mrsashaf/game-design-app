#!/usr/bin/env node
/**
 * sync_dialogues.js — ABERRATION Dialogue Sync Tool
 *
 * Читает диалоги из level_design.html и CSV-мастера.
 * Генерирует:
 *   1. dialogues_from_app.csv    — все строки из приложения (чистый список)
 *   2. dialogue_merge_report.txt — отчёт: что совпало, что новое, что устарело
 *
 * Использование:
 *   node sync_dialogues.js            # полный прогон
 *   node sync_dialogues.js --apply    # обновить мастер-CSV (осторожно!)
 */

'use strict';
const fs   = require('fs');
const path = require('path');
const vm   = require('vm');

// ─── ФАЙЛЫ ───────────────────────────────────────────────────────────────────
const DIR         = __dirname;
const HTML_FILE   = path.join(DIR, 'level_design.html');
const CSV_FILE    = path.join(DIR, 'ABERRATION_Dialogue_v3_full - Dialogue.csv');
const OUT_APP_CSV = path.join(DIR, 'dialogues_from_app.csv');
const OUT_REPORT  = path.join(DIR, 'dialogue_merge_report.txt');
const KEY_MAP     = path.join(DIR, 'dialogue_key_map.json');
const APPLY_MODE  = process.argv.includes('--apply');

// ─── SPEAKER MAP ─────────────────────────────────────────────────────────────
const SPEAKER_MAP = {
  'МЭГ':           { en: 'Meg',     code: 'MEG'   },
  'СТИВЕНС':       { en: 'Steve',   code: 'STEVE' },
  'СТИВ':          { en: 'Steve',   code: 'STEVE' },
  'ДЖЕК':          { en: 'Jack',    code: 'JACK'  },
  'НЕИЗВЕСТНЫЙ':   { en: 'Unknown', code: 'UNK'   },
  'ГОЛОС В ГОЛОВЕ':{ en: 'Unknown', code: 'UNK'   },
  'ГОЛОС':         { en: 'Unknown', code: 'UNK'   },
  'ТОТЕМ':         { en: 'Unknown', code: 'UNK'   },
  'EYESYS':        { en: 'Unknown', code: 'UNK'   },
};

// ─── SCENE META (для генерации ключей) ───────────────────────────────────────
const SCENE_META = {
  scene_00: { code: 'P00', title: 'Пролог' },
  scene_01: { code: 'P01', title: 'Переулок 01' },
  scene_02: { code: 'U01', title: 'Улица 01' },
  scene_03: { code: 'P02', title: 'Переулок 02' },
  scene_04: { code: 'S04', title: 'Заброшенный магазин' },
  scene_05: { code: 'U02', title: 'Улица 02' },
  scene_06: { code: 'CAV', title: 'Пещера Камня' },
  scene_07: { code: 'TUN', title: 'Туннель' },
  scene_08: { code: 'U03', title: 'Улица 03' },
  scene_09: { code: 'FIN', title: 'Финал' },
};

// ─── 1. ИЗВЛЕЧЬ КОНСТАНТЫ ИЗ HTML ────────────────────────────────────────────
function extractConstantsFromHTML(html) {
  const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/);
  if (!scriptMatch) throw new Error('Не найден тег <script> в HTML');
  const fullScript = scriptMatch[1];

  // Берём только блок с данными (до buildDefaultProject)
  const dataPart = fullScript.match(/([\s\S]*?)\/\/ ─── BUILD DEFAULT PROJECT/);
  if (!dataPart) throw new Error('Не найден маркер "BUILD DEFAULT PROJECT" в скрипте');

  // vm.runInNewContext не экспортирует const/let — заменяем на var
  const dataCode = dataPart[1].replace(/\bconst\b/g, 'var').replace(/\blet\b/g, 'var');
  const sandbox = {};
  try {
    vm.runInNewContext(dataCode, sandbox);
  } catch (e) {
    throw new Error('Ошибка выполнения скрипта данных: ' + e.message);
  }

  return {
    LAYER_DEFS:       sandbox.LAYER_DEFS,
    DEFAULT_BOARD:    sandbox.DEFAULT_BOARD,
    SCENE_GROUPING:   sandbox.SCENE_GROUPING,
    SUBNODE_TO_LAYER: sandbox.SUBNODE_TO_LAYER,
    DEFAULT_BEATS:    sandbox.DEFAULT_BEATS || {},
  };
}

// ─── 2. ПАРСЕР СТРОК ДИАЛОГА ──────────────────────────────────────────────────
/**
 * Разбирает текст вида:
 *   МЭГ (тихо): «Чёрт... Надо найти чем сломать эту цепь.»
 *   ДЖЕК (рация): «Мэг, приём...»
 * Возвращает массив { speakerRu, speakerEn, speakerCode, stageDir, russian }
 */
function parseDialogueLines(text) {
  if (!text || !text.trim()) return [];
  const lines = [];
  // Поддерживаем «», "", '' и обычные кавычки
  const pattern = /^(МЭГ|СТИВЕНС|СТИВ|ДЖЕК|НЕИЗВЕСТНЫЙ|ГОЛОС В ГОЛОВЕ|ГОЛОС|ТОТЕМ|EYESYS)(\s*\([^)]*\))?\s*:\s*[«"'„"](.+?)[»"'"][ \t]*$/gm;
  let m;
  while ((m = pattern.exec(text)) !== null) {
    const speakerRu   = m[1].trim();
    const stageDir    = m[2] ? m[2].trim().replace(/^\(|\)$/g, '').trim() : '';
    const russian     = m[3].trim();
    const info        = SPEAKER_MAP[speakerRu] || { en: 'Unknown', code: 'UNK' };
    lines.push({ speakerRu, speakerEn: info.en, speakerCode: info.code, stageDir, russian });
  }
  return lines;
}

// ─── 3. СОБРАТЬ ВСЕ ДИАЛОГИ ИЗ DEFAULT_BOARD ─────────────────────────────────
function collectAppDialogues(constants) {
  const { DEFAULT_BOARD, SCENE_GROUPING } = constants;

  // Построить обратную карту nodeId → sceneId
  const nodeToScene = {};
  for (const scene of SCENE_GROUPING) {
    for (const nid of scene.nodeIds) nodeToScene[nid] = scene.id;
  }

  // Построить карту nodeId → node object
  const nodeById = {};
  for (const node of DEFAULT_BOARD.nodes) nodeById[node.id] = node;

  // Собираем: для каждой ноды берём subNode с суффиксом _l6 (диалоги)
  const result = []; // { nodeId, sceneId, eventTitle, lines: [...] }

  for (const node of DEFAULT_BOARD.nodes) {
    const sceneId = nodeToScene[node.id] || 'unknown';
    const eventTitle = (node.title || '').replace(/^СЦЕНА\s*\d+\s*[—–-]\s*/i, '').trim();

    // Ищем subNode с диалогами (_l6)
    const dialogueSub = (node.subNodes || []).find(s => s.id && s.id.endsWith('_l6'));
    if (!dialogueSub || !dialogueSub.content || !dialogueSub.content.trim()) continue;

    const lines = parseDialogueLines(dialogueSub.content);
    if (lines.length === 0) continue;

    result.push({ nodeId: node.id, sceneId, eventTitle, content: dialogueSub.content, lines });
  }

  return result;
}

// ─── 4. ГЕНЕРАЦИЯ КЛЮЧА ───────────────────────────────────────────────────────
function generateKey(nodeId, speakerCode, index) {
  const nodeCode = nodeId.replace(/[^a-zA-Z0-9]/g, '_').toUpperCase();
  return `${speakerCode}_${nodeCode}_${String(index).padStart(2, '0')}`;
}

// ─── 5. CSV ПАРСЕР ────────────────────────────────────────────────────────────
function parseCSV(text) {
  const rows = [];
  let i = 0;

  function parseField() {
    if (text[i] === '"') {
      i++; // skip opening quote
      let val = '';
      while (i < text.length) {
        if (text[i] === '"' && text[i+1] === '"') { val += '"'; i += 2; }
        else if (text[i] === '"') { i++; break; }
        else { val += text[i++]; }
      }
      return val;
    } else {
      let val = '';
      while (i < text.length && text[i] !== ',' && text[i] !== '\n' && text[i] !== '\r') val += text[i++];
      return val;
    }
  }

  while (i < text.length) {
    const row = [];
    while (i < text.length && text[i] !== '\n' && text[i] !== '\r') {
      row.push(parseField());
      if (text[i] === ',') i++;
    }
    if (text[i] === '\r') i++;
    if (text[i] === '\n') i++;
    if (row.length > 1 || (row.length === 1 && row[0] !== '')) rows.push(row);
  }
  return rows;
}

// ─── 6. CSV СЕРИАЛИЗАТОР ──────────────────────────────────────────────────────
function csvField(val) {
  const s = String(val || '');
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}
function serializeCSV(rows) {
  return rows.map(row => row.map(csvField).join(',')).join('\r\n') + '\r\n';
}

// ─── 7. FUZZY MATCH РУССКОГО ТЕКСТА ──────────────────────────────────────────
function normalize(s) {
  return (s || '').toLowerCase().replace(/[^а-яёa-z0-9]/g, '').slice(0, 40);
}
function fuzzyMatch(appText, csvText) {
  const a = normalize(appText);
  const b = normalize(csvText);
  if (!a || !b) return 0;
  if (a === b) return 1;
  // Считаем совпадение первых N символов
  let match = 0;
  const len = Math.min(a.length, b.length);
  for (let i = 0; i < len; i++) { if (a[i] === b[i]) match++; }
  return match / Math.max(a.length, b.length);
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
function main() {
  console.log('═══════════════════════════════════════════════');
  console.log(' ABERRATION Dialogue Sync Tool');
  console.log('═══════════════════════════════════════════════\n');

  // 1. Читаем HTML
  if (!fs.existsSync(HTML_FILE)) { console.error('❌ Не найден файл:', HTML_FILE); process.exit(1); }
  const html = fs.readFileSync(HTML_FILE, 'utf8');
  let constants;
  try { constants = extractConstantsFromHTML(html); }
  catch (e) { console.error('❌ Ошибка извлечения данных из HTML:', e.message); process.exit(1); }
  console.log('✅ HTML прочитан, данные извлечены');

  // 2. Собираем диалоги из приложения
  const appGroups = collectAppDialogues(constants);
  const allAppLines = [];

  // Считаем индексы per speaker per node
  for (const group of appGroups) {
    const speakerCounters = {};
    for (const line of group.lines) {
      const key = line.speakerCode;
      speakerCounters[key] = (speakerCounters[key] || 0) + 1;
      const autoKey = generateKey(group.nodeId, line.speakerCode, speakerCounters[key]);
      allAppLines.push({
        autoKey,
        speakerEn: line.speakerEn,
        speakerRu: line.speakerRu,
        stageDir:  line.stageDir,
        russian:   line.russian,
        nodeId:    group.nodeId,
        sceneId:   group.sceneId,
        eventTitle:group.eventTitle,
      });
    }
  }
  console.log(`📄 Найдено диалоговых строк в приложении: ${allAppLines.length}`);

  // 3. Читаем мастер-CSV
  let csvRows = [];
  let csvHeaders = [];
  if (fs.existsSync(CSV_FILE)) {
    const csvText = fs.readFileSync(CSV_FILE, 'utf8');
    const parsed  = parseCSV(csvText);
    csvHeaders    = parsed[0];
    csvRows       = parsed.slice(1).filter(r => r[0] && r[0].trim());
    console.log(`📊 Строк в мастер-CSV: ${csvRows.length}`);
  } else {
    console.log('⚠️  Мастер-CSV не найден, создаём новый');
    csvHeaders = ['KEY','TAG','SPEAKER','TONE OF VOICE','ENGLISH','RUSSIAN','GERMAN','SPANISH','CHINESE','AUDIO FILE','STATUS','NOTE'];
  }

  // Индекс CSV по ключу и по русскому тексту
  const csvByKey = {};
  const csvByRu  = {};
  for (const row of csvRows) {
    const key = row[0]; const ru = row[5];
    if (key) csvByKey[key] = row;
    if (ru && ru.trim()) csvByRu[normalize(ru)] = row;
  }

  // 4. Загружаем или создаём key map (appAutoKey → csvKey)
  let keyMap = {};
  if (fs.existsSync(KEY_MAP)) {
    try { keyMap = JSON.parse(fs.readFileSync(KEY_MAP, 'utf8')); }
    catch (e) { keyMap = {}; }
  }

  // 5. Строим строки для нового CSV из приложения
  const reportLines   = [];
  const newCsvRows    = [csvHeaders.slice()];
  let countMatched    = 0;
  let countNew        = 0;
  let countFuzzy      = 0;

  for (const appLine of allAppLines) {
    const sceneMeta  = SCENE_META[appLine.sceneId] || { code: appLine.sceneId, title: appLine.sceneId };
    const note       = `[${sceneMeta.code}] ${appLine.eventTitle}${appLine.stageDir ? ' — ' + appLine.stageDir : ''}`;

    // Пробуем найти соответствие в CSV
    // Сначала проверяем keyMap — описательный ключ из маппинга
    const mappedKey  = keyMap[appLine.autoKey] || null;
    let matchedKey   = mappedKey;
    let matchedRow   = matchedKey ? csvByKey[matchedKey] : null;
    let matchScore   = matchedRow ? 1 : 0;
    let matchType    = matchedRow ? 'MAPPED' : (mappedKey ? 'MAPPED_NEW' : 'NONE');

    if (!matchedRow && !mappedKey) {
      // Fuzzy match по русскому тексту (только если нет маппинга в keyMap)
      let bestScore = 0; let bestRow = null;
      for (const [normRu, row] of Object.entries(csvByRu)) {
        const score = fuzzyMatch(appLine.russian, normRu);
        if (score > bestScore) { bestScore = score; bestRow = row; }
      }
      if (bestScore >= 0.65) {
        matchedRow  = bestRow;
        matchedKey  = bestRow[0];
        matchScore  = bestScore;
        matchType   = 'FUZZY';
        countFuzzy++;
      }
    }

    let finalKey, status, tov, audioFile, english, german, spanish, chinese;

    if (matchedRow) {
      countMatched++;
      finalKey   = matchedRow[0];
      tov        = matchedRow[3];
      english    = matchedRow[4];
      german     = matchedRow[6];
      spanish    = matchedRow[7];
      chinese    = matchedRow[8];
      audioFile  = matchedRow[9];
      status     = matchedRow[10] || 'DRAFT';
      // Обновляем ключ в key map
      keyMap[appLine.autoKey] = finalKey;
    } else if (mappedKey) {
      // Есть описательный ключ в keyMap, но нет строки в CSV — используем описательный ключ
      countNew++;
      finalKey   = mappedKey;
      tov        = '';
      english    = '';
      german     = '';
      spanish    = '';
      chinese    = '';
      audioFile  = '';
      status     = 'DRAFT';
    } else {
      countNew++;
      finalKey   = appLine.autoKey;
      tov        = '';
      english    = '';
      german     = '';
      spanish    = '';
      chinese    = '';
      audioFile  = '';
      status     = 'DRAFT';
    }

    // Собираем строку CSV
    newCsvRows.push([
      finalKey,
      'MAIN',
      appLine.speakerEn,
      tov,
      english,
      appLine.russian,   // ← текст из приложения
      german,
      spanish,
      chinese,
      audioFile,
      status,
      note,
    ]);

    // Строка отчёта
    const matchLabel = matchType === 'NONE' ? '🆕 НОВАЯ' :
                       matchType === 'FUZZY' ? `🔀 FUZZY(${Math.round(matchScore*100)}%)` : '✅ СОВПАЛА';
    reportLines.push(
      `${matchLabel.padEnd(18)} | KEY: ${finalKey.padEnd(35)} | ${appLine.speakerRu.padEnd(10)} | ${appLine.russian.slice(0,60)}${appLine.russian.length>60?'…':''}`
    );
  }

  // 6. Строки из CSV, которых нет в приложении (устаревшие?)
  const appAutoKeys = new Set(allAppLines.map(l => l.autoKey));
  const usedCsvKeys = new Set(newCsvRows.slice(1).map(r => r[0]));
  const orphanRows  = csvRows.filter(r => r[0] && !usedCsvKeys.has(r[0]));

  // 7. Если --apply: добавляем orphan-строки в конец (чтобы не потерять их)
  //    и перезаписываем мастер-CSV
  if (APPLY_MODE) {
    for (const orphan of orphanRows) newCsvRows.push(orphan);
    fs.writeFileSync(CSV_FILE, serializeCSV(newCsvRows), 'utf8');
    console.log(`\n✅ Мастер-CSV ОБНОВЛЁН (${newCsvRows.length - 1} строк)`);
  }

  // 8. Всегда пишем dialogues_from_app.csv
  fs.writeFileSync(OUT_APP_CSV, serializeCSV(newCsvRows), 'utf8');
  console.log(`\n✅ Записан: dialogues_from_app.csv (${newCsvRows.length - 1} строк)`);

  // 9. Сохраняем обновлённый key map
  fs.writeFileSync(KEY_MAP, JSON.stringify(keyMap, null, 2), 'utf8');
  console.log('✅ Записан: dialogue_key_map.json');

  // 10. Пишем отчёт
  const orphanReport = orphanRows.length > 0
    ? '\n═══ СТРОКИ В CSV, КОТОРЫХ НЕТ В ПРИЛОЖЕНИИ (возможно устарели) ═══\n' +
      orphanRows.map(r => `  ⚠️  ${r[0]} | ${r[2]} | ${(r[5]||'').slice(0,60)}`).join('\n')
    : '\n(Устаревших строк не найдено)';

  const report = [
    '═══════════════════════════════════════════════════════════════════',
    ' ABERRATION Dialogue Sync — Отчёт',
    `═══════════════════════════════════════════════════════════════════`,
    ``,
    `Всего строк из приложения : ${allAppLines.length}`,
    `  ✅ Совпало с CSV        : ${countMatched} (из них fuzzy: ${countFuzzy})`,
    `  🆕 Новые (нет в CSV)   : ${countNew}`,
    `  ⚠️  Orphan в CSV        : ${orphanRows.length}`,
    ``,
    `Режим: ${APPLY_MODE ? '⚠️  APPLY (мастер-CSV обновлён!)' : 'DRY RUN (мастер-CSV НЕ изменён)'}`,
    ``,
    '═══ ВСЕ СТРОКИ ИЗ ПРИЛОЖЕНИЯ ═══════════════════════════════════',
    '',
    ...reportLines,
    orphanReport,
    '',
    '═══════════════════════════════════════════════════════════════════',
    '',
    'Для применения изменений в мастер-CSV запусти:',
    '  node sync_dialogues.js --apply',
    '',
    'Для проверки совпадений отредактируй dialogue_key_map.json:',
    '  { "APP_AUTO_KEY": "EXISTING_CSV_KEY", ... }',
  ].join('\n');

  fs.writeFileSync(OUT_REPORT, report, 'utf8');
  console.log('✅ Записан: dialogue_merge_report.txt');

  // 11. Консольный итог
  console.log(`
═══════════════════════════════════════
  Итог:
  ✅ Совпало с CSV    : ${countMatched}
  🔀 Fuzzy совпадений: ${countFuzzy}
  🆕 Новых строк     : ${countNew}
  ⚠️  Orphan в CSV    : ${orphanRows.length}
═══════════════════════════════════════

${APPLY_MODE
  ? '⚠️  APPLY MODE: мастер-CSV перезаписан!'
  : 'ℹ️  DRY RUN: чтобы применить, запусти с флагом --apply'}

Следующий шаг:
  1. Открой dialogue_merge_report.txt
  2. Проверь FUZZY-совпадения
  3. Поправь dialogue_key_map.json если нужно
  4. Запусти: node sync_dialogues.js --apply
`);
}

main();
