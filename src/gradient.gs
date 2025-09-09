/*********** настройки ***********/
const SHEET_NAME = 'Свод';      // поменяй при необходимости
const A1_RANGE   = 'J2:P100';   // диапазон, где нужен градиент

// цвета, как на скрине (red → yellow → green)
const COL_MIN = '#EA4335';
const COL_MID = '#FBBC04';
const COL_MAX = '#34A853';

/*********** построчный градиент (правило на строку) ***********/
function cfCreateRowGradients() {
  const ss = SpreadsheetApp.getActive();
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  const rng = sh.getRange(A1_RANGE);

  // очистить старые УФ, пересекающиеся с диапазоном
  clearCfRulesForRange_(sh, rng);

  const rStart = rng.getRow();
  const cStart = rng.getColumn();
  const rCount = rng.getNumRows();
  const cCount = rng.getNumColumns();

  const rules = sh.getConditionalFormatRules();

  for (let i = 0; i < rCount; i++) {
    const rowRange = sh.getRange(rStart + i, cStart, 1, cCount);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([rowRange])
      .setGradientMinpointWithValue(COL_MIN, SpreadsheetApp.InterpolationType.MIN, '0')
      .setGradientMidpointWithValue(COL_MID, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
      .setGradientMaxpointWithValue(COL_MAX, SpreadsheetApp.InterpolationType.MAX, '0')
      .build();
    rules.push(rule);
  }

  sh.setConditionalFormatRules(rules);
}

/*********** утилита: удалить правила УФ, пересекающиеся с диапазоном ***********/
function clearCfRulesForRange_(sheet, targetRange) {
  const tRow = targetRange.getRow();
  const tCol = targetRange.getColumn();
  const tBottom = tRow + targetRange.getNumRows() - 1;
  const tRight  = tCol + targetRange.getNumColumns() - 1;

  const keep = sheet.getConditionalFormatRules().filter(rule => {
    return !rule.getRanges().some(r => {
      const rRow = r.getRow(), rCol = r.getColumn();
      const rBottom = rRow + r.getNumRows() - 1;
      const rRight  = rCol + r.getNumColumns() - 1;
      const overlap = !(rRight < tCol || rCol > tRight || rBottom < tRow || rRow > tBottom);
      return overlap;
    });
  });
  sheet.setConditionalFormatRules(keep);
}
