// --- 1. UI & SETUP ---

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Configure Auto-Fill Rules', 'showSidebar')
    .addItem('Setup Trigger (Run Once)', 'createInstallableTrigger')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Auto-Fill Rules');
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function createInstallableTrigger() {
  const sheet = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processEdits')
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert("Background trigger installed successfully!");
}

// --- 2. DATA MANAGEMENT & UTILS ---

function getRules() {
  const props = PropertiesService.getDocumentProperties();
  return JSON.parse(props.getProperty('autoFillRules') || '[]');
}

function saveRule(rule) {
  const props = PropertiesService.getDocumentProperties();
  let rules = JSON.parse(props.getProperty('autoFillRules') || '[]');
  
  if (rule.id) {
    const index = rules.findIndex(r => r.id === rule.id);
    if (index > -1) {
      rules[index] = rule;
    } else {
      rules.push(rule);
    }
  } else {
    rule.id = 'rule_' + new Date().getTime();
    rules.push(rule);
  }
  
  props.setProperty('autoFillRules', JSON.stringify(rules));
  return true;
}

function deleteRule(id) {
  const props = PropertiesService.getDocumentProperties();
  let rules = JSON.parse(props.getProperty('autoFillRules') || '[]');
  rules = rules.filter(r => r.id !== id);
  props.setProperty('autoFillRules', JSON.stringify(rules));
  return true;
}

function toggleRule(id, isEnabled) {
  const props = PropertiesService.getDocumentProperties();
  let rules = JSON.parse(props.getProperty('autoFillRules') || '[]');
  const index = rules.findIndex(r => r.id === id);
  if (index > -1) {
    rules[index].enabled = isEnabled;
    props.setProperty('autoFillRules', JSON.stringify(rules));
  }
  return true;
}

function saveImportedRules(jsonString) {
  const props = PropertiesService.getDocumentProperties();
  try {
    const parsed = JSON.parse(jsonString);
    if (Array.isArray(parsed)) {
      props.setProperty('autoFillRules', jsonString);
      return true;
    }
  } catch (e) {}
  return false;
}

function getActiveSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

// Helper: Converts a 1-based column index to A1 notation letters (e.g., 3 -> "C")
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// --- 3. THE EVALUATION ENGINE ---

function processEdits(e) {
  if (!e || !e.range) return; 
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  const editStartRow = e.range.getRow();
  const editEndRow = editStartRow + e.range.getNumRows() - 1;
  const editStartCol = e.range.getColumn();
  const editEndCol = editStartCol + e.range.getNumColumns() - 1;
  
  const props = PropertiesService.getDocumentProperties();
  const rules = JSON.parse(props.getProperty('autoFillRules') || '[]');
  
  rules.forEach(rule => {
    if (rule.enabled === false) return; 
    if (rule.sheetScope !== 'All Sheets' && rule.sheetScope !== sheetName) return;

    const detRangeObj = sheet.getRange(rule.detectionRange);
    const ruleStartRow = detRangeObj.getRow();
    const ruleEndRow = detRangeObj.getLastRow();
    const ruleStartCol = detRangeObj.getColumn();
    const ruleEndCol = detRangeObj.getLastColumn();
    
    if (editStartCol <= ruleEndCol && editEndCol >= ruleStartCol) {
      const startOverlap = Math.max(editStartRow, ruleStartRow);
      const endOverlap = Math.min(editEndRow, ruleEndRow);
      
      for (let r = startOverlap; r <= endOverlap; r++) {
        evaluateRuleForRow(sheet, rule, r, ruleStartCol, ruleEndCol);
      }
    }
  });
}

function forceRecalculate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  const props = PropertiesService.getDocumentProperties();
  const rules = JSON.parse(props.getProperty('autoFillRules') || '[]');
  
  let appliedRulesCount = 0;

  rules.forEach(rule => {
    if (rule.enabled === false) return; 
    if (rule.sheetScope !== 'All Sheets' && rule.sheetScope !== sheetName) return;

    const detRangeObj = sheet.getRange(rule.detectionRange);
    const ruleStartRow = detRangeObj.getRow();
    const ruleEndRow = detRangeObj.getLastRow(); 
    const ruleStartCol = detRangeObj.getColumn();
    const ruleEndCol = detRangeObj.getLastColumn();
    
    for (let r = ruleStartRow; r <= ruleEndRow; r++) {
      evaluateRuleForRow(sheet, rule, r, ruleStartCol, ruleEndCol);
    }
    appliedRulesCount++;
  });
  
  return appliedRulesCount;
}

function evaluateRuleForRow(sheet, rule, editRow, startCol, endCol) {
  const rowDetectionRange = sheet.getRange(editRow, startCol, 1, (endCol - startCol) + 1);
  const rowValues = rowDetectionRange.getValues()[0];
  
  const isAllEmpty = rowValues.every(val => val === "");
  const isAllFull = rowValues.every(val => val !== "");
  const isAnyChecked = rowValues.some(val => val === true); 
  const isAllUnchecked = rowValues.every(val => val === false || val === "");
  const hasSpecificMatch = rowValues.some(val => val.toString() === rule.targetValue); 
  
  const targetRangeObj = sheet.getRange(rule.fillRange);
  const targetCol = targetRangeObj.getColumn();
  const targetCell = sheet.getRange(editRow, targetCol);
  const currentTargetValue = targetCell.getValue();
  
  let fillValue = "";
  let isFormula = false; // Flag to determine how we write to the cell
  
  switch (rule.fillType) {
    case 'Date': fillValue = new Date(); break;
    case 'Static': fillValue = rule.staticTextValue || ""; break;
    case 'UUID': fillValue = Utilities.getUuid(); break;
    case 'Email': fillValue = Session.getActiveUser().getEmail(); break;
    case 'Checkbox': fillValue = true; break; 
    case 'CopyCol': 
      if (rule.copyColIndex) {
         // Create a dynamic formula like "=C5"
         const colLetter = columnToLetter(parseInt(rule.copyColIndex));
         fillValue = "=" + colLetter + editRow;
         isFormula = true;
      }
      break;
  }
  
  const applyFill = (val) => {
    // STATE PROTECTION: Do not overwrite existing Dates/UUIDs unless the rule demands constant updates
    if ((rule.fillType === 'Date' || rule.fillType === 'UUID') && currentTargetValue !== "") {
      if (rule.detectionType !== "Fill and Update" && rule.detectionType !== "Fill on Any Edit") {
        return; // Break out of the function, leaving the cell exactly as it is
      }
    }

    const dv = targetCell.getDataValidation();
    if (rule.fillType === 'Checkbox') {
      if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        targetCell.insertCheckboxes();
      }
      targetCell.setValue(true);
    } else {
      if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        targetCell.clearDataValidations();
      }
      
      // Inject formula if CopyCol, otherwise inject raw value
      if (isFormula) {
        targetCell.setFormula(val);
      } else {
        targetCell.setValue(val);
      }
    }
  };
  
  const clearTarget = () => {
    const dv = targetCell.getDataValidation();
    if (rule.fillType === 'Checkbox') {
      if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        targetCell.insertCheckboxes();
      }
      targetCell.setValue(false); 
    } else {
      if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        targetCell.clearDataValidations();
      }
      targetCell.clearContent();
    }
  };
  
  switch (rule.detectionType) {
    case "Fill and Update":
      if (isAllEmpty) clearTarget();
      else applyFill(fillValue);
      break;
      
    case "Fill When Full":
      if (isAllFull) applyFill(fillValue);
      else clearTarget();
      break;
      
    case "Fill When Full, Clear when Empty":
      if (isAllFull) applyFill(fillValue);
      else if (isAllEmpty) clearTarget();
      break;
      
    case "Fill Once":
      if (isAllEmpty) {
        clearTarget();
      } else if (currentTargetValue === "" || currentTargetValue === false) {
        applyFill(fillValue);
      }
      break;
      
    case "Specific Value Match":
      if (hasSpecificMatch) applyFill(fillValue);
      else if (isAllEmpty) clearTarget();
      break;
      
    case "Is Checked":
      if (isAnyChecked) applyFill(fillValue);
      else if (isAllUnchecked) clearTarget();
      break;
      
    case "Fill on Any Edit":
      if (!isAllEmpty) applyFill(fillValue);
      else clearTarget();
      break;
  }
}
