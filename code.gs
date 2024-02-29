// 取得設定檔
const config = getConfig();

// 格式化日期
function formatDate(timestamp) {
  var hour =
    timestamp.getHours() < 10
      ? `0${timestamp.getHours()}`
      : timestamp.getHours();
  var minute =
    timestamp.getMinutes() < 10
      ? `0${timestamp.getMinutes()}`
      : timestamp.getMinutes();
  var second =
    timestamp.getSeconds() < 10
      ? `0${timestamp.getSeconds()}`
      : timestamp.getSeconds();

  return (
    `${timestamp.getFullYear()}年${timestamp.getMonth() + 1}月${timestamp.getDate()}日 ` +
    `${hour}:${minute}:${second}`
  );
}

// 取得試算表內指定名稱的工作表
function getWorksheet(workSheetName) {
  var sheets = SpreadsheetApp.openById(config.spreadsheetId).getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === workSheetName) {
      return sheets[i];
    }
  }
  return null;
}

// 取得表單的 headers(資料的意義)及最新一筆回覆資料
function getFormLatestData() {
  var form = FormApp.openById(config.formId);
  var headers = ["date", "email", ...form.getItems().map((e) => e.getTitle())];
  var values = [
    headers,
    ...form
      .getResponses()
      .map((formResponse) => {
        const time = formatDate(formResponse.getTimestamp());
        return formResponse.getItemResponses().reduce(
          (o, itemResponse) => {
            const response = itemResponse.getResponse();
            return Object.assign(o, {
              [itemResponse.getItem().getTitle()]: Array.isArray(response) ? response.join(",") : response,
            });
          },
          { email: formResponse.getRespondentEmail(), date: time }
        );
      })
      .map((o) => headers.map((t) => o[t] || "")),
  ];

  return [headers, values[values.length - 1]];
}

// 以指定工作表內的特定行(以表頭名稱指定)為來源更新表單選項
function updateChoice(targetItemTitle, worksheetName, colName) {
  var items = FormApp.openById(config.formId).getItems();
  for(var i = 0; i < items.length; i++) {
    if(items[i].getTitle() === targetItemTitle) {
      // 取得儲存選項資料的工作表
      var sheet = getWorksheet(worksheetName);
      if(!sheet) {
        return;
      }
      var item = items[i].asMultipleChoiceItem();
      var sheetData = sheet.getDataRange().getValues();
      var choiceColIndex = sheetData[0].indexOf(colName)
      var choices = [];
      for(var j = 1; j < sheetData.length; j++) {
        // 如果顯示選項為空(超過上限人數)就不新增選項
        if(sheetData[j][choiceColIndex].length !== 0) {
          choices.push(item.createChoice(sheetData[j][choiceColIndex]))
        }
      }
      // 設定新選項
      if(choices.length > 0) {
        item.setChoices(choices);
      }
      break;
    }
  }
}

// 回傳工作表內關閉表單欄位的值
function isFormNeedClose(workSheetName, colName) {
  var sheet = getWorksheet(workSheetName);
  if(!sheet) {
    return false;
  }
  var sheetData = sheet.getDataRange().getValues();
  var index = sheetData[0].indexOf(colName) + 1;
  return sheetData[0][index];
}

// 關閉表單回覆並設定提示訊息
function closeForm() {
  var form = FormApp.openById(config.formId);
  form.setCustomClosedFormMessage(config.formClosedMessage);
  form.setAcceptingResponses(false);
}

// 更新選項 & 檢查是否關閉表單(可設定依時間自動觸發)
function autoUpdateForm() {
  // 更新所有在 config.gs 內設定的問題及選項
  for(var i = 0; i< config.updateTargets.length ; i++) {
    var target = config.updateTargets[i];
    updateChoice(target[0], target[1], target[2]);
  }
  // 若啟用自動關閉表單才執行
  if(config.enableAutoCloseForm && isFormNeedClose(config.closeFormCellData[0], config.closeFormCellData[1])) {
    closeForm();
  }
}

// 提交表單時觸發
function onSubmit(event) {
  var latestData = getFormLatestData();
  // 取得儲存表單資料的工作表
  var sheet = getWorksheet(config.formDataWorksheet);
  if (sheet) {
    // 若啟用刪除重複回應才執行
    if(config.enableDeleteDuplicateResponse) {
      var email = latestData[1][1].trim();
      var sheetData = sheet.getDataRange().getValues();
      // 取得工作表內 email 的行數並進行比較，重複則刪除舊資料
      var emailIndex = sheetData[0].indexOf(config.formDataEmailColName);
      for (var i = 1; i < sheetData.length; i++) {
        if (sheetData[i][emailIndex].trim() === email) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }
    // 將資料寫入工作表
    var lastRow = parseInt(sheet.getLastRow() + 1);
    sheet
      .getRange(lastRow, 1, 1, latestData[0].length)
      .setValues([latestData[1]]);
  }
  // 更新選項 & 檢查是否需要關閉表單
  autoUpdateForm();
}