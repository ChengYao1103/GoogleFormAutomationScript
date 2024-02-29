function getConfig() {
  const config = {
    formId: "",
    spreadsheetId: "",
    formDataWorksheet: "表單回應 1",
    formDataEmailColName: "電子郵件地址",
    enableDeleteDuplicateResponse: true,
    enableAutoCloseForm: true,
    closeFormCellData: ["工作表1", "關閉表單"],
    formClosedMessage: "現已無法填寫這份表單",
    updateTargets:[
      ["問題1", "工作表1", "顯示選項"]
    ]
  }


  return config;
}
