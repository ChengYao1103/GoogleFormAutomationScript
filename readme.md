## Google Form 自動化腳本

### 功能

透過 Google App Script 實作以下功能：

1. 依據試算表內容動態更新表單選項
2. 根據 email 判斷是否為重複填寫，若為重複填寫則將舊資料刪除，達到避免重複填寫的效果
3. 讀取試算表特定欄位，依據該欄位關閉表單並設定提示訊息

---

### 設定方法

1. 建立表單與試算表，表單**不可**與試算表連結
2. 試算表內格式需參照[此影片](https://youtu.be/CPa26jRHHhQ?si=XOVOXD1joU-yFW-L)，表頭列名稱可彈性設定
3. 若要啟用刪除重複填寫功能則表單需設定開啟收集電子郵件地址
4. 在表單頁開啟**指令碼編輯器**
5. 將 [code.gs](./code.gs) 複製到 `程式碼.gs / code.gs`
6. 新增 `config.gs`，並將 [config.gs](./config.gs) 貼上
7. 修改 `config.gs` 內的參數
8. 新增觸發條件
   - 執行的功能為 `onSubmit`、 活動類型為「表單提交時」
   - 若有定時更新的需求可多新增一個觸發條件
     - 執行的功能為 `autoUpdateForm`、活動來源為「時間驅動」
     - 依需求設定執行頻率

---

### config 參數說明

|             Name              |  Type   | Description                                           | Example Value                       |
| :---------------------------: | :-----: | :---------------------------------------------------- | :---------------------------------- |
|            formId             | string  | 表單 ID                                               | ![](./image/formId.jpg)             |
|         spreadsheetId         | string  | 試算表 ID                                             | ![](./image/spreadsheetId.jpg)      |
|       formDataWorksheet       | string  | 儲存表單資料的工作表名稱                              | "表單回覆 1"                        |
|     formDataEmailColName      | string  | 表單資料工作表內 email 行的表頭名稱                   | "電子郵件地址"                      |
| enableDeleteDuplicateResponse | boolean | 啟用/停用刪除重複填寫資料功能                         | true                                |
|      enableAutoCloseForm      | boolean | 啟用/停用自動關閉表單功能                             | true                                |
|       closeFormCellData       |  array  | 用以判斷是否關閉表單的儲存格位置                      | ["工作表 1", "關閉表單"]            |
|       formClosedMessage       | string  | 表單關閉回應後的提示訊息                              | "現已無法填寫這份表單"              |
|         updateTargets         |  array  | 需要動態生成/更新選項的問題及儲存對應資料的工作表資訊 | [["問題1", "工作表 1", "顯示選項"]] |

#### array 型態參數補充

1. closeFormCellPosition
   - 固定為 2 個元素的一維陣列
   - `closeFormCellData[0]` 為儲存格所在之工作表名稱
   - `closeFormCellData[1]` 為儲存格所在之表頭名稱
2. updateTargets
   - 二維陣列，若陣列為空則不啟用動態生成/更新功能
   - 內部的陣列固定為 3 個元素
     - `updateTargets[0]` 為欲動態生成/更新選項的問題
     - `updateTargets[1]` 為選項資料所在之工作表名稱
     - `updateTargets[2]` 為選項資料所在之表頭名稱
   - 可設定複數個問題進行動態生成/更新選項
