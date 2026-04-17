function createNextMeetingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("公版"); 
  const indexSheet = ss.getSheetByName("目錄");

  if (!templateSheet) {
    Logger.log("找不到名為 '公版' 的工作表");
    return;
  }

  // 1. 取得下週四的日期
  const today = new Date();
  const nextThursday = new Date(); // 雖然內容會變，但變數指向同一個 Date 物件，可用 const

  // 計算距離下個週四還有幾天 (週四是 4)
  const daysUntilThursday = (4 - today.getDay() + 7) % 7 || 7; 
  // 註：圖片中是用 if，這裡示範簡潔寫法，原 logic 改 const 亦可
  
  nextThursday.setDate(today.getDate() + daysUntilThursday);

  // 2. 格式化日期為 YYYY/MM/DD
  const year = nextThursday.getFullYear();
  const month = ("0" + (nextThursday.getMonth() + 1)).slice(-2);
  const day = ("0" + nextThursday.getDate()).slice(-2);
  const newSheetName = year + "/" + month + "/" + day;

  // 3. 檢查工作表是否存在
  if (ss.getSheetByName(newSheetName)) {
    Logger.log("工作表 " + newSheetName + " 已存在，跳過建立。");
    return;
  }

  // 4. 複製公版
  const newSheet = templateSheet.copyTo(ss);
  newSheet.setName(newSheetName);

  // 5. 排序：放在「公版」之後
  const templateIndex = templateSheet.getIndex();
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(templateIndex + 1);

  Logger.log("已成功建立新表：" + newSheetName + "，並排在公版之後。");
  // 6. 同步回目錄頁並建立超連結
  if (indexSheet) {
    // --- 動態尋找插入位置 ---
    const values = indexSheet.getRange("A1:A20").getValues(); // 取得前 20 列的內容來掃描
    let targetRow = 1; // 預設從第 1 列開始找
    
    // 遍歷 A 欄，尋找第一個包含 "20" 開頭的儲存格 (例如 2026/...)
    for (let i = 0; i < values.length; i++) {
      if (values[i][0].toString().indexOf("20") !== -1) {
        targetRow = i + 1; // 找到了！這就是目前最新的日期所在行
        break;
      }
    }
    
    // 如果沒找到任何日期，就維持你原本想要的起始位置 (例如第 7 列)
    if (targetRow === 1) targetRow = 7; 

    const sheetId = newSheet.getSheetId();
    
    // 在「最舊的最新日期」上方插入新的一列
    indexSheet.insertRowBefore(targetRow);
    
    // 建立超連結公式
    const hyperlinkFormula = `=HYPERLINK("#gid=${sheetId}", "${newSheetName}")`;
    
    // 寫入 A 欄對應位置
    indexSheet.getRange(targetRow, 1).setFormula(hyperlinkFormula);
    
    // 自動複製下一列的格式 (確保框線跟字體顏色一致)
    indexSheet.getRange(targetRow + 1, 1).copyTo(indexSheet.getRange(targetRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    SpreadsheetApp.getUi().alert(`已成功建立分頁「${newSheetName}」，並自動插在目錄第 ${targetRow} 列！`);
  }
}
