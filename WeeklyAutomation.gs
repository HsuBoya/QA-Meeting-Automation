function createNextMeetingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("公版"); 
  const indexSheet = ss.getSheetByName("目錄");

  if (!templateSheet) {
    console.error("找不到名為 '公版' 的工作表");
    return;
  }

  // 1. 取得下週四的日期
  const today = new Date();
  const nextThursday = new Date(today.getTime());
  const daysUntilThursday = (4 - today.getDay() + 7) % 7 || 7; 
  nextThursday.setDate(today.getDate() + daysUntilThursday);

  // 2. 格式化日期為 YYYY/MM/DD
  const year = nextThursday.getFullYear();
  const month = ("0" + (nextThursday.getMonth() + 1)).slice(-2);
  const day = ("0" + nextThursday.getDate()).slice(-2);
  const newSheetName = `${year}/${month}/${day}`;

  // 3. 檢查工作表是否存在
  if (ss.getSheetByName(newSheetName)) {
    console.log("工作表 " + newSheetName + " 已存在，跳過建立。");
    return;
  }

  // 4. 複製公版
  const newSheet = templateSheet.copyTo(ss);
  newSheet.setName(newSheetName);

  // 5. 排序：放在「公版」之後
  const templateIndex = templateSheet.getIndex();
  newSheet.activate(); // 在背景執行時這行沒影響，但保留結構
  ss.moveActiveSheet(templateIndex + 1);

  // 6. 同步回目錄頁
  if (indexSheet) {
    const values = indexSheet.getRange("A1:A30").getValues(); 
    let targetRow = 0; 
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][0].toString().indexOf("20") !== -1) {
        targetRow = i + 1; 
        break;
      }
    }
    
    if (targetRow === 0) targetRow = 7; 

    const sheetId = newSheet.getSheetId();
    indexSheet.insertRowBefore(targetRow);
    
    const hyperlinkFormula = `=HYPERLINK("#gid=${sheetId}", "${newSheetName}")`;
    indexSheet.getRange(targetRow, 1).setFormula(hyperlinkFormula);
    
    // 複製格式
    indexSheet.getRange(targetRow + 1, 1).copyTo(indexSheet.getRange(targetRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    
    // 紀錄到後台日誌
    console.log(`[自動化成功] 已在目錄第 ${targetRow} 列插入新連結：${newSheetName}`);
  }

  // 徹底移除 UI.alert，這樣就算沒開視窗也會靜默執行完成
  console.log(`任務完成：${newSheetName} 已建立。`);
}