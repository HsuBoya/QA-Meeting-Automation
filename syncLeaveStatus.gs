const scriptProperties = PropertiesService.getScriptProperties();
const TRELLO_KEY = scriptProperties.getProperty('TRELLO_KEY');
const TRELLO_TOKEN = scriptProperties.getProperty('TRELLO_TOKEN');
const BOARD_ID = scriptProperties.getProperty('BOARD_ID');


function syncLeaveStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  
  try {
    // 1. 初始化日期：將分頁名稱 (2026/04/23) 轉為時間
    const meetingDate = new Date(sheet.getName());
    meetingDate.setHours(0, 0, 0, 0); // 標準化為當天凌晨 0 點
    
    const attendeesRange = sheet.getRange("B21");
    const leaveRange = sheet.getRange("B22");
    
    const attendeesText = attendeesRange.getValue().toString();
    const attendeesArray = attendeesText.split(/[、，\s]+/).filter(String); 
    const leavePeople = []; // 使用 const 或 let 皆可，實務上常用 const

    // 2. 呼叫 Trello API
    //const url = "https://trello.com" + "/1/lists/" + LIST_ID + "/cards?key=" + TRELLO_KEY + "&token=" + TRELLO_TOKEN;
    const listIds = scriptProperties.getProperty('LIST_ID').split(',');
    let cards = [];

    listIds.forEach(id => {
    const url = "https://api.trello.com/1/lists/" + id + "/cards?key=" + TRELLO_KEY + "&token=" + TRELLO_TOKEN;
    try {
    const resp = UrlFetchApp.fetch(url);
    const listCards = JSON.parse(resp.getContentText());
    cards = cards.concat(listCards); 
  } catch(e) {
    console.log("抓取列表 " + id + " 出錯");
  }
});

    for (let i = 0; i < cards.length; i++) { // 迴圈計數器建議用 let
      const card = cards[i];

      // 只有標題含「請假」且有截止日期的卡片才處理
      if (card.due && card.name.indexOf("請假") !== -1) {
        const endDate = new Date(card.due);
        endDate.setHours(0, 0, 0, 0);
        
        // 如果有開始日期就用開始日期，沒有就用截止日期當作開始
        const startDate = card.start ? new Date(card.start) : new Date(card.due);
        startDate.setHours(0, 0, 0, 0);

        // 核心邏輯：判斷會議日期是否落在請假區間內 (包含頭尾)
        if (meetingDate.getTime() >= startDate.getTime() && meetingDate.getTime() <= endDate.getTime()) {
          // 過濾掉「請假」、「一日/三日」、括號與空格，只留人名
          const cleanName = card.name.replace(/請假|[\d一二三四五六七八九十]+日|【|】|\s/g, ""); 
          if (cleanName) {
            leavePeople.push(cleanName);
          }
        }
      }
    }

    // 4. 如果有找到請假的人
    if (leavePeople.length > 0) {
      const uniqueLeavePeople = []; // 最終的請假名單（完整姓名）
      const updatedAttendees = [];  // 剩餘的與會人員

      // 1. 遍歷所有的與會人員
      for (let j = 0; j < attendeesArray.length; j++) {
        const person = attendeesArray[j];
        let isLeaving = false; // 此變數會被修改，必須用 let

        // 2. 拿目前的與會者去比對 Trello 的請假人
        for (let k = 0; k < leavePeople.length; k++) {
          if (person.indexOf(leavePeople[k]) !== -1) {
            isLeaving = true;
            if (uniqueLeavePeople.indexOf(person) === -1) {
              uniqueLeavePeople.push(person);
            }
            break; 
          }
        }

        // 3. 如果這個人沒請假，才放回與會名單
        if (!isLeaving) {
          updatedAttendees.push(person);
        }
      }

      // 處理請假人員：合併「原本格內名單」與「Trello新抓到名單」
      const existingLeave = leaveRange.getValue();
      const currentLeaveArray = existingLeave ? existingLeave.split("、") : [];
      // 使用 Set 確保人名不重複
      const finalLeavePeople = [...new Set([...currentLeaveArray, ...uniqueLeavePeople])];
      leaveRange.setValue(finalLeavePeople.join("、"));
      attendeesRange.setValue(updatedAttendees.join("、"));
      
      SpreadsheetApp.getUi().alert("同步成功！今日請假人員：" + uniqueLeavePeople.join(", "));

    } else {
      SpreadsheetApp.getUi().alert("找不到日期為 " + sheet.getName() + " 的請假卡片");
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert("程式執行出錯：" + e.message);
  }
}


