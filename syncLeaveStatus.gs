const scriptProperties = PropertiesService.getScriptProperties();
const TRELLO_KEY = scriptProperties.getProperty('TRELLO_KEY');
const TRELLO_TOKEN = scriptProperties.getProperty('TRELLO_TOKEN');
const BOARD_ID = scriptProperties.getProperty('BOARD_ID');
const EXCEL_ID = scriptProperties.getProperty('EXCEL_ID');

function syncLeaveStatus() {
  const ss = SpreadsheetApp.openById(EXCEL_ID);
  // 取得今天日期字串
  const todayName = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
  // 根據日期名稱尋找分頁
  let sheet = ss.getSheetByName(todayName);
  
  // 如果找不到今天的分頁（例如還沒建立），則抓取當前活動分頁（手動時有效）或第一個分頁
  if (!sheet) {
    sheet = ss.getActiveSheet() || ss.getSheets()[0];
  }

  try {
    // 1. 初始化日期
    const meetingDate = new Date(sheet.getName());
    meetingDate.setHours(0, 0, 0, 0);

    const attendeesRange = sheet.getRange("B21");
    const leaveRange = sheet.getRange("B22");
    const attendeesText = attendeesRange.getValue().toString();
    const attendeesArray = attendeesText.split(/[、，\s]+/).filter(String);
    const leavePeople = [];

    // 2. 呼叫 Trello API
    const listIds = scriptProperties.getProperty('LIST_ID').split(',');
    let cards = [];
    listIds.forEach(id => {
      const url = "https://api.trello.com/1/lists/" + id + "/cards?key=" + TRELLO_KEY + "&token=" + TRELLO_TOKEN;
      try {
        const resp = UrlFetchApp.fetch(url);
        const listCards = JSON.parse(resp.getContentText());
        cards = cards.concat(listCards);
      } catch (e) {
        console.log("抓取列表 " + id + " 出錯");
      }
    });

    for (let i = 0; i < cards.length; i++) {
      const card = cards[i];
      if (card.due && card.name.indexOf("請假") !== -1) {
        const endDate = new Date(card.due);
        endDate.setHours(0, 0, 0, 0);
        const startDate = card.start ? new Date(card.start) : new Date(card.due);
        startDate.setHours(0, 0, 0, 0);

        if (meetingDate.getTime() >= startDate.getTime() && meetingDate.getTime() <= endDate.getTime()) {
          const cleanName = card.name.replace(/請假|[\d一二三四五六七八九十]+日|【|】|\s/g, "");
          if (cleanName) { leavePeople.push(cleanName); }
        }
      }
    }

    // 4. 如果有找到請假的人
    if (leavePeople.length > 0) {
      const uniqueLeavePeople = [];
      const updatedAttendees = [];

      for (let j = 0; j < attendeesArray.length; j++) {
        const person = attendeesArray[j];
        let isLeaving = false;
        for (let k = 0; k < leavePeople.length; k++) {
          if (person.indexOf(leavePeople[k]) !== -1) {
            isLeaving = true;
            if (uniqueLeavePeople.indexOf(person) === -1) {
              uniqueLeavePeople.push(person);
            }
            break;
          }
        }
        if (!isLeaving) { updatedAttendees.push(person); }
      }

      const existingLeave = leaveRange.getValue();
      const currentLeaveArray = existingLeave ? existingLeave.split("、") : [];
      const finalLeavePeople = [...new Set([...currentLeaveArray, ...uniqueLeavePeople])];
      leaveRange.setValue(finalLeavePeople.join("、"));
      attendeesRange.setValue(updatedAttendees.join("、"));

    //  showSafeAlert("同步成功！今日請假人員：" + uniqueLeavePeople.join("、"));
    } else {
    //  showSafeAlert("找不到日期為 " + sheet.getName() + " 的請假卡片");
    }

  } catch (e) {
    console.log("發生錯誤: " + e.message);
  } // <--- 補上結束 try 的括號
} // <--- 補上結束 syncLeaveStatus 的括號

// 輔助函式，獨立於主函式外
function showSafeAlert(msg) {
  try {
    // 改用 toast，會在右下角顯示通知，5秒後自動消失
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "系統通知", 5);
  } catch (e) {
    console.log("執行紀錄: " + msg);
  }
}

function setupTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncLeaveStatus') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('syncLeaveStatus')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(7)
    .create();
  
  // 改用 log，不會彈窗打擾
  console.log("排程已設定完成");
}
