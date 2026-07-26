const SPREADSHEET_ID = "1jCyiI4yQQCdW-894UwMGtejmWGNHCtT3Q1rbP0-Hvlg";

function doGet(e) {
  try {
    const action = e ? e.parameter.action : "";
    const initDataStr = e ? e.parameter.initData : "";
    const userId = parseTelegramUserId(initDataStr);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName("UserDB");
    const actSheet = ss.getSheetByName("ActivityLog");
    const calSheet = ss.getSheetByName("Calendar");

    // 1. 유저 인증 + 기본 일정 데이터 및 주간 미션 수치 조회
    if (action === "checkUser" || action === "getInitialData") {
      const targetDate = e.parameter.date || getTodayString();

      let isRegistered = false;
      let userObj = { id: userId, name: '', group: '', region: '', role: '', bookCount: 0, isExempt: false };
      const unregisteredList = [];

      const allRegionsList = [];
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          const cellId = uData[i][0] ? uData[i][0].toString().trim() : "";
          const cellName = uData[i][1] ? uData[i][1].toString().trim() : "";
          const cellGroup = formatGroupString(uData[i][2]);
          const cellRegion = uData[i][3] ? uData[i][3].toString().trim() : "";
          const cellRole = uData[i][4] ? uData[i][4].toString().trim() : "";

          const cellBookCount = parseInt(uData[i][13] || 0, 10);
          const cellIsExempt = uData[i][14] ? (uData[i][14].toString().trim().toUpperCase() === 'Y') : false;

          if (cellRegion) {
            allRegionsList.push(cellRegion);
          }

          if (cellId && cellId === userId.toString().trim()) {
            isRegistered = true;
            userObj = {
              id: cellId,
              name: cellName,
              group: cellGroup,
              region: cellRegion,
              role: cellRole,
              bookCount: cellBookCount,
              isExempt: cellIsExempt
            };
          }

          if (!cellId && cellName) {
            unregisteredList.push({
              row: i + 1,
              name: cellName,
              info: `(${cellGroup} ${cellRegion} ${cellRole})`.trim()
            });
          }
        }
      }

      const allRegions = Array.from(new Set(allRegionsList));

      const roleStr = userObj.role || "";
      const isSuperAdmin = roleStr.includes("전체관리자") || roleStr.includes("관리자");
      const isGroupAdmin = roleStr.includes("조장") || roleStr.includes("부조장");

      const activities = [];
      let weeklyMissionCount = 0;

      if (actSheet && isRegistered) {
        const actData = actSheet.getDataRange().getValues();

        const baseDate = new Date(targetDate);
        const dayOfWeek = baseDate.getDay();
        const sunday = new Date(baseDate);
        sunday.setDate(baseDate.getDate() - dayOfWeek);
        const saturday = new Date(sunday);
        saturday.setDate(sunday.getDate() + 6);

        const sundayStr = formatDateStr(sunday);
        const saturdayStr = formatDateStr(saturday);

        for (let i = 1; i < actData.length; i++) {
          let rowDate = actData[i][6];
          if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            rowDate = rowDate ? rowDate.toString().trim() : "";
          }

          const rowUserId = actData[i][1] ? actData[i][1].toString().trim() : "";
          const rowGroup = formatGroupString(actData[i][3]);
          const rowStatus = actData[i][11];
          const rowResultData = actData[i][12] ? actData[i][12].toString() : "";

          if (rowUserId === userId.toString().trim() && rowStatus === "completed") {
            if (rowDate >= sundayStr && rowDate <= saturdayStr) {
              if (rowResultData) {
                try {
                  const parsed = JSON.parse(rowResultData);
                  if (Array.isArray(parsed)) {
                    parsed.forEach(r => {
                      if (r.category === "찾기" && (r.type === "온만찾" || r.type === "오프만찾")) {
                        weeklyMissionCount += (r.count !== undefined ? parseInt(r.count, 10) : 1);
                      }
                    });
                  }
                } catch (e) { }
              }
            }
          }

          if (rowDate === targetDate) {
            let canAccess = false;
            if (isSuperAdmin) {
              canAccess = true;
            } else if (isGroupAdmin) {
              canAccess = (rowGroup === userObj.group);
            } else {
              canAccess = (rowUserId === userId.toString().trim());
            }

            if (canAccess) {
              activities.push({
                id: actData[i][0],
                userId: actData[i][1],
                name: actData[i][2],
                group: rowGroup,
                region: actData[i][4],
                role: actData[i][5],
                date: rowDate,
                startTime: formatTimeString(actData[i][7]),
                endTime: formatTimeString(actData[i][8]),
                location: actData[i][9],
                content: actData[i][10] ? actData[i][10].toString() : "찾기",
                status: rowStatus,
                resultData: rowResultData,
                resultText: actData[i][13] ? actData[i][13].toString() : ""
              });
            }
          }
        }
      }

      return makeJsonResponse({
        isRegistered: isRegistered,
        user: userObj,
        activities: activities,
        unregisteredList: unregisteredList,
        isAdmin: (isSuperAdmin || isGroupAdmin),
        weeklyMissionCount: weeklyMissionCount,
        currentSemester: getCurrentSemesterObj(),
        allRegions: allRegions
      });
    }

    // 2. 관리자 대시보드 API
    if (action === "getDashboardData") {
      const rangeType = e.parameter.range || "today";
      const baseDateStr = e.parameter.date || getTodayString();

      let userObj = { id: userId, group: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][0] && uData[i][0].toString().trim() === userId.toString().trim()) {
            userObj = {
              id: userId,
              group: formatGroupString(uData[i][2]),
              role: uData[i][4] ? uData[i][4].toString().trim() : ""
            };
            break;
          }
        }
      }

      const roleStr = userObj.role || "";
      const isSuperAdmin = roleStr.includes("전체관리자") || roleStr.includes("관리자");
      const isGroupAdmin = roleStr.includes("조장") || roleStr.includes("부조장");

      if (!isSuperAdmin && !isGroupAdmin) {
        return makeJsonResponse({ result: "fail", message: "관리자 권한이 없습니다." });
      }

      const rawList = [];
      if (actSheet) {
        const actData = actSheet.getDataRange().getValues();
        const dates = calculateRangeDates(baseDateStr, rangeType);

        for (let i = 1; i < actData.length; i++) {
          let rowDate = actData[i][6];
          if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            rowDate = rowDate ? rowDate.toString().trim() : "";
          }

          if (dates.includes(rowDate)) {
            const rowGroup = formatGroupString(actData[i][3]);
            let canAccess = isSuperAdmin ? true : (rowGroup === userObj.group);

            if (canAccess) {
              rawList.push({
                id: actData[i][0],
                userId: actData[i][1],
                name: actData[i][2],
                group: rowGroup,
                region: actData[i][4],
                role: actData[i][5],
                date: rowDate,
                startTime: formatTimeString(actData[i][7]),
                endTime: formatTimeString(actData[i][8]),
                location: actData[i][9],
                content: actData[i][10] ? actData[i][10].toString() : "찾기",
                status: actData[i][11],
                resultData: actData[i][12] ? actData[i][12].toString() : "",
                resultText: actData[i][13] ? actData[i][13].toString() : ""
              });
            }
          }
        }
      }

      return makeJsonResponse({ result: "success", range: rangeType, list: rawList });
    }

    // 💡 3. 미취합자(선택 날짜 당일 결과 + 내일 계획) 및 미달성자, 휴무일 스킵 API
    if (action === "getRiskData") {
      const targetDateStr = e.parameter.date || getTodayString();
      const targetDate = new Date(targetDateStr);

      // 내일 날짜 계산
      const tomorrow = new Date(targetDate);
      tomorrow.setDate(targetDate.getDate() + 1);
      const tomorrowStr = formatDateStr(tomorrow);

      // 이번 주 일요일 ~ 토요일 계산
      const dayOfWeek = targetDate.getDay();
      const sunday = new Date(targetDate);
      sunday.setDate(targetDate.getDate() - dayOfWeek);
      const saturday = new Date(sunday);
      saturday.setDate(sunday.getDate() + 6);
      const sundayStr = formatDateStr(sunday);
      const saturdayStr = formatDateStr(saturday);

      // 📅 Calendar 탭 검사 (선택 날짜 및 내일 날짜 휴무 여부)
      let targetDateOff = false;
      let targetOffReason = "";
      let tomorrowDateOff = false;
      let tomorrowOffReason = "";

      if (calSheet) {
        const cData = calSheet.getDataRange().getValues();
        for (let i = 1; i < cData.length; i++) {
          let cDate = cData[i][0];
          if (cDate instanceof Date) {
            cDate = Utilities.formatDate(cDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            cDate = cDate ? cDate.toString().trim() : "";
          }

          const statusVal = cData[i][1] ? cData[i][1].toString().trim().toUpperCase() : "";
          const noteVal = cData[i][2] ? cData[i][2].toString().trim() : "";

          if (cDate === targetDateStr && statusVal === "OFF") {
            targetDateOff = true;
            targetOffReason = noteVal || "공식 비계수일/휴무일";
          }
          if (cDate === tomorrowStr && statusVal === "OFF") {
            tomorrowDateOff = true;
            tomorrowOffReason = noteVal || "공식 비계수일/휴무일";
          }
        }
      }

      const userMap = {};
      const allUsers = [];

      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          const uId = uData[i][0] ? uData[i][0].toString().trim() : "";
          const name = uData[i][1] ? uData[i][1].toString().trim() : "";
          const group = formatGroupString(uData[i][2]);
          const region = uData[i][3] ? uData[i][3].toString().trim() : "";
          const bookCount = parseInt(uData[i][13] || 0, 10);
          const isExempt = uData[i][14] ? (uData[i][14].toString().trim().toUpperCase() === 'Y') : false;

          if (name) {
            const uObj = {
              id: uId,
              name: name,
              group: group,
              region: region,
              bookCount: bookCount,
              isExempt: isExempt,
              hasTodayActivity: false,
              isTodayCompleted: false,
              hasTomorrowActivity: false,
              weeklyFindCount: 0
            };
            userMap[name] = uObj;
            allUsers.push(uObj);
          }
        }
      }

      if (actSheet) {
        const actData = actSheet.getDataRange().getValues();
        for (let i = 1; i < actData.length; i++) {
          let rowDate = actData[i][6];
          if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            rowDate = rowDate ? rowDate.toString().trim() : "";
          }

          const name = actData[i][2] ? actData[i][2].toString().trim() : "";
          const status = actData[i][11];
          const resultData = actData[i][12] ? actData[i][12].toString() : "";

          if (userMap[name]) {
            // 당일 활동 파악
            if (rowDate === targetDateStr) {
              userMap[name].hasTodayActivity = true;
              if (status === 'completed') {
                userMap[name].isTodayCompleted = true;
              }
            }
            // 내일 활동 파악
            if (rowDate === tomorrowStr) {
              userMap[name].hasTomorrowActivity = true;
            }
            // 주간 찾기 성과 계산
            if (rowDate >= sundayStr && rowDate <= saturdayStr && status === 'completed' && resultData) {
              try {
                const parsed = JSON.parse(resultData);
                if (Array.isArray(parsed)) {
                  parsed.forEach(r => {
                    if (r.category === '찾기' && (r.type === '온만찾' || r.type === '오프만찾')) {
                      userMap[name].weeklyFindCount += (r.count !== undefined ? parseInt(r.count, 10) : 1);
                    }
                  });
                }
              } catch (e) { }
            }
          }
        }
      }

      // 1. [선택 날짜 + 1일] 내일 계획 미취합자 (내일 일정 아예 없음 & 예외 N)
      const tomorrowUncollected = tomorrowDateOff ? [] : allUsers.filter(u => !u.hasTomorrowActivity && !u.isExempt);

      // 2. [선택 날짜] 당일 결과 미취합자 (오늘 일정 없음 또는 오늘 완료 상태가 아님 & 예외 N)
      const todayUncollected = targetDateOff ? [] : allUsers.filter(u => (!u.hasTodayActivity || !u.isTodayCompleted) && !u.isExempt);

      // 3. 이번 주 찾기 미달성자 (복방 0개 & 이번 주 찾기 6개 미만)
      const unachievedUsers = allUsers.filter(u => u.bookCount === 0 && u.weeklyFindCount < 6);

      const allRegions = Array.from(new Set(allUsers.map(u => u.region).filter(Boolean)));

      return makeJsonResponse({
        result: "success",
        targetDateStr: targetDateStr,
        tomorrowStr: tomorrowStr,
        targetDateOff: targetDateOff,
        targetOffReason: targetOffReason,
        tomorrowDateOff: tomorrowDateOff,
        tomorrowOffReason: tomorrowOffReason,
        todayUncollected: todayUncollected,
        tomorrowUncollected: tomorrowUncollected,
        unachievedUsers: unachievedUsers,
        allUsers: allUsers,
        allRegions: allRegions
      });
    }

    // 4. 신규 유저 등록
    if (action === "registerUser") {
      const selectedRow = parseInt(e.parameter.selectedRow, 10);
      if (userSheet && selectedRow && userId) {
        userSheet.getRange(selectedRow, 1).setValue(userId.toString());
        const rowValues = userSheet.getRange(selectedRow, 1, 1, 5).getValues()[0];
        return makeJsonResponse({
          result: "success",
          userName: rowValues[1] || "",
          group: formatGroupString(rowValues[2]),
          region: rowValues[3] || "",
          role: rowValues[4] || ""
        });
      }
      return makeJsonResponse({ result: "fail", message: "잘못된 요청입니다." });
    }

    // 5. 활동 추가 및 수정
    if (action === "saveActivity") {
      const mode = e.parameter.mode;
      const actId = e.parameter.actId;
      const datePart = e.parameter.date || getTodayString();
      const startTime = e.parameter.startTime || "09:00";
      const endTime = e.parameter.endTime || "10:00";

      let userInfo = { name: '', group: '', region: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][0].toString().trim() === userId.toString().trim()) {
            userInfo = {
              name: uData[i][1],
              group: formatGroupString(uData[i][2]),
              region: uData[i][3],
              role: uData[i][4]
            };
            break;
          }
        }
      }

      if (actSheet) {
        if (mode === 'add') {
          actSheet.appendRow([
            actId,
            userId,
            userInfo.name,
            userInfo.group,
            userInfo.region,
            userInfo.role,
            datePart,
            startTime,
            endTime,
            e.parameter.location,
            e.parameter.content || "찾기",
            e.parameter.status,
            e.parameter.resultData || "",
            e.parameter.resultText || "",
            new Date()
          ]);
        } else {
          const actData = actSheet.getDataRange().getValues();
          for (let i = 1; i < actData.length; i++) {
            if (actData[i][0].toString() === actId.toString()) {
              actSheet.getRange(i + 1, 7).setValue(datePart);
              actSheet.getRange(i + 1, 8).setValue(startTime);
              actSheet.getRange(i + 1, 9).setValue(endTime);
              actSheet.getRange(i + 1, 10).setValue(e.parameter.location);
              actSheet.getRange(i + 1, 11).setValue(e.parameter.content || "찾기");
              actSheet.getRange(i + 1, 12).setValue(e.parameter.status);
              actSheet.getRange(i + 1, 13).setValue(e.parameter.resultData || "");
              actSheet.getRange(i + 1, 14).setValue(e.parameter.resultText || "");
              break;
            }
          }
        }
      }

      return makeJsonResponse({ result: "success" });
    }

    // 6. 관리자 전용: 유저 복방 수량 및 예외 처리 수정 API
    if (action === "updateUserStatus") {
      const targetName = e.parameter.targetName;
      const bookCount = e.parameter.bookCount;
      const isExempt = e.parameter.isExempt;

      if (userSheet && targetName) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][1] && uData[i][1].toString().trim() === targetName.toString().trim()) {
            if (bookCount !== undefined) {
              userSheet.getRange(i + 1, 14).setValue(parseInt(bookCount, 10));
            }
            if (isExempt !== undefined) {
              userSheet.getRange(i + 1, 15).setValue(isExempt === "true" ? "Y" : "N");
            }
            return makeJsonResponse({ result: "success" });
          }
        }
      }
      return makeJsonResponse({ result: "fail", message: "유저를 찾을 수 없습니다." });
    }

    // 7. 개강 사이클 조회
    if (action === "getSemesterData") {
      return getSemesterData();
    }

    // 8. 개강 사이클 업데이트
    if (action === "updateSemester") {
      return updateSemester(e.parameter.name, e.parameter.startDate, e.parameter.endDate);
    }

  } catch (err) {
    return makeJsonResponse({ error: err.message });
  }
}

function calculateRangeDates(baseDateStr, rangeType) {
  const parts = baseDateStr.split('-');
  const base = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const results = [];

  if (rangeType === "today") {
    results.push(baseDateStr);
  } else if (rangeType === "week") {
    const dayOfWeek = base.getDay();
    const diffToMon = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    const monday = new Date(base);
    monday.setDate(base.getDate() + diffToMon);

    for (let i = 0; i < 7; i++) {
      const d = new Date(monday);
      d.setDate(monday.getDate() + i);
      results.push(formatDateStr(d));
    }
  } else if (rangeType === "month") {
    const y = base.getFullYear();
    const m = base.getMonth();
    const lastDay = new Date(y, m + 1, 0).getDate();

    for (let i = 1; i <= lastDay; i++) {
      const d = new Date(y, m, i);
      results.push(formatDateStr(d));
    }
  }
  return results;
}

function formatDateStr(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function makeJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseTelegramUserId(initDataStr) {
  if (!initDataStr) return "";
  try {
    const decoded = decodeURIComponent(initDataStr);
    const userParam = decoded.split('&').find(p => p.startsWith('user='));
    if (userParam) {
      return JSON.parse(userParam.split('user=')[1]).id.toString();
    }
  } catch (e) {
    return "";
  }
  return "";
}

function getTodayString() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function formatTimeString(val) {
  if (!val) return "00:00";
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }
  return val.toString().trim();
}

function formatGroupString(val) {
  if (!val && val !== 0) return "";
  const str = val.toString().trim();
  if (!str) return "";
  return str.endsWith("조") ? str : `${str}조`;
}

function getSemesterData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let semSheet = ss.getSheetByName("Semesters");

  if (!semSheet) {
    semSheet = ss.insertSheet("Semesters");
    semSheet.appendRow(["사이클명", "시작일", "종료일", "현재여부"]);
    semSheet.appendRow(["2026-2학기 개강", "2026-09-01", "2026-12-31", "Y"]);
  }

  const data = semSheet.getDataRange().getValues();
  let currentSem = { name: "2026-2학기 개강", startDate: "2026-09-01", endDate: "2026-12-31" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toUpperCase() === "Y") {
      currentSem = {
        name: String(data[i][0]),
        startDate: formatDateStr(new Date(data[i][1])),
        endDate: formatDateStr(new Date(data[i][2]))
      };
      break;
    }
  }

  return makeJsonResponse({ result: "success", semester: currentSem });
}

function updateSemester(name, startDate, endDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let semSheet = ss.getSheetByName("Semesters");

  if (!semSheet) {
    semSheet = ss.insertSheet("Semesters");
    semSheet.appendRow(["사이클명", "시작일", "종료일", "현재여부"]);
  }

  // 기존 Y를 N으로 바꾸고 신규 사이클 Y로 설정
  const data = semSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    semSheet.getRange(i + 1, 4).setValue("N");
  }

  semSheet.appendRow([name, startDate, endDate, "Y"]);
  return makeJsonResponse({ result: "success", message: "개강 사이클이 변경되었습니다." });
}

// 순수 객체 반환 헬퍼 (getInitialData 등에서 재사용)
function getCurrentSemesterObj() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const semSheet = ss.getSheetByName("Semesters");

  if (!semSheet) {
    return { name: "2026-2학기 개강", startDate: "2026-09-01", endDate: "2026-12-31" };
  }

  const data = semSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toUpperCase() === "Y") {
      return {
        name: String(data[i][0]),
        startDate: formatDateStr(new Date(data[i][1])),
        endDate: formatDateStr(new Date(data[i][2]))
      };
    }
  }
  return { name: "2026-2학기 개강", startDate: "2026-09-01", endDate: "2026-12-31" };
}