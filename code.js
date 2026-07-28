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
    const noticeSheet = ss.getSheetByName("Notices");

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
          const rowRegion = actData[i][4] ? actData[i][4].toString().trim() : "";
          const rowStatus = actData[i][11];
          const rowResultData = actData[i][12] ? actData[i][12].toString() : "";

          if (rowUserId === userId.toString().trim() && rowStatus === "completed") {
            if (rowDate >= sundayStr && rowDate <= saturdayStr) {
              if (rowResultData) {
                try {
                  const parsed = JSON.parse(rowResultData);
                  if (Array.isArray(parsed)) {
                    parsed.forEach(r => {
                      // 기존 code.js line 87 부근
                      if ((r.category === "찾기" || r.category === "찾기(오프라인)" || r.category === "찾기(온라인)") &&
                        (r.type === "온만찾" || r.type === "오프만찾")) { // ✅ 오프번찾 제거 확인
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
              canAccess = (rowRegion === userObj.region);
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

      // 💡 공지사항 조회 연동
      const notices = [];
      if (noticeSheet) {
        const nData = noticeSheet.getDataRange().getValues();
        for (let i = 1; i < nData.length; i++) {
          const row = nData[i];
          if (row.length < 9) continue;
          const nId = row[0] ? row[0].toString().trim() : "";
          const nType = row[1] ? row[1].toString().trim() : "";
          const nAuthorId = row[2] ? row[2].toString().trim() : "";
          const nAuthorName = row[3] ? row[3].toString().trim() : "";
          const nGroup = row[4] ? formatGroupString(row[4]) : "";
          const nTitle = row[5] ? row[5].toString().trim() : "";
          const nContent = row[6] ? row[6].toString().trim() : "";
          const nRegion = row[7] ? row[7].toString().trim() : "";
          const nVisible = row[8] ? row[8].toString().trim().toUpperCase() : "";
          let nCreatedAt = row[9];

          if (nTitle && nContent) {
            let show = false;
            const targetRegions = nRegion ? nRegion.split(",").map(r => r.trim()) : [];
            const isTargetRegion = (nRegion === "ALL" || !nRegion || targetRegions.includes(userObj.region));

            if (isSuperAdmin) {
              show = true;
            } else if (isGroupAdmin) {
              // 조장/부조장은 본인 지역 대상 공지사항(비공개 포함) 조회 가능
              if (isTargetRegion) {
                show = true;
              }
            } else {
              // 일반 조원은 오직 공개된(Visible === Y) 본인 지역 공지만 조회 가능
              if (nVisible === "Y" && isTargetRegion) {
                show = true;
              }
            }

            if (show) {
              let dateStr = "";
              if (nCreatedAt instanceof Date) {
                dateStr = Utilities.formatDate(nCreatedAt, Session.getScriptTimeZone(), "MM-dd HH:mm");
              } else {
                dateStr = nCreatedAt ? nCreatedAt.toString().trim() : "";
              }
              const isImp = row[10] ? row[10].toString().trim().toUpperCase() : "N"; // K열
              const authRole = row[11] ? row[11].toString().trim() : "관리자"; // L열

              notices.push({
                id: nId,
                type: nType,
                authorName: nAuthorName,
                authorRole: authRole,
                group: nGroup,
                title: nTitle,
                content: nContent,
                region: nRegion,
                visible: nVisible,
                isImportant: isImp,
                createdAt: dateStr
              });
            }
          }
        }
      }

      // 중요 공지 우선 정렬 및 최신 등록순 정렬
      notices.sort((a, b) => {
        if (a.isImportant === "Y" && b.isImportant !== "Y") return -1;
        if (a.isImportant !== "Y" && b.isImportant === "Y") return 1;
        return parseInt(b.id) - parseInt(a.id);
      });

      return makeJsonResponse({
        isRegistered: isRegistered,
        user: userObj,
        activities: activities,
        unregisteredList: unregisteredList,
        isAdmin: (isSuperAdmin || isGroupAdmin),
        weeklyMissionCount: weeklyMissionCount,
        currentSemester: getCurrentSemesterObj(),
        allRegions: allRegions,
        weekName: getWeekNameForDate(targetDate),
        semestersList: getSemestersList(),
        notices: notices
      });
    }

    // 2. 관리자 대시보드 API
    if (action === "getDashboardData") {
      const rangeType = e.parameter.range || "today";
      const baseDateStr = e.parameter.date || getTodayString();

      let userObj = { id: userId, group: '', region: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][0] && uData[i][0].toString().trim() === userId.toString().trim()) {
            userObj = {
              id: userId,
              group: formatGroupString(uData[i][2]),
              region: uData[i][3] ? uData[i][3].toString().trim() : "",
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
            const rowRegion = actData[i][4] ? actData[i][4].toString().trim() : "";
            let canAccess = isSuperAdmin ? true : (rowRegion === userObj.region);

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

    // 💡 2.5. 나의 활동 추적 API
    if (action === "getUserTrackerData") {
      const startDate = e.parameter.startDate;
      const endDate = e.parameter.endDate;

      let isRegistered = false;
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          const cellId = uData[i][0] ? uData[i][0].toString().trim() : "";
          if (cellId && cellId === userId.toString().trim()) {
            isRegistered = true;
            break;
          }
        }
      }

      if (!isRegistered) {
        return makeJsonResponse({ result: "fail", message: "미등록 사용자입니다." });
      }

      const trackerList = [];
      if (actSheet) {
        const actData = actSheet.getDataRange().getValues();
        for (let i = 1; i < actData.length; i++) {
          let rowDate = actData[i][6];
          if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            rowDate = rowDate ? rowDate.toString().trim() : "";
          }

          const rowUserId = actData[i][1] ? actData[i][1].toString().trim() : "";

          if (rowUserId === userId.toString().trim() && rowDate >= startDate && rowDate <= endDate) {
            trackerList.push({
              id: actData[i][0],
              userId: rowUserId,
              name: actData[i][2],
              group: formatGroupString(actData[i][3]),
              region: actData[i][4],
              role: actData[i][5],
              date: rowDate,
              startTime: formatTimeString(actData[i][7]),
              endTime: formatTimeString(actData[i][8]),
              location: actData[i][9],
              content: actData[i][10] ? actData[i][10].toString() : "찾기(오프라인)",
              status: actData[i][11],
              resultData: actData[i][12] ? actData[i][12].toString() : "",
              resultText: actData[i][13] ? actData[i][13].toString() : ""
            });
          }
        }
      }

      return makeJsonResponse({ result: "success", list: trackerList });
    }

    // 💡 2.6. 공지사항 저장 API
    if (action === "saveNotice") {
      const title = e.parameter.title ? e.parameter.title.toString().trim() : "";
      const content = e.parameter.content ? e.parameter.content.toString().trim() : "";
      const type = e.parameter.type ? e.parameter.type.toString().trim() : "admin";
      const region = e.parameter.region ? e.parameter.region.toString().trim() : "ALL";
      const visible = e.parameter.visible ? e.parameter.visible.toString().trim().toUpperCase() : "Y";
      const isImportant = e.parameter.isImportant ? e.parameter.isImportant.toString().trim().toUpperCase() : "N";

      if (!title || !content) {
        return makeJsonResponse({ result: "fail", message: "제목과 내용을 입력해주세요." });
      }

      let isRegistered = false;
      let userObj = { id: userId, name: '', group: '', region: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          const cellId = uData[i][0] ? uData[i][0].toString().trim() : "";
          if (cellId && cellId === userId.toString().trim()) {
            isRegistered = true;
            userObj = {
              id: cellId,
              name: uData[i][1] ? uData[i][1].toString().trim() : "",
              group: formatGroupString(uData[i][2]),
              region: uData[i][3] ? uData[i][3].toString().trim() : "",
              role: uData[i][4] ? uData[i][4].toString().trim() : ""
            };
            break;
          }
        }
      }

      if (!isRegistered) {
        return makeJsonResponse({ result: "fail", message: "미등록 사용자입니다." });
      }

      const roleStr = userObj.role || "";
      const isSuperAdmin = roleStr.includes("전체관리자") || roleStr.includes("관리자");
      const isGroupAdmin = roleStr.includes("조장") || roleStr.includes("부조장");

      if (!isSuperAdmin && !isGroupAdmin) {
        return makeJsonResponse({ result: "fail", message: "공지사항 저장 권한이 없습니다." });
      }

      let finalType = type;
      let finalGroup = "";
      let finalRegion = region;

      if (isSuperAdmin) {
        finalType = type;
        if (type === "group") {
          finalGroup = userObj.group;
        }
        finalRegion = region;
      } else if (isGroupAdmin) {
        finalType = "group";
        finalGroup = userObj.group;
        finalRegion = userObj.region; // 조장은 본인의 지역으로만 공지 작성 강제
      }

      let noticeSheet = ss.getSheetByName("Notices");
      if (!noticeSheet) {
        noticeSheet = ss.insertSheet("Notices");
        noticeSheet.appendRow(["ID", "Type", "Author ID", "Author Name", "Group", "Title", "Content", "Region", "Visible", "Created At", "Is Important", "Author Role"]);
      }

      const nextId = noticeSheet.getLastRow() > 0 ? noticeSheet.getLastRow() : 1;
      const createdAt = new Date();

      noticeSheet.appendRow([
        nextId,
        finalType,
        userId.toString(),
        userObj.name,
        finalGroup,
        title,
        content,
        finalRegion,
        visible,
        createdAt,
        isImportant,
        userObj.role
      ]);

      return makeJsonResponse({ result: "success", message: "공지사항이 성공적으로 등록되었습니다." });
    }

    // 💡 2.7. 공지사항 공개/비공개 토글 API
    if (action === "toggleNoticeVisibility") {
      const noticeId = e.parameter.noticeId ? e.parameter.noticeId.toString().trim() : "";
      const newVisible = e.parameter.visible ? e.parameter.visible.toString().trim().toUpperCase() : "Y";

      let isRegistered = false;
      let userObj = { id: userId, group: '', region: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][0] && uData[i][0].toString().trim() === userId.toString().trim()) {
            isRegistered = true;
            userObj = {
              id: userId,
              group: formatGroupString(uData[i][2]),
              region: uData[i][3] ? uData[i][3].toString().trim() : "",
              role: uData[i][4] ? uData[i][4].toString().trim() : ""
            };
            break;
          }
        }
      }

      if (!isRegistered) {
        return makeJsonResponse({ result: "fail", message: "미등록 사용자입니다." });
      }

      const roleStr = userObj.role || "";
      const isSuperAdmin = roleStr.includes("전체관리자") || roleStr.includes("관리자");
      const isGroupAdmin = roleStr.includes("조장") || roleStr.includes("부조장");

      if (!isSuperAdmin && !isGroupAdmin) {
        return makeJsonResponse({ result: "fail", message: "권한이 없습니다." });
      }

      let noticeSheet = ss.getSheetByName("Notices");
      if (noticeSheet && noticeId) {
        const nData = noticeSheet.getDataRange().getValues();
        for (let i = 1; i < nData.length; i++) {
          const rowId = nData[i][0] ? nData[i][0].toString().trim() : "";
          if (rowId === noticeId) {
            const noticeType = nData[i][1] ? nData[i][1].toString().trim() : "";
            const noticeRegion = nData[i][7] ? nData[i][7].toString().trim() : "";

            let hasAuth = false;
            if (isSuperAdmin) {
              hasAuth = true;
            } else if (isGroupAdmin && noticeRegion === userObj.region) {
              hasAuth = true;
            }

            if (!hasAuth) {
              return makeJsonResponse({ result: "fail", message: "해당 공지사항을 수정할 권한이 없습니다." });
            }

            noticeSheet.getRange(i + 1, 9).setValue(newVisible); // Column I (9열)
            return makeJsonResponse({ result: "success", message: "공지 공개 여부가 수정되었습니다." });
          }
        }
      }

      return makeJsonResponse({ result: "fail", message: "공지사항을 찾을 수 없습니다." });
    }

    // 💡 2.8. 공지사항 삭제 API
    if (action === "deleteNotice") {
      const noticeId = e.parameter.noticeId ? e.parameter.noticeId.toString().trim() : "";

      let isRegistered = false;
      let userObj = { id: userId, group: '', region: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][0] && uData[i][0].toString().trim() === userId.toString().trim()) {
            isRegistered = true;
            userObj = {
              id: userId,
              group: formatGroupString(uData[i][2]),
              region: uData[i][3] ? uData[i][3].toString().trim() : "",
              role: uData[i][4] ? uData[i][4].toString().trim() : ""
            };
            break;
          }
        }
      }

      if (!isRegistered) {
        return makeJsonResponse({ result: "fail", message: "미등록 사용자입니다." });
      }

      const roleStr = userObj.role || "";
      const isSuperAdmin = roleStr.includes("전체관리자") || roleStr.includes("관리자");
      const isGroupAdmin = roleStr.includes("조장") || roleStr.includes("부조장");

      if (!isSuperAdmin && !isGroupAdmin) {
        return makeJsonResponse({ result: "fail", message: "권한이 없습니다." });
      }

      let noticeSheet = ss.getSheetByName("Notices");
      if (noticeSheet && noticeId) {
        const nData = noticeSheet.getDataRange().getValues();
        for (let i = 1; i < nData.length; i++) {
          const rowId = nData[i][0] ? nData[i][0].toString().trim() : "";
          if (rowId === noticeId) {
            const noticeType = nData[i][1] ? nData[i][1].toString().trim() : "";
            const noticeRegion = nData[i][7] ? nData[i][7].toString().trim() : "";

            let hasAuth = false;
            if (isSuperAdmin) {
              hasAuth = true;
            } else if (isGroupAdmin && noticeRegion === userObj.region) {
              hasAuth = true;
            }

            if (!hasAuth) {
              return makeJsonResponse({ result: "fail", message: "해당 공지사항을 삭제할 권한이 없습니다." });
            }

            noticeSheet.deleteRow(i + 1);
            return makeJsonResponse({ result: "success", message: "공지사항이 삭제되었습니다." });
          }
        }
      }

      return makeJsonResponse({ result: "fail", message: "공지사항을 찾을 수 없습니다." });
    }

    // 💡 3. 미취합자(선택 날짜 당일 결과 + 내일 계획) 및 미달성자, 휴무일 스킵 API
    if (action === "getRiskData") {
      const targetDateStr = e.parameter.date || getTodayString();
      const targetDate = new Date(targetDateStr);

      let userObj = { id: userId, group: '', region: '', role: '' };
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (uData[i][0] && uData[i][0].toString().trim() === userId.toString().trim()) {
            userObj = {
              id: userId,
              group: formatGroupString(uData[i][2]),
              region: uData[i][3] ? uData[i][3].toString().trim() : "",
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
            // 조장 권한인 경우 본인 지역 소속의 유저들만 통계 수집
            if (!isSuperAdmin && region !== userObj.region) {
              continue;
            }

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
                    // code.js & index.html 공통 검증 코드
                    // 기존 code.js line 454 부근
                    if ((r.category === '찾기' || r.category === '찾기(오프라인)' || r.category === '찾기(온라인)') &&
                      (r.type === '온만찾' || r.type === '오프만찾')) { // ✅ 오프번찾 제거 확인
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
// 순수 객체 반환 헬퍼 (getInitialData 등에서 재사용, 5열/4열 하이브리드 지원)
function getCurrentSemesterObj() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const semSheet = ss.getSheetByName("Semesters");

  if (!semSheet) {
    return { name: "2026-2학기 개강", startDate: "2026-09-01", endDate: "2026-12-31" };
  }

  const data = semSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length < 4) continue;

    let name = String(row[0]).trim();
    let weekName = "";
    let startVal = null;
    let endVal = null;
    let isCurrent = false;

    if (row.length >= 5) {
      weekName = String(row[1]).trim();
      startVal = row[2];
      endVal = row[3];
      isCurrent = String(row[4]).trim().toUpperCase() === "Y";
    } else {
      startVal = row[1];
      endVal = row[2];
      isCurrent = String(row[3]).trim().toUpperCase() === "Y";
    }

    if (isCurrent && startVal && endVal) {
      try {
        return {
          name: weekName ? `${name} ${weekName}` : name,
          startDate: formatDateStr(new Date(startVal)),
          endDate: formatDateStr(new Date(endVal))
        };
      } catch (e) { }
    }
  }
  return { name: "2026-2학기 개강", startDate: "2026-09-01", endDate: "2026-12-31" };
}

// 📅 선택된 날짜가 속한 주차(또는 사이클) 이름을 가져오는 헬퍼 함수
function getWeekNameForDate(targetDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const semSheet = ss.getSheetByName("Semesters");
  if (!semSheet) return "";

  const data = semSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length < 4) continue;

    let name = String(row[0]).trim();
    let weekName = "";
    let startVal = null;
    let endVal = null;

    if (row.length >= 5) {
      weekName = String(row[1]).trim();
      startVal = row[2];
      endVal = row[3];
    } else {
      startVal = row[1];
      endVal = row[2];
    }

    if (startVal && endVal) {
      try {
        const startStr = formatDateStr(new Date(startVal));
        const endStr = formatDateStr(new Date(endVal));
        if (targetDate >= startStr && targetDate <= endStr) {
          return weekName ? weekName : name;
        }
      } catch (e) { }
    }
  }
  return "";
}

// 📅 대시보드/활동 추적의 주차 목록을 조회하는 헬퍼 함수
function getSemestersList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const semSheet = ss.getSheetByName("Semesters");
  const list = [];
  if (!semSheet) return list;

  const data = semSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length < 4) continue;

    let name = String(row[0]).trim();
    let weekName = "";
    let startVal = null;
    let endVal = null;
    let isCurrent = false;

    if (row.length >= 5) {
      weekName = String(row[1]).trim();
      startVal = row[2];
      endVal = row[3];
      isCurrent = String(row[4]).trim().toUpperCase() === "Y";
    } else {
      startVal = row[1];
      endVal = row[2];
      isCurrent = String(row[3]).trim().toUpperCase() === "Y";
    }

    if (startVal && endVal) {
      try {
        list.push({
          name: weekName ? `${name} ${weekName}` : name,
          startDate: formatDateStr(new Date(startVal)),
          endDate: formatDateStr(new Date(endVal)),
          isCurrent: isCurrent
        });
      } catch (e) { }
    }
  }
  return list;
}

// 💡 Supabase 접속 정보 및 텔레그램 설정
const SUPABASE_URL = "https://qrrnvmskijxwtutgfrhf.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFycm52bXNraWp4d3R1dGdmcmhmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODUxNDAyNTUsImV4cCI6MjEwMDcxNjI1NX0.-EcLv73jL7yw3vwGQqIQ330ajq6_HHy2UDVNxmpzGa4";
const TELEGRAM_BOT_TOKEN = "YOUR_TELEGRAM_BOT_TOKEN"; // 텔레그램 봇 토큰 입력
const TELEGRAM_CHAT_ID = "-1003955494530";             // 텔레그램 그룹방 ID

// ⏰ 구글 서버가 주기적으로(예: 5분마다) 자동 실행할 함수
function checkAndSendTelegramAlarms() {
  const now = new Date();
  const currentHHmm = Utilities.formatDate(now, "Asia/Seoul", "HH:mm");
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");

  const headers = {
    "apikey": SUPABASE_ANON_KEY,
    "Authorization": `Bearer ${SUPABASE_ANON_KEY}`
  };

  // 1. Supabase의 alarm_schedules 테이블에서 활성화(is_active=true) 및 현재 시각(currentHHmm) 일치 알람 조회
  const alarmUrl = `${SUPABASE_URL}/rest/v1/alarm_schedules?is_active=eq.true&alarm_time=eq.${currentHHmm}`;
  const alarmRes = UrlFetchApp.fetch(alarmUrl, { headers: headers, muteHttpExceptions: true });

  if (alarmRes.getResponseCode() !== 200) return;
  const alarms = JSON.parse(alarmRes.getContentText());

  if (!alarms || alarms.length === 0) return; // 실행할 알람 없음

  // 2. 전체 유저 목록 및 당일 활동 기록 조회
  const usersRes = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/users?select=*`, { headers: headers });
  const allUsers = JSON.parse(usersRes.getContentText());

  const actsRes = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/activities?activity_date=eq.${todayStr}`, { headers: headers });
  const todayActs = JSON.parse(actsRes.getContentText());

  // 3. 알람 조건에 따른 미취합자 추출 및 메시지 생성
  alarms.forEach(alarm => {
    let uncollectedList = [];
    let titleText = "";

    if (alarm.category === 'today') {
      titleText = "[⚠️ 오늘 결과 미입력자 알림]\n\n오늘 활동 결과를 아직 입력하지 않은 멤버입니다. 마감 전 입력 바랍니다!";
      uncollectedList = (allUsers || []).filter(u => {
        if (u.is_exempt) return false;
        const act = (todayActs || []).find(a => a.name === u.name);
        return !act || act.status !== 'completed';
      });
    } else if (alarm.category === 'tomorrow') {
      titleText = "[📢 내일 일정 미입력자 알림]\n\n내일 일정을 아직 등록하지 않은 멤버입니다. 확인 후 작성 바랍니다!";
      // 내일 미취합자 체크 필요시 추가
    }

    if (uncollectedList.length > 0) {
      // 텔레그램 봇 유저네임 동적 감지 (버튼 링크 생성용)
      let botUsername = "Edu_Pom_test_bot";
      try {
        const meRes = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getMe`);
        const meData = JSON.parse(meRes.getContentText());
        if (meData.ok && meData.result.username) {
          botUsername = meData.result.username;
        }
      } catch (e) {}

      // 지역별 그룹화
      const grouped = {};
      uncollectedList.forEach(u => {
        const reg = u.region || '미정';
        if (!grouped[reg]) grouped[reg] = [];
        
        // 마크다운 문법 충돌 방지를 위한 특수문자 제거
        const cleanName = String(u.name).replace(/[\[\]\(\)\_\*]/g, "").trim();
        const mention = (u.id && !isNaN(u.id) && u.id.toString() !== "GUEST_USER" && !u.id.toString().startsWith("UNREG_"))
            ? `[${cleanName}](tg://user?id=${u.id})`
            : cleanName;
        grouped[reg].push(mention);
      });

      const lines = Object.keys(grouped).map(reg => `📍 [${reg}]: ${grouped[reg].join(', ')}`);
      const finalMsg = `${titleText}\n\n${lines.join('\n')}`;

      // 텔레그램 직발송
      UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ 
          chat_id: TELEGRAM_CHAT_ID, 
          text: finalMsg,
          parse_mode: 'Markdown',
          reply_markup: {
            inline_keyboard: [
              [
                { text: "📝 결과 입력하러 가기 (알림 자동 활성화)", web_app: { url: "https://msgo0303.github.io/eva_test_1/" } }
              ]
            ]
          }
        })
      });
    }
  });
}