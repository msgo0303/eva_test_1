// ============================================================
// eva.gs — 실제 사용 코드만 정리한 버전
// 역할 1: doGet()     → index.html에서 알람 등록 시 구글 시트 동기화
// 역할 2: checkAndSendTelegramAlarms() → 5분 트리거로 텔레그램 발송
// ============================================================

const SPREADSHEET_ID = "1gbKtDSrsSbrI4G0YqXEZZ2k0tOfIw_tMZLkAPPsDzCw";

// 🔐 모든 민감 정보는 AlarmSchedules 시트 X열에서 읽어옴 (코드에 직접 입력 금지)
//    X1 : 텔레그램 봇 토큰
//    X2 : 전송 대상 그룹방 Chat ID
//    X3 : Supabase URL
//    X4 : Supabase Anon Key
function getConfig() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("AlarmSchedules");
  return {
    botToken   : sheet.getRange("X1").getValue().toString().trim(),
    chatId     : sheet.getRange("X2").getValue().toString().trim(),
    supabaseUrl: sheet.getRange("X3").getValue().toString().trim(),
    supabaseKey: sheet.getRange("X4").getValue().toString().trim()
  };
}


// ============================================================
// [1] doGet — index.html이 알람 추가 시 구글 시트(AlarmSchedules 탭)에 동기화
//     호출: fetch(`${GAS_URL}?action=addAlarmSchedule&category=...&alarmTime=...`)
// ============================================================
function doGet(e) {
  try {
    const action = e ? e.parameter.action : "";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === "addAlarmSchedule") {
      const category = e.parameter.category;
      const alarmTime = e.parameter.alarmTime;
      const isActive = e.parameter.isActive || "TRUE";

      let alarmSheet = ss.getSheetByName("AlarmSchedules");
      if (!alarmSheet) {
        alarmSheet = ss.insertSheet("AlarmSchedules");
        alarmSheet.appendRow(["id", "category", "alarm_time", "is_active", "created_at"]);
      }

      const nextId = alarmSheet.getLastRow() > 0 ? alarmSheet.getLastRow() : 1;
      alarmSheet.appendRow([
        nextId,
        category,
        alarmTime,
        isActive.toString().toUpperCase(),
        new Date()
      ]);

      return makeJsonResponse({ result: "success", message: "시트에 알람이 등록되었습니다." });
    }

    if (action === "deleteAlarmSchedule") {
      const category = e.parameter.category;
      const alarmTime = e.parameter.alarmTime;
      
      let alarmSheet = ss.getSheetByName("AlarmSchedules");
      if (alarmSheet) {
        const data = alarmSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][1] === category && data[i][2] === alarmTime) {
            alarmSheet.deleteRow(i + 1);
            break;
          }
        }
      }
      return makeJsonResponse({ result: "success", message: "시트에서 알람이 삭제되었습니다." });
    }

    if (action === "toggleAlarmActive") {
      const category = e.parameter.category;
      const alarmTime = e.parameter.alarmTime;
      const isActive = e.parameter.isActive || "TRUE";
      
      let alarmSheet = ss.getSheetByName("AlarmSchedules");
      if (alarmSheet) {
        const data = alarmSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][1] === category && data[i][2] === alarmTime) {
            alarmSheet.getRange(i + 1, 4).setValue(isActive.toString().toUpperCase());
            break;
          }
        }
      }
      return makeJsonResponse({ result: "success", message: "시트 알람 상태가 변경되었습니다." });
    }

    if (action === "logRegisterClick") {
      const userId = e.parameter.userId;
      const name = e.parameter.name;
      const username = e.parameter.username || "";

      let clickSheet = ss.getSheetByName("RegisterClicks");
      if (!clickSheet) {
        clickSheet = ss.insertSheet("RegisterClicks");
        clickSheet.appendRow(["user_id", "name", "username", "clicked_at"]);
      }

      clickSheet.appendRow([
        userId,
        name,
        username,
        new Date()
      ]);

      return makeJsonResponse({ result: "success", message: "클릭 기록이 시트에 저장되었습니다." });
    }

    if (action === "getWeeklyReportTemplate") {
      const range = e.parameter.range || "week";
      let templateKey = "weekly_report";
      if (range === "today") templateKey = "daily_report";
      if (range === "month") templateKey = "monthly_report";

      let templateSheet = ss.getSheetByName("MessageTemplates");
      let templateText = "";
      if (templateSheet) {
        const data = templateSheet.getDataRange().getValues();
        // 1. 요청된 특정 키로 검색
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === templateKey) {
            templateText = data[i][1].toString();
            break;
          }
        }
        // 2. 만약 해당 키가 시트에 없으면 기본 weekly_report로 폴백
        if (!templateText && templateKey !== "weekly_report") {
          for (let i = 1; i < data.length; i++) {
            if (data[i][0] === "weekly_report") {
              templateText = data[i][1].toString();
              break;
            }
          }
        }
      }
      
      if (!templateText) {
        if (range === "today") {
          templateText = 
            "📢 [일일 활동 리포트]\n" +
            "일자: {start_date}\n\n" +
            "📈 활동 결과 비중:\n" +
            "- 찾기(오프): {find_off}건\n" +
            "- 찾기(온): {find_on}건\n" +
            "- 매칭: {match}건\n" +
            "- 잎사귀: {leaf}건\n" +
            "- 복방: {book}건\n\n" +
            "📊 일일 활동 현황:\n" +
            "{weekly_regions_summary}\n\n" +
            "🔥 오늘도 수고하셨습니다!";
        } else if (range === "month") {
          templateText = 
            "📢 [월간 활동 리포트]\n" +
            "기간: {start_date} ~ {end_date}\n\n" +
            "📈 활동 결과 비중:\n" +
            "- 찾기(오프): {find_off}건\n" +
            "- 찾기(온): {find_on}건\n" +
            "- 매칭: {match}건\n" +
            "- 잎사귀: {leaf}건\n" +
            "- 복방: {book}건\n\n" +
            "📊 월간 미션 달성률: {mission_achieved}/{mission_target}개 ({mission_pct}%)\n" +
            "🍎 복음방 미션 달성률: {book_achieved}/{book_target}개 ({book_pct}%)\n\n" +
            "🔥 이번 개강 사이클도 수고하셨습니다!";
        } else {
          templateText = 
            "📢 [주간 활동 리포트]\n" +
            "기간: {start_date} ~ {end_date}\n\n" +
            "📈 활동 결과 비중:\n" +
            "- 찾기(오프): {find_off}건\n" +
            "- 찾기(온): {find_on}건\n" +
            "- 매칭: {match}건\n" +
            "- 잎사귀: {leaf}건\n" +
            "- 복방: {book}건\n\n" +
            "📊 주간 미션 달성률: {mission_achieved}/{mission_target}개 ({mission_pct}%)\n" +
            "🍎 복음방 미션 달성률: {book_achieved}/{book_target}개 ({book_pct}%)\n\n" +
            "🔥 이번 주도 수고하셨습니다!";
        }
      }
      
      return makeJsonResponse({ result: "success", template: templateText });
    }

    if (action === "registerUser") {
      const config = getConfig();
      let tgUserId = "GUEST_USER";

      if (!config.botToken) {
        Logger.log("⚠️ WARNING: botToken is not configured in AlarmSchedules sheet (X1). Skipping security verification.");
        tgUserId = parseTelegramUserIdFromInitData(e.parameter.initData || "");
      } else {
        const validatedUser = safeParseTelegramUser(e.parameter.initData || "", config.botToken);
        if (validatedUser) {
          tgUserId = validatedUser.id;
        } else {
          return makeJsonResponse({ result: "fail", message: "유효하지 않은 요청입니다. (인증 실패)" });
        }
      }

      const selectedUserId = e.parameter.selectedRow;

      const headers = {
        "apikey": config.supabaseKey,
        "Authorization": `Bearer ${config.supabaseKey}`,
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
      };

      const payload = {
        id: tgUserId
      };

      const url = `${config.supabaseUrl}/rest/v1/users?id=eq.${encodeURIComponent(selectedUserId)}`;
      const response = UrlFetchApp.fetch(url, {
        method: "patch",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      if (code >= 200 && code < 300) {
        return makeJsonResponse({ result: "success", message: "사용자 등록 완료" });
      } else {
        return makeJsonResponse({ result: "fail", message: "Supabase 업데이트 실패: " + response.getContentText() });
      }
    }

    if (action === "registerCustomUser") {
      const config = getConfig();
      let tgUserId = "GUEST_USER";

      if (!config.botToken) {
        Logger.log("⚠️ WARNING: botToken is not configured in AlarmSchedules sheet (X1). Skipping security verification.");
        tgUserId = parseTelegramUserIdFromInitData(e.parameter.initData || "");
      } else {
        const validatedUser = safeParseTelegramUser(e.parameter.initData || "", config.botToken);
        if (validatedUser) {
          tgUserId = validatedUser.id;
        } else {
          return makeJsonResponse({ result: "fail", message: "유효하지 않은 요청입니다. (인증 실패)" });
        }
      }

      const name = e.parameter.name;

      const headers = {
        "apikey": config.supabaseKey,
        "Authorization": `Bearer ${config.supabaseKey}`,
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
      };

      const payload = {
        id: tgUserId,
        name: name,
        role: "부원",
        region: "",
        group_name: "",
        book_count: 0,
        is_exempt: false
      };

      const url = `${config.supabaseUrl}/rest/v1/users`;
      const response = UrlFetchApp.fetch(url, {
        method: "post",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      if (code >= 200 && code < 300) {
        return makeJsonResponse({ result: "success", message: "사용자 가입 신청 완료" });
      } else {
        return makeJsonResponse({ result: "fail", message: "Supabase 저장 실패: " + response.getContentText() });
      }
    }

    return makeJsonResponse({ result: "fail", message: "지원하지 않는 action입니다." });

  } catch (err) {
    return makeJsonResponse({ error: err.message });
  }
}


// ============================================================
// [2] checkAndSendTelegramAlarms
//     GAS 트리거로 5분마다 자동 실행
//     → Supabase alarm_schedules 스캔 → 현재 시각 일치 알람 발견 시
//     → Supabase users/activities 조회 → 미취합자 계산 → 텔레그램 발송
// ============================================================
function checkAndSendTelegramAlarms() {
  const { botToken: TELEGRAM_BOT_TOKEN, chatId: TELEGRAM_CHAT_ID,
          supabaseUrl: SUPABASE_URL, supabaseKey: SUPABASE_ANON_KEY } = getConfig();

  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
    Logger.log("봇 토큰 또는 챗 ID가 설정되지 않았습니다. AlarmSchedules 시트 X1, X2 셀을 확인하세요.");
    return;
  }
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
    Logger.log("Supabase URL 또는 Key가 설정되지 않았습니다. AlarmSchedules 시트 X3, X4 셀을 확인하세요.");
    return;
  }

  const now = new Date();
  const currentHHmm = Utilities.formatDate(now, "Asia/Seoul", "HH:mm");
  const todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");

  const headers = {
    "apikey": SUPABASE_ANON_KEY,
    "Authorization": `Bearer ${SUPABASE_ANON_KEY}`
  };

  // 1. Supabase alarm_schedules에서 활성화 & 현재 시각 일치 알람 조회
  const alarmUrl = `${SUPABASE_URL}/rest/v1/alarm_schedules?is_active=eq.true&alarm_time=eq.${currentHHmm}`;
  const alarmRes = UrlFetchApp.fetch(alarmUrl, { headers: headers, muteHttpExceptions: true });

  if (alarmRes.getResponseCode() !== 200) return;
  const alarms = JSON.parse(alarmRes.getContentText());

  if (!alarms || alarms.length === 0) return; // 현재 시각에 실행할 알람 없음

  // 2. 전체 유저 목록 및 당일 활동 기록 조회
  const usersRes = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/users?select=*`, { headers: headers });
  const allUsers = JSON.parse(usersRes.getContentText());

  const actsRes = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/activities?activity_date=eq.${todayStr}`, { headers: headers });
  const todayActs = JSON.parse(actsRes.getContentText());

  // 3. 알람 조건별 미취합자 추출 및 메시지 발송
  alarms.forEach(alarm => {
    let uncollectedList = [];
    let titleHeader = "";

    if (alarm.category === 'today') {
      titleHeader = "⚠️ 오늘 결과 미입력자 알림\n\n오늘 활동 결과를 아직 입력하지 않은 인원입니다. 입력 바랍니다!";
      uncollectedList = (allUsers || []).filter(u => {
        if (u.is_exempt) return false;
        const act = (todayActs || []).find(a => a.name === u.name);
        return !act || act.status !== 'completed';
      });
    } else if (alarm.category === 'tomorrow') {
      titleHeader = "📢 내일 일정 미입력자 알림\n\n내일 일정을 아직 등록하지 않은 인원입니다. 작성 바랍니다!";
      // TODO: 내일 날짜 기준 activities 조회 후 미취합자 계산 로직 필요 시 추가
    }

    if (uncollectedList.length > 0) {
      const REGION_HEARTS = {
        "사당": "❤️", "안양": "🩷", "신림": "🧡", "신사": "💛", "금천": "💚",
        "군포": "🩵", "인덕원": "💙", "잠실": "💜", "양재": "🖤", "약수": "🩶",
        "서울시흥": "🤍", "서울역": "🤎", "새신자": "💖", "대학": "❣️"
      };

      // 1. 지역별로 유저 객체 그룹핑
      const groupedByRegion = {};
      uncollectedList.forEach(u => {
        const reg = u.region || '미정';
        if (!groupedByRegion[reg]) groupedByRegion[reg] = [];
        groupedByRegion[reg].push(u);
      });

      // 2. 최대 50명씩 끊어서 청크 나누기 (지역 보존 그리디 패킹)
      const chunks = [];
      let currentChunk = [];
      let currentChunkSize = 0;

      const regions = Object.keys(groupedByRegion);
      for (let j = 0; j < regions.length; j++) {
        const reg = regions[j];
        const regionUsers = groupedByRegion[reg];
        const count = regionUsers.length;

        if (count > 50) {
          if (currentChunk.length > 0) {
            chunks.push(currentChunk);
            currentChunk = [];
            currentChunkSize = 0;
          }
          // 50명씩 나눔
          for (let i = 0; i < count; i += 50) {
            chunks.push(regionUsers.slice(i, i + 50));
          }
        } else {
          if (currentChunkSize + count <= 50) {
            currentChunk.push(...regionUsers);
            currentChunkSize += count;
          } else {
            if (currentChunk.length > 0) {
              chunks.push(currentChunk);
            }
            currentChunk = [...regionUsers];
            currentChunkSize = count;
          }
        }
      }
      if (currentChunk.length > 0) {
        chunks.push(currentChunk);
      }

      // 3. 각 청크별 메시지 빌드 및 발송
      chunks.forEach((chunk, index) => {
        // 지역별 묶기
        const chunkGrouped = {};
        chunk.forEach(u => {
          const reg = u.region || '미정';
          if (!chunkGrouped[reg]) chunkGrouped[reg] = [];
          
          const cleanName = String(u.name).replace(/[\[\]\(\)\_\*]/g, "").trim();
          
          // 태그 처리: Markdown 파싱 모드에서 [이름](tg://user?id=아이디)
          const mention = (u.id && !isNaN(u.id) && u.id !== "GUEST_USER" && String(u.id).trim() !== "")
            ? `[${cleanName}](tg://user?id=${u.id})`
            : cleanName;
            
          chunkGrouped[reg].push(mention);
        });

        const lines = Object.keys(chunkGrouped).map(reg => {
          const heart = REGION_HEARTS[reg] || "📍";
          return `${heart} ${reg}\n${chunkGrouped[reg].join(', ')}`;
        });

        let chunkHeader = titleHeader;
        if (chunks.length > 1) {
          chunkHeader = `[메시지 ${index + 1}/${chunks.length} - ${chunk.length}명]\n${titleHeader}`;
        }

        const finalMsg = `${chunkHeader}\n\n${lines.join('\n\n')}`;

        // 텔레그램 메인방 발송 (Markdown 파싱 및 콜백 버튼 포함)
        const payload = {
          chat_id: TELEGRAM_CHAT_ID,
          text: finalMsg,
          parse_mode: 'Markdown',
          reply_markup: {
            inline_keyboard: [
              [
                { text: "내 정보 등록하기 (최초 1회는 클릭 필요)", callback_data: "register_click" }
              ]
            ]
          }
        };

        const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        Logger.log(`Scheduled alarm chunk ${index + 1} send response: ${res.getContentText()}`);

        // 레이트 리밋 방지를 위한 약간의 딜레이
        if (chunks.length > 1 && index < chunks.length - 1) {
          Utilities.sleep(200);
        }
      });
    }
  });
}


// ============================================================
// [헬퍼] makeJsonResponse — doGet 응답 생성
// ============================================================
function makeJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// [3] doPost — 텔레그램 웹훅 업데이트 처리 (callback_query 수집)
// ⚠️ 주의: 기존 출결 시스템의 doPost와 병합하여 사용하는 경우, 
// 아래 코드는 중복 선언 방지를 위해 주석 처리되었습니다. 병합본을 사용해 주세요.
// ============================================================
/*
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return makeJsonResponse({ result: "fail", message: "데이터가 없습니다." });
    }
    
    const postData = JSON.parse(e.postData.contents);
    
    if (postData.callback_query) {
      const callbackQuery = postData.callback_query;
      const data = callbackQuery.data;
      const from = callbackQuery.from;
      
      if (data === "register_click") {
        const userId = from.id ? from.id.toString() : "";
        const name = [from.first_name, from.last_name].filter(Boolean).join(" ") || "알수없음";
        const username = from.username || "";
        
        const config = getConfig();
        const botToken = config.botToken;
        
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let clickSheet = ss.getSheetByName("RegisterClicks");
        if (!clickSheet) {
          clickSheet = ss.insertSheet("RegisterClicks");
          clickSheet.appendRow(["user_id", "name", "username", "clicked_at"]);
        }
        
        clickSheet.appendRow([
          userId,
          name,
          username,
          new Date()
        ]);
        
        if (botToken) {
          const answerUrl = `https://api.telegram.org/bot${botToken}/answerCallbackQuery`;
          UrlFetchApp.fetch(answerUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
              callback_query_id: callbackQuery.id,
              text: `${name}님, 최초 1회 정보 등록이 완료되었습니다.`,
              show_alert: true
            })
          });
        }
      }
    }
    
    return makeJsonResponse({ result: "success" });
  } catch (err) {
    return makeJsonResponse({ error: err.message });
  }
}
*/


// ============================================================
// [4] sendWeeklyTelegramReport — 매주 수요일/일요일 오전 9시 주간 리포트 발송
// ============================================================
function sendWeeklyTelegramReport() {
  const { botToken, chatId, supabaseUrl, supabaseKey } = getConfig();

  if (!botToken || !chatId) {
    Logger.log("봇 토큰 또는 챗 ID가 설정되지 않았습니다.");
    return;
  }
  if (!supabaseUrl || !supabaseKey) {
    Logger.log("Supabase URL 또는 Key가 설정되지 않았습니다.");
    return;
  }

  // 1. KST 기준 날짜 계산
  const now = new Date();
  const kstDateStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");
  const parts = kstDateStr.split('-');
  const kstYear = parseInt(parts[0], 10);
  const kstMonth = parseInt(parts[1], 10) - 1;
  const kstDay = parseInt(parts[2], 10);
  const kstNow = new Date(kstYear, kstMonth, kstDay);

  const day = kstNow.getDay(); // 0: Sun, 1: Mon, ..., 6: Sat
  
  const sunday = new Date(kstNow);
  sunday.setDate(kstNow.getDate() - day);
  
  const saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() + 6);
  
  const startStr = Utilities.formatDate(sunday, "Asia/Seoul", "yyyy-MM-dd");
  const endStr = Utilities.formatDate(saturday, "Asia/Seoul", "yyyy-MM-dd");

  const headers = {
    "apikey": supabaseKey,
    "Authorization": `Bearer ${supabaseKey}`
  };

  // 2. 전체 유저 목록 조회
  const usersUrl = `${supabaseUrl}/rest/v1/users?select=*`;
  const usersRes = UrlFetchApp.fetch(usersUrl, { headers: headers });
  const allUsers = JSON.parse(usersRes.getContentText()) || [];

  // 3. 이번 주(일요일~토요일) 활동 조회
  const actsUrl = `${supabaseUrl}/rest/v1/activities?activity_date=gte.${startStr}&activity_date=lte.${endStr}`;
  const actsRes = UrlFetchApp.fetch(actsUrl, { headers: headers });
  const weeklyActs = JSON.parse(actsRes.getContentText()) || [];

  // 4. 활동 결과 집계
  const resultCounts = { '찾기(오프)': 0, '찾기(온)': 0, '매칭': 0, '잎사귀': 0, '복방': 0 };
  const regionFindResult = {};
  
  weeklyActs.forEach(item => {
    if (item.status === 'completed' && item.result_data) {
      let parsed = [];
      try {
        parsed = typeof item.result_data === 'string' ? JSON.parse(item.result_data) : item.result_data;
      } catch (e) { }
      
      if (Array.isArray(parsed)) {
        parsed.forEach(r => {
          const countVal = (r.count !== undefined ? parseInt(r.count, 10) : 1);
          
          if (r.type === '합자찾(오프)') {
            resultCounts['찾기(오프)'] += countVal;
          }
          if (r.type === '합자찾(온)') {
            resultCounts['찾기(온)'] += countVal;
          }
          if (r.category === '매칭' && r.type !== '취소') {
            resultCounts['매칭'] += (countVal || 1);
          }
          if (r.category === '잎사귀') {
            resultCounts['잎사귀'] += (countVal || 1);
          }
          if (r.category === '복방' && r.type !== '취소') {
            resultCounts['복방'] += (countVal || 1);
          }
          
          if (r.type === '합자찾(오프)' || r.type === '합자찾(온)') {
            const reg = item.region || '미정';
            regionFindResult[reg] = (regionFindResult[reg] || 0) + countVal;
          }
        });
      }
    }
  });

  // 5. 달성률 분모/분자 계산
  const regionTargets = {};
  const bookTargets = {};
  const bookAchieved = {};
  const FIXED_REGIONS = [
    "사당", "안양", "신림", "신사", "금천", "군포", "인덕원",
    "잠실", "양재", "약수", "서울시흥", "서울역", "새신자", "대학"
  ];

  allUsers.forEach(u => {
    if (!u.is_exempt) {
      const reg = u.region || '미정';
      // 주간 미션
      if ((u.book_count || 0) === 0) {
        regionTargets[reg] = (regionTargets[reg] || 0) + 6;
      }
      // 복방
      bookTargets[reg] = (bookTargets[reg] || 0) + 1;
      bookAchieved[reg] = (bookAchieved[reg] || 0) + (u.book_count || 0);
    }
  });

  const totalFindAchieved = Object.values(regionFindResult).reduce((a, b) => a + b, 0);
  const totalFindTarget = Object.values(regionTargets).reduce((a, b) => a + b, 0);
  const totalBookAchieved = Object.values(bookAchieved).reduce((a, b) => a + b, 0);
  const totalBookTarget = Object.values(bookTargets).reduce((a, b) => a + b, 0);

  // 6. QuickChart.io 차트 이미지 URL 생성
  // 차트 1: 도넛 차트 (결과 비중)
  const doughnutConfig = {
    type: 'doughnut',
    data: {
      labels: ['찾기(오프)', '찾기(온)', '매칭', '잎사귀', '복방'],
      datasets: [{
        data: [
          resultCounts['찾기(오프)'] || 0,
          resultCounts['찾기(온)'] || 0,
          resultCounts['매칭'] || 0,
          resultCounts['잎사귀'] || 0,
          resultCounts['복방'] || 0
        ],
        backgroundColor: ['#2563eb', '#60a5fa', '#16a34a', '#d97706', '#9333ea'],
        borderWidth: 2,
        borderColor: '#ffffff'
      }]
    },
    options: {
      cutout: '70%',
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { family: 'Noto Sans KR', size: 11, weight: 'bold' },
            color: '#18181b',
            boxWidth: 12,
            padding: 10
          }
        },
        title: {
          display: true,
          text: '📈 활동 결과 비중',
          font: { family: 'Noto Sans KR', size: 14, weight: 'bold' },
          color: '#18181b',
          padding: 15
        },
        datalabels: {
          display: false
        }
      }
    }
  };
  const doughnutChartUrl = `https://quickchart.io/chart?bkg=%23ffffff&w=400&h=250&c=${encodeURIComponent(JSON.stringify(doughnutConfig))}`;

  // 차트 2: 가로 막대 차트 (주간 미션)
  const weeklyAchievedData = FIXED_REGIONS.map(r => regionFindResult[r] || 0);
  const weeklyTargetsData = FIXED_REGIONS.map(r => regionTargets[r] || 0);
  const weeklyPctData = FIXED_REGIONS.map(r => {
    const t = regionTargets[r] || 0;
    const a = regionFindResult[r] || 0;
    return t > 0 ? Math.min(Math.round((a / t) * 100), 100) : 0;
  });
  const weeklyBackgroundColors = weeklyPctData.map(v => v >= 100 ? '#16a34a' : v >= 60 ? '#60a5fa' : '#f87171');
  const weeklyConfig = {
    type: 'bar',
    data: {
      labels: FIXED_REGIONS,
      datasets: [{
        label: '달성률(%)',
        data: weeklyPctData,
        achieved: weeklyAchievedData,
        targets: weeklyTargetsData,
        backgroundColor: weeklyBackgroundColors,
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      layout: {
        padding: { left: 10, right: 40, top: 10, bottom: 10 }
      },
      plugins: {
        legend: { display: false },
        title: {
          display: true,
          text: '📊 주간 미션 달성 현황(%)',
          font: { family: 'Noto Sans KR', size: 14, weight: 'bold' },
          color: '#18181b',
          padding: 10
        },
        datalabels: {
          display: true,
          formatter: `function(value, context) {
            var index = context.dataIndex;
            var achieved = context.dataset.achieved[index];
            var target = context.dataset.targets[index];
            if (value === 0 && achieved === 0) return '';
            return achieved + '/' + target + '개 (' + value + '%)';
          }`,
          anchor: 'end',
          align: `function(context) {
            return context.dataset.data[context.dataIndex] >= 30 ? 'start' : 'end';
          }`,
          color: `function(context) {
            return context.dataset.data[context.dataIndex] >= 30 ? '#ffffff' : '#374151';
          }`,
          font: { family: 'Noto Sans KR', size: 10, weight: 'bold' }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          max: 100,
          ticks: {
            callback: `function(value) { return value + '%'; }`,
            font: { family: 'Noto Sans KR', size: 10 }
          },
          grid: {
            color: '#e4e4e7',
            borderDash: [4, 4]
          }
        },
        y: {
          ticks: {
            font: { family: 'Noto Sans KR', size: 10, weight: 'bold' },
            color: '#18181b'
          },
          grid: {
            display: false
          }
        }
      }
    }
  };
  const weeklyChartUrl = `https://quickchart.io/chart?bkg=%23ffffff&w=500&h=450&c=${encodeURIComponent(JSON.stringify(weeklyConfig))}`;

  // 차트 3: 가로 막대 차트 (복방 미션)
  const bookAchievedData = FIXED_REGIONS.map(r => bookAchieved[r] || 0);
  const bookTargetsData = FIXED_REGIONS.map(r => bookTargets[r] || 0);
  const bookPctData = FIXED_REGIONS.map(r => {
    const t = bookTargets[r] || 0;
    const a = bookAchieved[r] || 0;
    return t > 0 ? Math.min(Math.round((a / t) * 100), 100) : 0;
  });
  const bookBackgroundColors = bookPctData.map(v => v >= 100 ? '#16a34a' : v >= 60 ? '#60a5fa' : '#f87171');
  const bookConfig = {
    type: 'bar',
    data: {
      labels: FIXED_REGIONS,
      datasets: [{
        label: '달성률(%)',
        data: bookPctData,
        achieved: bookAchievedData,
        targets: bookTargetsData,
        backgroundColor: bookBackgroundColors,
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      layout: {
        padding: { left: 10, right: 40, top: 10, bottom: 10 }
      },
      plugins: {
        legend: { display: false },
        title: {
          display: true,
          text: '🍎 복음방 미션 달성 현황(%)',
          font: { family: 'Noto Sans KR', size: 14, weight: 'bold' },
          color: '#18181b',
          padding: 10
        },
        datalabels: {
          display: true,
          formatter: `function(value, context) {
            var index = context.dataIndex;
            var achieved = context.dataset.achieved[index];
            var target = context.dataset.targets[index];
            if (value === 0 && achieved === 0) return '';
            return achieved + '/' + target + '개 (' + value + '%)';
          }`,
          anchor: 'end',
          align: `function(context) {
            return context.dataset.data[context.dataIndex] >= 30 ? 'start' : 'end';
          }`,
          color: `function(context) {
            return context.dataset.data[context.dataIndex] >= 30 ? '#ffffff' : '#374151';
          }`,
          font: { family: 'Noto Sans KR', size: 10, weight: 'bold' }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          max: 100,
          ticks: {
            callback: `function(value) { return value + '%'; }`,
            font: { family: 'Noto Sans KR', size: 10 }
          },
          grid: {
            color: '#e4e4e7',
            borderDash: [4, 4]
          }
        },
        y: {
          ticks: {
            font: { family: 'Noto Sans KR', size: 10, weight: 'bold' },
            color: '#18181b'
          },
          grid: {
            display: false
          }
        }
      }
    }
  };
  const bookChartUrl = `https://quickchart.io/chart?bkg=%23ffffff&w=500&h=450&c=${encodeURIComponent(JSON.stringify(bookConfig))}`;

  // 7. 구글 시트에서 메시지 템플릿 로드
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let templateSheet = ss.getSheetByName("MessageTemplates");
  if (!templateSheet) {
    templateSheet = ss.insertSheet("MessageTemplates");
    templateSheet.appendRow(["template_key", "template_text", "description"]);
    
    const defaultTemplate = 
      "📢 [주간 활동 리포트]\n" +
      "기간: {start_date} ~ {end_date}\n\n" +
      "📈 활동 결과 비중:\n" +
      "- 찾기(오프): {find_off}건\n" +
      "- 찾기(온): {find_on}건\n" +
      "- 매칭: {match}건\n" +
      "- 잎사귀: {leaf}건\n" +
      "- 복방: {book}건\n\n" +
      "📊 주간 미션 달성률: {mission_achieved}/{mission_target}개 ({mission_pct}%)\n" +
      "🍎 복음방 미션 달성률: {book_achieved}/{book_target}개 ({book_pct}%)\n\n" +
      "🔥 이번 주도 수고하셨습니다!";
      
    templateSheet.appendRow([
      "weekly_report",
      defaultTemplate,
      "수요일/일요일 오전 9시 주간 리포트 발송용 템플릿"
    ]);
  }

  const data = templateSheet.getDataRange().getValues();
  let templateText = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "weekly_report") {
      templateText = data[i][1].toString();
      break;
    }
  }

  if (!templateText) {
    templateText = 
      "📢 [주간 활동 리포트]\n" +
      "기간: {start_date} ~ {end_date}\n\n" +
      "📈 활동 결과:\n" +
      "- 찾기(오프): {find_off}건\n" +
      "- 찾기(온): {find_on}건\n" +
      "- 매칭: {match}건\n" +
      "- 잎사귀: {leaf}건\n" +
      "- 복방: {book}건\n\n" +
      "📊 주간 미션 달성률: {mission_achieved}/{mission_target}개 ({mission_pct}%)\n" +
      "🍎 복음방 미션 달성률: {book_achieved}/{book_target}개 ({book_pct}%)\n\n" +
      "🔥 이번 주도 수고하셨습니다!";
  }

  // 8. 템플릿 변수 치환
  const reportDateStr = Utilities.formatDate(kstNow, "Asia/Seoul", "yyyy.MM.dd");
  
  const weeklyRegionsSummary = FIXED_REGIONS.map(r => {
    const t = regionTargets[r] || 0;
    const a = regionFindResult[r] || 0;
    const p = t > 0 ? Math.round((a / t) * 100) : 0;
    return `- ${r}: ${a}/${t}개 (${p}%)`;
  }).join('\n');

  const bookRegionsSummary = FIXED_REGIONS.map(r => {
    const t = bookTargets[r] || 0;
    const a = bookAchieved[r] || 0;
    const p = t > 0 ? Math.round((a / t) * 100) : 0;
    return `- ${r}: ${a}/${t}개 (${p}%)`;
  }).join('\n');

  const formattedMsg = templateText
    .replace(/{start_date}/g, startStr.replace(/-/g, '.')) // yyyy.MM.dd 포맷으로 통일
    .replace(/{end_date}/g, endStr.replace(/-/g, '.'))     // yyyy.MM.dd 포맷으로 통일
    .replace(/{report_date}/g, reportDateStr)
    .replace(/{weekly_regions_summary}/g, weeklyRegionsSummary)
    .replace(/{book_regions_summary}/g, bookRegionsSummary)
    .replace(/{find_off}/g, (resultCounts['찾기(오프)'] || 0).toString())
    .replace(/{find_on}/g, (resultCounts['찾기(온)'] || 0).toString())
    .replace(/{match}/g, (resultCounts['매칭'] || 0).toString())
    .replace(/{leaf}/g, (resultCounts['잎사귀'] || 0).toString())
    .replace(/{book}/g, (resultCounts['복방'] || 0).toString())
    .replace(/{mission_achieved}/g, totalFindAchieved.toString())
    .replace(/{mission_target}/g, totalFindTarget.toString())
    .replace(/{mission_pct}/g, (totalFindTarget > 0 ? Math.round((totalFindAchieved / totalFindTarget) * 100) : 0).toString())
    .replace(/{book_achieved}/g, totalBookAchieved.toString())
    .replace(/{book_target}/g, totalBookTarget.toString())
    .replace(/{book_pct}/g, (totalBookTarget > 0 ? Math.round((totalBookAchieved / totalBookTarget) * 100) : 0).toString());

  // 9. 텔레그램 sendMediaGroup API 전송
  const media = [
    {
      type: 'photo',
      media: doughnutChartUrl,
      caption: formattedMsg,
      parse_mode: 'Markdown'
    },
    {
      type: 'photo',
      media: weeklyChartUrl
    },
    {
      type: 'photo',
      media: bookChartUrl
    }
  ];

  const payload = {
    chat_id: chatId,
    media: JSON.stringify(media)
  };

  const response = UrlFetchApp.fetch(`https://api.telegram.org/bot${botToken}/sendMediaGroup`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log("sendWeeklyTelegramReport Response: " + response.getContentText());
}

// 수동 테스트 실행용 헬퍼 함수
function testSendWeeklyTelegramReport() {
  sendWeeklyTelegramReport();
}

function parseTelegramUserIdFromInitData(initDataStr) {
  if (!initDataStr) return "GUEST_USER";
  try {
    const decoded = decodeURIComponent(initDataStr);
    const userParam = decoded.split('&').find(p => p.startsWith('user='));
    if (userParam) {
      const userObj = JSON.parse(userParam.split('user=')[1]);
      return userObj.id.toString();
    }
  } catch (e) { }
  return "GUEST_USER";
}

function verifyTelegramInitData(initDataStr, botToken) {
  if (!initDataStr || !botToken) return false;

  try {
    const params = {};
    const pairs = initDataStr.split('&');
    let providedHash = "";

    for (let i = 0; i < pairs.length; i++) {
      const pair = pairs[i].split('=');
      if (pair.length < 2) continue;
      const key = decodeURIComponent(pair[0]);
      const val = decodeURIComponent(pair.slice(1).join('='));

      if (key === 'hash') {
        providedHash = val;
      } else {
        params[key] = val;
      }
    }

    if (!providedHash) {
      Logger.log("No hash found in initData");
      return false;
    }

    const sortedKeys = Object.keys(params).sort();
    const dataCheckArr = [];
    for (let i = 0; i < sortedKeys.length; i++) {
      dataCheckArr.push(sortedKeys[i] + '=' + params[sortedKeys[i]]);
    }
    const dataCheckString = dataCheckArr.join('\n');

    const secretKeyBytes = Utilities.computeHmacSignature(
      Utilities.MacAlgorithm.HMAC_SHA_256,
      botToken,
      "WebAppData",
      Utilities.Charset.UTF_8
    );

    const computedHashBytes = Utilities.computeHmacSignature(
      Utilities.MacAlgorithm.HMAC_SHA_256,
      dataCheckString,
      secretKeyBytes,
      Utilities.Charset.UTF_8
    );

    let computedHashHex = "";
    for (let i = 0; i < computedHashBytes.length; i++) {
      let byteVal = computedHashBytes[i];
      if (byteVal < 0) byteVal += 256;
      let hexByte = byteVal.toString(16);
      if (hexByte.length === 1) hexByte = "0" + hexByte;
      computedHashHex += hexByte;
    }

    return computedHashHex === providedHash;
  } catch (e) {
    Logger.log("verifyTelegramInitData error: " + e.message);
    return false;
  }
}

function safeParseTelegramUser(initDataStr, botToken) {
  if (!initDataStr) {
    Logger.log("initDataStr is empty");
    return null;
  }
  
  const isValid = verifyTelegramInitData(initDataStr, botToken);
  if (!isValid) {
    Logger.log("Telegram initData verification failed");
    return null;
  }
  
  try {
    const decoded = decodeURIComponent(initDataStr);
    const pairs = decoded.split('&');
    for (let i = 0; i < pairs.length; i++) {
      const pair = pairs[i].split('=');
      if (pair.length < 2) continue;
      const key = pair[0];
      const value = pair.slice(1).join('=');
      if (key === 'user') {
        const userObj = JSON.parse(value);
        return {
          id: userObj.id ? userObj.id.toString() : "",
          name: userObj.first_name || userObj.username || ""
        };
      }
    }
  } catch (e) {
    Logger.log("Failed to parse user from initData: " + e.message);
  }
  return null;
}