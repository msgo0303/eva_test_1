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

      const grouped = {};
      uncollectedList.forEach(u => {
        const reg = u.region || '미정';
        if (!grouped[reg]) grouped[reg] = [];
        grouped[reg].push(u.name);
      });

      const lines = Object.keys(grouped).map(reg => {
        const heart = REGION_HEARTS[reg] || "📍";
        return `${heart} ${reg}\n${grouped[reg].join(', ')}`;
      });

      const finalMsg = `${titleHeader}\n\n${lines.join('\n\n')}`;

      // 텔레그램 메인방 직발송
      UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: finalMsg })
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