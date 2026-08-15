import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2.39.8";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function parseTelegramUser(initDataStr: string) {
  if (!initDataStr) return null;
  try {
    const params = new URLSearchParams(initDataStr);
    const userStr = params.get("user");
    if (userStr) {
      return JSON.parse(userStr);
    }
  } catch (e) {
    console.error("⚠️ parseTelegramUser error:", e);
  }
  return null;
}

async function verifyTelegramInitData(initDataStr: string, botToken: string): Promise<boolean> {
  try {
    const params = new URLSearchParams(initDataStr);
    const providedHash = params.get("hash");
    if (!providedHash) return false;

    // hash 파라미터 제외
    params.delete("hash");

    // 알파벳 순서대로 정렬 후 key=value\n 조립
    const keys = Array.from(params.keys()).sort();
    const dataCheckArr: string[] = [];
    for (const key of keys) {
      const val = params.get(key);
      if (val !== null) {
        dataCheckArr.push(`${key}=${val}`);
      }
    }
    const dataCheckString = dataCheckArr.join("\n");

    const encoder = new TextEncoder();
    
    // WebAppData 시크릿 키 생성 (HMAC-SHA256)
    const secretKeyMaterial = await crypto.subtle.importKey(
      "raw",
      encoder.encode("WebAppData"),
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["sign"]
    );
    const secretKeyBytes = await crypto.subtle.sign(
      "HMAC",
      secretKeyMaterial,
      encoder.encode(botToken)
    );

    // data_check_string 서명용 hmacKey 로드
    const hmacKey = await crypto.subtle.importKey(
      "raw",
      secretKeyBytes,
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["sign"]
    );

    const signature = await crypto.subtle.sign(
      "HMAC",
      hmacKey,
      encoder.encode(dataCheckString)
    );

    const signatureHex = Array.from(new Uint8Array(signature))
      .map(b => b.toString(16).padStart(2, "0"))
      .join("");

    return signatureHex === providedHash;
  } catch (e) {
    console.error("⚠️ verifyTelegramInitData error:", e);
    return false;
  }
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function getActivityResultSummary(act: any): string {
  const resultData = act.result_data || act.resultData;
  if (resultData) {
    try {
      const parsed = typeof resultData === "string" ? JSON.parse(resultData) : resultData;
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map((r: any) => {
          const category = r.category || "";
          const type = r.type || "";
          const count = r.count !== undefined ? parseInt(r.count, 10) : 0;

          if (type === "합자찾(오프)" || type === "합자찾(온)") {
            const prefix = category ? `${category} | ` : "";
            return `${prefix}${type} | ${count}개`;
          } else {
            if (category && type) {
              return `${category} | ${type}`;
            }
            return type || category || "";
          }
        }).join(", ");
      }
    } catch (e) {
      console.error("⚠️ Failed to parse result_data in helper:", e);
    }
  }
  return act.result_text || act.content || "내용 없음";
}

interface ResultItem {
  key: string;
  isCounted: boolean;
  count: number;
}

function getActivityResultItems(act: any): ResultItem[] {
  const resultData = act.result_data || act.resultData;
  if (resultData) {
    try {
      const parsed = typeof resultData === "string" ? JSON.parse(resultData) : resultData;
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map((r: any) => {
          const category = r.category || "";
          const type = r.type || "";
          const count = r.count !== undefined ? parseInt(r.count, 10) : 0;

          const isCounted = type === "합자찾(오프)" || type === "합자찾(온)";
          const prefix = category ? `${category} | ` : "";
          const key = `${prefix}${type}`;

          if (isCounted) {
            return { key, isCounted, count };
          } else {
            const fullKey = category && type ? `${category} | ${type}` : (type || category || "");
            return { key: fullKey, isCounted: false, count: 1 };
          }
        });
      }
    } catch (e) {
      console.error("⚠️ Failed to parse result_data in helper:", e);
    }
  }
  const fallbackText = act.result_text || act.content || "내용 없음";
  return [{ key: fallbackText, isCounted: false, count: 1 }];
}

function isCanceledActivity(act: any): boolean {
  const resultData = act.result_data || act.resultData;
  if (resultData) {
    try {
      const parsed = typeof resultData === "string" ? JSON.parse(resultData) : resultData;
      if (Array.isArray(parsed)) {
        return parsed.some((r: any) => {
          const category = r.category || "";
          const type = r.type || "";
          return (category === "복방" || category === "매칭") && type === "취소";
        });
      }
    } catch (e) {
      console.error("⚠️ Failed to parse result_data in cancellation check:", e);
    }
  }
  return false;
}

async function sendRealtimeActivityNotification(
  supabase: any,
  date: string,
  dbUser: any,
  payload: any,
  botToken: string,
  chatId: string
) {
  try {
    const { data: todayActs, error } = await supabase
      .from("activities")
      .select("name, region, content, result_text, result_data")
      .eq("activity_date", date)
      .eq("status", "completed");

    if (error) {
      console.error("⚠️ Failed to fetch today's activities for notification:", error);
      return;
    }

    // 복방/매칭 취소 활동은 목록에서 제외
    const activeTodayActs = (todayActs || []).filter(act => !isCanceledActivity(act));

    const userRegion = dbUser.region || "미정";
    const userName = dbUser.name || "알 수 없음";

    // 결과 중심의 요약 생성
    const activitySummary = getActivityResultSummary(payload);
    let summaryDetail = activitySummary;
    if (payload.result_text && payload.result_text !== activitySummary) {
      summaryDetail += ` (메모: ${payload.result_text})`;
    }

    const REGION_HEARTS: Record<string, string> = {
      "사당": "❤️", "안양": "🩷", "신림": "🧡", "신사": "💛", "금천": "💚",
      "군포": "🩵", "인덕원": "💙", "잠실": "💜", "양재": "🖤", "약수": "🩶",
      "서울시흥": "🤍", "서울역": "🤎", "새신자": "💖", "대학": "❣️"
    };

    const regionOrder = [
      "사당", "안양", "신림", "신사", "금천", "군포", "인덕원",
      "잠실", "양재", "약수", "서울시흥", "서울역", "새신자", "대학"
    ];

    const actsByRegion: Record<string, any[]> = {};
    for (const act of activeTodayActs) {
      const reg = act.region || "미정";
      if (!actsByRegion[reg]) actsByRegion[reg] = [];
      actsByRegion[reg].push(act);
    }

    const uniqueUserNames = new Set(activeTodayActs.map(act => act.name));
    const totalUniqueCount = uniqueUserNames.size;

    let message = `🏃 <b>실시간 활동 결과 등록 알림!</b>\n`;
    message += `[${escapeHtml(userRegion)}] <b>${escapeHtml(userName)}</b>님이 방금 결과를 등록했습니다.\n\n`;
    message += `💬 내용: ${escapeHtml(summaryDetail)}\n\n`;
    message += `──────────────────\n\n`;
    message += `📋 <b>오늘 완료자 명단 (총 ${totalUniqueCount}명)</b>\n`;

    const completedRegions = Object.keys(actsByRegion).sort((a, b) => {
      let idxA = regionOrder.indexOf(a);
      let idxB = regionOrder.indexOf(b);
      if (idxA === -1) idxA = 999;
      if (idxB === -1) idxB = 999;
      return idxA - idxB;
    });

    for (const reg of completedRegions) {
      const heart = REGION_HEARTS[reg] || "💙";
      message += `\n${heart} <b>${escapeHtml(reg)}</b>\n`;

      // 동일인별로 활동 결과를 병합하고 개수 합산
      const userNames: string[] = [];
      const userActs: Record<string, any[]> = {};
      for (const act of actsByRegion[reg]) {
        const name = act.name || "알 수 없음";
        if (!userActs[name]) {
          userActs[name] = [];
          userNames.push(name);
        }
        userActs[name].push(act);
      }

      for (const name of userNames) {
        const aggregated: Record<string, { isCounted: boolean; count: number }> = {};
        const keyOrder: string[] = [];

        for (const act of userActs[name]) {
          const items = getActivityResultItems(act);
          for (const item of items) {
            if (!aggregated[item.key]) {
              aggregated[item.key] = { isCounted: item.isCounted, count: 0 };
              keyOrder.push(item.key);
            }
            aggregated[item.key].count += item.count;
          }
        }

        const formattedResults = keyOrder.map((key) => {
          const agg = aggregated[key];
          if (agg.isCounted) {
            return `${key} | ${agg.count}개`;
          } else {
            if (agg.count > 1) {
              return `${key} | ${agg.count}개`;
            }
            return key;
          }
        });

        const summaryStr = formattedResults.join(", ");
        message += `- ${escapeHtml(name)} | ${escapeHtml(summaryStr)}\n`;
      }
    }

    const telegramUrl = `https://api.telegram.org/bot${botToken}/sendMessage`;
    const response = await fetch(telegramUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        chat_id: chatId,
        text: message,
        parse_mode: "HTML",
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`⚠️ Telegram API returned error: ${response.status} ${errorText}`);
    }
  } catch (err) {
    console.error("⚠️ sendRealtimeActivityNotification error:", err);
  }
}

serve(async (req) => {
  // CORS preflight 요청 처리
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { action, initData, params } = await req.json();

    const supabaseUrl = Deno.env.get("SUPABASE_URL")!;
    const supabaseServiceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const botToken = Deno.env.get("TELEGRAM_BOT_TOKEN");

    if (!botToken) {
      return new Response(
        JSON.stringify({ result: "fail", message: "서버 설정 에러: TELEGRAM_BOT_TOKEN 비밀값이 누락되었습니다." }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // 1. 보안 검증 파이프라인
    const isValid = await verifyTelegramInitData(initData || "", botToken);
    if (!isValid) {
      return new Response(
        JSON.stringify({ result: "fail", message: "유효하지 않은 텔레그램 인증 정보입니다." }),
        { status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const tgUser = parseTelegramUser(initData || "");
    if (!tgUser || !tgUser.id) {
      return new Response(
        JSON.stringify({ result: "fail", message: "텔레그램 유저 정보를 추출할 수 없습니다." }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }
    const verifiedTgUserId = tgUser.id.toString();

    // Supabase Service Role Client 생성 (서버 측에서 RLS를 우회하고 권한 검증 및 제어 진행)
    const supabase = createClient(supabaseUrl, supabaseServiceKey);

    // [보안 CUD 위임 API] 1. 활동 저장 및 취소 (saveActivity)
    if (action === "saveActivity") {
      const { mode, actId, date, startTime, endTime, location, content, status, resultData, resultText } = params;

      // 요청한 유저의 DB 프로필 조회 (F12 권한 위조 원천 차단)
      const { data: dbUsers, error: userErr } = await supabase
        .from("users")
        .select("*")
        .eq("id", verifiedTgUserId);

      if (userErr || !dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "등록되지 않은 사용자입니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const payload: any = {
        activity_date: date,
        start_time: startTime || "09:00",
        end_time: endTime || "10:00",
        location: location || "",
        content: content || "찾기",
        status: status,
        result_data: resultData ? (typeof resultData === 'string' ? JSON.parse(resultData) : resultData) : null,
        result_text: resultText || ""
      };

      if (mode === "edit") {
        const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
        const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");

        let query = supabase.from("activities").update(payload).eq("id", actId);
        // 일반 관리자나 조장이 아니면 본인 활동만 수정 가능하도록 보안 잠금
        if (!isSuper && !isLeader) {
          query = query.eq("user_id", verifiedTgUserId);
        }
        const { error } = await query;
        if (error) {
          return new Response(
            JSON.stringify({ result: "fail", message: "활동 수정 실패: " + error.message }),
            { headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
      } else {
        // 신규 추가 시에는 필수 Primary Key인 id 지정 및 유저 식별자 바인딩
        const { error } = await supabase.from("activities").insert([{
          ...payload,
          id: actId || ("ACT_" + Date.now()),
          user_id: verifiedTgUserId,
          name: dbUser.name,
          group_name: dbUser.group_name || "",
          region: dbUser.region || "",
          role: dbUser.role || ""
        }]);
        if (error) {
          return new Response(
            JSON.stringify({ result: "fail", message: "활동 추가 실패: " + error.message }),
            { headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
      }

      if (status === "completed" && !isCanceledActivity(payload)) {
        const now = new Date();
        const kstTime = new Date(now.getTime() + (9 * 60 * 60 * 1000));
        const todayKST = kstTime.toISOString().split("T")[0];

        if (date === todayKST) {
          const targetChatId = Deno.env.get("TELEGRAM_CHAT_ID") || "-1003736767935";
          sendRealtimeActivityNotification(supabase, date, dbUser, payload, botToken, targetChatId);
        }
      }

      return new Response(
        JSON.stringify({ result: "success", message: "활동이 성공적으로 저장되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 2. 사용자 정보 변경 (updateUserStatus: 복방 개수 / 면제 상태)
    if (action === "updateUserStatus") {
      const { targetName, bookCount, isExempt } = params;

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");
      if (!isSuper && !isLeader) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      // 대상 사용자 조회
      const { data: targets } = await supabase.from("users").select("*").eq("name", targetName);
      if (!targets || targets.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "대상 사용자를 찾을 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const targetUser = targets[0];

      // 조장/부조장은 본인 지역 소속의 유저만 관리 가능
      if (!isSuper && targetUser.region !== dbUser.region) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const updatePayload: any = {};
      if (bookCount !== undefined) {
        updatePayload.book_count = parseInt(bookCount, 10);
      }
      if (isExempt !== undefined) {
        updatePayload.is_exempt = isExempt === true || isExempt === "true";
      }

      const { error } = await supabase.from("users").update(updatePayload).eq("id", targetUser.id);
      if (error) {
        return new Response(
          JSON.stringify({ result: "fail", message: "상태 변경 실패: " + error.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "정보가 성공적으로 변경되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 3. 사용자 권한 수정 (updateUserRole)
    if (action === "updateUserRole") {
      const { targetName, role } = params;

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");
      if (!isSuper && !isLeader) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { data: targets } = await supabase.from("users").select("*").eq("name", targetName);
      if (!targets || targets.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "대상 사용자를 찾을 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const targetUser = targets[0];

      if (!isSuper && targetUser.region !== dbUser.region) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { error } = await supabase.from("users").update({ role }).eq("id", targetUser.id);
      if (error) {
        return new Response(
          JSON.stringify({ result: "fail", message: "권한 변경 실패: " + error.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "권한이 성공적으로 변경되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 4. 개강 사이클 관리 (saveSemester)
    if (action === "saveSemester") {
      const { mode, id, name, startDate, endDate, isActive } = params;

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      if (!isSuper) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      if (mode === "delete") {
        const { error } = await supabase.from("semesters").delete().eq("id", id);
        if (error) {
          return new Response(
            JSON.stringify({ result: "fail", message: "사이클 삭제 실패: " + error.message }),
            { headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
        return new Response(
          JSON.stringify({ result: "success", message: "사이클이 삭제되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      if (mode === "add") {
        await supabase.from("semesters").update({ is_active: false }).neq("id", 0);

        const { error } = await supabase.from("semesters").insert([{
          name,
          start_date: startDate,
          end_date: endDate,
          is_active: true
        }]);
        if (error) {
          return new Response(
            JSON.stringify({ result: "fail", message: "사이클 추가 실패: " + error.message }),
            { headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
        return new Response(
          JSON.stringify({ result: "success", message: "새 사이클이 등록되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      if (mode === "edit") {
        if (isActive === true || isActive === "true") {
          await supabase.from("semesters").update({ is_active: false }).neq("id", id);
        }

        const { error } = await supabase.from("semesters").update({
          name,
          start_date: startDate,
          end_date: endDate,
          is_active: isActive === true || isActive === "true"
        }).eq("id", id);
        if (error) {
          return new Response(
            JSON.stringify({ result: "fail", message: "사이클 수정 실패: " + error.message }),
            { headers: { ...corsHeaders, "Content-Type": "application/json" } }
          );
        }
        return new Response(
          JSON.stringify({ result: "success", message: "사이클이 수정되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
    }

    // [보안 CUD 위임 API] 5. 공지사항 등록 (saveNotice)
    if (action === "saveNotice") {
      const { title, content, type, region, visible, isImportant } = params;

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");
      if (!isSuper && !isLeader) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      let finalRegion = region || "ALL";
      if (!isSuper) {
        finalRegion = dbUser.region;
      }

      const { error } = await supabase.from("notices").insert([{
        type,
        author_id: verifiedTgUserId,
        author_name: dbUser.name,
        author_role: dbUser.role,
        group_name: dbUser.group_name || "",
        title,
        content,
        region: finalRegion,
        visible: visible === true || visible === "Y" || visible === "true",
        is_important: isImportant === true || isImportant === "Y" || isImportant === "true"
      }]);

      if (error) {
        return new Response(
          JSON.stringify({ result: "fail", message: "공지사항 저장 실패: " + error.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "공지사항이 등록되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 6. 공지사항 상태 수정 (toggleNoticeVisibility)
    if (action === "toggleNoticeVisibility") {
      const { noticeId, visible } = params;

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const { data: notices } = await supabase.from("notices").select("*").eq("id", noticeId);
      if (!notices || notices.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "공지사항을 찾을 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const notice = notices[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");
      if (!isSuper && (!isLeader || notice.region !== dbUser.region)) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { error } = await supabase.from("notices").update({
        visible: visible === true || visible === "Y" || visible === "true"
      }).eq("id", noticeId);

      if (error) {
        return new Response(
          JSON.stringify({ result: "fail", message: "공지 상태 변경 실패: " + error.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "공지 상태가 수정되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 7. 공지사항 삭제 (deleteNotice)
    if (action === "deleteNotice") {
      const { noticeId } = params;

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const { data: notices } = await supabase.from("notices").select("*").eq("id", noticeId);
      if (!notices || notices.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "공지사항을 찾을 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const notice = notices[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");
      if (!isSuper && (!isLeader || notice.region !== dbUser.region)) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { error } = await supabase.from("notices").delete().eq("id", noticeId);

      if (error) {
        return new Response(
          JSON.stringify({ result: "fail", message: "공지사항 삭제 실패: " + error.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "공지사항이 삭제되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 8. 1:1 문의/건의 접수 (submitInquiry)
    if (action === "submitInquiry") {
      const { category, title, content } = params;

      if (!category || !title || !content) {
        return new Response(
          JSON.stringify({ result: "fail", message: "필수 입력 필드가 누락되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      // 작성자 텔레그램 프로필 정보 조회
      const { data: dbUsers, error: userErr } = await supabase
        .from("users")
        .select("*")
        .eq("id", verifiedTgUserId);

      if (userErr || !dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "등록되지 않은 사용자입니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      // 문의 테이블에 삽입
      const { error: insertErr } = await supabase
        .from("inquiries")
        .insert([{
          user_id: verifiedTgUserId,
          name: dbUser.name,
          region: dbUser.region || "미정",
          role: dbUser.role || "조원",
          category,
          title,
          content,
          status: "pending"
        }]);

      if (insertErr) {
        return new Response(
          JSON.stringify({ result: "fail", message: "문의 접수 실패: " + insertErr.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "문의사항이 성공적으로 접수되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 9. 문의/건의 답변 등록 및 수정 (replyInquiry)
    if (action === "replyInquiry") {
      const { inquiryId, replyText } = params;

      if (!inquiryId || replyText === undefined) {
        return new Response(
          JSON.stringify({ result: "fail", message: "필수 파라미터가 누락되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      // 답변자 권한 확인
      const { data: dbUsers, error: userErr } = await supabase
        .from("users")
        .select("*")
        .eq("id", verifiedTgUserId);

      if (userErr || !dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "작성자를 확인할 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");

      if (!isSuper && !isLeader) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      // 해당 문의 건 확인
      const { data: inquiries, error: inquiryErr } = await supabase
        .from("inquiries")
        .select("*")
        .eq("id", inquiryId);

      if (inquiryErr || !inquiries || inquiries.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "문의를 찾을 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const inquiry = inquiries[0];

      // 조장의 경우 동일 지역 소속 사용자의 문의인지 검증
      if (!isSuper && inquiry.region !== dbUser.region) {
        return new Response(
          JSON.stringify({ result: "fail", message: "타 지역 사용자의 문의에 답할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      // 답변 업데이트
      const updateData = {
        reply: replyText,
        replied_by: dbUser.name,
        replied_at: new Date().toISOString(),
        status: replyText.trim() === "" ? "pending" : "replied"
      };

      const { error: updateErr } = await supabase
        .from("inquiries")
        .update(updateData)
        .eq("id", inquiryId);

      if (updateErr) {
        return new Response(
          JSON.stringify({ result: "fail", message: "답변 등록 실패: " + updateErr.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "답변이 성공적으로 등록되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 10. 거점 등록 (saveTodayBase)
    if (action === "saveTodayBase") {
      const { place, baseTime, baseDate } = params;

      if (!place || !baseTime || !baseDate) {
        return new Response(
          JSON.stringify({ result: "fail", message: "필수 파라미터가 누락되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { data: dbUsers, error: userErr } = await supabase
        .from("users")
        .select("*")
        .eq("id", verifiedTgUserId);

      if (userErr || !dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "등록되지 않은 사용자입니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const { error: insertErr } = await supabase
        .from("today_bases")
        .insert([{
          region: dbUser.region || "미정",
          creator_id: verifiedTgUserId,
          creator_name: dbUser.name,
          place,
          base_time: baseTime,
          base_date: baseDate,
          is_disbanded: false
        }]);

      if (insertErr) {
        return new Response(
          JSON.stringify({ result: "fail", message: "거점 등록 실패: " + insertErr.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "거점이 성공적으로 등록되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 11. 거점 해산 및 활성화 (disbandTodayBase)
    if (action === "disbandTodayBase") {
      const { baseId, isDisbanded } = params;

      if (!baseId || isDisbanded === undefined) {
        return new Response(
          JSON.stringify({ result: "fail", message: "필수 파라미터가 누락되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const { data: bases, error: baseErr } = await supabase.from("today_bases").select("*").eq("id", baseId);
      if (baseErr || !bases || bases.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "거점을 찾을 수 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const base = bases[0];

      const isSuper = dbUser.role.includes("관리자") || dbUser.role.includes("전체관리자");
      const isLeader = dbUser.role.includes("조장") || dbUser.role.includes("부조장");
      const isCreator = base.creator_id === verifiedTgUserId;

      if (!isSuper && !isCreator && (!isLeader || base.region !== dbUser.region)) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const targetDisbanded = isDisbanded === true || isDisbanded === "true";

      const { error: updateErr } = await supabase
        .from("today_bases")
        .update({ is_disbanded: targetDisbanded })
        .eq("id", baseId);

      if (updateErr) {
        return new Response(
          JSON.stringify({ result: "fail", message: "거점 상태 변경 실패: " + updateErr.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: targetDisbanded ? "거점이 해산되었습니다." : "거점이 다시 활성화되었습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 12. 거점 도착 (joinTodayBase)
    if (action === "joinTodayBase") {
      const { baseId } = params;

      if (!baseId) {
        return new Response(
          JSON.stringify({ result: "fail", message: "필수 파라미터가 누락되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { data: dbUsers } = await supabase.from("users").select("*").eq("id", verifiedTgUserId);
      if (!dbUsers || dbUsers.length === 0) {
        return new Response(
          JSON.stringify({ result: "fail", message: "해당 작업을 수행할 권한이 없습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const dbUser = dbUsers[0];

      const { error: insertErr } = await supabase
        .from("base_members")
        .upsert([{
          base_id: baseId,
          user_id: verifiedTgUserId,
          user_name: dbUser.name
        }], { onConflict: "base_id,user_id" });

      if (insertErr) {
        return new Response(
          JSON.stringify({ result: "fail", message: "거점 도착 기록 실패: " + insertErr.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "거점에 도착했습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // [보안 CUD 위임 API] 13. 거점 떠나기 (leaveTodayBase)
    if (action === "leaveTodayBase") {
      const { baseId } = params;

      if (!baseId) {
        return new Response(
          JSON.stringify({ result: "fail", message: "필수 파라미터가 누락되었습니다." }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const { error: deleteErr } = await supabase
        .from("base_members")
        .delete()
        .eq("base_id", baseId)
        .eq("user_id", verifiedTgUserId);

      if (deleteErr) {
        return new Response(
          JSON.stringify({ result: "fail", message: "거점 떠나기 기록 실패: " + deleteErr.message }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({ result: "success", message: "거점을 떠났습니다." }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    return new Response(
      JSON.stringify({ result: "fail", message: "지원하지 않는 action입니다." }),
      { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );

  } catch (err: any) {
    return new Response(
      JSON.stringify({ result: "fail", message: err.message }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
