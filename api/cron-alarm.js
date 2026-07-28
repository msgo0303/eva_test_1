// api/cron-alarm.js
import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY;
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

export default async function handler(req, res) {
  try {
    // 한국 표준시 (KST) 기준 현재 시간 구하기 (HH:mm)
    const now = new Date();
    const kstOffset = 9 * 60 * 60 * 1000;
    const kstDate = new Date(now.getTime() + kstOffset);
    const currentHHmm = kstDate.toISOString().substring(11, 16); // '21:00'
    const todayStr = kstDate.toISOString().substring(0, 10);      // '2026-07-28'

    // 1. 활성화(is_active = true) 및 현재 시각과 일치하는 알람 조회
    const { data: alarms, error: alarmErr } = await supabase
      .from('alarm_schedules')
      .select('*')
      .eq('is_active', true)
      .eq('alarm_time', currentHHmm);

    if (alarmErr) throw alarmErr;
    if (!alarms || alarms.length === 0) {
      return res.status(200).json({ message: "실행할 알람이 없습니다.", time: currentHHmm });
    }

    // 2. 유저 전체 명단 및 당일 활동 조회
    const { data: allUsers } = await supabase.from('users').select('*');
    const { data: todayActs } = await supabase.from('activities').select('*').eq('activity_date', todayStr);

    for (const alarm of alarms) {
      let uncollectedList = [];
      let titleText = "";

      if (alarm.category === 'today') {
        titleText = "[⚠️ 오늘 결과 미입력자 알림]\n\n오늘 활동 결과를 아직 입력하지 않은 멤버입니다. 마감 전 입력 바랍니다!";
        uncollectedList = (allUsers || []).filter(u => {
          if (u.is_exempt) return false;
          const act = (todayActs || []).find(a => a.name === u.name);
          return !act || act.status !== 'completed';
        });
      } else {
        titleText = "[📢 내일 일정 미입력자 알림]\n\n내일 일정을 아직 등록하지 않은 멤버입니다. 확인 후 작성 바랍니다!";
        // 내일 날짜 계산 후 내일 일정 조회 로직
      }

      if (uncollectedList.length > 0) {
        // 지역별 묶기
        const grouped = {};
        uncollectedList.forEach(u => {
          const reg = u.region || '미정';
          if (!grouped[reg]) grouped[reg] = [];
          grouped[reg].push(u.name);
        });

        const lines = Object.keys(grouped).map(reg => `📍 [${reg}]: ${grouped[reg].join(', ')}`);
        const finalMsg = `${titleText}\n\n${lines.join('\n')}`;

        // 텔레그램 General 방 발송
        await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: finalMsg })
        });
      }
    }

    return res.status(200).json({ success: true, executedAlarms: alarms.length });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}