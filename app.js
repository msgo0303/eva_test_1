// 텔레그램 WebApp 초기화
const tg = window.Telegram.WebApp;
tg.ready();
tg.expand();

// ⚠️ 본인의 구글 앱스 스크립트 웹앱 URL (/exec로 끝나는 주소)
const GAS_URL = "https://script.google.com/macros/s/AKfycbwOnzceBFUvKYlqB0r-MBTYWC3jfv6ebqhuy2UV9alE5S83isURYGSbG4de0pFWsGMM/exec";

async function checkUser() {
    const initData = tg.initData;
    const loadingEl = document.getElementById('loading-screen');
    const registerEl = document.getElementById('register-screen');

    // 텔레그램 환경이 아닐 때
    if (!initData) {
        loadingEl.innerHTML = "<h3>텔레그램 모바일 앱에서 접속해주세요.</h3>";
        return;
    }

    try {
        // 구글 서버에 사용자 가입 여부 확인
        const res = await fetch(`${GAS_URL}?action=checkUser&initData=${encodeURIComponent(initData)}`);
        const data = await res.json();

        loadingEl.style.display = 'none';

        if (data.isRegistered) {
            enterMainScreen(data.userName);
        } else {
            registerEl.style.display = 'block';
            document.getElementById('register-btn').onclick = registerUser;
        }
    } catch (err) {
        // 통신 오류 시 일단 등록 화면을 띄워주는 안전장치
        loadingEl.style.display = 'none';
        registerEl.style.display = 'block';
        document.getElementById('register-btn').onclick = registerUser;
    }
}

async function registerUser() {
    const name = document.getElementById('user-name-input').value.trim();
    if (!name) return alert("이름을 입력해주세요.");

    document.getElementById('register-screen').style.display = 'none';
    document.getElementById('loading-screen').style.display = 'block';
    document.getElementById('loading-screen').innerText = "등록 처리 중...";

    try {
        // GET 방식을 활용해 CORS 문제 없이 안전하게 데이터를 전송
        const registerUrl = `${GAS_URL}?action=registerUser&initData=${encodeURIComponent(tg.initData)}&userName=${encodeURIComponent(name)}`;

        await fetch(registerUrl, { mode: 'no-cors' });

        setTimeout(() => {
            document.getElementById('loading-screen').style.display = 'none';
            enterMainScreen(name);
        }, 1500);

    } catch (err) {
        alert("등록 실패: " + err.message);
        document.getElementById('loading-screen').style.display = 'none';
        document.getElementById('register-screen').style.display = 'block';
    }
}

function enterMainScreen(userName) {
    document.getElementById('main-screen').style.display = 'block';
    document.getElementById('welcome-msg').innerText = `안녕하세요, ${userName}님!`;
}

// 앱 실행 시작
checkUser();