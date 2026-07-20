const tg = window.Telegram.WebApp;
tg.ready();
tg.expand();

// ⚠️ 본인의 구글 웹앱 URL (/exec로 끝나는 것)을 적어주세요.
const GAS_URL = "https://script.google.com/macros/s/AKfycbwOnzceBFUvKYlqB0r-MBTYWC3jfv6ebqhuy2UV9alE5S83isURYGSbG4de0pFWsGMM/exec";

async function checkUser() {
    const initData = tg.initData; 
    
    // 환경 체크: 텔레그램 내부가 아니라면 경고 메시지 출력
    if (!initData) {
        document.getElementById('loading-screen').innerHTML = "<h3>텔레그램 모바일 앱에서 미니앱을 열어주세요!</h3>";
        return;
    }

    try {
        // 구글 웹앱 리다이렉션 이슈를 피하기 위해 JSONP 또는 폼 전송 방식과 유사하게 fetch 옵션 조정
        const response = await fetch(`${GAS_URL}?action=checkUser&initData=${encodeURIComponent(initData)}`, {
            method: 'GET',
            headers: {
                'Accept': 'application/json'
            }
        });
        
        const result = await response.json();
        document.getElementById('loading-screen').style.display = 'none';

        if (result.isRegistered) {
            enterMainScreen(result.userName);
        } else {
            document.getElementById('register-screen').style.display = 'block';
            document.getElementById('register-btn').onclick = registerUser;
        }
    } catch (error) {
        console.error(error);
        // 에러가 나더라도 사용자가 진행할 수 있도록 신규 등록 화면을 열어주는 안전장치(Fallback)
        document.getElementById('loading-screen').style.display = 'none';
        document.getElementById('register-screen').style.display = 'block';
        document.getElementById('register-btn').onclick = registerUser;
    }
}

async function registerUser() {
    const name = document.getElementById('user-name-input').value.trim();
    if (!name) return alert("이름을 입력해주세요.");

    document.getElementById('register-screen').style.display = 'none';
    document.getElementById('loading-screen').style.display = 'block';
    document.getElementById('loading-screen').innerText = "등록 중...";

    try {
        // GAS로의 POST 요청 안정성 확보
        await fetch(GAS_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: new URLSearchParams({
                action: 'registerUser',
                initData: tg.initData,
                userName: name
            })
        });
        
        // 안전하게 화면 전환
        setTimeout(() => {
            document.getElementById('loading-screen').style.display = 'none';
            enterMainScreen(name);
        }, 1000);

    } catch (error) {
        alert("등록 중 오류가 발생했습니다. 다시 시도해주세요.");
        document.getElementById('loading-screen').style.display = 'none';
        document.getElementById('register-screen').style.display = 'block';
    }
}

function enterMainScreen(userName) {
    document.getElementById('main-screen').style.display = 'block';
    document.getElementById('welcome-msg').innerText = `안녕하세요, ${userName}님!`;
}

// 앱 실행
checkUser();