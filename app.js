// 텔레그램 WebApp 객체 초기화
const tg = window.Telegram.WebApp;
tg.ready(); // 미ni앱이 준비되었음을 텔레그램에 알림
tg.expand(); // 화면을 최대 크기로 확장

const GAS_URL = "https://script.google.com/macros/s/AKfycbwOnzceBFUvKYlqB0r-MBTYWC3jfv6ebqhuy2UV9alE5S83isURYGSbG4de0pFWsGMM/exec";

// 미니앱이 켜지자마자 실행되는 함수
async function checkUser() {
    // 텔레그램이 제공하는 유저 정보 문자열 (보안 검증용)
    const initData = tg.initData; 
    
    // 만약 PC 브라우저 등 텔레그램 외 환경에서 접속했다면 에러 처리
    if (!initData) {
        document.getElementById('loading-screen').innerText = "텔레그램 모바일 앱에서 접속해주세요.";
        return;
    }

    try {
        // 백엔드(GAS)에 유저 상태 조회 요청
        const response = await fetch(`${GAS_URL}?action=checkUser&initData=${encodeURIComponent(initData)}`);
        const result = await response.json();

        document.getElementById('loading-screen').style.display = 'none';

        if (result.isRegistered) {
            // Case A: 이미 등록된 유저 -> 바로 메인 화면으로 진입 (자동 로그인 완료)
            enterMainScreen(result.userName);
        } else {
            // Case B: 처음 온 유저 -> 실명 입력 폼 표시
            document.getElementById('register-screen').style.display = 'block';
            
            // 등록 버튼 클릭 이벤트 설정
            document.getElementById('register-btn').onclick = () => registerUser(initData);
        }
    } catch (error) {
        console.error(error);
        document.getElementById('loading-screen').innerText = "서버 연결 오류가 발생했습니다.";
    }
}

// 신규 유저 등록 함수
async function registerUser(initData) {
    const name = document.getElementById('user-name-input').value.trim();
    if (!name) return alert("이름을 입력해주세요.");

    document.getElementById('register-screen').style.display = 'none';
    document.getElementById('loading-screen').style.display = 'block';
    document.getElementById('loading-screen').innerText = "등록 중...";

    try {
        const response = await fetch(GAS_URL, {
            method: 'POST',
            mode: 'no-cors', // GAS의 CORS 특성상 응답을 받지 않고 보낼 때 주로 사용
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                action: 'registerUser',
                initData: initData,
                userName: name
            })
        });
        
        // no-cors 모드는 응답 값을 읽을 수 없으므로 잠시 후 성공으로 가정하고 화면 전환
        setTimeout(() => {
            document.getElementById('loading-screen').style.display = 'none';
            enterMainScreen(name);
        }, 1500);

    } catch (error) {
        alert("등록 실패. 다시 시도해주세요.");
    }
}

function enterMainScreen(userName) {
    document.getElementById('main-screen').style.display = 'block';
    document.getElementById('welcome-msg').innerText = `안녕하세요, ${userName}님!`;
}

// 실행
checkUser();