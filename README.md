# KRAFTON BOD PPT Translator

KRAFTON 이사회 PPT 자동 번역 웹앱

## 파일 구조
```
app.py            ← 메인 앱 코드
requirements.txt  ← 필요 패키지
README.md         ← 이 파일
```

## 배포 방법 (Streamlit Cloud)

### 1단계: GitHub 업로드
1. https://github.com 에서 무료 계정 만들기
2. New Repository 클릭 → 이름: `krafton-ppt-translator` → Public or Private
3. 이 3개 파일 업로드: `app.py`, `requirements.txt`, `README.md`

### 2단계: Streamlit Cloud 배포
1. https://share.streamlit.io 접속 (GitHub 계정으로 로그인)
2. "New app" 클릭
3. Repository: `krafton-ppt-translator` 선택
4. Main file path: `app.py`
5. Deploy 클릭 → 자동으로 앱 빌드 (2~3분)
6. 주소 생성됨: `https://xxx.streamlit.app`

### 3단계: 사내 접근 제한 (선택)
- Streamlit Cloud → Settings → Sharing
- "Only specific people can view this app" 선택
- `@krafton.com` 이메일만 허용 설정

## 사용 방법
1. 사이드바에서 Claude API Key 입력
2. 번역 언어 선택 (English / Japanese / Chinese)
3. PPT 파일 업로드
4. "번역 시작" 클릭
5. 완료 후 번역된 파일 다운로드
