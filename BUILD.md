# 개발자 가이드 (BUILD)

이 문서는 **코드를 수정하고 exe를 배포하는 사람(개발 담당자)**용입니다.
파트너에게 전달할 설명서는 `README.md` 입니다.

---

## 파일 구성

```
rehab_crawler/
├── app.py                      # GUI (PySide6) — 실행 진입점
├── crawler_v10.py              # 크롤링 엔진 (로직 전부)
├── requirements.txt            # 의존성
├── .env_verify.example         # 인증 키 예시 (실제 .env_verify는 커밋 금지)
├── .gitignore
├── README.md                   # 파트너용 설명서
├── BUILD.md                    # 이 문서
└── .github/workflows/build.yml # 자동 exe 빌드
```

`app.py` 는 `crawler_v10.py` 의 함수를 import 해서 씁니다. 로직 수정은 `crawler_v10.py` 에서,
화면 수정은 `app.py` 에서 하면 됩니다.

---

## 로컬에서 실행 (개발 중 테스트)

```bash
pip install -r requirements.txt
python app.py
```

---

## exe 배포 방법 (권장: GitHub Actions 자동 빌드)

로컬에서 직접 빌드할 필요 없이, 깃허브가 클라우드에서 윈도우 exe를 만들어 줍니다.

### 최초 1회 세팅

1. GitHub에 새 저장소(repo)를 만들고 이 폴더 전체를 올립니다.
   (웹 업로드 또는 git push)

2. **인증 키를 GitHub Secrets에 저장** (exe에 키를 안전하게 내장하기 위함)
   - 저장소 **Settings → Secrets and variables → Actions**
   - **New repository secret** 클릭
   - Name: `VERIFY_API_KEY`
   - Secret: 2captcha에서 발급받은 실제 API 키
   - **Add secret**

   > 이렇게 하면 빌드할 때만 키가 `_key.py`로 생성되어 exe에 포함됩니다.
   > 저장소 코드에는 키가 남지 않고, GitHub에도 암호화되어 저장됩니다.
   > 파트너는 프로그램 실행 시 **키 입력 없이 바로 사용**할 수 있습니다.
   > 키를 바꾸려면 이 Secret 값만 수정한 뒤 새 태그로 다시 빌드하면 됩니다.

### exe 만들기 (배포할 때마다)

**방법 A — 태그 push (릴리스 자동 생성, 권장)**
```bash
git tag v1.0
git push origin v1.0
```
→ 몇 분 뒤 저장소 **Releases** 에 `회생사건조회.exe` 가 자동으로 올라갑니다.
   파트너에게는 이 Releases 링크만 보내면 됩니다.

**방법 B — 수동 실행**
→ 저장소 **Actions** 탭 → "Build Windows EXE" → **Run workflow** 클릭
→ 완료 후 Artifacts 에서 exe 다운로드 (릴리스는 안 만들어짐)

### 로직을 수정한 뒤 새 버전 배포
```bash
# 코드 수정 후
git add .
git commit -m "수정 내용"
git push
git tag v1.1        # 버전 올리고
git push origin v1.1
```

---

## exe 로컬 빌드 (윈도우 PC가 있을 때, 선택)

굳이 로컬에서 만들고 싶으면 **윈도우 PC에서**:
```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name "회생사건조회" app.py
```
→ `dist/회생사건조회.exe` 생성.
(리눅스·맥에서는 윈도우 exe가 만들어지지 않습니다. 반드시 윈도우에서.)

---

## 주의사항

- **`.env_verify` (실제 키가 든 파일)는 절대 커밋하지 마세요.** `.gitignore` 에 이미 등록돼 있습니다.
- 파트너 배포용 exe에는 키를 넣지 말고, 파트너가 화면에서 직접 입력하게 하세요.
- 크롬 브라우저가 설치된 PC에서만 동작합니다. (webdriver-manager 가 드라이버 자동 설치)
