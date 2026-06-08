# Gmarket 자동 수집 배포 안내

## 사용자에게 배포할 파일

빌드 후 아래 폴더 전체를 직원 PC에 전달합니다.

```text
dist\GmarketCrawler
```

직원은 이 폴더 안의 `GmarketCrawler.exe`만 실행하면 됩니다.

```

## 빌드 방법

배포 담당자 PC에서 한 번 실행합니다.

```powershell
.\build_exe.ps1
```
빌드 결과는 아래 위치에 만들어집니다.

```text
dist\GmarketCrawler\GmarketCrawler.exe
```

설치 파일까지 만들려면 `build_exe.ps1` 실행 후 Inno Setup Compiler에서 `installer.iss`를 엽니다.

## 직원 PC 사용 방법

1. 엑셀파일 판매자, 상품명열에 자신이 크롤링 하고싶은 제품의 판매자, 상품명을 작성합니다. 
2. Chrome이 설치되어 있는지 확인합니다.
3. `GmarketCrawlerSetup.exe`를 실행하여 `GmarketCrawler.exe`를 다운받습니다.
4. `GmarketCrawler.exe`를 실행합니다.
5. `찾기` 버튼으로 사용할 엑셀 파일을 선택합니다.
6. 매일 실행할 시간을 선택합니다.
7. `저장하고 예약 등록`을 누릅니다.
8. 바로 테스트하려면 `지금 실행`을 누릅니다.

## 자동 실행 조건

- Windows 작업 스케줄러에 `GmarketDailyCrawler` 작업이 등록됩니다.
- 사용자가 Windows에 로그인되어 있을 때 실행되는 것을 기준으로 합니다.
- 실행 시 Chrome 브라우저가 열릴 수 있습니다.
- ChromeDriver는 먼저 자동 감지를 시도하고, 실패하면 설치된 Chrome 버전을 읽어서 다시 시도합니다.
- 엑셀 파일이 열려 있으면 저장에 실패할 수 있으므로, 예약 시간에는 엑셀을 닫아 두는 것이 좋습니다.
- 일반모델의 경우 직배수모델과 중복될 수 있으니 유의


## 설정 저장 위치

선택한 엑셀 파일과 실행 시간은 아래 파일에 저장됩니다.

```text
%APPDATA%\GmarketCrawler\config.json
```
