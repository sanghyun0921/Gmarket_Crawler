#define MyAppName "Gmarket Crawler"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Gmarket Crawler"
#define MyAppExeName "GmarketCrawler.exe"

[Setup]
AppId={{7DD1BC89-10A6-49F7-B94A-6900DCC263D4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\GmarketCrawler
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=GmarketCrawlerSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 만들기"; GroupDescription: "추가 아이콘:"; Flags: unchecked

[Files]
Source: "{#SourcePath}\dist\GmarketCrawler\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Gmarket 자동 수집 설정 실행"; Flags: nowait postinstall skipifsilent
