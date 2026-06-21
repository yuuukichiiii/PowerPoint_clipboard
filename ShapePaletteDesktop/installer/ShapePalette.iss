; Shape Palette — per-user インストーラー（管理者不要）
; ビルド: "C:\Users\yukit\AppData\Local\Programs\Inno Setup 6\ISCC.exe" ShapePalette.iss

#define AppName "Shape Palette"
#define AppVersion "1.0.0"
#define Publisher "Yuki"
#define ConnectClsid "{{352F4B34-C24C-42FB-AF36-FE303D48BB57}"
#define HostClsid "{{7F4C3A87-FB96-43D8-8B52-B9D92EF2A81B}"
#define AsmName "ShapePalette, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
#define Runtime "v4.0.30319"

[Setup]
AppId={{AE373F0A-7F0B-4825-95DF-3AA56DF42318}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
DefaultDirName={localappdata}\ShapePalette
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no
Uninstallable=yes
OutputDir=Output
OutputBaseFilename=ShapePaletteSetup
SolidCompression=yes
Compression=lzma2
WizardStyle=modern

[Languages]
Name: "ja"; MessagesFile: "compiler:Languages\Japanese.isl"

[Files]
Source: "..\ShapePalette\bin\Release\ShapePalette.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\ShapePalette\bin\Release\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; ---- COM: Connect ----
Root: HKCU; Subkey: "Software\Classes\ShapePalette.Connect"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\ShapePalette.Connect\CLSID"; ValueType: string; ValueData: "{#ConnectClsid}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}"; ValueType: string; ValueData: "ShapePalette.Connect"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\ProgId"; ValueType: string; ValueData: "ShapePalette.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ShapePalette.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "{#Runtime}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBaseUri}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Class"; ValueData: "ShapePalette.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "{#Runtime}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ConnectClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBaseUri}"

; ---- COM: PaletteHost ----
Root: HKCU; Subkey: "Software\Classes\ShapePalette.PaletteHost"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\ShapePalette.PaletteHost\CLSID"; ValueType: string; ValueData: "{#HostClsid}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}"; ValueType: string; ValueData: "ShapePalette.UI.PaletteHost"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\ProgId"; ValueType: string; ValueData: "ShapePalette.PaletteHost"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ShapePalette.UI.PaletteHost"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "{#Runtime}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBaseUri}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Class"; ValueData: "ShapePalette.UI.PaletteHost"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmName}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "{#Runtime}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#HostClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBaseUri}"

; ---- PowerPoint アドイン登録 ----
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Shape Palette"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect"; ValueType: string; ValueName: "Description"; ValueData: "図形をタブ別に並べ、クリックでスライドに挿入するパレット"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Code]
function CodeBaseUri(Param: String): String;
var
  s: String;
begin
  s := ExpandConstant('{app}\ShapePalette.dll');
  StringChangeEx(s, '\', '/', True);
  Result := 'file:///' + s;
end;
