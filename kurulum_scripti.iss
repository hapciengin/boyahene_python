; Inno Setup Script - {src} Değişkeni ile Garantili Çözüm

#define MyAppName "Boya Barkod Uygulamasi"
#define MyAppVersion "1.1" // Versiyonu güncelledik
#define MyAppExeName "BoyaBarkodYonetimiFinal.exe"
#define MyAppDistFolder "BoyaBarkodYonetimiFinal" // PyInstaller'ın --name ile oluşturduğu klasörün adı

[Setup]
AppId={{YENI-UNIQUE-APP-ID-1234-ABCD-5678}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=yes
OutputBaseFilename=BoyaBarkodUygulamasi_Kurulum
OutputDir=.\install_output_final
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
; Hata almamak için bu satırı devre dışı bırakıyoruz.
; Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";

[Files]
; --- BURASI EN ÖNEMLİ DEĞİŞİKLİK ---
; Kaynak olarak, script'in kendi klasörünü referans alıyoruz ({src}).
; Bu, yolun yanlış olma ihtimalini ortadan kaldırır.
; Anlamı: "Bu script'in olduğu klasörün içindeki 'dist' klasörünün içindeki
; '{#MyAppDistFolder}' klasörünün içindeki her şeyi al."
Source: "{src}\dist\{#MyAppDistFolder}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent