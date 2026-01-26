#define Publisher "BasicNullification"
#define AssemblyName "Archiver"
#define AppId "{{0250B316-6712-4D7B-BBD3-33114DCA208F}"
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppVersion3 Copy(AppVersion, 1, RPos(".", AppVersion)-1)

[Setup]
AppId={#AppId}
AppName={#AssemblyName}
AppVersion={#AppVersion3}
AppPublisher={#Publisher}

OutputDir=..\..\dist\installer
OutputBaseFilename={#AssemblyName}-Setup-v{#AppVersion3}

DefaultDirName={autopf}\{#Publisher}\{#AssemblyName}

DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=admin

ArchitecturesAllowed=x86 x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

UninstallDisplayName={#AssemblyName} COM Library
Compression=lzma2
SolidCompression=yes
ChangesEnvironment=no
UsePreviousAppDir=yes

[Files]
; x86 build -> Program Files (x86) on 64-bit, Program Files on 32-bit
Source: "..\bin\x86\Release\{#AssemblyName}.dll"; DestDir: "{autopf32}\{#Publisher}\{#AssemblyName}"; Flags: ignoreversion overwritereadonly restartreplace uninsrestartdelete

; x64 build -> only on 64-bit OS
Source: "..\bin\x64\Release\{#AssemblyName}.dll"; DestDir: "{autopf}\{#Publisher}\{#AssemblyName}"; Flags: ignoreversion overwritereadonly restartreplace uninsrestartdelete; Check: IsWin64

[Code]
function RegAsm32Path(): string;
begin
  Result := ExpandConstant('{win}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe');
end;

function RegAsm64Path(): string;
begin
  Result := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe');
end;

function DllPathX86(): string;
begin
  Result := ExpandConstant('{autopf32}\{#Publisher}\{#AssemblyName}\{#AssemblyName}.dll');
end;

function DllPathX64(): string;
begin
  Result := ExpandConstant('{autopf}\{#Publisher}\{#AssemblyName}\{#AssemblyName}.dll');
end;

procedure ExecOrLog(const Exe, Params: string);
var
  ResultCode: Integer;
begin
  if FileExists(Exe) then begin
    Exec(Exe, Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Log(Format('Exec: %s %s -> rc=%d', [Exe, Params, ResultCode]));
  end else begin
    Log(Format('Missing: %s (skipping)', [Exe]));
  end;
end;

procedure UnregisterOldIfPresent();
begin
  // Unregister x86 if it exists
  if FileExists(DllPathX86()) then
    ExecOrLog(RegAsm32Path(), '/nologo /silent /unregister "' + DllPathX86() + '"');

  // Unregister x64 if it exists (64-bit only)
  if IsWin64 and FileExists(DllPathX64()) then
    ExecOrLog(RegAsm64Path(), '/nologo /silent /unregister "' + DllPathX64() + '"');
end;

procedure RegisterNew();
begin
  // Register x86
  if FileExists(DllPathX86()) then
    ExecOrLog(RegAsm32Path(), '/nologo /silent /codebase /tlb "' + DllPathX86() + '"');

  // Register x64 (64-bit only)
  if IsWin64 and FileExists(DllPathX64()) then
    ExecOrLog(RegAsm64Path(), '/nologo /silent /codebase /tlb "' + DllPathX64() + '"');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then begin
    // Upgrade-safe: unregister previous registrations before files get replaced
    UnregisterOldIfPresent();
  end;

  if CurStep = ssPostInstall then begin
    // After files are in place, register the new ones
    RegisterNew();
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then begin
    // Uninstall: unregister both
    UnregisterOldIfPresent();
  end;
end;
