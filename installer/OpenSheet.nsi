Unicode True

!define APP_NAME      "OpenSheet"
!define APP_VERSION   "1.0.0"
!define APP_EXE       "OpenSheet.exe"
!define INST_DIR      "$LOCALAPPDATA\${APP_NAME}"
!define UNINST_KEY    "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

Name              "${APP_NAME} ${APP_VERSION}"
OutFile           "OpenSheet-Setup.exe"
InstallDir        "${INST_DIR}"
InstallDirRegKey  HKCU "${UNINST_KEY}" "InstallLocation"
RequestExecutionLevel user
SetCompressor     lzma

!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON   "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Custom welcome page text
!define MUI_WELCOMEPAGE_TEXT "Setup will install ${APP_NAME} ${APP_VERSION} on your computer.$\r$\n$\r$\nClick Next to continue."

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

; ── Helper: kill running instance before installing ──────────────────────────
!macro KillRunningApp
  ; Try graceful close first
  FindWindow $0 "" "${APP_NAME}"
  IntCmp $0 0 +3
    SendMessage $0 ${WM_CLOSE} 0 0
    Sleep 1500

  ; Force-kill if still running
  nsExec::ExecToStack 'taskkill /F /IM "${APP_EXE}" /T'
  Pop $0  ; exit code
  Pop $1  ; output
  Sleep 500
!macroend

Section "Install"
  ; ── Close any running instance ─────────────────────────────────────────────
  !insertmacro KillRunningApp

  SetOutPath "$INSTDIR"
  SetOverwrite try          ; try to overwrite, retry on failure
  File /r "*.*"

  ; Start Menu
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut  "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"
  CreateShortcut  "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"   "$INSTDIR\Uninstall.exe"
  CreateShortcut  "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"

  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Add/Remove Programs (user-level)
  WriteRegStr   HKCU "${UNINST_KEY}" "DisplayName"      "${APP_NAME}"
  WriteRegStr   HKCU "${UNINST_KEY}" "DisplayVersion"   "${APP_VERSION}"
  WriteRegStr   HKCU "${UNINST_KEY}" "InstallLocation"  "$INSTDIR"
  WriteRegStr   HKCU "${UNINST_KEY}" "UninstallString"  "$INSTDIR\Uninstall.exe"
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoModify"         1
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoRepair"         1

  ; .xlsx file association
  WriteRegStr HKCU "Software\Classes\.xlsx\OpenWithProgids" "${APP_NAME}.xlsx" ""
  WriteRegStr HKCU "Software\Classes\${APP_NAME}.xlsx" "" "OpenSheet Document"
  WriteRegStr HKCU "Software\Classes\${APP_NAME}.xlsx\shell\open\command" "" '"$INSTDIR\${APP_EXE}" "%1"'
SectionEnd

Section "Uninstall"
  ; Close app before uninstalling
  !insertmacro KillRunningApp

  RMDir /r "$INSTDIR"
  RMDir /r "$SMPROGRAMS\${APP_NAME}"
  Delete    "$DESKTOP\${APP_NAME}.lnk"
  DeleteRegKey HKCU "${UNINST_KEY}"
  DeleteRegKey HKCU "Software\Classes\${APP_NAME}.xlsx"
SectionEnd
