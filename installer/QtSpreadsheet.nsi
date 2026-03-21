Unicode True

!define APP_NAME      "QtSpreadsheet"
!define APP_VERSION   "1.0.0"
!define APP_EXE       "QtSpreadsheet.exe"

; Install to user's local AppData — no admin rights needed
!define INST_DIR      "$LOCALAPPDATA\${APP_NAME}"
!define UNINST_KEY    "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

Name              "${APP_NAME} ${APP_VERSION}"
OutFile           "QtSpreadsheet-Setup.exe"
InstallDir        "${INST_DIR}"
InstallDirRegKey  HKCU "${UNINST_KEY}" "InstallLocation"

; User level — no UAC prompt, no admin password
RequestExecutionLevel user

SetCompressor lzma

!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON   "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Section "Install"
  SetOutPath "$INSTDIR"
  File /r "*.*"

  ; Start Menu shortcut (user-level, no admin needed)
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut  "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"
  CreateShortcut  "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"   "$INSTDIR\Uninstall.exe"

  ; Desktop shortcut
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"

  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Register in Add/Remove Programs — HKCU instead of HKLM (no admin needed)
  WriteRegStr   HKCU "${UNINST_KEY}" "DisplayName"      "${APP_NAME}"
  WriteRegStr   HKCU "${UNINST_KEY}" "DisplayVersion"   "${APP_VERSION}"
  WriteRegStr   HKCU "${UNINST_KEY}" "InstallLocation"  "$INSTDIR"
  WriteRegStr   HKCU "${UNINST_KEY}" "UninstallString"  "$INSTDIR\Uninstall.exe"
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoModify"         1
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoRepair"         1

  ; .xlsx file association — HKCU (no admin needed)
  WriteRegStr HKCU "Software\Classes\.xlsx\OpenWithProgids" "${APP_NAME}.xlsx" ""
  WriteRegStr HKCU "Software\Classes\${APP_NAME}.xlsx" "" "QtSpreadsheet Document"
  WriteRegStr HKCU "Software\Classes\${APP_NAME}.xlsx\shell\open\command" "" '"$INSTDIR\${APP_EXE}" "%1"'
SectionEnd

Section "Uninstall"
  RMDir /r "$INSTDIR"
  RMDir /r "$SMPROGRAMS\${APP_NAME}"
  Delete    "$DESKTOP\${APP_NAME}.lnk"
  DeleteRegKey HKCU "${UNINST_KEY}"
  DeleteRegKey HKCU "Software\Classes\${APP_NAME}.xlsx"
SectionEnd
