!include "LogicLib.nsh"

!macro customHeader
  Caption "lookup v${VERSION} 업데이트"
  BrandingText "lookup v${VERSION} 업데이트"
!macroend


!ifndef BUILD_UNINSTALLER
  Function LookupCleanupLegacy
    ; Run after the standard upgrade uninstaller and before copying new files.
    ; The helper preserves user data and only removes verified lookup artifacts.
    InitPluginsDir
    SetOutPath "$PLUGINSDIR"
    File /oname=lookup-legacy-cleanup.ps1 "${BUILD_RESOURCES_DIR}\legacy-cleanup.ps1"
    DetailPrint "이전 lookup 설치 정리 중"
    nsExec::ExecToStack '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "$PLUGINSDIR\lookup-legacy-cleanup.ps1" -InstallDir "$INSTDIR" -Phase Apply -AllowElevation'
    Pop $R6
    Pop $R7
    ${If} $R6 != 0
      DetailPrint "$R7"
      MessageBox MB_OK|MB_ICONEXCLAMATION "이전 lookup 설치 정리에 실패했습니다.$\r$\n실행 중인 lookup을 닫고 관리자 권한 확인 후 다시 설치해 주세요." /SD IDOK
      SetErrorLevel 2
      Quit
    ${EndIf}
    ClearErrors
  FunctionEnd

  !macro customUnInstallCheck
    ; Keep electron-builder's failure handling for a running/broken uninstaller.
    ; A missing old uninstaller is handled by the verified cleanup helper.
    ${IfNot} ${Errors}
      ${If} $R0 != 0
        MessageBox MB_OK|MB_ICONEXCLAMATION "$(uninstallFailed): $R0" /SD IDOK
        SetErrorLevel 2
        Quit
      ${EndIf}
    ${EndIf}
    ${If} $installMode != "all"
      Call LookupCleanupLegacy
    ${EndIf}
  !macroend

  !macro customUnInstallCheckCurrentUser
    ${IfNot} ${Errors}
      ${If} $R0 != 0
        MessageBox MB_OK|MB_ICONEXCLAMATION "$(uninstallFailed): $R0" /SD IDOK
        SetErrorLevel 2
        Quit
      ${EndIf}
    ${EndIf}
    Call LookupCleanupLegacy
  !macroend

  !macro customInstall
    ; Shortcut/icon refresh to avoid stale Electron cache.
    Delete "$DESKTOP\lookup.lnk"
    Delete "$SMPROGRAMS\lookup.lnk"
    Delete "$SMPROGRAMS\lookup\lookup.lnk"

    CreateDirectory "$SMPROGRAMS\lookup"
    CreateShortCut "$DESKTOP\lookup.lnk" "$appExe" "" "$INSTDIR\resources\icon.ico" 0
    CreateShortCut "$SMPROGRAMS\lookup\lookup.lnk" "$appExe" "" "$INSTDIR\resources\icon.ico" 0
    CreateShortCut "$SMPROGRAMS\lookup\lookup Uninstall.lnk" "$INSTDIR\Uninstall lookup.exe" "" "$INSTDIR\resources\icon.ico" 0

    ; Make lookup appear in Windows "Open with" app list for supported formats.
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe" "FriendlyAppName" "lookup"
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe" "ApplicationIcon" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\shell\open\command" "" "$\"$appExe$\" $\"%1$\""
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\App Paths\lookup.exe" "" "$appExe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\App Paths\lookup.exe" "Path" "$INSTDIR"

    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".pdf" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".hwp" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".hwpx" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".doc" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".docx" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".xls" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".xlsx" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".xlsm" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".accdb" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".ppt" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".pptx" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".html" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".htm" ""

    WriteRegStr HKCU "Software\Classes\lookup.PDF" "" "lookup PDF"
    WriteRegStr HKCU "Software\Classes\lookup.PDF\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.PDF\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.HWP" "" "lookup HWP"
    WriteRegStr HKCU "Software\Classes\lookup.HWP\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.HWP\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.HWPX" "" "lookup HWPX"
    WriteRegStr HKCU "Software\Classes\lookup.HWPX\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.HWPX\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.DOC" "" "lookup DOC"
    WriteRegStr HKCU "Software\Classes\lookup.DOC\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.DOC\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.DOCX" "" "lookup DOCX"
    WriteRegStr HKCU "Software\Classes\lookup.DOCX\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.DOCX\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.XLS" "" "lookup XLS"
    WriteRegStr HKCU "Software\Classes\lookup.XLS\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.XLS\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.XLSX" "" "lookup XLSX"
    WriteRegStr HKCU "Software\Classes\lookup.XLSX\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.XLSX\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.PPT" "" "lookup PPT"
    WriteRegStr HKCU "Software\Classes\lookup.PPT\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.PPT\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\lookup.PPTX" "" "lookup PPTX"
    WriteRegStr HKCU "Software\Classes\lookup.PPTX\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\lookup.PPTX\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\.pdf\OpenWithProgids" "lookup.PDF" ""
    WriteRegStr HKCU "Software\Classes\.hwp\OpenWithProgids" "lookup.HWP" ""
    WriteRegStr HKCU "Software\Classes\.hwpx\OpenWithProgids" "lookup.HWPX" ""
    WriteRegStr HKCU "Software\Classes\.doc\OpenWithProgids" "lookup.DOC" ""
    WriteRegStr HKCU "Software\Classes\.docx\OpenWithProgids" "lookup.DOCX" ""
    WriteRegStr HKCU "Software\Classes\.xls\OpenWithProgids" "lookup.XLS" ""
    WriteRegStr HKCU "Software\Classes\.xlsx\OpenWithProgids" "lookup.XLSX" ""
    WriteRegStr HKCU "Software\Classes\.ppt\OpenWithProgids" "lookup.PPT" ""
    WriteRegStr HKCU "Software\Classes\.pptx\OpenWithProgids" "lookup.PPTX" ""
    System::Call "shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)"
  !macroend
!else
  !macro customUnInstall
    ; Remove the exact shortcuts created by customInstall, including when
    ; electron-builder's standard shortcut creation is disabled.
    Delete "$DESKTOP\lookup.lnk"
    Delete "$SMPROGRAMS\lookup.lnk"
    Delete "$SMPROGRAMS\lookup\lookup.lnk"
    Delete "$SMPROGRAMS\lookup\lookup Uninstall.lnk"
    RMDir "$SMPROGRAMS\lookup"

    DeleteRegKey HKCU "Software\Classes\Applications\lookup.exe"
    DeleteRegKey HKCU "Software\Classes\lookup.PDF"
    DeleteRegKey HKCU "Software\Classes\lookup.HWP"
    DeleteRegKey HKCU "Software\Classes\lookup.HWPX"
    DeleteRegKey HKCU "Software\Classes\lookup.DOC"
    DeleteRegKey HKCU "Software\Classes\lookup.DOCX"
    DeleteRegKey HKCU "Software\Classes\lookup.XLS"
    DeleteRegKey HKCU "Software\Classes\lookup.XLSX"
    DeleteRegKey HKCU "Software\Classes\lookup.PPT"
    DeleteRegKey HKCU "Software\Classes\lookup.PPTX"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\App Paths\lookup.exe"

    DeleteRegValue HKCU "Software\Classes\.pdf\OpenWithProgids" "lookup.PDF"
    DeleteRegValue HKCU "Software\Classes\.hwp\OpenWithProgids" "lookup.HWP"
    DeleteRegValue HKCU "Software\Classes\.hwpx\OpenWithProgids" "lookup.HWPX"
    DeleteRegValue HKCU "Software\Classes\.doc\OpenWithProgids" "lookup.DOC"
    DeleteRegValue HKCU "Software\Classes\.docx\OpenWithProgids" "lookup.DOCX"
    DeleteRegValue HKCU "Software\Classes\.xls\OpenWithProgids" "lookup.XLS"
    DeleteRegValue HKCU "Software\Classes\.xlsx\OpenWithProgids" "lookup.XLSX"
    DeleteRegValue HKCU "Software\Classes\.ppt\OpenWithProgids" "lookup.PPT"
    DeleteRegValue HKCU "Software\Classes\.pptx\OpenWithProgids" "lookup.PPTX"
    System::Call "shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)"
  !macroend
!endif


