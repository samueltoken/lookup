!include "LogicLib.nsh"


!ifndef BUILD_UNINSTALLER
  !macro customInstall
    ; Shortcut/icon refresh to avoid stale Electron cache.
    Delete "$DESKTOP\lookup*.lnk"
    Delete "$DESKTOP\lookup.lnk"
    Delete "$SMPROGRAMS\lookup*.lnk"
    Delete "$SMPROGRAMS\lookup.lnk"
    Delete "$SMPROGRAMS\lookup\lookup.lnk"

    CreateDirectory "$SMPROGRAMS\lookup"
    CreateShortCut "$DESKTOP\lookup.lnk" "$appExe" "" "$INSTDIR\resources\icon.ico" 0
    CreateShortCut "$SMPROGRAMS\lookup\lookup.lnk" "$appExe" "" "$INSTDIR\resources\icon.ico" 0
    CreateShortCut "$SMPROGRAMS\lookup\lookup Uninstall.lnk" "$INSTDIR\Uninstall lookup.exe" "" "$INSTDIR\resources\icon.ico" 0

    ; Make lookup appear in Windows "Open with" app list for supported formats.
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe" "FriendlyAppName" "lookup"
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\DefaultIcon" "" "$INSTDIR\resources\icon.ico,0"
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\shell\open\command" "" "$\"$appExe$\" $\"%1$\""

    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".pdf" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".hwp" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".hwpx" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".doc" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".docx" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".xls" ""
    WriteRegStr HKCU "Software\Classes\Applications\lookup.exe\SupportedTypes" ".xlsx" ""

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

    WriteRegStr HKCU "Software\Classes\.pdf\OpenWithProgids" "lookup.PDF" ""
    WriteRegStr HKCU "Software\Classes\.hwp\OpenWithProgids" "lookup.HWP" ""
    WriteRegStr HKCU "Software\Classes\.hwpx\OpenWithProgids" "lookup.HWPX" ""
    WriteRegStr HKCU "Software\Classes\.doc\OpenWithProgids" "lookup.DOC" ""
    WriteRegStr HKCU "Software\Classes\.docx\OpenWithProgids" "lookup.DOCX" ""
    WriteRegStr HKCU "Software\Classes\.xls\OpenWithProgids" "lookup.XLS" ""
    WriteRegStr HKCU "Software\Classes\.xlsx\OpenWithProgids" "lookup.XLSX" ""
  !macroend
!else
  !macro customUnInstall
    DeleteRegKey HKCU "Software\Classes\Applications\lookup.exe"
    DeleteRegKey HKCU "Software\Classes\lookup.PDF"
    DeleteRegKey HKCU "Software\Classes\lookup.HWP"
    DeleteRegKey HKCU "Software\Classes\lookup.HWPX"
    DeleteRegKey HKCU "Software\Classes\lookup.DOC"
    DeleteRegKey HKCU "Software\Classes\lookup.DOCX"
    DeleteRegKey HKCU "Software\Classes\lookup.XLS"
    DeleteRegKey HKCU "Software\Classes\lookup.XLSX"

    DeleteRegValue HKCU "Software\Classes\.pdf\OpenWithProgids" "lookup.PDF"
    DeleteRegValue HKCU "Software\Classes\.hwp\OpenWithProgids" "lookup.HWP"
    DeleteRegValue HKCU "Software\Classes\.hwpx\OpenWithProgids" "lookup.HWPX"
    DeleteRegValue HKCU "Software\Classes\.doc\OpenWithProgids" "lookup.DOC"
    DeleteRegValue HKCU "Software\Classes\.docx\OpenWithProgids" "lookup.DOCX"
    DeleteRegValue HKCU "Software\Classes\.xls\OpenWithProgids" "lookup.XLS"
    DeleteRegValue HKCU "Software\Classes\.xlsx\OpenWithProgids" "lookup.XLSX"
  !macroend
!endif

