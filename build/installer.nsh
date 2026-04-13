!include "LogicLib.nsh"
!include "nsDialogs.nsh"
!include "FileAssociation.nsh"

!ifndef BUILD_UNINSTALLER
  Var PdfAssocCheckbox
  Var ShouldAssociatePdf

  !macro customInit
    StrCpy $ShouldAssociatePdf "0"
  !macroend

  !macro customPageAfterChangeDir
    Page custom PdfAssocPageCreate PdfAssocPageLeave
  !macroend

  Function PdfAssocPageCreate
    nsDialogs::Create 1018
    Pop $0
    ${If} $0 == error
      Abort
    ${EndIf}

    ${NSD_CreateLabel} 0u 0u 100% 24u "체크하면 PDF 파일 더블클릭 시 lookup으로 열립니다."
    Pop $0
    ${NSD_CreateCheckbox} 0u 32u 100% 12u ".pdf 파일을 lookup으로 열기"
    Pop $PdfAssocCheckbox
    ${NSD_SetState} $PdfAssocCheckbox ${BST_UNCHECKED}
    nsDialogs::Show
  FunctionEnd

  Function PdfAssocPageLeave
    ${NSD_GetState} $PdfAssocCheckbox $0
    ${If} $0 == ${BST_CHECKED}
      StrCpy $ShouldAssociatePdf "1"
    ${Else}
      StrCpy $ShouldAssociatePdf "0"
    ${EndIf}
  FunctionEnd

  !macro customInstall
    ${If} $ShouldAssociatePdf == "1"
      !insertmacro APP_ASSOCIATE "pdf" "lookup.PDF" "PDF 문서" "$appExe,0" "Open with lookup" "$\"$appExe$\" $\"%1$\""
      !insertmacro UPDATEFILEASSOC
    ${EndIf}
  !macroend
!else
  !macro customUnInstall
    !insertmacro APP_UNASSOCIATE "pdf" "lookup.PDF"
    !insertmacro UPDATEFILEASSOC
  !macroend
!endif
