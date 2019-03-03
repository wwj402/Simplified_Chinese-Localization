; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"4.0.0.372"					; version of launcher
!define APPNAME 	"sisulizer"					; complete name of program
!define APP 		"sisulizer"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"sisulizer.exe"				; main exe name
!define APPEXE64 	"sisulizer.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"
!define CONFIGDIR	"$DOCUMENTS\Sisulizer 4"
!define CONFIGFILE	"Sisulizer.slo"

; **************************************************************************
; === Best Compression ===
; **************************************************************************
SetCompressor /SOLID lzma
SetCompressorDictSize 32

; **************************************************************************
; === Includes ===
; **************************************************************************

!include "LogicLib.nsh"
!include "x64.nsh"
; !include "SetEnvironmentVariable.nsh"
; !include "ForEachPath.nsh"
!include "FileFunc.nsh"
!include "WordFunc.nsh"
!include "NewTextReplace.nsh"


; **************************************************************************
; === Set basic information ===
; **************************************************************************
Name "${APP} Launcher"
OutFile ".\${APP}Launcher.exe"
Icon ".\${APP}.ico"
SilentInstall silent

; **************************************************************************
; === Set version information ===
; **************************************************************************
Caption "${APPNAME} Launcher"
VIProductVersion "${VER}"
VIAddVersionKey ProductName "${APPNAME}"
VIAddVersionKey Comments "${APPNAME} Localize your software."
VIAddVersionKey CompanyName "Sisulizer Ltd & Co KG"
VIAddVersionKey LegalCopyright ""
VIAddVersionKey FileDescription "${APPNAME}"
VIAddVersionKey FileVersion "${VER}"
VIAddVersionKey ProductVersion "${VER}"
VIAddVersionKey InternalName "${APPNAME}"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey OriginalFilename "${APP}Launcher.exe"

; **************************************************************************
; === Other Actions ===
; **************************************************************************

Var uilanguage
Var AppID
Var homedir
Var docdir
Var projectfile
LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	ReadEnvStr $homedir SISUHOME
	${IfThen} $homedir == "" ${|} StrCpy $homedir "${APPDIR}" ${|}
	; MessageBox MB_OK "$homedir"
	StrCpy $docdir "${CONFIGDIR}"
	StrCpy $AppID "Portable"
	nsisXML::create
	; nsisXML::load "$homedir\${CONFIGFILE}"
	nsisXML::load "${CONFIGDIR}\${CONFIGFILE}"
	nsisXML::select '/options/languages'
	IntCmp $2 0 notFound1
	; MessageBox MB_OK "$0/$1/$2"
	nsisXML::getAttribute "ui"
	; MessageBox MB_OK "$3"
	StrCpy $uilanguage $3
	notFound1:
	nsisXML::select '/options/general/dirs/item'
	IntCmp $2 0 +2
	nsisXML::setText "$homedir"
	nsisXML::select '/options/projects/item'
	IntCmp $2 0 +3
	nsisXML::getText
	StrCpy $projectfile $3
	nsisXML::select '/options/translationengines/engine/database'
	IntCmp $2 0 notFound2
	nsisXML::setAttribute "directory" "${CONFIGDIR}\TM"
	nsisXML::save "${CONFIGDIR}\${CONFIGFILE}"
	notFound2:
	; MessageBox  MB_OK "$projectfile"
	${WordFind} "${CONFIGDIR}" "\" "#" $0
	${WordFind} "$projectfile" "\" "+$0{*" $1
	; MessageBox  MB_OK "$1"
	${If} "$1" != "${CONFIGDIR}"
		${textreplace::ReplaceInFile} "${CONFIGDIR}\${CONFIGFILE}" "${CONFIGDIR}\${CONFIGFILE}" \
		"$1" "${CONFIGDIR}" "/U=0" $0
	${EndIf}
	; MessageBox  MB_OK "$LANGUAGE"
	${If} $uilanguage == ""
	${AndIf} $LANGUAGE == 2052
/* 		${ForEachFile} $1 $2 "${APPDIR}\zh_cn\*.*"
			; MessageBox MB_OK "$1, $2, $0"
			; ForEachPath TYPE FOUND_PATH FILE_NAME SEARCH_PATH
			; !define ForEachFile '!insertmacro ForEachPath FILES'

			${IfNot} ${FileExists} "$homedir\$2"
				; MessageBox MB_OK "$1"
				CopyFiles /SILENT "$1" "$homedir"
			${Else}
				md5dll::GetMD5File "$1"
				Pop $3
				md5dll::GetMD5File "$homedir\$2"
				Pop $4
				; MessageBox MB_OK "$3||$4"
				${If} $3 != $4
					Delete "$homedir\$2"
					; MessageBox MB_OK "delete ok"
					CopyFiles /SILENT "$1" "$homedir"
					; MessageBox MB_OK "copy ok"
				${EndIf}
			${EndIf}
		${NextFile}
		${ForEachFile} $1 $2 "${APPDIR}\zh_cn\doc\*.*"
			${IfNot} ${FileExists} "$docdir\$2"
				CopyFiles /SILENT "$1" "$docdir"
			${Else}
				md5dll::GetMD5File "$1"
				Pop $3
				md5dll::GetMD5File "$docdir\$2"
				Pop $4
				${If} $3 != $4
					Delete "$docdir\$2"
					CopyFiles /SILENT "$1" "$docdir"
				${EndIf}
			${EndIf}
		${NextFile} */
		; ExecWait "$homedir\pro_zh_cn.exe"
		; ExecWait "$homedir\doc_zh_cn.exe"
		SetOverwrite on
		SetOutPath $EXEDIR
		File /r /x doc "d:\SnapShot\sus\Files\ProgramFilesX86\Sisulizer 4\zh_cn\*.*"
		SetOutPath "$DOCUMENTS\Sisulizer 4"
		File /r "d:\SnapShot\sus\Files\ProgramFilesX86\Sisulizer 4\zh_cn\doc\*.*"
	${ElseIf} $uilanguage == "CH"
/* 		${ForEachFile} $1 $2 "${APPDIR}\zh_cn\*.*"
			; MessageBox MB_OK "$1, $2, $0"
			; ForEachPath TYPE FOUND_PATH FILE_NAME SEARCH_PATH
			; !define ForEachFile '!insertmacro ForEachPath FILES'

			${IfNot} ${FileExists} "$homedir\$2"
				; MessageBox MB_OK "$1"
				CopyFiles /SILENT "$1" "$homedir"
			${Else}
				md5dll::GetMD5File "$1"
				Pop $3
				md5dll::GetMD5File "$homedir\$2"
				Pop $4
				; MessageBox MB_OK "$3||$4"
				${If} $3 != $4
					Delete "$homedir\$2"
					; MessageBox MB_OK "delete ok"
					CopyFiles /SILENT "$1" "$homedir"
					; MessageBox MB_OK "copy ok"
				${EndIf}
			${EndIf}
		${NextFile}
		${ForEachFile} $1 $2 "${APPDIR}\zh_cn\doc\*.*"
			${IfNot} ${FileExists} "$docdir\$2"
				; MessageBox MB_OK "$1"
				CopyFiles /SILENT "$1" "$docdir"
			${Else}
				md5dll::GetMD5File "$1"
				Pop $3
				md5dll::GetMD5File "$docdir\$2"
				Pop $4
				${If} $3 != $4
					Delete "$docdir\$2"
					CopyFiles /SILENT "$1" "$docdir"
				${EndIf}
			${EndIf}
		${NextFile} */
		; ExecWait "$homedir\pro_zh_cn.exe"
		; ExecWait "$homedir\doc_zh_cn.exe"
		SetOverwrite on
		SetOutPath $EXEDIR
		File /r /x doc "d:\SnapShot\sus\Files\ProgramFilesX86\Sisulizer 4\zh_cn\*.*"
		SetOutPath "$DOCUMENTS\Sisulizer 4"
		File /r "d:\SnapShot\sus\Files\ProgramFilesX86\Sisulizer 4\zh_cn\doc\*.*"
	${Else}
/* 		${ForEachFile} $1 $2 "${APPDIR}\en\*.*"
			; MessageBox MB_OK "$1, $2, $0"
			; ForEachPath TYPE FOUND_PATH FILE_NAME SEARCH_PATH
			; !define ForEachFile '!insertmacro ForEachPath FILES'

			${IfNot} ${FileExists} "$homedir\$2"
				; MessageBox MB_OK "$1"
				CopyFiles /SILENT "$1" "homedir"
			${Else}
				md5dll::GetMD5File "$1"
				Pop $3
				md5dll::GetMD5File "$homedir\$2"
				Pop $4
				; MessageBox MB_OK "$3||$4"
				${If} $3 != $4
					Delete "$homedir\$2"
					; MessageBox MB_OK "delete ok"
					CopyFiles /SILENT "$1" "$homedir"
					; MessageBox MB_OK "copy ok"
				${EndIf}
			${EndIf}
		${NextFile}
		${ForEachFile} $1 $2 "${APPDIR}\en\doc\*.*"
			${IfNot} ${FileExists} "$docdir\$2"
				CopyFiles /SILENT "$1" "$docdir"
			${Else}
				md5dll::GetMD5File "$1"
				Pop $3
				md5dll::GetMD5File "$docdir\$2"
				Pop $4
				${If} $3 != $4
					Delete "$docdir\$2"
					CopyFiles /SILENT "$1" "$docdir"
				${EndIf}
			${EndIf}
		${NextFile} */
		; ExecWait "$homedir\pro_en.exe"
		; ExecWait "$homedir\doc_en.exe"
		SetOverwrite on
		SetOutPath $EXEDIR
		File /r /x doc "d:\SnapShot\sus\Files\ProgramFilesX86\Sisulizer 4\en\*.*"
		SetOutPath "$DOCUMENTS\Sisulizer 4"
		File /r "d:\SnapShot\sus\Files\ProgramFilesX86\Sisulizer 4\en\doc\*.*"
	${EndIf}

	${GetParameters} $0
	${If} $0 == ""
		Exec "$homedir\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		Exec '"$homedir\${APPEXE}" "$0"'
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" "$0"' '' ''
	${EndIf}

SectionEnd

