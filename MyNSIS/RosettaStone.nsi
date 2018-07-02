; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"5.0.37.0"					; version of launcher
!define APPNAME 	"RosettaStone"					; complete name of program
!define APP 		"RosettaStone"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"Rosetta Stone.exe"				; main exe name
!define APPEXE64 	"Rosetta Stone.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR\Rosetta Stone\Rosetta Stone Language Training"					; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"


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
!include "ForEachPath.nsh"
!include "FileFunc.nsh"


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
LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	ReadEnvStr $homedir SISUHOME
	${IfThen} $homedir == "" ${|} StrCpy $homedir "${APPDIR}" ${|}
	; MessageBox MB_OK "$homedir"
	StrCpy $AppID "Portable"
	nsisXML::create
	nsisXML::load "$homedir\${CONFIGFILE}"
	nsisXML::select '/options/languages'
	IntCmp $2 0 notFound
	; MessageBox MB_OK "$0/$1/$2"
	nsisXML::getAttribute "ui"
	; MessageBox MB_OK "$3"
	StrCpy $uilanguage $3
	notFound:
	; MessageBox  MB_OK "$LANGUAGE"
	${If} $uilanguage == ""
	${AndIf} $LANGUAGE == 2052
		${ForEachFile} $1 $2 "${APPDIR}\zh_cn\*.*"
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
	${ElseIf} $uilanguage == "CH"
		${ForEachFile} $1 $2 "${APPDIR}\zh_cn\*.*"
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
	${Else}
		${ForEachFile} $1 $2 "${APPDIR}\en\*.*"
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
	${EndIf}

	${GetParameters} $0
	${If} $0 == ""
		Exec "$homedir\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		Exec '"${APPDIR}\$homedir" "$0"'
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" "$0"' '' ''
	${EndIf}


SectionEnd

