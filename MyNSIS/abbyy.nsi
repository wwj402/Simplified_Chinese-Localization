; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"12.0.104.799"					; version of launcher
!define APPNAME 	"ABBYY PDF Transformer+"					; complete name of program
!define APP 		"PDFTransformer"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"Transformer.exe"				; main exe name
!define APPEXE64 	"Transformer.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"
!define LANGUAGEREGROOT	"HKCU"
!define LANGUAGEREGSUB	"Software\ABBYY\PDFTransformer\12.00"
!define LANGUAGEREGKEY	"InterfaceLanguage"
!define LANGUAGECONFIGSEC	"languages"
!define LANGUAGECONFIGKEY	"current"
!define USERLANGUAGEFILE "${APP}_lng.ini"

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
!include "textfunc.nsh"
!include "WordFunc.nsh"
; !include "NewTextReplace.nsh"


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
VIAddVersionKey Comments "${APPNAME} launcher"
VIAddVersionKey CompanyName "ABBYY Software Co.,Ltd."
VIAddVersionKey LegalCopyright "Copyright © 2003-2018 ABBYY. All Rights Reserved."
VIAddVersionKey FileDescription "ABBYY ${APP}"
VIAddVersionKey FileVersion "${VER}"
VIAddVersionKey ProductVersion "${VER}"
VIAddVersionKey InternalName "${APPEXE}"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey OriginalFilename "${APP}Launcher.exe"

; **************************************************************************
; === Other Actions ===
; **************************************************************************

Var userlanguagepath
Var languageinistr
LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	ReadEnvStr $userlanguagepath SPOONAPPDIR
	${IfThen} $userlanguagepath == "" ${|} StrCpy $userlanguagepath "${APPDIR}" ${|}
	; MessageBox MB_OK "$userlanguagepath"
	StrCpy $userlanguagepath "$userlanguagepath\${USERLANGUAGEFILE}"
	${IfNot} ${FileExists} "$userlanguagepath"
		SetOverwrite on
		SetOutPath $EXEDIR
		File "/oname=$userlanguagepath" "d:\APPs\PortableApps\ABBYYFineReaderPortable\ChooseLang.ini"
		${LineFind} "$userlanguagepath" "/NUL" "1:-1" "FindLanguageStr"
		${IfThen} $languageinistr == "" ${|} StrCpy $languageinistr "0" ${|}
		WriteINIStr "$userlanguagepath" "${LANGUAGECONFIGSEC}" "${LANGUAGECONFIGKEY}" "$languageinistr"
	${EndIf}
	ReadINIStr $0 "$userlanguagepath" "${LANGUAGECONFIGSEC}" "${LANGUAGECONFIGKEY}"
	${if} $0 != ""
		WriteRegExpandStr ${LANGUAGEREGROOT} "${LANGUAGEREGSUB}" "${LANGUAGEREGKEY}" $0
	${EndIf}

	${GetParameters} $0
	${If} $0 == ""
		Exec "${APPDIR}\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		Exec '"${APPDIR}\${APPEXE}" "$0"'
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" "$0"' '' ''
	${EndIf}

SectionEnd

Function "FindLanguageStr"
	; $R9       current line
	; $R8       current line number
	; $R7       current line negative number
	; $R6       current range of lines
	; $R5       handle of a file opened to read
	; $R4       handle of a file opened to write ($R4="" if "/NUL")

	; you can use any string functions
	; $R0-$R3  are not used (save data in them).
	; ...
	${TrimNewLines} '$R9' $R9
	${WordFind} '$R9' '=' '+1' $0
	${WordFind2X} '$0' '(' ')' '-1' $0
	${If} $0 == $LANGUAGE
		${WordFind} '$R9' '=' '-1' $languageinistr
		StrCpy $1 "StopLineFind"
	${Else}
		${IfThen} $1 == "StopLineFind" ${|} StrCpy $1 "" ${|}
	${EndIf}
	; MessageBox MB_OK "$0;$LANGUAGE;$languageinistr"
	Push $1      ; If $var="StopLineFind"  Then exit from function
	               ; If $var="SkipWrite"     Then skip current line (ignored if "/NUL")
FunctionEnd


