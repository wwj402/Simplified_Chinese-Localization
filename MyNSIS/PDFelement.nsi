; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"6.8.0.3523"					; version of launcher
!define APPNAME 	"Wondershare PDFelement"					; complete name of program
!define APP 		"PDFelement"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"PDFelement.exe"				; main exe name
!define APPEXE64 	"PDFelement.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"
!define LANGUAGECONFIGDIR	"$EXEDIR"
!define LANGUAGECONFIGFILE	"Lanch.dat"
!define LANGUAGECONFIGSEC	"SYSTEM"
!define LANGUAGECONFIGKEY	"DefLanguage"
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
VIAddVersionKey CompanyName "Wondershare Software Co.,Ltd."
VIAddVersionKey LegalCopyright "Copyright © 2003-2018 Wondershare. All Rights Reserved."
VIAddVersionKey FileDescription "Wondershare ${APP}"
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
		WriteINIStr "$userlanguagepath" languages forceuserlanguage "false"
		WriteINIStr "$userlanguagepath" languages language ""
		WriteINIStr "$userlanguagepath" languages English(1033) "ENG"
		WriteINIStr "$userlanguagepath" languages Japanese(1041) "JPN"
		WriteINIStr "$userlanguagepath" languages German(1031) "DEU"
		WriteINIStr "$userlanguagepath" languages French(1036) "FRA"
		WriteINIStr "$userlanguagepath" languages Portuguese(2070) "PTG"
		WriteINIStr "$userlanguagepath" languages Spanish(1034) "ESP"
		WriteINIStr "$userlanguagepath" languages Italian(1040) "ITA"
		WriteINIStr "$userlanguagepath" languages Dutch(1043) "NLD"
		WriteINIStr "$userlanguagepath" languages Russian(1049) "RUS"
		WriteINIStr "$userlanguagepath" languages SimpChinese(2052) "CHS"
		WriteINIStr "$userlanguagepath" languages TradChinese(1028) "CHS-CN"
	${EndIf}
	ReadINIStr $0 "$userlanguagepath" languages forceuserlanguage
	${if} $0 == true
		ReadINIStr $languageinistr "$userlanguagepath" languages language
		${IfThen} $languageinistr == "" ${|} StrCpy $languageinistr "ENG" ${|}
	${Else}
		${LineFind} "$userlanguagepath" "/NUL" "2:-1" "FindLanguageStr"
		${IfThen} $languageinistr == "" ${|} StrCpy $languageinistr "ENG" ${|}
	${EndIf}
	WriteINIStr "${LANGUAGECONFIGDIR}\${LANGUAGECONFIGFILE}" "${LANGUAGECONFIGSEC}" "${LANGUAGECONFIGKEY}" "$languageinistr"



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
	${WordFind2X} '$R9' '(' ')' '+1' $0
	${If} $0 == $LANGUAGE
		${WordFind} '$R9' '=' '+2' $languageinistr
		StrCpy $1 "StopLineFind"
	${Else}
		${IfThen} $1 == "StopLineFind" ${|} StrCpy $1 "" ${|}
	${EndIf}
	; MessageBox MB_OK "$0;$LANGUAGE;$languageinistr"
	Push $1      ; If $var="StopLineFind"  Then exit from function
	               ; If $var="SkipWrite"     Then skip current line (ignored if "/NUL")
FunctionEnd


