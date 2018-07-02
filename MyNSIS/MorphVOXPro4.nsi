; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"4.4.71.0"				; version of launcher
!define APPNAME 	"MorphVOX Pro4"			; complete name of program
!define APP 		"MorphVOXPro"			; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"MorphVOXPro.exe"		; main exe name
!define APPEXE64 	"MorphVOXPro.exe"		; main exe 64 bit name
!define APPDIR 		"$EXEDIR"				; main exe relative path
!define APPSWITCH 	``
!define DRVSETUPEXE 	"SBAudioInstall.exe"
!define DRVSETUPEXE64 	"SBAudioInstallx64.exe"
!define DRVSETUPDIR		"$EXEDIR\drivers"
!define DRVSETUPWITCH	`"$usedinf" "*ScreamingBAudio"`

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
; !include "ForEachPath.nsh"
!include "FileFunc.nsh"
!include "WinVer.nsh"


; **************************************************************************
; === Set basic information ===
; **************************************************************************
Name "${APPNAME} Launcher"
OutFile ".\${APP}Launcher.exe"
Icon ".\${APP}.ico"
SilentInstall silent

; **************************************************************************
; === Set version information ===
; **************************************************************************
Caption "${APPNAME} Launcher"
VIProductVersion "${VER}"
VIAddVersionKey ProductName "${APPNAME}"
VIAddVersionKey Comments "${APPNAME} launcher."
VIAddVersionKey CompanyName "Screaming Bee"
VIAddVersionKey LegalCopyright ""
VIAddVersionKey FileDescription "${APPNAME} ${VER}"
VIAddVersionKey FileVersion "${VER}"
VIAddVersionKey ProductVersion "${VER}"
VIAddVersionKey InternalName "${APPNAME}"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey OriginalFilename "${APP}Launcher.exe"

; **************************************************************************
; === Other Actions ===
; **************************************************************************

Var usedappexe
Var usedinf
Var useddrvexe
LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${If} ${RunningX64}
		GetFullPathName $usedappexe "${APPDIR}\${APPEXE64}"
		GetFullPathName $useddrvexe "${DRVSETUPDIR}\${DRVSETUPEXE64}"
	${Else}
		GetFullPathName $usedappexe "${APPDIR}\${APPEXE}"
		GetFullPathName $useddrvexe "${DRVSETUPDIR}\${DRVSETUPEXE}"
	${EndIf}
	; MessageBox MB_OK "$usedappexe"

	${If} ${AtMostWinXP}
	${orIf} ${AtMostWin2003}
		${If} ${RunningX64}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x64\XP\SBAudio.inf"
		${Else}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x86\XP\SBAudio.inf"
		${EndIf}
	${EndIf}
	${If} ${AtLeastWinVista}
	${AndIf} ${AtMostWin8.1}
		${If} ${RunningX64}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x64\W7\SBAudio.inf"
		${Else}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x86\W7\SBAudio.inf"
		${EndIf}
	${EndIf}
	${If} ${AtLeastWin2008}
	${AndIf} ${AtMostWin2012R2}
		${If} ${RunningX64}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x64\W7\SBAudio.inf"
		${Else}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x86\W7\SBAudio.inf"
		${EndIf}
	${EndIf}
	${If} $usedinf == ""
		${If} ${RunningX64}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x64\W10\SBAudio.inf"
		${Else}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x86\W10\SBAudio.inf"
		${EndIf}
	${EndIf}
/* 	${If} ${AtLeastWin10}
	${OrIf} ${AtLeastWin2016}
		${If} ${RunningX64}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x64\W10\SBAudio.inf"
		${Else}
			GetFullPathName $usedinf "${DRVSETUPDIR}\x86\W10\SBAudio.inf"
		${EndIf}
	${EndIf} */
	; MessageBox MB_OK '"$useddrvexe" i ${DRVSETUPWITCH}'
	ExecDos::exec /TOSTACK '"$useddrvexe" i ${DRVSETUPWITCH}' '' ''
	ExecWait "$usedappexe"
	; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" $0' '' ''
	ExecDos::exec /TOSTACK '"$useddrvexe" u ${DRVSETUPWITCH}' '' ''

SectionEnd

