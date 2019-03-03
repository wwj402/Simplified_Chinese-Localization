; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"9.2.7.4225 T3"					; version of launcher
!define APPNAME 	"Kerio Connect"					; complete name of program
!define APP 		"KerioConnect"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"mailctrl.exe"				; main exe name
!define APPEXE64 	"mailctrl.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR\MailServer"					; main exe relative path
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
; !include "x64.nsh"
; !include "SetEnvironmentVariable.nsh"
; !include "ForEachPath.nsh"
!include "FileFunc.nsh"
; !include "WordFunc.nsh"
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
VIAddVersionKey Comments "${APPNAME} mail server."
VIAddVersionKey CompanyName "Kerio Technologies Inc"
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

; Var AppID

LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${If} $LANGUAGE == 2052

		SetOverwrite on
		SetOutPath "${APPDIR}"
		File /r "d:\SnapShot\kerio\Files\@PROGRAMFILES@\Kerio\MailServer\zh-CN\mailctrl.exe"

	${Else}
		SetOverwrite on
		SetOutPath "${APPDIR}"
		File /r "d:\SnapShot\kerio\Files\@PROGRAMFILES@\Kerio\MailServer\mailctrl.exe"
	${EndIf}

	${GetParameters} $0
	${If} $0 == ""
		Exec "${APPDIR}\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"${APPDIR}\${APPEXE}"' '' ''
	${Else}
		Exec '"${APPDIR}\${APPEXE}" "$0"'
		; ExecDos::exec /ASYNC /TOSTACK '"${APPDIR}\${APPEXE}" "$0"' '' ''
	${EndIf}

SectionEnd

