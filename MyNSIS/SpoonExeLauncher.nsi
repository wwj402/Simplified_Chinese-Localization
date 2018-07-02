; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"2016.1.4.0"			; version of launcher
!define APPNAME 	"JetBrains Pycharm"		; complete name of program
!define APP 		"Pycharm"				; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"pycharm.exe"			; main exe name
!define APPEXE64 	"pycharm64.exe"			; main exe 64 bit name
!define APPDIR 		"$EXEDIR"				; main exe relative path
!define APPSWITCH 	``						; some default Parameters

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
VIAddVersionKey Comments "${APPNAME} to be run from launcher."
VIAddVersionKey CompanyName ""
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




; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${If} ${RunningX64}
		Exec "${APPDIR}\${APPEXE64}"
	${Else}
		Exec "${APPDIR}\${APPEXE}"
	${EndIf}

SectionEnd

