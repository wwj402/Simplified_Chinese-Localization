; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"14.1.17.0"					; version of launcher
!define APPNAME 	"JetBrains IntelliJIDEA"		; complete name of program
!define APP 		"IntelliJIDEA"							; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"idea.exe"						; main exe name
!define APPEXE64 	"idea64.exe"					; main exe 64 bit name
!define APPDIR 		"$EXEDIR"						; main exe relative path
!define APPSWITCH 	``								; some default Parameters
!define JAVAHOME	"jre"
!define JAVAHOME64	"jre64"

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
!include "SetEnvironmentVariable.nsh"


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

	GetFullPathName $0 "$EXEDIR\.."
	${If} ${RunningX64}
		${SetEnvironmentVariable} JAVA_HOME "$0\${JAVAHOME64}"
		Exec "${APPDIR}\${APPEXE64}"
	${Else}
		${SetEnvironmentVariable} JAVA_HOME "$0\${JAVAHOME}"
		Exec "${APPDIR}\${APPEXE}"
	${EndIf}

SectionEnd

