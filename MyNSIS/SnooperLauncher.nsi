; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"1.48.4.0"			; version of launcher
!define APPNAME 	"Snooper"		; complete name of program
!define APP 		"Snooper"				; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"snpr.exe"				; main exe name
!define APPEXE64 	"snpr.exe"			; main exe 64 bit name
!define APPDIR 		"$EXEDIR"				; main exe relative path
!define APPSWITCH 	``						; some default Parameters
!define CONFIGINI	"$LOCALAPPDATA\Snooper\snooper.ini"
!define INISECTION	"GENERAL"
!define INIKEY	"DEFAULTDIR"
!define INIDEFVALUE "$DOCUMENTS\Snooper"


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
VIAddVersionKey CompanyName "Snooper"
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

	${If} ${FileExists} "${CONFIGINI}"
		ReadINIStr $0 "${CONFIGINI}" "${INISECTION}" "${INIKEY}"
		${If} $0 == ""
		${OrIfNot} $0 == ${INIDEFVALUE}
			WriteINIStr "${CONFIGINI}" "${INISECTION}" "${INIKEY}" ${INIDEFVALUE}
		${EndIf}
	${EndIf}

	${If} ${RunningX64}
	${AndIf} ${FileExists} "${APPDIR}\${APPEXE64}"
		Exec "${APPDIR}\${APPEXE64}"
	${Else}
		Exec "${APPDIR}\${APPEXE}"
	${EndIf}

SectionEnd

