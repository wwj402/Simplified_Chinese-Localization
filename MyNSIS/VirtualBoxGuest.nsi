; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"1.0.0.0"			; version of launcher
!define APPNAME 	"VBOX GUEST"		; complete name of program
!define APP 		"VBOXGUEST"			; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"devcon_x86.exe"	; main exe name
!define APPEXE64 	"devcon_x64.exe"	; main exe 64 bit name
!define APPDIR 		"$EXEDIR"			; main exe relative path
!define APPSWITCH 	``					; some default Parameters
!define SERVICENAME "VBoxGuest"
!define DRIVERINF "VBoxGuest.inf"

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
!include "servicelib.nsh"


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

	!insertmacro SERVICE "status" "${SERVICENAME}" ""
	Pop $0
	; MessageBox MB_OK "$0"
	${Switch} $0
	${Case} unknown
		!insertmacro SERVICE "stop" "${SERVICENAME}" ""
	${Case} false
		${If} ${RunningX64}
		${AndIf} ${FileExists} "${APPDIR}\${APPEXE64}"
			ExecDos::exec /TOSTACK '"${APPDIR}\${APPEXE64}" install "${APPDIR}\${DRIVERINF}" "PCI\VEN_80ee&DEV_cafe"' '' ''
		${Else}
			ExecDos::exec /TOSTACK '"${APPDIR}\${APPEXE}" install "${APPDIR}\${DRIVERINF}" "PCI\VEN_80ee&DEV_cafe"' '' ''
		${EndIf}
		!insertmacro SERVICE "start" "${SERVICENAME}" ""
		ExecDos::exec /TOSTACK "$SYSDIR\VBoxTray.exe" "" ""
		${Break}
	${Case}	running
		${Break}
	${CaseElse}
		!insertmacro SERVICE "start" "${SERVICENAME}" ""
	${EndSwitch}

SectionEnd

