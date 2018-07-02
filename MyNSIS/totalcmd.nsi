; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"9.0.0.1"					; version of launcher
!define APPNAME 	"Total Commander"			; complete name of program
!define APP 		"TOTALCMD"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"TOTALCMD.EXE"				; main exe name
!define APPEXE64 	"TOTALCMD64.EXE"			; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"
!define APPCONFIGINI "wincmd.ini"
!define APPCONFIGINISEC "Configuration"
!define APPCONFIGINIKEY "languageini"
!define APPLANGCN "wcmd_chn.lng"
!define APPLANGEN "wcmd_eng.lng"

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
VIAddVersionKey Comments "${APPNAME} ."
VIAddVersionKey CompanyName " Ghisler Software GmbH"
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

Var UsedExe
LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${If} ${RunningX64}
		StrCpy $UsedExe "${APPDIR}\${APPEXE64}"
	${Else}
		StrCpy $UsedExe "${APPDIR}\${APPEXE}"
	${EndIf}

	${If} $LANGUAGE == 2052
		ReadINIStr $0 "${APPDIR}\${APPCONFIGINI}" "${APPCONFIGINISEC}" "${APPCONFIGINIKEY}"
		${IfNot} ${Errors}
		${AndIf} $0 != ${APPLANGCN}
			WriteINIStr "${APPDIR}\${APPCONFIGINI}" "${APPCONFIGINISEC}" "${APPCONFIGINIKEY}" "${APPLANGCN}"
		${EndIf}
	${Else}
		ReadINIStr $0 "${APPDIR}\${APPCONFIGINI}" "${APPCONFIGINISEC}" "${APPCONFIGINIKEY}"
		${IfNot} ${Errors}
		${AndIf} $0 == ${APPLANGCN}
			WriteINIStr "${APPDIR}\${APPCONFIGINI}" "${APPCONFIGINISEC}" "${APPCONFIGINIKEY}" "${APPLANGEN}"
		${EndIf}
	${EndIf}
		
	
	${GetParameters} $0
	${If} $0 == ""
		Exec "$UsedExe"
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		Exec '"$UsedExe" $0'
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" $0' '' ''
	${EndIf}


SectionEnd

