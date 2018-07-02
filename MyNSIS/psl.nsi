; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"18.0.56.0"					; version of launcher
!define APPNAME 	"Passolo"					; complete name of program
!define APP 		"Passolo"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"psl.exe"				; main exe name
!define APPEXE64 	"psl.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"
!define PSLREGROOT "HKCU"
!define PSLREGSUB "Software\SDL\Passolo 2018\System"
!define PSLREGNAME "Language"
!define IDFILEDIR "System\DnAndroidParser"
!define IDFILE "DnAndroidParser.dll"
!define CNDIR "System_cn"
!define ENDIR "System_en"
!define LANGUAGEFILES "system.exe"
!define USERLANGUAGEFILES "user.exe"


; **************************************************************************
; === Best Compression ===
; **************************************************************************
; Unicode true
SetCompressor /SOLID lzma
SetCompressorDictSize 32

; **************************************************************************
; === Includes ===
; **************************************************************************

!include "LogicLib.nsh"
!include "x64.nsh"
!include "SetEnvironmentVariable.nsh"
; !include "ForEachPath.nsh"
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
VIAddVersionKey CompanyName "SDL Passolo GmbH"
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

	ReadEnvStr $homedir PSLHOME
	${IfThen} $homedir == "" ${|} StrCpy $homedir "${APPDIR}" ${|}
	; MessageBox MB_OK "$homedir"
	StrCpy $AppID "Portable"
	${SetEnvironmentVariable} PSLHOME $homedir
	ReadRegDWORD $uilanguage ${PSLREGROOT} "${PSLREGSUB}" ${PSLREGNAME}
	; MessageBox MB_OK "$uilanguage"
	${If} $uilanguage == 2052

		md5dll::GetMD5File "$homedir\${IDFILEDIR}\${IDFILE}"
		Pop $3
		md5dll::GetMD5File "$homedir\${CNDIR}\${IDFILE}"
		Pop $4
		; MessageBox MB_OK "$3||$4"
		${If} $3 != $4
			ExecWait "$homedir\${CNDIR}\${LANGUAGEFILES}"
			ExecWait "$homedir\${CNDIR}\${USERLANGUAGEFILES}"
		${EndIf}

	${Else}
		md5dll::GetMD5File "$homedir\${IDFILEDIR}\${IDFILE}"
		Pop $3
		md5dll::GetMD5File "$homedir\${ENDIR}\${IDFILE}"
		Pop $4
		; MessageBox MB_OK "$3||$4"
		${If} $3 != $4
			ExecWait "$homedir\${ENDIR}\${LANGUAGEFILES}"
			ExecWait "$homedir\${ENDIR}\${USERLANGUAGEFILES}"
		${EndIf}

	${EndIf}

	${GetParameters} $0
	${If} $0 == ""
		Exec "$homedir\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		Exec '"$homedir\${APPEXE}" "$0"'
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" "$0"' '' ''
	${EndIf}

SectionEnd

