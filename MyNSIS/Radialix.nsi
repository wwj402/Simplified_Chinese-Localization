; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"3.0.0.486"				; version of launcher
!define APPNAME 	"Radialix 3"			; complete name of program
!define APP 		"Radialix"				; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"x86\rdl.exe"			; main exe name
!define APPEXE64 	"x64\rdl.exe"			; main exe 64 bit name
!define APPDIR 		"$EXEDIR"				; main exe relative path
!define APPSWITCH 	``
!define IDAEXE 		"idaq.exe"
!define IDAEXE64 	"idaq64.exe"
!define IDADIR		"$EXEDIR\..\IDAPro61"
!define IDASWITCH	``
!define APPEXELNK	"Radialix 3 32-bit.lnk"
!define APPEXELNK64	"Radialix 3 64-bit.lnk"

; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"

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
VIAddVersionKey CompanyName "Radialix Software"
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

Var UsedAppExe
Var UsedAppExeLnk
Var UsedIdaExe
LangString Message 1033 "English message"
LangString Message 2052 "Simplified Chinese message"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${If} ${RunningX64}
		GetFullPathName $UsedAppExe "${APPDIR}\${APPEXE64}"
		StrCpy $UsedAppExeLnk "${APPDIR}\${APPEXELNK64}"
		; GetFullPathName $UsedIdaExe "${IDADIR}\${IDAEXE64}"
	${Else}
		GetFullPathName $UsedAppExe "${APPDIR}\${APPEXE}"
		StrCpy $UsedAppExeLnk "${APPDIR}\${APPEXELNK}"
		; GetFullPathName $UsedIdaExe "${IDADIR}\${IDAEXE}"
	${EndIf}
	; MessageBox MB_OK "$UsedIdaExe||$UsedAppExe"


	${GetParameters} $0
	${Select} $0
	${Case} "/ida"
	; ${If} $0 == "/ida"
		Push $OUTDIR
		${GetParent} "$UsedAppExe" $OUTDIR
		${SetEnvironmentVariable} RDLWORKDIR $OUTDIR
		; ExecShell "open" "$UsedAppExeLnk"
		Exec "$UsedAppExe"
		GetFullPathName $UsedIdaExe "${IDADIR}\${IDAEXE}"
		${GetParent} "$UsedIdaExe" $OUTDIR
		Exec "$UsedIdaExe"
		Pop $OUTDIR
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Case} "/ida64"
	; ${ElseIf} $0 == "/ida64"
		Push $OUTDIR
		${GetParent} "$UsedAppExe" $OUTDIR
		${SetEnvironmentVariable} RDLWORKDIR $OUTDIR
		; ExecShell "open" "$UsedAppExeLnk"
		Exec "$UsedAppExe"
		GetFullPathName $UsedIdaExe "${IDADIR}\${IDAEXE64}"
		${GetParent} "$UsedIdaExe" $OUTDIR
		Exec "$UsedIdaExe"
		Pop $OUTDIR
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Case} "/x86"
		Push $OUTDIR
		${GetParent} "${APPDIR}\${APPEXE}" $OUTDIR
		${SetEnvironmentVariable} RDLWORKDIR $OUTDIR
		; ExecShell "open" "${APPDIR}\${APPEXELNK}"
		Exec "${APPDIR}\${APPEXE}"
		Pop $OUTDIR
	${Case} "/x64"
		Push $OUTDIR
		${GetParent} "${APPDIR}\${APPEXE64}" $OUTDIR
		${SetEnvironmentVariable} RDLWORKDIR $OUTDIR
		; ExecShell "open" "${APPDIR}\${APPEXELNK64}"
		Exec "${APPDIR}\${APPEXE64}"
		Pop $OUTDIR
	${Default}
	; ${Else}
		Push $OUTDIR
		${GetParent} "$UsedAppExe" $OUTDIR
		${SetEnvironmentVariable} RDLWORKDIR $OUTDIR
		; ExecShell "open" "$UsedAppExeLnk"
		Exec "$UsedAppExe"
		Pop $OUTDIR
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" $0' '' ''
	${EndSelect}
	; ${EndIf}


SectionEnd

