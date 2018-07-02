; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER				"2017.3.0.0"				; version of launcher
!define APPNAME			"JetBrains Pycharm"			; complete name of program
!define APP				"Pycharm"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE			"pycharm.exe"				; main exe name
!define APPEXE64		"pycharm64.exe"				; main exe 64 bit name
!define APPDIR			"PyCharm 2017.3\bin"		; main exe relative path
!define APPSWITCH		``							; some default Parameters
!define APPCONFIG		"pycharm.exe.vmoptions"
!define APPCONFIG64		"pycharm64.exe.vmoptions"
!define APPCONFIGENT	"-javaagent:"
!define LICJAR			"JetbrainsCrack-2.6.2.jar"
!define JAVAHOME		"jre32"
!define JAVAHOME64		"jre64"
!define LICDIR			"dvt-license_server\windows"
!define LICEXE			"dvt-jb_licsrv.386.exe"
!define LICEXE64		"dvt-jb_licsrv.amd64.exe"
!define LICSWITCH		`-mode start`
!define APPHOME			"$PROFILE\.PyCharm2017.3\system\.home"
!define APPLICFILE		"$PROFILE\.PyCharm2017.3\config\pycharm.key"
!define APPLICID		"$APPDATA\JetBrains\PermanentUserId"
!define APPLICREG		"HKCU\Software\JavaSoft\Prefs\jetbrains"
!define APPLICREGKEY	"licenseserverticket"

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
!include "FileFunc.nsh"
!include "TextFunc.nsh"
!include "Registry.nsh"
!include "ProcFunc.nsh"
!include "WordFunc.nsh"



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

Var LauncherSwitch
Var AppBaseDir
Var AppHomeDir
Var AppExeUsed
Var AppHomeConfig
Var AppLicConfig
Var LicFlag

	Function "AppInitial"
	; $R9    "path\name"
	; $R8    "path"
	; $R7    "name"
	; $R6    "size"  ($R6="" if directory, $R6="0" if file with /S=)

	; $R0-$R5  are not used (save data in them).
	; ...
	${If} ${FileExists} $R9\bin\${APPEXE}
	${OrIf} ${FileExists} $R9\bin\${APPEXE}
		StrCpy $AppBaseDir "$R9\bin"
		StrCpy $AppHomeDir "$EXEDIR"
		${WordReplace} "$R7" " " "" "+*" $0
		StrCpy $AppHomeConfig "$PROFILE\.$0\system\.home"
		StrCpy $AppLicConfig "$PROFILE\.$0\config\pycharm.key"
		Push "StopLocate"    ; If $var="StopLocate" Then exit from function
	${EndIf}
	; MessageBox MB_OK "$R9, $R8, $R7"
	FunctionEnd


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${Locate} "$EXEDIR" "/L=D /M=PyCharm* /G=0" "AppInitial"


	${IfThen} $AppBaseDir != "" ${|} Goto ParSelect ${|}
	${IfThen} $AppHomeConfig == "" ${|} StrCpy $AppHomeConfig "${APPHOME}" ${|}
	${If} ${FileExists} $AppHomeConfig
		${LineRead} "$AppHomeConfig" "1" $AppHomeDir
		; MessageBox MB_OK "$AppHomeDir"
		IfFileExists "$AppHomeDir\*.*" 0 +3
		StrCpy $AppBaseDir "$AppHomeDir\bin"
		GetFullPathName $AppHomeDir "$AppHomeDir\.."
		GetFullPathName $AppHomeDir "$EXEDIR"
		StrCpy $AppBaseDir "$EXEDIR\${APPDIR}"
	${Else}
		GetFullPathName $AppHomeDir "$EXEDIR"
		StrCpy $AppBaseDir "$EXEDIR\${APPDIR}"
	${EndIf}
	; MessageBox MB_OK "$AppBaseDir;$AppHomeDir"
	ParSelect:

	${GetParameters} $LauncherSwitch
	${Select} "$LauncherSwitch"
	${Case} "/x86"
		StrCpy $AppExeUsed "$AppBaseDir\${APPEXE}"
	${Case} "/x64"
		StrCpy $AppExeUsed "$AppBaseDir\${APPEXE64}"
	${Case} "/code"
		StrCpy $LicFlag "code"
		${If} ${RunningX64}
			StrCpy $AppExeUsed "$AppBaseDir\${APPEXE64}"
		${Else}
			StrCpy $AppExeUsed "$AppBaseDir\${APPEXE}"
		${EndIf}
		Goto RUNAPP
	${Case} "/server"
		StrCpy $LicFlag "server"
		${If} ${RunningX64}
			StrCpy $AppExeUsed "$AppBaseDir\${APPEXE64}"
		${Else}
			StrCpy $AppExeUsed "$AppBaseDir\${APPEXE}"
		${EndIf}
		Goto RUNAPP
	${CaseElse}
		${If} ${RunningX64}
			StrCpy $AppExeUsed "$AppBaseDir\${APPEXE64}"
		${Else}
			StrCpy $AppExeUsed "$AppBaseDir\${APPEXE}"
		${EndIf}
	${EndSelect}

	${registry::Open} "${APPLICREG}" "/K=1 /V=1 /S=0 /N='${APPLICREGKEY}'" $0
	${registry::Find} "$0" $1 $2 $3 $4
	${If} $LauncherSwitch != "/server"
	${AndIf} $4 != ""
		StrCpy $LicFlag "server"
	${EndIf}
	${registry::Close} "$0"
	${registry::Unload}
	; MessageBox MB_OK "$1$\n$2$\n$3$\n$4"

	${IfThen} $AppLicConfig == "" ${|} StrCpy $AppLicConfig "${APPLICFILE}" ${|}
	${If} ${FileExists} $AppLicConfig
		${LineSum} "$AppLicConfig" $0
		StrCpy $1 1
		StrCpy $2 ""
		${Do}
			${IfThen} $1 > $0 ${|} ${ExitDo} ${|}
			${LineRead} "$AppLicConfig" "$1" $3
			StrCpy $2 "$2$3"
			IntOp $1 $1 + 1
		${LoopWhile} $1 < 5
		; MessageBox MB_OK "$0;$1;$2;$3"
	${EndIf}

	${WordFind} "$2" "url:" "*" $3

	; MessageBox MB_OK "|$2|;$3"
	${If} ${FileExists} $AppLicConfig
	${AndIf} $3 < 1
	${AndIf} $LauncherSwitch != "/code"
	${AndIf} $LicFlag == ""
		StrCpy $LicFlag "code"
	${EndIf}

	; MessageBox MB_OK "$AppExeUsed;$LicFlag"
	RUNAPP:

	${Select} "$LicFlag"
	${Case} "server"
		${If} ${FileExists} "$AppBaseDir\${APPCONFIG}"
			${ConfigWrite} "$AppBaseDir\${APPCONFIG}" "${APPCONFIGENT}" "" $0
		${EndIf}
		${If} ${FileExists} "$AppBaseDir\${APPCONFIG64}"
			${ConfigWrite} "$AppBaseDir\${APPCONFIG64}" "${APPCONFIGENT}" "" $0
		${EndIf}
/* 		${GetProcessPID} "${LICEXE}" $0
		${GetProcessPID} "${LICEXE64}" $1
		${If} ${RunningX64}
			${If} $0 = 0
			${AndIf} $1 = 0
				ExecDos::exec /ASYNC /TOSTACK '"$EXEDIR\${LICDIR}\${LICEXE64}" ${LICSWITCH}' '' ''
			${EndIf}
		${Else}
			${If} $0 = 0
			${AndIf} $1 = 0
				ExecDos::exec /ASYNC /TOSTACK '"$EXEDIR\${LICDIR}\${LICEXE}" ${LICSWITCH}' '' ''
			${EndIf}
		${EndIf} */
	${Case} "code"
		${If} ${FileExists} "$AppBaseDir\${APPCONFIG}"
		${AndIf} ${FileExists} "$AppHomeDir\${LICJAR}"
			${ConfigWrite} "$AppBaseDir\${APPCONFIG}" "${APPCONFIGENT}" "$AppHomeDir\${LICJAR}" $0
		${EndIf}
		${If} ${FileExists} "$AppBaseDir\${APPCONFIG64}"
		${AndIf} ${FileExists} "$AppHomeDir\${LICJAR}"
			${ConfigWrite} "$AppBaseDir\${APPCONFIG64}" "${APPCONFIGENT}" "$AppHomeDir\${LICJAR}" $0
		${EndIf}
	${CaseElse}
		${If} ${FileExists} "$AppBaseDir\${APPCONFIG}"
		${AndIf} ${FileExists} "$AppHomeDir\${LICJAR}"
			${ConfigWrite} "$AppBaseDir\${APPCONFIG}" "${APPCONFIGENT}" "$AppHomeDir\${LICJAR}" $0
		${EndIf}
		${If} ${FileExists} "$AppBaseDir\${APPCONFIG64}"
		${AndIf} ${FileExists} "$AppHomeDir\${LICJAR}"
			${ConfigWrite} "$AppBaseDir\${APPCONFIG64}" "${APPCONFIGENT}" "$AppHomeDir\${LICJAR}" $0
		${EndIf}
		${GetProcessPID} "${LICEXE}" $0
		${GetProcessPID} "${LICEXE64}" $1
		${If} ${RunningX64}
			${If} $0 = 0
			${AndIf} $1 = 0
				ExecDos::exec /ASYNC /TOSTACK '"$AppHomeDir\${LICDIR}\${LICEXE64}" ${LICSWITCH}' '' ''
			${EndIf}
		${Else}
			${If} $0 = 0
			${AndIf} $1 = 0
				ExecDos::exec /ASYNC /TOSTACK '"$AppHomeDir\${LICDIR}\${LICEXE}" ${LICSWITCH}' '' ''
			${EndIf}
		${EndIf}
	${EndSelect}
	; MessageBox MB_OK "$0,$1,$AppExeUsed"

	WriteRegExpandStr HKLM "SOFTWARE\Classes\Applications\pycharm.exe\shell\open\command" "" '"$AppExeUsed" "%1"'
	ReadRegStr $0 HKLM "SOFTWARE\Classes\.ipr" ""
	${If} $0 != ""
		WriteRegExpandStr HKLM "SOFTWARE\Classes\$0\DefaultIcon" "" "$AppExeUsed,0"
		WriteRegExpandStr HKLM "SOFTWARE\Classes\IntelliJIdeaProjectFile\shell\open\command" "" '"$AppExeUsed" "%1"'
	${EndIf}
	ReadRegStr $0 HKLM "SOFTWARE\Classes\.py" ""
	${If} $0 != ""
		WriteRegExpandStr HKLM "SOFTWARE\Classes\$0\DefaultIcon" "" "$AppExeUsed,0"
		WriteRegExpandStr HKLM "SOFTWARE\Classes\$0\shell\open\command" "" '"$AppExeUsed" "%1"'
	${EndIf}

	; ${Execute} '"$AppExeUsed"' "$AppBaseDir" $0
	Exec '"$AppExeUsed"'
	; ExecWait '"$AppExeUsed"'
	; ExecDos::exec /TOSTACK '"$AppExeUsed"' '' ''


SectionEnd

