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
!define CNDIR "System_zh-CN"
!define ENDIR "System_en-US"
!define LANGUAGEFILES "system.exe"
!define USERLANGUAGEFILES "user.exe"
!define LANGUAGEROOT "HKLM"
!define LANGUAGESUB "SOFTWARE\Classes\MIME\Database\Rfc1766"


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
!include "ProcFunc.nsh"
!include "WordFunc.nsh"


; **************************************************************************
; === Set basic information ===
; **************************************************************************
Name "${APP} Launcher"
OutFile ".\${APP}Launcher2018.exe"
Icon ".\${APP}2018.ico"
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
Var PslFindOutput
Var PslFindPath
Var PslFindPid
; Var PslSpoonFileName
; Var PslSpoonPid
LangString Message1 1033 "${APPEXE} is running. Path $PslFindPath. $\r$\n\
							Please close the program, then run again. "
LangString Message1 2052 "${APPEXE} 正在运行。路径 $PslFindPath。$\r$\n\
							请退出正在运行的程序后，再次运行。"
LangString Message2 1033 "${APPEXE} is running. Some files need to be copied. $\r$\n\
							Please close the program, then run again. "
LangString Message2 2052 "${APPEXE} 正在运行。需要拷贝文件。$\r$\n\
							请退出正在运行的程序后，再次运行。"


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
	IntFmt $0 "0x%04x" "$uilanguage"
	System::Call "kernel32::LCIDToLocaleName(i r0, t .r1, i 0, l 0) i .r2"
	System::Call "kernel32::LCIDToLocaleName(i r0, t .r1, i r2, l 0) i .r2"
	; MessageBox MB_OK "$uilanguage, $0, $1, $2"
	${If} $2 == 0
		IntFmt $0 "%04x" "$uilanguage"
		ReadRegStr $1 ${LANGUAGEROOT} "${LANGUAGESUB}" "$0"
		${WordFind} "$1" ";" "+1" $1
		; MessageBox MB_OK "$uilanguage, $0, $1"
	${EndIf}

	${Select} $1
	${Case} "zh-CN"
		SetOverwrite on
		SetOutPath "$APPDATA\SDL\Passolo 2018"
		File /r "d:\SnapShot\sdl\psl2018\Files\@PROGRAMFILESX86@\SDL Passolo 2018\zh-CN\*.dat"
	${Case} "de-DE"
		SetOverwrite on
		SetOutPath "$APPDATA\SDL\Passolo 2018"
		File /r "d:\SnapShot\sdl\psl2018\Files\@PROGRAMFILESX86@\SDL Passolo 2018\de-DE\*.dat"
	${Case} "de"
		SetOverwrite on
		SetOutPath "$APPDATA\SDL\Passolo 2018"
		File /r "d:\SnapShot\sdl\psl2018\Files\@PROGRAMFILESX86@\SDL Passolo 2018\de-DE\*.dat"
	${CaseElse}
		SetOverwrite on
		SetOutPath "$APPDATA\SDL\Passolo 2018"
		File /r "d:\SnapShot\sdl\psl2018\Files\@PROGRAMFILESX86@\SDL Passolo 2018\en-US\*.dat"		
	${EndSelect}

	${If} $uilanguage == 2052

		md5dll::GetMD5File "$homedir\${IDFILEDIR}\${IDFILE}"
		Pop $3
		md5dll::GetMD5File "$homedir\${CNDIR}\${IDFILE}"
		Pop $4
		; MessageBox MB_OK "$3||$4"
		${If} $3 != $4

			${EnumProcessPaths} "PslFind" $PslFindOutput
			${GetProcessPath} "${APPEXE}" $0
			; MessageBox MB_OK "$0"
			${If} $0 != "$homedir\${APPEXE}"
				ExecWait "$homedir\${CNDIR}\${LANGUAGEFILES}"
				ExecWait "$homedir\${CNDIR}\${USERLANGUAGEFILES}"
			${Else}
/* 				MessageBox MB_YESNO "$(Message2)" IDYES WaitAndRun1 IDNO 0
				Abort
				WaitAndRun1:
				${ProcessWaitClose} "${APPEXE}" "-1" $2
				${ProcessWaitClose} "$PslSpoonFileName" "-1" $2
				ExecWait "$homedir\${CNDIR}\${LANGUAGEFILES}"
				ExecWait "$homedir\${CNDIR}\${USERLANGUAGEFILES}" */
				MessageBox MB_OK "$(Message2)"
				Abort
			${EndIf}

		${EndIf}

	${Else}
		md5dll::GetMD5File "$homedir\${IDFILEDIR}\${IDFILE}"
		Pop $3
		md5dll::GetMD5File "$homedir\${ENDIR}\${IDFILE}"
		Pop $4
		; MessageBox MB_OK "$3||$4"
		${If} $3 != $4
			${EnumProcessPaths} "PslFind" $PslFindOutput
			${GetProcessPath} "${APPEXE}" $0
			; MessageBox MB_OK "$0"
			${If} $0 != "$homedir\${APPEXE}"
				ExecWait "$homedir\${ENDIR}\${LANGUAGEFILES}"
				ExecWait "$homedir\${ENDIR}\${USERLANGUAGEFILES}"
			${Else}
/* 				MessageBox MB_YESNO "$(Message2)" IDYES WaitAndRun2 IDNO 0
				Abort
				WaitAndRun2:
				${ProcessWaitClose} "${APPEXE}" "-1" $2
				${ProcessWaitClose} "$PslSpoonFileName" "-1" $2
				ExecWait "$homedir\${ENDIR}\${LANGUAGEFILES}"
				ExecWait "$homedir\${ENDIR}\${USERLANGUAGEFILES}" */
				MessageBox MB_OK "$(Message2)"
				Abort
			${EndIf}

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

	Function "PslFind"
		Pop $PslFindPath			; matching path string
		Pop $PslFindPid			; matching process PID
		; MessageBox MB_OK "$PslFindPath; $PslFindPid"
		${If} $PslFindPath == "$homedir\${APPEXE}"
/* 			MessageBox MB_YESNO "$(Message1)" IDYES PslFindYes1 IDNO 0
			Abort
			PslFindYes1:
			MessageBox MB_OK "$PslFindPid; $PslSpoonFileName"
			${TerminateProcess} "$PslFindPid" $R1
			${ProcessWaitClose} "$PslSpoonFileName" "10000" $R1 */
			MessageBox MB_OK "$(Message1)"
			Push 0
			Abort
		${EndIf}
		Push 1			; must return 1 on the stack to continue
							; must return some value or corrupt the stack
							; DO NOT save data in $0-$9
	FunctionEnd

