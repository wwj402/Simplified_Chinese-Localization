; **************************************************************************
; === Define constants ===
; **************************************************************************
!define EXEFULLDIR "d:\SnapShot\trados\Files\@PROGRAMFILESX86@\SDL\SDL Trados Studio\Studio15"
!define EXENAME "SDLTradosStudio.exe"
!define USERDIR "d:\SnapShot\trados\Fix"

!ifdef NSIS_UNICODE
	!define /file_version MAJOR "${EXEFULLDIR}\${EXENAME}" 0
	!define /file_version MINOR "${EXEFULLDIR}\${EXENAME}" 1
	!define /file_version OPTION "${EXEFULLDIR}\${EXENAME}" 2
	!define /file_version BUILD "${EXEFULLDIR}\${EXENAME}" 3
	!if ${MAJOR} == 0
		!define VER "15.1.2.48878"		; version of launcher
	!else
		!define VER ${MAJOR}.${MINOR}.${OPTION}.${BUILD}
	!endif
	!undef MAJOR
	!undef MINOR
	!undef OPTION
	!undef BUILD
!else
	!echo "${NSIS_VERSION}"
	!getdllversion "${EXEFULLDIR}\${EXENAME}" Expv_
	!define VER "${Expv_1}.${Expv_2}.${Expv_3}.${Expv_4}"
!endif

!execute '"ProductInfo.exe" "${EXEFULLDIR}\${EXENAME}"'
!searchparse /file "ProductInfo.ini" "Comments={" COMMENTS "}, CompanyName={" COMPANYNAME \
				"}, FileDescription={" FILEDESCRIPTION "}, FileVersion={" FILEVERSION "}, "
!searchparse /file "ProductInfo.ini" "LegalCopyright={" LEGALCOPYRIGHT "}, LegalTrademarks={" \
				LEGALTRADEMARKS "}, OriginalFileName={" ORIGINALFILENAME "}, PrivateBuild={" PRIVATEBUILD "}, "
!searchparse /file "ProductInfo.ini" "ProductName={" PRODUCTNAME "}, ProductVersion={" PRODUCTVERSION \
				"}, SpecialBuild={" SPECIALBUILD "},"

!undef EXENAME

!define APPNAME "SDLTradosStudio"				; complete name of program
!define APP "SDLTradosStudio"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE "SDLTradosStudio.exe"				; main exe name
!define APPEXE64 "SDLTradosStudio.exe"				; main exe 64 bit name
!define APPDIR "$EXEDIR"				; main exe relative path
!define APPSWITCH 	``
!define LANGUAGEROOT "HKLM"
!define LANGUAGESUB "SOFTWARE\Classes\MIME\Database\Rfc1766"
!define ADDINDIR "SDL plugin installer"


; **************************************************************************
; === Best Compression ===
; **************************************************************************
!ifndef NSIS_UNICODE
	Unicode true
!endif
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
OutFile ".\${APP}Launcher${VER}.exe"
Icon ".\${APP}2019.ico"
SilentInstall silent

; **************************************************************************
; === Set version information ===
; **************************************************************************
Caption "${PRODUCTNAME}"
VIProductVersion "${VER}"
VIAddVersionKey ProductName "${PRODUCTNAME}"
VIAddVersionKey Comments "${COMMENTS}"
VIAddVersionKey CompanyName "${COMPANYNAME}"
VIAddVersionKey LegalCopyright "${LEGALCOPYRIGHT}"
VIAddVersionKey FileDescription "${FILEDESCRIPTION}"
VIAddVersionKey FileVersion "${FILEVERSION}"
VIAddVersionKey ProductVersion "${PRODUCTVERSION}"
VIAddVersionKey InternalName "${ORIGINALFILENAME}"
VIAddVersionKey LegalTrademarks "${LEGALTRADEMARKS}"
VIAddVersionKey OriginalFilename "${ORIGINALFILENAME}"

; **************************************************************************
; === Other Actions ===
; **************************************************************************

Var AppID
Var homedir
Var SdlFindOutput
Var SdlFindPath
Var SdlFindPid

LangString Message1 1033 "${APPEXE} is running. Path $SdlFindPath. $\r$\n\
							Please close the program, then run again. "
LangString Message1 2052 "${APPEXE} 正在运行。路径 $SdlFindPath。$\r$\n\
							请退出正在运行的程序后，再次运行。"
LangString Message2 1033 "${APPEXE} is running. Some files need to be copied. $\r$\n\
							Please close the program, then run again. "
LangString Message2 2052 "${APPEXE} 正在运行。需要拷贝文件。$\r$\n\
							请退出正在运行的程序后，再次运行。"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	ReadEnvStr $homedir SDLHOME
	${IfThen} $homedir == "" ${|} StrCpy $homedir "${APPDIR}" ${|}
	; MessageBox MB_OK "$homedir"
	StrCpy $AppID "Portable"
	${SetEnvironmentVariable} SDLHOME $homedir

	${EnumProcessPaths} "SdlFind" $SdlFindOutput
	${GetProcessPath} "${APPEXE}" $0
	; MessageBox MB_OK "$0"

	${If} $0 != "$homedir\${APPEXE}"
		SetOverwrite on
		SetOutPath "${APPDIR}"
		File /r "${USERDIR}\*.dll"
	${Else}
		MessageBox MB_OK "$(Message2)"
		Abort
	${EndIf}

	${Select} $LANGUAGE
	${Case} 2052
		SetOverwrite on
		SetOutPath "${APPDIR}\..\..\${ADDINDIR}"
		File /r "${USERDIR}\zh-CN\*.*"
	${CaseElse}
		SetOverwrite on
		SetOutPath "${APPDIR}\..\..\${ADDINDIR}"
		File /r "${USERDIR}\en-US\*.*"
	${EndSelect}

	${GetParameters} $0
	${If} $0 == ""
		Exec "$homedir\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		Exec '"$homedir\${APPEXE}" $0'
		; ExecDos::exec /ASYNC /TOSTACK '"$homedir\${APPEXE}" "$0"' '' ''
	${EndIf}

SectionEnd

	Function "SdlFind"
		Pop $SdlFindPath			; matching path string
		Pop $SdlFindPid			; matching process PID
		; MessageBox MB_OK "$SdlFindPath; $SdlFindPid"
		${If} $SdlFindPath == "$homedir\${APPEXE}"
/* 			MessageBox MB_YESNO "$(Message1)" IDYES SdlFindYes1 IDNO 0
			Abort
			SdlFindYes1:
			MessageBox MB_OK "$SdlFindPid; $PslSpoonFileName"
			${TerminateProcess} "$SdlFindPid" $R1
			${ProcessWaitClose} "$PslSpoonFileName" "10000" $R1 */
			MessageBox MB_OK "$(Message1)"
			Push 0
			Abort
		${EndIf}
		Push 1			; must return 1 on the stack to continue
							; must return some value or corrupt the stack
							; DO NOT save data in $0-$9
	FunctionEnd

