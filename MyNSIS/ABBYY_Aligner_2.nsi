; **************************************************************************
; === Define constants ===
; **************************************************************************
!define EXEFULLDIR "d:\SnapShot\abbyya\Files\@PROGRAMFILESX86@\ABBYY Aligner 2.0"
!define EXENAME "Aligner.exe"
!define USERDIR "d:\SnapShot\abbyya\Patch"

!ifdef NSIS_UNICODE
	!define /file_version MAJOR "${EXEFULLDIR}\${EXENAME}" 0
	!define /file_version MINOR "${EXEFULLDIR}\${EXENAME}" 1
	!define /file_version OPTION "${EXEFULLDIR}\${EXENAME}" 2
	!define /file_version BUILD "${EXEFULLDIR}\${EXENAME}" 3
	!define VER ${MAJOR}.${MINOR}.${OPTION}.${BUILD}
	!undef MAJOR
	!undef MINOR
	!undef OPTION
	!undef BUILD
!else
	!echo "${NSIS_VERSION}"
	!getdllversion "${EXEFULLDIR}\${EXENAME}" Expv_
	!define VER "${Expv_1}.${Expv_2}.${Expv_3}.${Expv_4}"
!endif
!if ${VER} == "..."
	!undef VER
	!define VER "1.0.6.59"		; version of launcher
!endif

!execute '"ProductInfo.exe" "${EXEFULLDIR}\${EXENAME}"'
!searchparse /file "ProductInfo.ini" "Comments={" COMMENTS "}, CompanyName={" COMPANYNAME \
				"}, FileDescription={" FILEDESCRIPTION "}, FileVersion={" FILEVERSION "}, "
!searchparse /file "ProductInfo.ini" "LegalCopyright={" LEGALCOPYRIGHT "}, LegalTrademarks={" \
				LEGALTRADEMARKS "}, OriginalFileName={" ORIGINALFILENAME "}, PrivateBuild={" PRIVATEBUILD "}, "
!searchparse /file "ProductInfo.ini" "ProductName={" PRODUCTNAME "}, ProductVersion={" PRODUCTVERSION \
				"}, SpecialBuild={" SPECIALBUILD "},"

!undef EXENAME

!define APPNAME "ABBYY Aligner"				; complete name of program
!define APP "Aligner"						; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE "Aligner.exe"				; main exe name
!define APPEXE64 "Aligner.exe"				; main exe 64 bit name
!define APPDIR "$EXEDIR"				; main exe relative path
!define APPSWITCH 	``
!define LANGUAGEROOT "HKLM"
!define LANGUAGESUB "SOFTWARE\Classes\MIME\Database\Rfc1766"
!define SDLREGROOT "HKCU"
!define PATCHDIR "Patch"


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
; !include "ForEachPath.nsh"
!include "FileFunc.nsh"
!include "ProcFunc.nsh"
!include "WordFunc.nsh"
; !include "Registry.nsh"


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

Var AppFindOutput
Var AppFindPath
Var AppFindPid

LangString Message1 1033 "${APPEXE} is running. Path $AppFindPath. $\r$\n\
							Please close the program, then run again. "
LangString Message1 2052 "${APPEXE} 正在运行。路径 $AppFindPath。$\r$\n\
							请退出正在运行的程序后，再次运行。"
LangString Message2 1033 "${APPEXE} is running. Some files need to be copied. $\r$\n\
							Please close the program, then run again. "
LangString Message2 2052 "${APPEXE} 正在运行。需要拷贝文件。$\r$\n\
							请退出正在运行的程序后，再次运行。"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

/* 
	IntFmt $0 "0x%04x" "$languagevalue"
	System::Call "kernel32::LCIDToLocaleName(i r0, t .r1, i 0, l 0) i .r2"
	System::Call "kernel32::LCIDToLocaleName(i r0, t .r1, i r2, l 0) i .r2"
	MessageBox MB_OK "$0, $1, $2"
	${If} $2 == 0
		IntFmt $0 "%04x" "$languagevalue"
		ReadRegStr $1 ${LANGUAGEROOT} "${LANGUAGESUB}" "$0"
		${WordFind} "$1" ";" "+1" $1
		; MessageBox MB_OK "$languagevalue, $0, $1"
	${EndIf}
	StrCpy $uilanguage $1
 */
	; ${Locate} "${APPDIR}\${PATCHDIR}" "" "PatchApp"
	SetOverwrite ifdiff
	SetOutPath "${APPDIR}"
	File /nonfatal "${USERDIR}\*.dll"
	File /nonfatal "${USERDIR}\Language\*.*"

	${GetParameters} $0
	${If} $0 == ""
		Exec "${APPDIR}\${APPEXE}"
		; ExecDos::exec /ASYNC /TOSTACK '"${APPDIR}\${APPEXE}"' '' ''
	${Else}
		Exec '"${APPDIR}\${APPEXE}" $0'
		; ExecDos::exec /ASYNC /TOSTACK '"${APPDIR}\${APPEXE}" "$0"' '' ''
	${EndIf}

SectionEnd

	Function "AppPathFind"
		Pop $AppFindPath		; matching path string
		Pop $AppFindPid			; matching process PID
		; MessageBox MB_OK "$AppFindPath; $AppFindPid"
		${If} $AppFindPath == "${APPDIR}\${APPEXE}"
/* 			MessageBox MB_YESNO "$(Message1)" IDYES AppPathFindYes1 IDNO 0
			Abort
			AppPathFindYes1:
			MessageBox MB_OK "$AppFindPid; $PslSpoonFileName"
			${TerminateProcess} "$AppFindPid" $R1
			${ProcessWaitClose} "$PslSpoonFileName" "10000" $R1 */
			MessageBox MB_OK "$(Message1)"
			Push 0
			Abort
		${EndIf}
		Push 1			; must return 1 on the stack to continue
							; must return some value or corrupt the stack
							; DO NOT save data in $0-$9
	FunctionEnd

	Function "PatchApp"
		; $R9    "path\name"
		; $R8    "path"
		; $R7    "name"
		; $R6    "size"  ($R6="" if directory, $R6="0" if file with /S=)

		; $R0-$R5  are not used (save data in them).
		; ...
		; MessageBox MB_OK "R9=$R9$\r$\nR8=$R8$\r$\nR7=$R7$\r$\nR6=$R6$\r$\n"
		${If} $R6 == ""
			Nop
		${Else}
			md5dll::GetMD5File "$R9"
			Pop $0
			md5dll::GetMD5File "${APPDIR}\$R7"
			Pop $1
			${If} $0 != $1
				${EnumProcessPaths} "AppPathFind" $AppFindOutput
				${GetProcessPath} "${APPEXE}" $0
				; MessageBox MB_OK "$0"

				${If} $0 != "${APPDIR}\${APPEXE}"
					CopyFiles /SILENT "$R9" "${APPDIR}"
				${Else}
					MessageBox MB_OK "$(Message2)"
					Abort
				${EndIf}
			${EndIf}
		${EndIf}

		Push $0    ; If $var="StopLocate" Then exit from function
	FunctionEnd
