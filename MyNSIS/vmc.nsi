; **************************************************************************
; === Define constants ===
; **************************************************************************
!define EXEFULLDIR "d:\SnapShot\vmc\Files\@PROGRAMFILESX86@\VMware\VMware vCenter Converter Standalone"
!define EXENAME "converter.exe"
!define USERDIR "d:\SnapShot\vmc\Files\@PROGRAMFILESX86@\VMware\VMware vCenter Converter Standalone\drivers"
!define DEFAULTDIR "$PROGRAMFILES\VMware\VMware vCenter Converter Standalone"

!ifdef NSIS_UNICODE
	!echo "${NSIS_VERSION}"
	!define /file_version MAJOR "${EXEFULLDIR}\${EXENAME}" 0
	!define /file_version MINOR "${EXEFULLDIR}\${EXENAME}" 1
	!define /file_version OPTION "${EXEFULLDIR}\${EXENAME}" 2
	!define /file_version BUILD "${EXEFULLDIR}\${EXENAME}" 3
	!if ${MAJOR} == 0
		!define VER "18.0.109.0"		; version of launcher
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

!define APPNAME "VMC"				; complete name of program
!define APP "VMC"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE "converter.exe"				; main exe name
!define APPEXE64 "converter.exe"				; main exe 64 bit name
!define APPDIR "$EXEDIR"				; main exe relative path
!define APPSWITCH 	``
; !define JAVAHOME	"jre"
; !define JAVAHOME64	"jre64"
!define LANGUAGEROOT "HKLM"
!define LANGUAGESUB "SOFTWARE\Classes\MIME\Database\Rfc1766"
!define SERVICEROOT "HKLM"
!define SERVICESUB "SYSTEM\CurrentControlSet\Services"
!define SERVICEDELETE "DeleteFlag"
!define SERVICENAME1 "vstor2-mntapi10-shared"
!define SERVICEIMAGE1 "vstor2-mntapi10-shared.sys"
!define SERVICEDISPLAY1 "Vstor2 MntApi 1.0 Driver (shared)"
!define SERVICESTARTTYPE1 "0x00000003"
!define SERVICETYPE1 "0x00000001"
!define SERVICENAME2 "bmdrvr"
!define SERVICEIMAGE2 "bmdrvr.sys"
!define SERVICEDISPLAY2 "Modified Clusters Tracking Driver"
!define SERVICESTARTTYPE2 "0x00000003"
!define SERVICETYPE2 "0x00000001"
!define SERVICEDIR "$EXEDIR\drivers"
!define WAITFORCLOSE1 "converter.exe"
!define WAITFORCLOSE2 "vmware-converter.exe"


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
; !include "SetEnvironmentVariable.nsh"
; !include "ForEachPath.nsh"
!include "FileFunc.nsh"
!include "ProcFunc.nsh"
!include "WordFunc.nsh"
!include "servicelib.nsh"


; **************************************************************************
; === Set basic information ===
; **************************************************************************
Name "${APP} Launcher"
OutFile ".\${APP}Launcher${VER}.exe"
Icon ".\${APP}.ico"
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

Var serviceflag1
Var serviceflag2
Var servicename
Var procpath
Var procid
Var procnum

LangString Message1 1033 "Service $servicename install failed. "
LangString Message1 2052 "服务 $servicename 安装失败。"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${If} ${RunningX64}
        ${EnableX64FSRedirection}
	${EndIf}

	StrCpy $R0 1
	ClearErrors
	${Do}
		${Select} $R0
		${Case} 1
			${IfThen} ${SERVICENAME1} == "" ${|} ${ExitDo} ${|}
			!insertmacro SERVICE "status" "${SERVICENAME1}" ""
			Pop $0
            ; MessageBox MB_OK "$0"
			${If} $0 == "false"
                IfFileExists "$SYSDIR\drivers\${SERVICEIMAGE1}" +3 0
                SetOutPath $SYSDIR\drivers
				File "${USERDIR}\${SERVICEIMAGE1}"
				!insertmacro SERVICE "create" "${SERVICENAME1}" 'path=$WINDIR\SysWOW64\drivers\${SERVICEIMAGE1};\
                starttype=${SERVICESTARTTYPE1};servicetype=${SERVICETYPE1};display=${SERVICEDISPLAY1};'
                Pop $0
                ${If} $0 != true
                    StrCpy $servicename ${SERVICENAME1}
                    MessageBox MB_OK "$(Message1)"
                ${EndIf}
			${Else}
				StrCpy $serviceflag1 "exist"
			${EndIf}
			!insertmacro SERVICE "start" "${SERVICENAME1}" ""
		${Case} 2
			${IfThen} ${SERVICENAME2} == "" ${|} ${ExitDo} ${|}
			!insertmacro SERVICE "status" "${SERVICENAME2}" ""
			Pop $0
			; MessageBox MB_OK "$0"
			${If} $0 == "false"
                IfFileExists "$SYSDIR\drivers\${SERVICEIMAGE2}" +3 0
                SetOutPath $SYSDIR\drivers
				File "${USERDIR}\${SERVICEIMAGE2}"
				!insertmacro SERVICE "create" "${SERVICENAME2}" 'path=$WINDIR\SysWOW64\drivers\${SERVICEIMAGE2};\
                starttype=${SERVICESTARTTYPE2};servicetype=${SERVICETYPE2};display=${SERVICEDISPLAY2};'
                Pop $0
                ${If} $0 != true
                    StrCpy $servicename ${SERVICENAME2}
                    MessageBox MB_OK "$(Message1)"
                ${EndIf}
			${Else}
				StrCpy $serviceflag2 "exist"
			${EndIf}
			!insertmacro SERVICE "start" "${SERVICENAME2}" ""
		${Default}
			${ExitDo}
		${EndSelect}
		IntOp $R0 $R0 + 1
		; MessageBox MB_OK "$R0"
	${Loop}

	${If} ${RunningX64}
		${GetParameters} $0
		${If} $0 == ""
			Exec "${APPDIR}\${APPEXE64}"
		${Else}
			Exec '"${APPDIR}\${APPEXE64}" $0'
		${EndIf}
	${Else}
		${GetParameters} $0
		${If} $0 == ""
			${Execute} "${APPDIR}\${APPEXE}" "" $procid
			; Exec "${APPDIR}\${APPEXE}"
		${Else}
			${Execute} '"${APPDIR}\${APPEXE}" $0' "" $procid
			; Exec '"${APPDIR}\${APPEXE}" $0'
		${EndIf}
	${EndIf}

	Sleep 30000

	StrCpy $procnum "1"
	${Do}
		${EnumProcessPaths} "waitForClose" $0
		; MessageBox  MB_OK "$0"
		${If} $0 == 1
			IntOp $procnum $procnum - 1
		${EndIf}
	${LoopWhile} $procnum > 0

	${Do}
		${ProcessWaitClose} "${WAITFORCLOSE2}" "6000" $0
		; MessageBox MB_OK "${WAITFORCLOSE2}, $0"
	${LoopUntil} $0 <= 0

	StrCpy $R0 1
	ClearErrors
	${Do}
		${Select} $R0
		${Case} 1
			${IfThen} ${SERVICENAME1} == "" ${|} ${ExitDo} ${|}
			${If} $serviceflag1 != "exist"
				!insertmacro SERVICE "status" "${SERVICENAME1}" ""
				Pop $0
				; MessageBox MB_OK "$0"
				${If} $0 != "stopped"
        	        !insertmacro SERVICE "stop" "${SERVICENAME1}" ""
				${EndIf}
				!insertmacro SERVICE "delete" "${SERVICENAME1}" ""
				Delete /REBOOTOK "$WINDIR\SysWOW64\drivers\${SERVICEIMAGE1}"
			${EndIf}
		${Case} 2
			${IfThen} ${SERVICENAME2} == "" ${|} ${ExitDo} ${|}
			${If} $serviceflag2 != "exist"
				!insertmacro SERVICE "status" "${SERVICENAME2}" ""
				Pop $0
				; MessageBox MB_OK "$0"
				${If} $0 != "stopped"
        	        !insertmacro SERVICE "stop" "${SERVICENAME2}" ""
				${EndIf}
				!insertmacro SERVICE "delete" "${SERVICENAME2}" ""
				Delete /REBOOTOK "$WINDIR\SysWOW64\drivers\${SERVICEIMAGE2}"
			${EndIf}
		${Default}
			${ExitDo}
		${EndSelect}
		IntOp $R0 $R0 + 1
		; MessageBox MB_OK "$R0"
	${Loop}

SectionEnd

	Function "waitForClose"
		Pop $procpath			; matching path string
		Pop $procid				; matching process PID
		; MessageBox MB_OK "$procpath/$procid"
		${If} $procpath == "${DEFAULTDIR}\${WAITFORCLOSE1}"
			IntOp $procnum $procnum + 1
			${ProcessWaitClose} "$procid" "-1" $0
		${EndIf}
		Push 1			; must return 1 on the stack to continue
							; must return some value or corrupt the stack
							; DO NOT save data in $0-$9
	FunctionEnd