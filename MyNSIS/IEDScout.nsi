; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"4.20.0464.0000"		; version of launcher
!define APPNAME 	"OMICRON IEDScout"  	; complete name of program
!define APP 		"IEDScout"				; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"IEDScout.exe"			; main exe name
!define APPEXE64 	"IEDScout.exe"			; main exe 64 bit name
!define APPDIR 		"$EXEDIR"	            ; main exe relative path
!define APPSWITCH 	``

!define SERVICENAME1 "NPF"
!define SERVICEIMAGE1 "$SYSDIR\drivers\npf.sys"
!define SERVICEDISPLAY1 "NetGroup Packet Filter Driver"
!define SERVICESTARTTYPE1 "0x00000003"
!define SERVICETYPE1 "0x00000001"
!define SERVICEVERSION1 "4.1.0.2980"

; **************************************************************************
; === Best Compression ===
; **************************************************************************
; Unicode true
SetCompressor /SOLID lzma
SetCompressorDictSize 32
; RequestExecutionLevel admin

; **************************************************************************
; === Includes ===
; **************************************************************************

!include "LogicLib.nsh"
!include "x64.nsh"
!include "servicelib.nsh"

!include "FileFunc.nsh"
!include "textfunc.nsh"
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
VIAddVersionKey Comments "${APPNAME} launcher"
VIAddVersionKey CompanyName "OMICRON electronics GmbH"
VIAddVersionKey LegalCopyright "© 2019 OMICRON"
VIAddVersionKey FileDescription "${APPNAME}"
VIAddVersionKey FileVersion "${VER}"
VIAddVersionKey ProductVersion "${VER}"
VIAddVersionKey InternalName "${APPEXE}"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey OriginalFilename "${APP}Launcher.exe"

; **************************************************************************
; === Other Actions ===
; **************************************************************************

Var service_status
Var service_path
Var used_exe

LangString Message1 1033 "Need install winpcap driver.$\r$\nYes, install;$\r$\nNo, skip install."
LangString Message1 2052 "需要安装 winpcap 驱动。$\r$\n是，安装；$\r$\n否，跳过。"
LangString Message2 1033 "Had installed winpcap driver.$\r$\nYes, uninstall;$\r$\nNo, skip uninstall."
LangString Message2 2052 "已经安装 winpcap 驱动。$\r$\n是，卸载；$\r$\n否，跳过。"

; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	!insertmacro SERVICE "status" "${SERVICENAME1}" ""
	Pop $0

	${IF} $0 != "false"
        ${If} ${RunningX64}
            ${DisableX64FSRedirection}
        ${EndIf}
		ReadRegStr $service_path HKLM SYSTEM\CurrentControlSet\services\${SERVICENAME1} ImagePath
		${IF} $service_path != ""
            Push $service_path
            Call ServiceImagePath
            Pop $1
            ; MessageBox MB_OK "$1"
            ${If} ${FileExists} "$1"
		        ReadRegStr $2 HKLM SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WinPcapInst DisplayVersion
                ; MessageBox MB_OK "$2"
                ${If} $2 != ${SERVICEVERSION1}
                    MessageBox MB_YESNO "$(Message1)" IDNO +2
                    ExecDos::exec /TOSTACK '"$EXEDIR\winpcap.exe"' '' ''
                ${Else}
                    StrCpy $service_status "exist"
                ${EndIf}
            ${Else}
                MessageBox MB_YESNO "$(Message1)" IDNO +2
                ExecDos::exec /TOSTACK '"$EXEDIR\winpcap.exe"' '' ''
            ${EndIf}
        ${Else}
            MessageBox MB_YESNO "$(Message1)" IDNO +2
            ExecDos::exec /TOSTACK '"$EXEDIR\winpcap.exe"' '' ''
        ${EndIf}
    ${Else}
        ExecDos::exec /TOSTACK '"$EXEDIR\winpcap.exe"' '' ''
    ${EndIf}

    ${If} ${RunningX64}
        StrCpy $used_exe "$EXEDIR\${APPEXE64}"
    ${Else}
        StrCpy $used_exe "$EXEDIR\${APPEXE}"
    ${EndIf}
    ; MessageBox MB_OK "$used_exe"
	${GetParameters} $0
	${If} $0 == ""
		; ExecWait "$used_exe"
		ExecDos::exec /TOSTACK '"$used_exe"' '' ''
	${Else}
		; ExecWait '"$used_exe" "$0"'
		ExecDos::exec /TOSTACK '"$used_exe" "$0"' '' ''
	${EndIf}

    ${If} $service_status != "exist"
        MessageBox MB_YESNO "$(Message2)" IDNO SKIPUNISTALL
        ReadRegStr $0  HKLM SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WinPcapInst UninstallString
		ExecDos::exec /TOSTACK '"$0"' '' ''
        SKIPUNISTALL:
    ${EndIf}

SectionEnd

Function "ServiceImagePath"

	Exch $0
	Push $1
	Push $2
	ExpandEnvStrings  $0 "$0"
	${GetFileName} $SYSDIR $1
	; MessageBox MB_OK "$0$\n$1"
	${WordFind} "$0" "\" "+1" $2
	${If} $1 == $2
		StrCpy $0 "$WINDIR\$0"
	${EndIf}
	${WordReplace} "$0" "\systemroot" "$WINDIR" "+*" $0
	${WordReplace} "$0" "\??\" "" "+1*" $0

	; MessageBox MB_OK "$0"
	Pop $2
	Pop $1
	Exch $0

FunctionEnd