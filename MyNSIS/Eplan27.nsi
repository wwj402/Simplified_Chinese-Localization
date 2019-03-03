; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"2.7.3.11746"		; version of launcher
!define APPNAME 	"EPLAN Electric P8"	; complete name of program
!define APP 		"EPLAN"				; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"EPLAN.exe"			; main exe name
!define APPEXE64 	"EPLAN.exe"			; main exe 64 bit name
!define APPDIR 		"$EXEDIR\..\Platform\2.7.3\Bin"	; main exe relative path
!define APPSWITCH 	`/Variant:"Electric P8"`

!define SERVICENAME1 "hardlock"
!define SERVICEIMAGE1 "hardlock.sys"
!define SERVICEDISPLAY1 "hardlock"
!define SERVICESTARTTYPE1 "0x00000002"
!define SERVICETYPE1 "0x00000001"
!define SERVICENAME2 "akshasp"
!define SERVICEIMAGE2 "akshasp.sys"
!define SERVICEDISPLAY2 "SafeNet Inc. HASP Key"
!define SERVICESTARTTYPE2 "0x00000003"
!define SERVICETYPE2 "0x00000001"
!define SERVICENAME3 "aksusb"
!define SERVICEIMAGE3 "aksusb.sys"
!define SERVICEDISPLAY3 "SafeNet Inc. USB Key"
!define SERVICESTARTTYPE3 "0x00000003"
!define SERVICETYPE3 "0x00000001"
!define HASPCMDEXE "HaspDriver\haspdinst.exe"
!define HASPSETUPEXE "HaspDriver\HASPUserSetup.exe"
!define HASPIDFILE "$SYSDIR\drivers\hardlock.sys"
!define HASPVERSION "3.91"

!define INFDIR "$EXEDIR\MultiKey32"
!define INFDIR64 "$EXEDIR\MultiKey64"
!define DEVEXEDIR "$EXEDIR\devcon"
!define DEVEXE "devcon_x86.exe"
!define DEVEXE64 "devcon_x64.exe"

!define DEVNAME1 "multikey"
!define DEVINF1 "multikey.inf"
!define DEVINFHWID1 "root\multikey"
!define DEVSERVERIMG1 "multikey.sys"
!define MULTIKEYIDFILE "$SYSDIR\drivers\multikey.sys"

!define MULTIKEYREG "HKLM\SYSTEM\CurrentControlSet\MultiKey"
!define MULTIKEYREGBAK "HKCU\Software\PortableApps.com\Keys\MultiKey"
!define MULTIKEYDIR "EPLAN\Common"
!define MULTIKEYDIRBAK "$EXEDIR\Eplankey"
!define KEYFILENAME "SN-U10066"

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
!include "Registry.nsh"


; **************************************************************************
; === Set basic information ===
; **************************************************************************
Name "${APP} Launcher"
OutFile ".\${APP}2.7.3Launcher.exe"
Icon ".\${APP}.ico"
SilentInstall silent

; **************************************************************************
; === Set version information ===
; **************************************************************************
Caption "${APPNAME} Launcher"
VIProductVersion "${VER}"
VIAddVersionKey ProductName "${APPNAME}"
VIAddVersionKey Comments "${APPNAME} launcher"
VIAddVersionKey CompanyName "EPLAN Software & Service GmbH & Co. KG"
VIAddVersionKey LegalCopyright "Copyright © EPLAN S&S GmbH & Co KG 2001"
VIAddVersionKey FileDescription "${APPNAME}"
VIAddVersionKey FileVersion "${VER}"
VIAddVersionKey ProductVersion "${VER}"
VIAddVersionKey InternalName "${APPEXE}"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey OriginalFilename "${APP}Launcher.exe"

; **************************************************************************
; === Other Actions ===
; **************************************************************************

Var useddevexe
Var usedinfdir
Var keydir
Var keyflag
Var drivername
LangString Message1 1033 "Need install HASP driver.$\r$\nYes, silent install;$\r$\nNo, GUI install."
LangString Message1 2052 "需要安装 HASP 驱动。$\r$\n是，静默安装；$\r$\n否，GUI 安装。"
LangString Message2 1033 "Fail to install HASP driver, Please manually install."
LangString Message2 2052 "安装 HASP 驱动失败，请手动安装。"
LangString Message3 1033 "Fail to backup, exit program.$\r$\nYes, exit;$\r$\nNo, continue."
LangString Message3 2052 "备份失败，退出程序。$\r$\n是，退出；$\r$\n否，继续。"
LangString Message4 1033 "Fail to restore key, Please manually check."
LangString Message4 2052 "恢复 key 备份失败，请手动检查。"
LangString Message5 1033 "Key installed, overwrite.$\r$\nYes, overwrite;$\r$\nNo, skip."
LangString Message5 2052 "Key 已安装，是否覆盖。$\r$\n是，覆盖；$\r$\n否，跳过。"
LangString Message6 1033 "$drivername driver version differnt, reinstall.$\r$\nYes, reinstall;$\r$\nNo, skip."
LangString Message6 2052 "$drivername 驱动版本不同，重新安装。$\r$\n是，重新安装；$\r$\n否，跳过安装。"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	StrCpy $R0 1
	ClearErrors
	${Do}
		${Select} $R0
		${Case} 1
			${IfThen} ${SERVICENAME1} == "" ${|} ${ExitDo} ${|}
			!insertmacro SERVICE "status" "${SERVICENAME1}" ""
			Pop $0
			${If} $0 == "unknown"
				Call haspinstall
			${Else}
				Call haspreinstall
			${EndIf}
		${Case} 2
			${IfThen} ${SERVICENAME2} == "" ${|} ${ExitDo} ${|}
			!insertmacro SERVICE "status" "${SERVICENAME2}" ""
			Pop $0
			${If} $0 == "unknown"
				Call haspinstall
			${EndIf}
		${Case} 3
			${IfThen} ${SERVICENAME3} == "" ${|} ${ExitDo} ${|}
			!insertmacro SERVICE "status" "${SERVICENAME3}" ""
			Pop $0
			${If} $0 == "unknown"
				Call haspinstall
			${EndIf}
		${Default}
			${ExitDo}
		${EndSelect}
		IntOp $R0 $R0 + 1
		; MessageBox MB_OK "$R0"
	${Loop}

	${If} ${RunningX64}
		StrCpy $useddevexe "${DEVEXEDIR}\${DEVEXE64}"
		StrCpy $usedinfdir "${INFDIR64}"
	${Else}
		StrCpy $useddevexe "${DEVEXEDIR}\${DEVEXE}"
		StrCpy $usedinfdir "${INFDIR}"
	${EndIf}

	ReadEnvStr $0 "PUBLIC"
	${If} $0 != ""
		StrCpy $keydir "$0\${MULTIKEYDIR}"
	${Else}
		ReadEnvStr $0 "ALLUSERSPROFILE"
		StrCpy $keydir "$0\${MULTIKEYDIR}"
	${EndIf}
	; MessageBox MB_OK "$0\${MULTIKEYDIR}, ${MULTIKEYIDFILE}"

	IfFileExists "$keydir\${KEYFILENAME}.EGF" 0 NOKEY
	; md5dll::GetMD5File "$keydir\${KEYFILENAME}.EGF"
	; Pop $0
	; md5dll::GetMD5File "${MULTIKEYDIRBAK}\${KEYFILENAME}.EGF"
	; Pop $1
	; ${If} $0 != $1
		MessageBox MB_YESNO $(Message5) IDNO SKIPKEY
		Call eplanbackup
		SetOutPath "$keydir"
		File /r "d:\PortableApps\PortableApps.comLauncher\Data\MyNSIS\Eplankey\${KEYFILENAME}.*"
		StrCpy $keyflag "update"
		ExecWait 'regedit.exe /s "$keydir\${KEYFILENAME}.REG"'
		Call mulitykeyrestart
		SKIPKEY:
	; ${EndIf}
	Goto KEYEND
	NOKEY:
	Call eplanbackup
	SetOutPath "$keydir"
	File /r "d:\PortableApps\PortableApps.comLauncher\Data\MyNSIS\Eplankey\${KEYFILENAME}.*"
	StrCpy $keyflag "copy"
	ExecWait 'regedit.exe /s "$keydir\${KEYFILENAME}.REG"'
	Call mulitykeyrestart
	KEYEND:

	!insertmacro SERVICE "status" "${DEVNAME1}" ""
	Pop $0
	${If} $0 == "unknown"
		Call mulitykeyinstall
	${Else}
		Call mulitykeyreinstall
	${EndIf}


	${GetParameters} $0
	${If} $0 == ""
		ExecWait '"${APPDIR}\${APPEXE}" ${APPSWITCH}'
		; ExecDos::exec /TOSTACK '"$homedir\${APPEXE}"' '' ''
	${Else}
		ExecWait '"${APPDIR}\${APPEXE}" ${APPSWITCH} $0'
		; ExecDos::exec /TOSTACK '"$homedir\${APPEXE}" "$0"' '' ''
	${EndIf}

	${If} $keyflag != ""
		Delete "$keydir\${KEYFILENAME}.REG"
		Delete "$keydir\${KEYFILENAME}.EGF"
		${If} $keyflag != "copy"
			${Registry::DeleteKey} "${MULTIKEYREGBAK}" $0
		${EndIf}
	${EndIf}
	Call eplanrestore

SectionEnd


Function "haspinstall"
	MessageBox MB_YESNO "$(Message1)" IDYES 0 IDNO RUNSETUP
	ExecDos::exec /TOSTACK '"$EXEDIR\${HASPCMDEXE}" -i -cm' '' ''
	Pop $0
	${If} $0 != 0
		MessageBox MB_OK "$(Message2)"
		ExecShell open "devmgmt.msc"
	${EndIf}
	Goto HASPEND
	RUNSETUP:
	ExecWait "$EXEDIR\${HASPSETUPEXE}"
	HASPEND:
FunctionEnd

Function "haspuninstall"
	ExecDos::exec /TOSTACK '"$EXEDIR\${HASPCMDEXE}" -r -cm' '' ''
FunctionEnd

Function "haspreinstall"
	MoreInfo::GetFileVersion "${HASPIDFILE}"
	Pop $0
	${VersionCompare} "$0" "${HASPVERSION}" $1
	; MessageBox MB_OK "$0, $1, ${HASPIDFILE}"
	${If} $1 != 0
		StrCpy $drivername "HASP"
		MessageBox MB_YESNO "$(Message6)" IDYES 0 IDNO HASPSKIP
		Call haspuninstall
		Call haspinstall
		HASPSKIP:
	${EndIf}
FunctionEnd

Function "mulitykeyinstall"
	ExecDos::exec /TOSTACK '"$useddevexe" install "$usedinfdir\${DEVINF1}" root\multikey' '' ''
	ExecDos::exec /TOSTACK '"$useddevexe" rescan' '' ''
	ExecDos::exec /TOSTACK '"$useddevexe" update "$usedinfdir\${DEVINF1}" root\multikey' '' ''
FunctionEnd

Function "mulitykeyrestart"
	ExecDos::exec /TOSTACK '"$useddevexe" restart root\multikey' '' ''
FunctionEnd

Function "mulitykeyuninstall"
	ExecDos::exec /TOSTACK '"$useddevexe" remove root\multikey' '' ''
	ExecDos::exec /TOSTACK '"$useddevexe" rescan' '' ''
	IfFileExists "${MULTIKEYIDFILE}" 0 +2
	Delete "${MULTIKEYIDFILE}"
FunctionEnd

Function "mulitykeyreinstall"
	md5dll::GetMD5File "${MULTIKEYIDFILE}"
	Pop $0
	md5dll::GetMD5File "$usedinfdir\${DEVSERVERIMG1}"
	Pop $1
	${If} $0 != $1
		StrCpy $drivername "Mulitykey"
		MessageBox MB_YESNO "$(Message6)" IDYES 0 IDNO MULITYSKIP
		Call mulitykeyuninstall
		Call mulitykeyinstall
		MULITYSKIP:
	${EndIf}
FunctionEnd

Function "eplanbackup"
	${Locate} "$keydir" "/L=F /M=*.EGF" "fileopertion"
	${registry::KeyExists} "${MULTIKEYREG}" $0
	${If} $0 == 0
		${registry::KeyExists} "${MULTIKEYREGBAK}" $0
		${If} $0 != 0
			${Registry::MoveKey} "${MULTIKEYREG}" "${MULTIKEYREGBAK}" $0
			${If} $0 != 0
				MessageBox MB_YESNO "$(Message3), key" IDYES BACKUPEXIT IDNO BACKUPCONTINUE
				BACKUPEXIT:
				Quit
				BACKUPCONTINUE:
			${EndIf}
		${EndIf}
	${EndIf}

FunctionEnd

Function "eplanrestore"
	${Locate} "$keydir" "/L=F /M=*.BAK" "fileopertion"
	${registry::KeyExists} "${MULTIKEYREGBAK}" $0
	${If} $0 == 0
		${Registry::MoveKey} "${MULTIKEYREGBAK}" "${MULTIKEYREG}" $0
		${If} $0 != 0
			MessageBox MB_OK "$(Message4)"
		${EndIf}
	${EndIf}
FunctionEnd

Function "fileopertion"
	; $R9    "path\name"
	; $R8    "path"
	; $R7    "name"
	; $R6    "size"  ($R6="" if directory, $R6="0" if file with /S=)

	; $R0-$R5  are not used (save data in them).
	; ...
	${GetFileExt} "$R9" $0
	${If} $0 == "EGF"
		Rename "$R9" "$R8\$R7.bak"
	${Else}
		StrLen $1 "$0"
		IntOp $1 $1 + 1
		StrCpy $0 $R9 -$1
		Rename "$R9" "$0"
	${EndIf}

	Push ""    ; If $var="StopLocate" Then exit from function
FunctionEnd
