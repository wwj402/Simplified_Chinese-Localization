; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"2018.10.09.0"					; version of launcher
!define APPNAME 	"MoreInfo"					; complete name of program
!define APP 		"MoreInfo"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"MoreInfo.exe"				; main exe name
!define APPEXE64 	"MoreInfo.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``


; **************************************************************************
; === Best Compression ===
; **************************************************************************
Unicode true
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
; !include "ProcFunc.nsh"


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
VIAddVersionKey Comments "${APPNAME} update your software."
VIAddVersionKey CompanyName "Home"
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

Var AppID
Var workingdir
Var homedir
Var funcdir
Var exe_ver
Var ini_ver

LangString PalMessage1 1033 "Version in ini:$ini_ver$\nVersion in exe:$exe_ver$\nReplace?"
LangString PalMessage1 2052 "ini文件中版本:$ini_ver$\nexe文件中版本:$exe_ver$\n是否替换？"


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	${GetParameters} $workingdir
	GetFullPathName $workingdir "$workingdir.."
	ReadEnvStr $homedir APPHOME
	${IfThen} $homedir == "" ${|} StrCpy $homedir "${APPDIR}" ${|}
	GetFullPathName $homedir "$homedir\.."
	; MessageBox MB_OK "||$homedir//$workingdir||"
	StrCpy $AppID "Portable"
	IfFileExists "$workingdir\App\AppInfo\appinfo.ini" 0 Noworkingdir
	StrCpy $funcdir "$workingdir"
	${Locate} "$workingdir" "/L=F /M=*Portable.exe  /G=0" "setdetails"
	ExecShell "open" "$workingdir\App\AppInfo\appinfo.ini"
	Noworkingdir:
	IfFileExists "$homedir\App\AppInfo\appinfo.ini" 0 Nohomedir
	StrCpy $funcdir "$homedir"
	${Locate} "$homedir" "/L=F /M=*Portable.exe  /G=0" "setdetails"
	ExecShell "open" "$workingdir\App\AppInfo\appinfo.ini"
	Nohomedir:

SectionEnd

Function setdetails

	; MessageBox MB_OK "$funcdir"
	WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Control Start $R7
	${GetBaseName} "$R9" $R1
	ReadINIStr $R2 "$funcdir\App\AppInfo\appinfo.ini" Details Path
	${If} $R2 != ""
		MoreInfo::GetFileVersion "$funcdir\App\$R2"
		Pop $exe_ver
		MoreInfo::GetCompanyName "$funcdir\App\$R2"
		Pop $R3
	${Else}
		${If} ${RunningX64}
			ReadINIStr $R2 "$funcdir\App\AppInfo\Launcher\$R1.ini" Launch ProgramExecutable64
			${If} $R2 == ""
				ReadINIStr $R2 "$funcdir\App\AppInfo\Launcher\$R1.ini" Launch ProgramExecutable
			${EndIf}
		${Else}
			ReadINIStr $R2 "$funcdir\App\AppInfo\Launcher\$R1.ini" Launch ProgramExecutable
		${EndIf}
		MoreInfo::GetFileVersion "$funcdir\App\$R2"
		Pop $exe_ver
		MoreInfo::GetCompanyName "$funcdir\App\$R2"
		Pop $R3
	${EndIf}

	MessageBox MB_OK "$exe_ver;$R3"
	${If} $R3 == ""
		MoreInfo::GetCompanyName "$R9"
		Pop $R3
		${If} $R3 == ""
			WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Details Publisher ""
		${Else}
			WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Details Publisher $R3
		${EndIf}
	${Else}
			WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Details Publisher $R3
	${EndIf}

	${If} $exe_ver != ""
		ReadINIStr $ini_ver "$funcdir\App\AppInfo\appinfo.ini" Version DisplayVersion
		${If} $exe_ver != $ini_ver
			MessageBox MB_YESNO "$(PalMessage1)" IDYES 0 IDNO VERNOCHANGE
			WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Version DisplayVersion $exe_ver
			WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Version PackageVersion $exe_ver
			VERNOCHANGE:
		${EndIf}
	${EndIf}
	WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Details AppID $R1
	MoreInfo::GetProductName "$R9"
	Pop $R1
	WriteINIStr "$funcdir\App\AppInfo\appinfo.ini" Details Name $R1
	; MessageBox MB_OK "$R7$\n$1$\n$exe_ver"
	StrCpy $R1 StopLocate
	Push $R1

FunctionEnd

