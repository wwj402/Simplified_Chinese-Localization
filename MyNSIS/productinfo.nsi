; **************************************************************************
; === Define constants ===
; **************************************************************************
!define VER 		"2019.04.05.0"					; version of launcher
!define APPNAME 	"ProductInfo"					; complete name of program
!define APP 		"ProductInfo"					; short name of program without space and accent  this one is used for the final executable an in the directory structure
!define APPEXE 		"ProductInfo.exe"				; main exe name
!define APPEXE64 	"ProductInfo.exe"				; main exe 64 bit name
!define APPDIR 		"$EXEDIR"					; main exe relative path
!define APPSWITCH 	``
!define INIFILE "$EXEDIR\ProductInfo.ini"
!define RESOURCEHACKER "${NSISDIR}\Packhdr\resource_hacker\ResourceHacker.exe"
!define RHSWITCH `-open "$workingfile" -save "$EXEDIR\$productname" -action extract -mask ICONGROUP,, -log CON`

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
; !include "ProcFunc.nsh"
!include "WordFunc.nsh"


; **************************************************************************
; === Set basic information ===
; **************************************************************************
Name "${APP} Launcher"
OutFile ".\${APP}.exe"
; Icon ".\${APP}.ico"
SilentInstall silent

; **************************************************************************
; === Set version information ===
; **************************************************************************
Caption "${APPNAME} Launcher"
VIProductVersion "${VER}"
VIAddVersionKey ProductName "${APPNAME}"
VIAddVersionKey Comments "Get exe file info."
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

Var workingfile
Var productstring
Var productname

LangString PalMessage1 1033 "File info: "
LangString PalMessage1 2052 "文件信息："


; **************************************************************************
; ==== Running ====
; **************************************************************************

Section "Main"

	StrCpy $productstring ""
	${GetParameters} $workingfile
	${WordFind2x} "$workingfile" "'" "'" "+1" $0
	MessageBox MB_OK "$0"
	GetFullPathName $0 "$0"
	${If} $0 == ""
		${WordFind2x} '$workingfile' '"' '"' '+1' $0
		GetFullPathName $0 "$0"
		${If} $0 == ""
			${WordFind2x} "$workingfile" "$\`" "$\`" "+1" $0
		${EndIf}
	${EndIf}
	GetFullPathName $workingfile $0
	${GetBaseName} "$workingfile" $productname
	MessageBox MB_OK "$0, $productname"

	${If} $productname == ""
		StrCpy $productname "Info"
	${EndIf}
	Exec '"${RESOURCEHACKER}" ${RHSWITCH}'
	MoreInfo::GetComments "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringComments={$0}, "
	MoreInfo::GetCompanyName "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringCompanyName={$0}, "
	MoreInfo::GetFileDescription "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringFileDescription={$0}, "
	MoreInfo::GetFileVersion "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringFileVersion={$0}, "
	MoreInfo::GetLegalCopyright "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringLegalCopyright={$0}, "
	MoreInfo::GetLegalTrademarks "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringLegalTrademarks={$0}, "
	MoreInfo::GetOriginalFileName "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringOriginalFileName={$0}, "
	MoreInfo::GetPrivateBuild "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringPrivateBuild={$0}, "
	MoreInfo::GetProductName "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringProductName={$0}, "
	MoreInfo::GetProductVersion "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringProductVersion={$0}, "
	MoreInfo::GetSpecialBuild "$workingfile"
	Pop $0
	StrCpy $productstring "$productstringSpecialBuild={$0}, "
	MessageBox MB_OK "$productstring"
	IfFileExists "${INIFILE}" 0 +2
	Delete "${INIFILE}"
	WriteINIStr "${INIFILE}" "$productname" "Product" "$productstring"

SectionEnd

