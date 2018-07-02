SetCompressor /SOLID lzma
SetCompressorDictSize 32
Unicode true

Name "Test Launcher"
OutFile ".\TestLauncher.exe"
; Icon ".\${APP}.ico"
SilentInstall silent

!include "FileFunc.nsh"
!define TESTSTR "..\..\jdk"

Section "Main"
MessageBox MB_OK "${TESTSTR}"
${GetRoot} "${TESTSTR}" $0
MessageBox MB_OK "$0,"
IfErrors 0 +2
MessageBox MB_OK "error"
MessageBox MB_OK "$0,"

SectionEnd