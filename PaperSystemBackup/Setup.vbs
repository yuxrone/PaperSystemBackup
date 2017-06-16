Option Explicit
	Dim objFSO
	Dim strPS1File1
	Dim strPS1File2
	Dim strVBSFile
	DIm strINIFile
	Dim strFolder
	Dim objFolder
	Dim objShell

	strFolder = "C:\Paper_backup\"
	strPS1File1 = "ArchiveRun1.ps1"
	strPS1File2 = "FileZipper.ps1"
	strINIFile = "ZipArchive.ini"
	strVBSFile = "main.vbs"

	' Create III Folder If Not Existiong
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FolderExists(strFolder) Then
		Set objFolder = objFSO.CreateFolder(strFolder)
	End If

	objFSO.CopyFile strPS1File1, strFolder, True
	objFSO.CopyFile strPS1File2, strFolder, True
	objFSO.CopyFile strINIFile, strFolder, True
	objFSO.CopyFile strVBSFile, strFolder, True
	
	Set objShell = CreateObject("Wscript.Shell")
	objShell.Run "SchTasks /Delete /TN PaperBackup /F"
	objShell.Run "SchTasks /Create /TN PaperBackup /XML ""./PaperBackup.xml"""
	Set objFSO = Nothing
	Set objShell = Nothing
