# |Info|
# Written by Bryan O'Connell, August 2013
# Purpose: Creates a .zip file of a file or folder.
#
# Sample: zipstuff.ps1 -target "C:\Projects\wsubi" -zip_to "C:\Users\Bryan\Desktop\wsubi" [-compression fast] [-timestamp] [-confirm]
#
# Params:
# -target: The file or folder you would like to zip.
#
# -zip_to: The location where the zip file will be created. If an old version
# exists, it will be deleted.
#
# -compression (optional): Sets the compression level for your zip file. Options:
# a. fast - Higher process speed, larger file size (default option).
# b. small - Slower process speed, smaller file size.
# c. none - Fastest process speed, largest file size.
#
# -add_timestamp (optional): Applies a timestamp to the .zip file name.
# By default, no timestamp is used.
#
# -confirm (optional): When provided, indicates that you would like to be
# prompted when the zip process is finished.
#
# |Info|

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true,Position=0)]
  [string]$target,

  [Parameter(Mandatory=$true,Position=1)]
  [string]$zip_to,

  [Parameter(Mandatory=$false,Position=2)]
  [ValidateSet("fast","small","none")]
  [string]$compression,

  [Parameter(Mandatory=$false,Position=3)]
  [bool]$timestamp,

  [Parameter(Mandatory=$false,Position=4)]
  [bool]$confirm
)

#-----------------------------------------------------------------------------#
function DeleteFileOrFolder
{ Param([string]$PathToItem)

  if (Test-Path $PathToItem)
  {
    Remove-Item ($PathToItem) -Force -Recurse;
  }
}

function DetermineCompressionLevel
{
  $CompressionToUse = $null;

  switch($compression)
  {
    "fast" {$CompressionToUse = [System.IO.Compression.CompressionLevel]::Fastest}
    "small" {$CompressionToUse = [System.IO.Compression.CompressionLevel]::Optimal}
    "none" {$CompressionToUse = [System.IO.Compression.CompressionLevel]::NoCompression}
    default {$CompressionToUse = [System.IO.Compression.CompressionLevel]::Fastest}
  }

  return $CompressionToUse;
}

#-----------------------------------------------------------------------------#
#Write-Output "Starting zip process...";
$returnMsg = "Starting zip process...<br/>"
if ((Get-Item $target).PSIsContainer)
{
  $zip_to = ($zip_to + "\" + (Split-Path $target -Leaf) + ".zip");
}
else{

  #So, the CreateFromDirectory function below will only operate on a $target
  #that's a Folder, which means some additional steps are needed to create a
  #new folder and move the target file into it before attempting the zip process. 
  $FileName = [System.IO.Path]::GetFileNameWithoutExtension($target);
  $NewFolderName = ($zip_to + "\" + $FileName)
  $FullFileName = [System.IO.Path]::GetFileNameWithoutExtension($target) + [System.IO.Path]::GetExtension($target);

  DeleteFileOrFolder($NewFolderName);

  md -Path $NewFolderName > $null;
  Copy-Item ($target) $NewFolderName;

  $file = Get-Item $target
  $fileSize = $file.length / 1mb -as [int]

  $target = $NewFolderName;
  $zip_to = $NewFolderName + ".zip";
}

$returnMsg += "[Source File]: " + $FullFileName + " <br/>";
$returnMsg += "[File Size]:   " + $fileSize + "MB <br/>";

DeleteFileOrFolder($zip_to);

if ($timestamp)
{
  $TimeInfo = New-Object System.Globalization.DateTimeFormatInfo;
  $CurrentTimestamp = Get-Date -Format $TimeInfo.SortableDateTimePattern;
  $CurrentTimestamp = $CurrentTimestamp.Replace(":", "-");
  $zip_to = $zip_to.Replace(".zip", ("-" + $CurrentTimestamp + ".zip"));
}

$FullFileName = [System.IO.Path]::GetFileNameWithoutExtension($zip_to)  + [System.IO.Path]::GetExtension($zip_to);

$returnMsg += "[Target File]:" + $FullFileName + "<br/>";

$Compression_Level = (DetermineCompressionLevel);
$IncludeBaseFolder = $false;

#add-type -assemblyName "System.ServiceProcess"
#[Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" );
#Add-Type -Path "C:\Program Files\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.IO.Compression.FileSystem.dll"
Add-Type -AssemblyName System.IO.Compression.FileSystem 
[System.IO.Compression.ZipFile]::CreateFromDirectory($target, $zip_to, $Compression_Level, $IncludeBaseFolder);

DeleteFileOrFolder($NewFolderName);

#Start-Sleep -s 1

$file = Get-Item $zip_to
$fileSize = $file.length / 1mb -as [int]
$returnMsg += "[File Size]:   " + $fileSize + "MB <br/>";

#Write-Output "Zip process complete.";
$returnMsg += "Zip process complete.<br/>";

if ($confirm)
{
  write-Output "Press any key to quit ...";
  $quit = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown");
}

return $returnMsg