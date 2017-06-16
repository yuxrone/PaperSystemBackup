set-executionpolicy -executionpolicy bypass
# Load assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function Parse-IniFile ($file) {
  $ini = @{}

 # Create a default section if none exist in the file. Like a java prop file.
 $section = "NO_SECTION"
 $ini[$section] = @{}

  switch -regex -file $file {
    "^\[(.+)\]$" {
      $section = $matches[1].Trim()
      $ini[$section] = @{}
    }
    "^\s*([^#].+?)\s*=\s*(.*)" {
      $name,$value = $matches[1..2]
      # skip comments that start with semicolon:
      if (!($name.StartsWith(";"))) {
        $ini[$section][$name] = $value.Trim()
      }
    }
  }
  $ini
}

#Set-Location $PWD
#Set-Location "H:\Docs\Work\++ Projects\Paper System\"

Set-Location "c:\paper_backup\"
$file = "ZipArchive.ini" 
$zipArchive = Parse-IniFile ($file)

#$oReturn=[System.Windows.Forms.Messagebox]::Show($zipArchive["sourceFile"]["sourceFile1"])


$now = Get-Date -format "yyyy-MM-dd HH:mm:ss"
$sourceFile1 = $zipArchive["sourceFile"]["sourceFile1"]
$sourceFile2 = $zipArchive["sourceFile"]["sourceFile2"]
$destinationFolder = $zipArchive["PWD"]["destinationFolder"]
$workingFolder = $zipArchive["PWD"]["workingDirectory"]
$tempZipFolder = $workingFolder + "TEMP\"

# email notification
$email = New-Object System.Net.Mail.MailMessage
$email.From = $zipArchive["email"]["emailFrom"]
$email.To.Add($zipArchive["email"]["emailTo"])
$email.IsBodyHtml = $true
$email.Priority = [System.Net.Mail.MailPriority]::High
$email.Subject = $zipArchive["email"]["emailSubject"]
$smtp = New-Object System.Net.Mail.SmtpClient
$smtp.Host = $zipArchive["email"]["smtpHost"]

#Set-Location $PWD
#Set-Location "H:\Docs\Work\++ Projects\Paper System\"
Set-Location $workingFolder

$email.Body = "Paper System Backup Log @ :" + $now + "<br/>"
$email.Body += "<br/>---------------------------------------------------------<br/>"
$email.Body += . .\FileZipper.ps1 -target $sourceFile1  -zip_to $tempZipFolder -compression fast -timestamp 1
$email.Body += "<br/>---------------------------------------------------------<br/>"
$email.Body += . .\FileZipper.ps1 -target $sourceFile2  -zip_to $tempZipFolder -compression fast -timestamp 1


#$oReturn=[System.Windows.Forms.Messagebox]::Show($email.Body)

Start-Sleep -s 1

#Clean up expired archive files
$email.Body += "<br/>---------------------------------------------------------<br/>"
$email.Body += "Transfering archive file to server...<br/>"
Get-ChildItem $tempZipFolder | Where-Object { $_ -is [System.IO.FileInfo] -and $_.extension -eq ".zip" -and $_.name -like "*$pattern*"} | ForEach-Object {
$destinationFile = $destinationFolder + $_.name
Move-Item $_.FullName -Destination $destinationFile
$email.Body += $_.Name + " has been archived. <br/>"
}

$email.Body += "<br/>---------------------------------------------------------<br/>"
$email.Body += "Clean up expired archive files.<br/>"

$now = Get-Date
$expireDays = $zipArchive["system"]["archiveExpirationDays"]
$lastWrite = $now.AddDays(-$expireDays)
$count=0

Set-Location $destinationFolder
$pattern = $zipArchive["system"]["fileNamePattern"]

Get-ChildItem $destinationFolder | Where-Object { $_ -is [System.IO.FileInfo] } | ForEach-Object {
	If ($_.LastWriteTime -lt $lastWrite -and $_.extension -eq ".zip" -and $_.name -like "*$pattern*")
	{
        Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue
        $email.Body += $_.Name + " has been REMOVED. <br/>"
        $count++
	}
}

If (!$count) 
{
    $email.Body += "No exired archive file found. <br/>"
}

$smtp.Send($email)

