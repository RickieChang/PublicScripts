# BarracudaExportEml.ps1 2015-10-07 Rickie Chang
#
# Problem: Decommissioned Barracuda Message Archiver 650 required export to eml format.
# The barracuda allows you to share out the raw archives via the
# Advanced->Backup->[Backup Archives via SMB] option and sets an "smb" user password
# Mount that location to a drive letter, in my case Z:\ and you will see a hierarchy of
# numbered and hexadecimal lettered subfolders. In my case, there were top level folders
# under one main "1" folder. I used these top level folders as groupings for the extract
# so I wouldn't get a million e-mails in a single folder in my R:\bma-smb folder. This
# script attempts to extract every single file from the share and renames the extracted
# file with an eml extension. If it is unable to extract, it copies and renames the file
# to eml format.
#
# Depending on your version of the Barracuda appliance that may be multiple files
# within archives which this script won't account for. I created it for my use case. If
# you have the other format, please refer to Charles Stemaly's blog which I started from:
# http://itstuffilearnedtoday.blogspot.com/2014/02/getting-your-email-out-of-barracuda.html
#
# Prereq: Install 7-Zip in the default location (use your OS architecture version)

# Mapped folder from bma-smb share
$Source = "Z:\1"

# Folder to extract eml files to. Make sure you have enough space for all expanded files
$Out = "R:\bma-smb"

$folderstoprocess = gci $Source

foreach ($folder in $folderstoprocess) {
	Write-Host "Processing" $folder.FullName
	$foldertocreate = $Out + "\" + $folder.Name
	Write-Host "Creating" $foldertocreate
	New-Item $foldertocreate -type directory | out-null

	gci $folder.FullName *.* -rec | where { $_.GetType().Name -eq "FileInfo" } | %{
		Write-Host "Processing" $_.FullName

		$filename = ($_.Name -split "\.",2)[0]
		$filename = $foldertocreate + "\" + $filename
		$filenametobe = $filename + ".eml"

		[Array]$7zparams = "x", $_.FullName, "-o$foldertocreate"

		&'C:\Progra~1\7-Zip\7z.exe' $7zparams | out-null

		If ($LastExitCode -eq 2) {
			Copy-Item $_.FullName $filenametobe
		}

		If ($LastExitCode -eq 0) {
			Rename-Item $filename $filenametobe
		}
	}
}
