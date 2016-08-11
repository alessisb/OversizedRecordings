<#
================================================================================
================================================================================
		   File:  Get-OversizedRecordings.ps1
		 Author:  Spencer Alessi		 
   Date Created:  4/3/2015
  Last Modified:  4/20/2015 
	Description:  Find oversized recordings on the voicemail server and move
				  them to the appropriate users U: drive. Then email the user
				  to notify them they have an oversized recording and where
				  they can find it.	   
================================================================================
#>

# Full Path to Voicemail Accounts
$voicemailPath = "\\path\to\your\voicemail\recordings"

# Full Path to Oversized Recordings Files Backup Location
$oversizedBackupPath = "\\your\backup\location"

# Grabs the whole path of the oversized recordings files
$fullPaths = Get-ChildItem $voicemailPath -rec | where {$_.Length -gt 10MB}


Foreach ($path in $fullPaths) {
	
	# Full Path to file
	$oversizedFilePath = $path.FullName

    # Date & time oversized recording started
    $fileStartDate = $path.CreationTime.DateTime

    # Date & time oversized recording finished
    $fileFinishedDate = $path.LastWriteTime.DateTime
	
	# Split the file off to get the username
	$oversizedFileUserPath = Split-Path $oversizedFilePath -parent
	$userName = Split-Path $oversizedFileUserPath -leaf
	
    # Get email address
    [string]$userEmail = $username + "@somedomain.com"

    # Creates a backup copy of the recorded statement and verify it gets copied there
    Copy-Item $oversizedFilePath -Destination $oversizedBackupPath\$userName
	
    # Sets the path to the users network drive
	[string]$userDrivePath = "\\server\drive\" + $userName
	
	# Set the Oversized file path on the users U: drive
	$userOversizedRecordingsFolder = [string]$userDrivePath + "\Oversized-Recordings"
	
	# Create an Oversized Recordings Folder, overwrite if it already exists, copy files to it
	New-Item $userOversizedRecordingsFolder -Type Directory -Force
	Move-Item $oversizedFilePath -Destination $userOversizedRecordingsFolder -Force


    # Log File
    $logPath = "$oversizedBackupPath\Logs"
    $logFile = "$logPath\$(Get-Date -Format yyy-mm-dd-hh-mm-ss).log"
    Write-Output "oversizedFilePath" $oversizedFilePath "oversizedFileUserPath" $oversizedFileUserPath "userName" $userName "userEmail" $userEmail "userDrivePath" $userDrivePath "userOversizedRecordingsFolder" $userOversizedRecordingsFolder "oversizedBackupPath" $oversizedBackupPath\$userName > $logFile
	
	Function Generate-Email {
		Write-Output "<html><head><title>Oversized Recordings Found - Automated Email</title></head><body>"
        Write-Output "<h3>This is an automated message, do not reply</h3><br/>"
        Write-Output "This is to notify you that your recent recorded statement that started on:<br/>"
        Write-Output "<b>$fileStartDate</b><br/>"
        Write-Output "and ended on:<br/>"
        Write-Output "<b>$fileFinishedDate</b><br/>" 
        Write-Output "was too large to email, so a copy of it was moved to your network Drive.<br/><br/>"
		Write-Output "<a href=""$userOversizedRecordingsFolder"">Click here</a> to go to your oversized recorded statement.<br/>"
		Write-Output "</body></html>" 
	}

	$SmtpClient = New-Object system.net.mail.smtpClient
	$SmtpClient.host = "yourmailserver.domain.com"   # Change to a SMTP server in your environment
	$MailMessage = New-Object system.net.mail.mailmessage
	$MailMessage.from = "helpdesk@somedomain.com"   # Change to email address you want emails to be coming from
	$MailMessage.To.add($userEmail)	# Change to email address you would like to receive emails.
    $MailMessage.To.add("someuser@somedomain.com")	
    $MailMessage.IsBodyHtml = 1
	$MailMessage.Subject = "Oversized Recordings Found - Automated Email"
	$MailMessage.Body = Generate-Email
	$SmtpClient.Send($MailMessage)

}