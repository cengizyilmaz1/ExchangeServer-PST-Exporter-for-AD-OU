<#
=============================================================================================
Name = Cengiz YILMAZ
Microsoft Certified Trainer (MCT) - Microsoft MVP
Date = 18.08.2023
www.cengizyilmaz.net
365gurusu.com
www.cozumpark.com/author/cengizyilmaz
============================================================================================
#>

# Start transcript
$transcriptFile = "C:\PST_Export_Transcript.txt"
Start-Transcript -Path $transcriptFile -Append

# Parameters
$ouPath = "OU=YourOU,DC=domain,DC=com" # OU path
$targetFolder = "C:\PST_Folder" # Target folder path

$ouPath = "OU=fixcloud-test.com,OU=HC-Systems,DC=fixcloud,DC=com,DC=tr"
$targetFolder = "C:\PSTExports"

if (-not (Test-Path -Path $targetFolder -PathType Container)) {
    New-Item -Path $targetFolder -ItemType Directory
}

$mailboxes = Get-Mailbox -OrganizationalUnit $ouPath

Write-Host "Retrieved $($mailboxes.Count) mailboxes from OU: $ouPath"

$mailboxes | ForEach-Object {
    $mailbox = $_
    $emailAddress = $mailbox.PrimarySmtpAddress.ToString().Replace("@", ".")
    $pstFile = Join-Path -Path $targetFolder -ChildPath ("$emailAddress.pst")
    $existingRequest = Get-MailboxExportRequest -Mailbox $mailbox.Identity | Where-Object { ($_.Status -ne "Completed") -and ($_.FilePath -eq $pstFile) }
    if (-not $existingRequest) {
        try {
            Write-Host "Starting PST export request for mailbox: $($mailbox.EmailAddresses[0].Address) - $pstFile"
            New-MailboxExportRequest -Mailbox $mailbox.Identity -FilePath $pstFile
            Write-Host "Started PST export request for mailbox: $($mailbox.EmailAddresses[0].Address) - $pstFile"
        } catch {
            Write-Host "Error: $($mailbox.EmailAddresses[0].Address) - $pstFile - Error Message: $_"
        }
    } else {
        Write-Host "Existing or completed export request found for mailbox: $($mailbox.EmailAddresses[0].Address) - $pstFile"
    }
}

# Stop transcript
Stop-Transcript

Write-Host "Process completed. Transcript file: $transcriptFile"
