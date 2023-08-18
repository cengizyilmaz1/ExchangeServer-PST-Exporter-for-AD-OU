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

# Check if target folder exists
Write-Host "Checking if target folder exists: $targetFolder"
if (-not (Test-Path -Path $targetFolder)) {
    # Create target folder
    Write-Host "Creating target folder: $targetFolder"
    New-Item -Path $targetFolder -ItemType Directory
    Write-Host "Target folder created: $targetFolder"
}

# Retrieve mailboxes from OU
Write-Host "Retrieving mailboxes from OU: $ouPath"
$mailboxes = Get-Mailbox -OrganizationalUnit $ouPath
Write-Host "Retrieved $($mailboxes.Count) mailboxes from OU: $ouPath"

# Export PST files for mailboxes
$mailboxes | ForEach-Object {
    $mailbox = $_
    $pstFile = Join-Path -Path $targetFolder -ChildPath ("$($mailbox.Alias).pst")
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