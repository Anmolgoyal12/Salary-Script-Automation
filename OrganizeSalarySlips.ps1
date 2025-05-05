Add-Type -AssemblyName System.Windows.Forms

# 🔐 Only run on last day of the month
$today = Get-Date
$lastDay = [DateTime]::DaysInMonth($today.Year, $today.Month)

if ($today.Day -ne $lastDay) {
    $override = [System.Windows.Forms.MessageBox]::Show(
        "Today is not the last day of the month ($($today.ToShortDateString())).`nDo you want to run the script anyway for past slips?",
        "Run Organizer Now?",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($override -ne "Yes") {
        Write-Host "⏹️ Script exited by user." -ForegroundColor Yellow
        exit
    }
}

# Month mapping for sorting
$monthOrder = @{
    "January" = 1; "February" = 2; "March" = 3; "April" = 4;
    "May" = 5; "June" = 6; "July" = 7; "August" = 8;
    "September" = 9; "October" = 10; "November" = 11; "December" = 12
}

# Folder picker dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing salary slips"

if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $sourceFolder = $folderBrowser.SelectedPath
    $logPath = Join-Path -Path $sourceFolder -ChildPath "SalarySlipLog.txt"
    Add-Content -Path $logPath -Value "----- Script run on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -----`n"

    # STEP 1: Get files, or pull from Outlook if folder is empty
    $pdfs = Get-ChildItem -Path $sourceFolder -Filter "*.pdf"

    if (-not $pdfs) {
        Write-Host "✉️ Checking Outlook..." -ForegroundColor Cyan

        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.Folders.Item("anmol.goyal@drogevate.com").Folders.Item("Inbox")

        Write-Host "📂 Available subfolders inside Inbox:"
        $inbox.Folders | ForEach-Object { Write-Host "  → $($_.Name)" }

        $salaryFolder = $inbox.Folders | Where-Object { $_.Name -ieq "Salary Recipt" }

        $emails = $salaryFolder.Items | Sort-Object -Property ReceivedTime -Descending

        foreach ($mail in $emails) {
            if ($mail.Subject -like "*Salary Slip*" -and $mail.Attachments.Count -gt 0) {
                foreach ($att in $mail.Attachments) {
                    if ($att.FileName -like "*.pdf") {
                        $dest = Join-Path $sourceFolder $att.FileName
                        $att.SaveAsFile($dest)
                        Add-Content -Path $logPath -Value "📥 Pulled from Outlook: $($att.FileName)"
                    }
                }
            }
        }
    }

    # STEP 2: Parse valid files
    $parsedFiles = @()
    Get-ChildItem -Path $sourceFolder -Filter "*.pdf" | ForEach-Object {
        if ($_.Name -match "Anmol Goyal Salary Slip (\w+) (\d{4})\.pdf") {
            $parsedFiles += [PSCustomObject]@{
                File     = $_
                Month    = $matches[1]
                Year     = $matches[2]
                MonthNum = $monthOrder[$matches[1]]
            }
        }
    }

    # STEP 3: Sort and move files
    $parsedFiles
    | Group-Object Year
    | Sort-Object Name
    | ForEach-Object {
        $year = $_.Name
        $entries = $_.Group | Sort-Object MonthNum

        $yearFolder = Join-Path -Path $sourceFolder -ChildPath $year

        if (-not (Test-Path -Path $yearFolder)) {
            New-Item -ItemType Directory -Path $yearFolder | Out-Null
            Add-Content -Path $logPath -Value "📁 Created folder: $yearFolder"
        } else {
            Add-Content -Path $logPath -Value "📁 Folder exists: $yearFolder"
        }

        foreach ($entry in $entries) {
            $destination = Join-Path -Path $yearFolder -ChildPath $entry.File.Name
            if (Test-Path -Path $destination) {
                Move-Item -Path $entry.File.FullName -Destination $destination -Force
                Add-Content -Path $logPath -Value "📝 Overwritten: $($entry.File.Name) → $year"
            } else {
                Move-Item -Path $entry.File.FullName -Destination $destination
                Add-Content -Path $logPath -Value "➡️ Moved: $($entry.File.Name) → $year"
            }
        }

        Add-Content -Path $logPath -Value ""
    }

    Add-Content -Path $logPath -Value "✅ Script Completed successfully.`n"
    Write-Host "✅ All slips organized in ascending month order (Jan → Dec)." -ForegroundColor Green

    # ✅ Toast (optional if BurntToast installed)
    try {
        Import-Module BurntToast -ErrorAction Stop
        New-BurntToastNotification -Text "Salary Slip Organizer", "✅ Salary slips have been organized successfully."
    } catch {
        Write-Host "⚠️ Notification skipped. BurntToast module not available." -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ Operation cancelled by user." -ForegroundColor Yellow
}
