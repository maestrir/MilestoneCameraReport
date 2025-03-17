<#
.SYNOPSIS
Genera un report completo delle telecamere Milestone con snapshot integrati in un file Excel.

.DESCRIPTION
Questo script si collega al Management Server di Milestone utilizzando PSTools, estrae dati dettagliati delle telecamere, compresi gli snapshot degli ultimi fotogrammi registrati, e genera un report Excel sintetico con immagini incorporate direttamente nel foglio di lavoro.

.REQUIREMENTS
- MilestonePSTools Module
- ImportExcel Module

.NOTES
Creato da Roby
#>

# Connessione al Server Milestone
Connect-ManagementServer -ShowDialog -AcceptEula

# Ottenere i dati dal server
$cameraReport = Get-VmsCameraReport -IncludeRetentionInfo
$cameraInfo = Get-VmsCamera

# Creare cartella snapshot
$snapshotFolder = Join-Path $PSScriptRoot "Snapshots_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $snapshotFolder -ItemType Directory -Force | Out-Null

# Funzione per aggiungere immagini in Excel
function Add-ExcelImage {
    param(
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,
        [System.Drawing.Image]$Image,
        [int]$Row,
        [int]$Column,
        [int]$Width = 150,
        [int]$Height = 100
    )

    $picture = $WorkSheet.Drawings.AddPicture((New-Guid).ToString(), $Image)
    $picture.SetPosition($Row - 1, 5, $Column - 1, 5)
    $picture.SetSize($Width, $Height)

    # Adatta dimensioni celle all'immagine
    $WorkSheet.Column($Column).Width = ($Width / 6)
    $WorkSheet.Row($Row).Height = ($Height * 0.75)
}

# Combina informazioni e ottiene snapshot
Write-Host "Sto elaborando le telecamere, attendere prego..."

$combinedReport = foreach ($cam in $cameraInfo) {

    $reportMatch = $cameraReport | Where-Object { $_.Name -eq $cam.Name } | Select-Object -First 1

    # Snapshot ultimo fotogramma registrato
    try {
        $snapshot = $cam | Get-Snapshot -Behavior GetEnd -ErrorAction Stop
        $snapshotPath = Join-Path $snapshotFolder "$($cam.ShortName)_$($cam.Id).jpg"
        [io.file]::WriteAllBytes($snapshotPath, $snapshot.Bytes)
    }
    catch {
        $snapshotPath = $null
    }

    [PSCustomObject]@{
        RecorderName        = $reportMatch.RecorderName
        Id                  = $cam.Id
        Name                = $cam.Name
        ShortName           = $cam.ShortName
        Description         = $cam.Description
        Address             = $reportMatch.Address
        Enabled             = if ($reportMatch) { $true } else { $false }
        LastModified        = $cam.LastModified
        MediaDatabaseBegin  = $reportMatch.MediaDatabaseBegin
        MediaDatabaseEnd    = $reportMatch.MediaDatabaseEnd
        UsedSpaceInGB       = $reportMatch.UsedSpaceInGB
        ActualRetentionDays = $reportMatch.ActualRetentionDays
        IsRecording         = $reportMatch.IsRecording
        SnapshotPath        = if ($snapshotPath) { $snapshotPath } else { "No Snapshot" }
    }
}

# Mostra dati in griglia
$combinedReport | Out-GridView -Title "Report Completo Telecamere Milestone"

# Esportazione Excel
$recorderName = ($cameraReport | Select-Object -First 1).RecorderName -replace '[^\w\-]','_'
$fileExcel = Join-Path $PSScriptRoot "${recorderName}_ReportTelecamere_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"

$excel = $combinedReport | Select RecorderName,Id,Name,ShortName,Description,Address,Enabled,LastModified,MediaDatabaseBegin,MediaDatabaseEnd,UsedSpaceInGB,ActualRetentionDays,IsRecording |
    Export-Excel -Path $fileExcel -AutoSize -WorksheetName "Telecamere" -PassThru

# Inserimento snapshot in Excel
$sheet = $excel.Workbook.Worksheets["Telecamere"]
$row = 2
foreach ($cam in $combinedReport) {
    if (Test-Path $cam.SnapshotPath) {
        $image = [System.Drawing.Image]::FromFile($cam.SnapshotPath)
        Add-ExcelImage -WorkSheet $sheet -Image $image -Row $row -Column 14
    }
    $row++
}

Close-ExcelPackage $excel

Write-Host "Report e snapshot generati con successo in: $fileExcel" -ForegroundColor Green
