$computerList = Get-Content -Path "C:\script\comp.txt"

$destinationFolder = "C:\komp-status"
New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null

$lastRunFile = "C:\script\last_run.txt"

$remotePath = "c$\ProgramData\ESET\...\status.html"

while ($true) {
    $connectionResults = @{} # 

    foreach ($computer in $computerList) {
        try {
            $windowsFolder = "\\$computer\c$\Windows"
            if (Test-Path $windowsFolder) {
                Write-Host "Połączono z komputerem: $computer" -ForegroundColor Green
                $connectionResults[$computer] = "TAK"
                
                $remoteFile = "\\$computer\$remotePath"
                $destinationFile = Join-Path -Path $destinationFolder -ChildPath "$computer-status.html"

                Copy-Item -Path $remoteFile -Destination $destinationFile -ErrorAction Stop
                Write-Host "Pobrano plik z komputera $computer i zapisano jako $destinationFile" -ForegroundColor Green
            } else {
                Write-Host "Brak dostępu do komputera $computer (nie znaleziono C:\Windows)" -ForegroundColor Yellow
                $connectionResults[$computer] = "NIE"
            }
        } catch {
            Write-Host "Błąd podczas łączenia z komputerem $computer : $_" -ForegroundColor Red
            $connectionResults[$computer] = "NIE"
        }
    }

    $currentTime = Get-Date -Format "HH:mm"
    $currentDate = Get-Date -Format "yyyy-MM-dd"

    if (Test-Path $lastRunFile) {
        $lastRunDate = Get-Content $lastRunFile
    } else {
        $lastRunDate = ""
    }

    if ($currentTime -ge "07:00" -and $currentDate -ne $lastRunDate) {
        Write-Host "Uruchamiam analizę i eksport do Excela..." -ForegroundColor Cyan

        Set-Content -Path $lastRunFile -Value $currentDate

        $inputFile = "C:\Temp\komputery.txt"
        $outputFile = "C:\Temp\wynik_komputery.xlsx"
        $statusFolder = "C:\komp-status"

        $computerList = Get-Content $inputFile | ForEach-Object { $_ -replace "\.ente\.local", "" } | Where-Object { $_ -ne "" }
        $statusComputers = Get-ChildItem -Path $statusFolder -Filter "*-status.html" | ForEach-Object { $_.BaseName -replace "-status", "" }

        if ($computerList.Count -eq 0) {
            Write-Host "Brak komputerów w pliku!"
            continue
        }

        $results = @()
        foreach ($computer in $computerList) {
            try {
                $computerInfo = Get-ADComputer -Identity $computer -Properties Created, LastLogonDate -ErrorAction Stop
                if ($computerInfo) {
                    $results += [PSCustomObject]@{
                        "Nazwa Komputera" = $computerInfo.Name
                        "Data Utworzenia" = $computerInfo.Created
                        "Ostatnie Logowanie" = $computerInfo.LastLogonDate
                        "Status" = if ($statusComputers -contains $computer) { "TAK" } else { "NIE" }
                        "Połączenie" = $connectionResults[$computer] -as [string]
                    }
                }
            } catch {
                Write-Warning "Nie znaleziono komputera: $computer"
                $results += [PSCustomObject]@{
                    "Nazwa Komputera" = $computer
                    "Data Utworzenia" = ""
                    "Ostatnie Logowanie" = ""
                    "Status" = "NIE"
                    "Połączenie" = $connectionResults[$computer] -as [string]
                }
            }
        }

        if ($results.Count -eq 0) {
            Write-Host "Nie znaleziono żadnych komputerów w AD!"
            continue
        }

        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "Instaluję moduł ImportExcel..."
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
            Import-Module ImportExcel
        }

        $excelParams = @{
            Path          = $outputFile
            WorksheetName = "Komputery"
            AutoSize      = $true
            BoldTopRow    = $true
            FreezeTopRow  = $true
            TableName     = "KomputeryData"
        }

        $sortedResults = $results | Sort-Object "Nazwa Komputera"
        $sortedResults | Export-Excel @excelParams -ClearSheet

        $excel = Open-ExcelPackage -Path $outputFile
        $sheet = $excel.Workbook.Worksheets["Komputery"]

        $range = $sheet.Cells["A1:E$($sortedResults.Count+1)"]
        $range.AutoFilter = $true

        $row = 2
        foreach ($item in $sortedResults) {
            if ($item.Status -eq "TAK") {
                $sheet.Cells["A$row:D$row"].Style.Fill.PatternType = "Solid"
                $sheet.Cells["A$row:D$row"].Style.Fill.BackgroundColor.SetColor("LightGreen")
            }
            
            if ($item.Połączenie -eq "TAK") {
                $sheet.Cells["E$row"].Style.Fill.PatternType = "Solid"
                $sheet.Cells["E$row"].Style.Fill.BackgroundColor.SetColor("LightGreen")
            } else {
                $sheet.Cells["E$row"].Style.Fill.PatternType = "Solid"
                $sheet.Cells["E$row"].Style.Fill.BackgroundColor.SetColor("LightCoral")
            }
            $row++
        }

        $excel.Save()
        Close-ExcelPackage $excel

        Write-Host "Dane zapisane i oznaczone w Excelu: $outputFile"
    }

    Write-Host "Czekam 60 minut przed kolejną iteracją..." -ForegroundColor Cyan
    Start-Sleep -Seconds 3600
}
