function Get-WindowsVersion {
    param(
        [string]$ReleaseId,
        [string]$CurrentBuild
    )
    
    # Windows 11 (Build 22000 i wyższe)
    if ($CurrentBuild -ge 22000) {
        return "Windows 11"
    }
    # Windows 10 (Build 10240 - 19045)
    elseif ($CurrentBuild -ge 10240 -and $CurrentBuild -lt 22000) {
        return "Windows 10"
    }
    # Windows 8.1 (Build 9600)
    elseif ($CurrentBuild -eq 9600) {
        return "Windows 8.1"
    }
    # Windows 8 (Build 9200)
    elseif ($CurrentBuild -eq 9200) {
        return "Windows 8"
    }
    # Windows 7 (Build 7600 - 7601)
    elseif ($CurrentBuild -ge 7600 -and $CurrentBuild -lt 8000) {
        return "Windows 7"
    }
    # Windows Vista (Build 6000 - 6002)
    elseif ($CurrentBuild -ge 6000 -and $CurrentBuild -lt 7000) {
        return "Windows Vista"
    }
    # Windows XP (Build 2600)
    elseif ($CurrentBuild -eq 2600) {
        return "Windows XP"
    }
    else {
        return "Unknown Windows Version"
    }
}

$path = ".\results.xls"
$Excel = new-object -comobject excel.application

if (Test-Path $path)
{ 
    $Workbook = $Excel.WorkBooks.Open($path) 
    $Worksheet = $Workbook.Worksheets.Item(1) 
}
else 
{ 
    $Workbook = $Excel.Workbooks.Add() 
    $Worksheet = $Workbook.Worksheets.Item(1)
}

$Excel.Visible = $True

# Rozszerzone nagłówki kolumn
$Worksheet.Cells.Item(1, 1) = "IP Address"           # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"               # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "Windows Version"      # Kolumna C - wersja Windows (11/10/8.1,..)
$Worksheet.Cells.Item(1, 4) = "ReleaseId"            # Kolumna D - ReleaseId (21H2, 22H2)
$Worksheet.Cells.Item(1, 5) = "Build"                # Kolumna E - numer kompilacji
$Worksheet.Cells.Item(1, 6) = "Edition"              # Kolumna F - edycja (Enterprise, Pro, Home,...)

# Formatowanie nagłówków (pogrubienie)
$Worksheet.Rows.Item(1).Font.Bold = $True

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()
$computers = gc "ścieżka do pliku .txt"

foreach ($computer in $computers)
{
    # Test połączenia z komputerem
    $ping = Test-Connection $computer -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    if($ping){
        
        # Zapis podstawowych informacji
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "UP"
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 4  # Zielony dla UP
        
        try {
            # Ścieżka do klucza rejestru z informacjami o systemie
            $KeyPathOS = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
            $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
            
            # Otwarcie połączenia z rejestrem zdalnego komputera
            $regOS = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $computer)
            $keyOS = $regOS.OpenSubKey($KeyPathOS)
            
            # Odczyt wartości z rejestru
            $resultReleaseId = $keyOS.GetValue('ReleaseId')        # Np. 22H2
            $resultBuild = $keyOS.GetValue('CurrentBuild')         # Np. 19045
            $resultEdition = $keyOS.GetValue('EditionID')          # Np. Professional
            $resultProductName = $keyOS.GetValue('ProductName')    # Pełna nazwa systemu
            
            # OKREŚLANIE WERSJI WINDOWS
            
            $windowsVersion = Get-WindowsVersion -ReleaseId $resultReleaseId -CurrentBuild $resultBuild
            
            # Dodatkowe sprawdzenie dla Windows 11 (mimo że ReleaseId może być takie samo jak Win10)
            if ($resultBuild -ge 22000) {
                $windowsVersion = "Windows 11"
            }
            elseif ($resultBuild -ge 10240 -and $resultBuild -lt 22000) {
                $windowsVersion = "Windows 10"
            }
            elseif ($resultBuild -ge 7600 -and $resultBuild -lt 8000) {
                $windowsVersion = "Windows 7"
            }
            
            # KOLOROWANIE W ZALEŻNOŚCI OD WERSJI WINDOWS
            
            # Zapis wersji Windows z odpowiednim kolorem
            switch ($windowsVersion) {
                "Windows 11" {
                    $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 5  # Niebieski
                    $Worksheet.Cells.Item($row, 3) = $windowsVersion
                }
                "Windows 10" {
                    $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 10 # Zielony
                    $Worksheet.Cells.Item($row, 3) = $windowsVersion
                }
                "Windows 7" {
                    $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 3  # Czerwony
                    $Worksheet.Cells.Item($row, 3) = $windowsVersion
                }
                default {
                    $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 1  # Czarny
                    $Worksheet.Cells.Item($row, 3) = $windowsVersion
                }
            }
            
            # Zapis pozostałych informacji
            $Worksheet.Cells.Item($row, 4) = "$resultReleaseId"
            $Worksheet.Cells.Item($row, 5) = "$resultBuild"
            $Worksheet.Cells.Item($row, 6) = "$resultEdition"
            
            # Dodatkowe ostrzeżenie dla Windows 7 (koniec wsparcia)
            if ($windowsVersion -eq "Windows 7") {
                $Worksheet.Cells.Item($row, 7) = "END OF LIFE - Upgrade required!"
                $Worksheet.Cells.Item($row, 7).Font.ColorIndex = 3  # Czerwony
            }
            
            # Zamknięcie połączenia z rejestrem
            $keyOS.Close()
            $regOS.Close()
        }
        catch {
            # Obsługa błędów (brak dostępu, brak klucza itp.)
            $Worksheet.Cells.Item($row, 3) = "Error"
            $Worksheet.Cells.Item($row, 4) = "N/A"
            $Worksheet.Cells.Item($row, 5) = "N/A"
            $Worksheet.Cells.Item($row, 6) = "N/A"
        }
        
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else {
        # Komputer niedostępny
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 3  # Czerwony
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        $Worksheet.Cells.Item($row, 3) = "N/A"
        $Worksheet.Cells.Item($row, 4) = "N/A"
        $Worksheet.Cells.Item($row, 5) = "N/A"
        $Worksheet.Cells.Item($row, 6) = "N/A"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}