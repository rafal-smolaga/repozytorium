# AdComputers-NetworkInventoryReport.ps1
# Skrypt do inwentaryzacji sieciowej stacji roboczych z dwĂłch jednostek organizacyjnych AD

# Ścieżka do pliku wynikowego Excel
$path = ".\NetworkInventory_Report.xls"

# Utworzenie obiektu aplikacji Excel
$Excel = New-Object -comobject excel.application

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

# Nagłówki kolumn raportu
$Worksheet.Cells.Item(1, 1) = "Hostname"
$Worksheet.Cells.Item(1, 2) = "Status"
$Worksheet.Cells.Item(1, 3) = "IP Address"
$Worksheet.Cells.Item(1, 4) = "DNS"
$Worksheet.Cells.Item(1, 5) = "Ou"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Pobranie komputerów z pierwszej jednostki organizacyjnej (np. Win11)
Write-Host "Pobieranie komputerów z OU: Win11..." -ForegroundColor Cyan
$OU1 = Get-ADComputer -Filter * -SearchBase "OU=Win11,OU=Workstations,OU=Computers,DC=example,DC=com" | select Name, DistinguishedName

# Pobranie komputerów z drugiej jednostki organizacyjnej (np. Win10)
Write-Host "Pobieranie komputerów z OU: Win10..." -ForegroundColor Cyan
$OU2 = Get-ADComputer -Filter * -SearchBase "OU=Win10,OU=Workstations,OU=Computers,DC=example,DC=com" | select Name, DistinguishedName

# PoĹ‚Ä…czenie i sortowanie listy komputerów
$computers = $OU1 + $OU2 | Sort-Object Name
Write-Host "Znaleziono łącznie: $($computers.Count) komputerów" -ForegroundColor Green

foreach($computer in $computers)
{
    if(($computer | select DistinguishedName -ExpandProperty DistinguishedName) -match "Win11")
    {
        $Worksheet.Cells.Item($row, 5) = "Win11"
    }
    else
    {
        $Worksheet.Cells.Item($row, 5) = "Win10"
    }
    
    $computerName = $computer | select Name -ExpandProperty Name
    Write-Host "Sprawdzanie: $computerName" -ForegroundColor Yellow
    
    # Test poĹ‚Ä…czenia ping (1 pakiet)
    $ping = Test-Connection $computerName -Count 1 -ea silentlycontinue

    if($ping){
        # Komputer odpowiada - zapis podstawowych informacji
        $Worksheet.Cells.Item($row, 1) = $computerName
        $Worksheet.Cells.Item($row, 2) = "UP"
        Write-Host "  $computerName - UP" -ForegroundColor Green
        
        # Pobranie adresu IPv4
        $IP = ((Test-Connection -ComputerName "$computerName" -Count 1).IPV4Address).IPAddressToString
        $Worksheet.Cells.Item($row, 3) = "$IP"
        
        # Odwrotne zapytanie DNS (pobranie nazwy hosta na podstawie adresu IP)
        try {
            $DNS = [System.Net.Dns]::GetHostByAddress($IP).HostName
            $Worksheet.Cells.Item($row, 4) = "$DNS"
            Write-Host "IP: $IP, DNS: $DNS" -ForegroundColor Gray
        }
        catch {
            $Worksheet.Cells.Item($row, 4) = "DNS lookup failed"
            Write-Host "IP: $IP, DNS: lookup failed" -ForegroundColor Red
        }
        
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else {
        # Komputer nie odpowiada - zapis tylko nazwy i statusu DOWN
        $Worksheet.Cells.Item($row, 1) = $computerName
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        Write-Host "  $computerName - DOWN" -ForegroundColor Red
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}

# Zapisanie i zakoĹ„czenie
Write-Host "Raport został‚ wygenerowany: $path" -ForegroundColor Green
Write-Host "łącznie przetworzono: $($row-2) komputerów" -ForegroundColor Cyan

# Automatyczne dopasowanie kolumn na koniec
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()
