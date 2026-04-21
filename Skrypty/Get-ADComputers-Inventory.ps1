# ============================================================
# Skrypt: Inwentaryzacja komputerów
# ============================================================

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

# Nagłówki
$Worksheet.Cells.Item(1, 1) = "HostName"
$Worksheet.Cells.Item(1, 2) = "Status"
$Worksheet.Cells.Item(1, 3) = "IP Address"
$Worksheet.Cells.Item(1, 4) = "DNS"
$Worksheet.Cells.Item(1, 5) = "OU"

# Formatowanie nagłówków
$Worksheet.Rows.Item(1).Font.Bold = $True

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Konfiguracja AD (DO MODYFIKACJI)
$adServer = "twoj.kontroler.domeny.com"
$basePath = "OU=TwojaJednostka,OU=Komputery,OU=Dzial,OU=Firma,DC=domena,DC=firma,DC=com"
$ouCommon = "OU=Wspolne,$basePath"
$ouWin10 = "OU=Windows10,$basePath"

try {
    $OU1 = Get-ADComputer -Server $adServer -Filter * -SearchBase $ouCommon -ErrorAction Stop | select Name, DistinguishedName
    $OU2 = Get-ADComputer -Server $adServer -Filter * -SearchBase $ouWin10 -ErrorAction Stop | select Name, DistinguishedName
}
catch {
    Write-Host "ERROR: Cannot connect to AD server: $adServer" -ForegroundColor Red
    Write-Host "Please verify the server name and your permissions." -ForegroundColor Yellow
    exit
}

$computers = $OU1 + $OU2 | Sort-Object Name

foreach($computer in $computers)
{
    $Hostname = $computer.Name
    $ping = Test-Connection $Hostname -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    if($ping){
        $Worksheet.Cells.Item($row, 1) = $Hostname
        $Worksheet.Cells.Item($row, 2) = "UP"
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 4  # Zielony
        
        # Pobranie IP
        try {
            $IP = (Test-Connection -ComputerName $Hostname -Count 1 -ErrorAction Stop).IPV4Address.IPAddressToString
            $Worksheet.Cells.Item($row, 3) = $IP
            
            # Odwrotne wyszukiwanie DNS
            $DNS = [System.Net.Dns]::GetHostEntry($IP).HostName
            $Worksheet.Cells.Item($row, 4) = $DNS
        }
        catch {
            $Worksheet.Cells.Item($row, 3) = "N/A"
            $Worksheet.Cells.Item($row, 4) = "N/A"
        }
        
        # Określenie OU
        if($computer.DistinguishedName -match "Common") {
            $Worksheet.Cells.Item($row, 5) = "Common"
        } else {
            $Worksheet.Cells.Item($row, 5) = "Win10"
        }
        
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else {
        $Worksheet.Cells.Item($row, 1) = $Hostname
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 3  # Czerwony
        $Worksheet.Cells.Item($row, 5) = if($computer.DistinguishedName -match "Common") {"Common"} else {"Win10"}
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}