# ścieżka do pliku wynikowego Excel
$path = ".\example_results.xls"

# utworzenie obiektu aplikacji Excel
$Excel = new-object -comobject excel.application

# sprawdzenie czy plik już istnieje
if (Test-Path $path)
{ 
    # otwarcie istniejącego pliku
    $Workbook = $Excel.WorkBooks.Open($path) 
    $Worksheet = $Workbook.Worksheets.Item(1) 
}
else 
{ 
    # utworzenie nowego pliku
    $Workbook = $Excel.Workbooks.Add() 
    $Worksheet = $Workbook.Worksheets.Item(1)
}

# widoczność arkusza Excel
$Excel.Visible = $True

# nagłówki kolumn raportu
$Worksheet.Cells.Item(1, 1) = "ExampleHostName"
$Worksheet.Cells.Item(1, 2) = "ExampleStatus"
$Worksheet.Cells.Item(1, 3) = "ExampleIPAddress"
$Worksheet.Cells.Item(1, 4) = "ExampleDNS"
$Worksheet.Cells.Item(1, 5) = "ExampleOU"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# pobranie komputerów z pierwszej jednostki organizacyjnej (OU_Common)
$OU1 = Get-ADComputer -Filter * -SearchBase "OU=ExampleCommon,OU=ExampleWorkstations,OU=ExampleComputers,OU=ExampleEMFP,OU=ExampleManufacturing,DC=exampledomain,DC=example,DC=com" | select Name, DistinguishedName

# pobranie komputerów z drugiej jednostki organizacyjnej (OU_Win10)
$OU2 = Get-ADComputer -Filter * -SearchBase "OU=ExampleWin10,OU=ExampleWorkstations,OU=ExampleComputers,OU=ExampleEMFP,OU=ExampleManufacturing,DC=exampledomain,DC=example,DC=com" | select Name, DistinguishedName

# połączenie i sortowanie listy komputerów
$computers = $OU1 + $OU2 | Sort-Object Name

# pętla po każdym komputerze
foreach($computer in $computers)
{
    # określenie przynależności do OU na podstawie DistinguishedName
    if(($computer | select DistinguishedName -ExpandProperty DistinguishedName) -match "ExampleCommon")
    {
        $Worksheet.Cells.Item($row, 5) = "Common"
    }
    else
    {
        $Worksheet.Cells.Item($row, 5) = "Win10"
    }
    
    # wyodrębnienie samej nazwy komputera
    $computer = $computer | select Name -ExpandProperty Name
    
    # test połączenia ping (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue

    if($ping){
        # komputer odpowiada - zapis podstawowych informacji
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Up"
        
        # pobranie adresu IPv4
        $IP = ''
        $IP = ((Test-Connection -ComputerName "$computer" -Count 1).IPV4Address).IPAddressToString
        $Worksheet.Cells.Item($row, 3) = "$IP"
        
        # odwrotne zapytanie DNS (pobranie nazwy hosta na podstawie adresu IP)
        $DNS = ''
        $DNS = [System.Net.Dns]::GetHostByAddress($IP).HostName
        $Worksheet.Cells.Item($row, 4) = "$DNS"
        
        # reset zmiennych
        $IP = ''
        $DNS = ''
        
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else {
        # komputer nie odpowiada - zapis tylko nazwy i statusu DOWN
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Down"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}