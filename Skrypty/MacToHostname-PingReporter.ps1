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
$Worksheet.Cells.Item(1, 1) = "ExampleMac"
$Worksheet.Cells.Item(1, 2) = "ExampleHostname"
$Worksheet.Cells.Item(1, 3) = "ExampleStatus"
$Worksheet.Cells.Item(1, 4) = "ExampleIPAddress"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# wczytanie listy adresów MAC z pliku źródłowego
$Macs = gc "C:\ExamplePath\example_devices_list.txt"

# pętla po każdym adresie MAC
foreach ($Mac in $Macs)
{
    # zapis adresu MAC
    $Worksheet.Cells.Item($row, 1) = $Mac
    
    # wyodrębnienie ostatnich 6 znaków adresu MAC
    $Name = $Mac.Substring($Mac.Length - 6)
    
    # utworzenie nazwy hosta z prefiksem (np. AVX + ostatnie 6 znaków MAC)
    $Hostname = "ExamplePrefix" + $Name
    $Worksheet.Cells.Item($row, 2) = $Hostname
    
    # test połączenia ping z nazwą hosta + domena
    $ping = (Test-Connection -ComputerName "$Hostname.example.domain.com" -Count 1).IPV4Address

    if($ping){
        # komputer odpowiada - pobranie adresu IP
        $IPAddress = $ping.IPAddressToString
        $Worksheet.Cells.Item($row, 3) = "ExampleStatus_Up"
        $Worksheet.Cells.Item($row, 4) = "$IPAddress"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else
    {
        # komputer nie odpowiada
        $Worksheet.Cells.Item($row, 3) = "ExampleStatus_Down"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}