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
$Worksheet.Cells.Item(1, 1) = "ExampleIPAddress"
$Worksheet.Cells.Item(1, 2) = "ExampleStatus"
$Worksheet.Cells.Item(1, 3) = "ExampleMAC"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# wczytanie listy komputerów z pliku źródłowego
$computers = gc "C:\ExamplePath\example_computers_list.txt"

# pętla po każdym komputerze
foreach ($computer in $computers)
{
    $MAC = $null
    
    # test połączenia ping (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue

    if($ping){
        # komputer odpowiada - zapis adresu IP i statusu UP
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Up"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # pobranie adresu MAC z aktywnej karty sieciowej Intel
        $MAC = Get-WmiObject win32_networkadapterconfiguration -Filter "IPEnabled='True'" -ComputerName $computer | where description -Match "example_vendor_pattern" | Select macaddress -ExpandProperty macaddress
        
        # zapis adresu MAC bez dwukropków
        $Worksheet.Cells.Item($row, 3) = $MAC -replace ':', ''

        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else {
        # komputer nie odpowiada - zapis adresu IP i statusu DOWN (czerwona czcionka)
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 3  # kolor czerwony
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Down"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}