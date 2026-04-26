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
$Worksheet.Cells.Item(1, 1) = "ExampleAccountName"
$Worksheet.Cells.Item(1, 2) = "ExampleLastLogonDate"

# wczytanie listy kont z pliku źródłowego
$machines = gc C:\ExamplePath\ExampleComputers.txt

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# pętla po każdym koncie
foreach ($machine in $machines)
{
    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    
    # zapis nazwy konta
    $Worksheet.Cells.Item($row, 1) = "$Machine"
    
    # pobranie daty ostatniego logowania z Active Directory
    $LLD = Get-ADUser -Identity "$Machine" -Properties * | Select LastLogonDate -ExpandProperty LastLogonDate

    # zapis daty ostatniego logowania
    $Worksheet.Cells.Item($row, 2) = "$LLD"

    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    $row++
}