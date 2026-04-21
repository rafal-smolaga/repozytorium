# ścieżka do pliku wynikowego
$path = ".\example_results.xls"

# utworzenie obiektu Excel
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

# widoczność arkusza
$Excel.Visible = $True

# nagłówki kolumn
$Worksheet.Cells.Item(1, 1) = "ExampleServiceAccount"
$Worksheet.Cells.Item(1, 2) = "ExampleAttributeLastSet"

# automatyczne dopasowanie szerokości kolumn
$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# wczytanie listy kont z pliku źródłowego
$serviceAccounts = Get-Content "C:\example_input\service_accounts_list.txt"

# przetwarzanie każdego konta
ForEach ($account in $serviceAccounts) 
{
    # pobranie danych z Active Directory
    $adResult = Get-ADUser -Identity $account -properties ExampleAttributeLastSet | select Name, ExampleAttributeLastSet
    
    # wyodrębnienie wartości
    $accountName = $adResult | select Name -ExpandProperty Name
    $attributeValue = $adResult | select ExampleAttributeLastSet -ExpandProperty ExampleAttributeLastSet

    # zapis do arkusza
    $Worksheet.Cells.Item($row, 1) = $accountName
    $Worksheet.Cells.Item($row, 2) = $attributeValue
    
    # dopasowanie kolumn po każdym wpisie
    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    $row++
}