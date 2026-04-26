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
$Worksheet.Cells.Item(1, 1) = "ExampleName"
$Worksheet.Cells.Item(1, 2) = "ExampleEmployeeMail"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# wczytanie listy użytkowników z pliku źródłowego
$users = gc "C:\ExamplePath\example_users_list.txt"

# pętla po każdym użytkowniku
foreach ($user in $users)
{
    # zapis nazwy użytkownika
    $Worksheet.Cells.Item($row, 1) = $user

    # pobranie adresu e-mail z Active Directory
    $employee = Get-ADUser -Identity "$user" -Properties * | select Mail -ExpandProperty Mail

    # zapis adresu e-mail
    $Worksheet.Cells.Item($row, 2) = "$employee"
    
    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    $row++
}