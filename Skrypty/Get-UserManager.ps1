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
$Worksheet.Cells.Item(1, 1) = "User Name"
$Worksheet.Cells.Item(1, 2) = "Manager DN"
$Worksheet.Cells.Item(1, 3) = "Manager Name"

# Formatowanie nagłówków
$Worksheet.Rows.Item(1).Font.Bold = $True

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()
$users = gc "ścieżka do pliku .txt"

foreach ($user in $users)
{
    $Worksheet.Cells.Item($row, 1) = $user
    
    try {
        # Pobranie menedżera z AD
        $managerDN = Get-ADUser -Identity $user -Properties Manager -ErrorAction Stop | Select-Object -ExpandProperty Manager
        
        if ($managerDN) {
            $Worksheet.Cells.Item($row, 2) = $managerDN
            
            $managerName = Get-ADUser -Identity $managerDN -Properties Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
            $Worksheet.Cells.Item($row, 3) = $managerName
        } else {
            $Worksheet.Cells.Item($row, 2) = "No manager assigned"
            $Worksheet.Cells.Item($row, 3) = "N/A"
        }
    }
    catch {
        $Worksheet.Cells.Item($row, 2) = "User not found"
        $Worksheet.Cells.Item($row, 3) = "N/A"
    }
    
    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    $row++
}