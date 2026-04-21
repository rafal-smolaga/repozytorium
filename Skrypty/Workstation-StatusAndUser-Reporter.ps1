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
$Worksheet.Cells.Item(1, 1) = "ExampleHostname"
$Worksheet.Cells.Item(1, 2) = "ExampleStatus"
$Worksheet.Cells.Item(1, 3) = "ExampleLoggedUser"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# pobranie listy komputerów z dwóch jednostek organizacyjnych AD
$OU1 = Get-ADComputer -Filter * -SearchBase "OU=ExampleOU_Common,OU=ExampleWorkstations,OU=ExampleComputers,OU=ExampleEMFP,OU=ExampleManufacturing,DC=exampledomain,DC=example,DC=com" | select Name -expandproperty Name

$OU2 = Get-ADComputer -Filter * -SearchBase "OU=ExampleOU_Win10,OU=ExampleWorkstations,OU=ExampleComputers,OU=ExampleEMFP,OU=ExampleManufacturing,DC=exampledomain,DC=example,DC=com" | select Name -expandproperty Name

# połączenie i sortowanie listy komputerów
$computers = $OU1 + $OU2 | sort

# pętla po każdym komputerze
foreach ($computer in $computers)
{
    # test połączenia ping (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue

    if($ping){
        # komputer odpowiada - zapis stanu UP
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Up"

        # pobranie zalogowanego użytkownika przez WMI
        $User = (get-wmiobject Win32_ComputerSystem -Computer $computer).UserName

        if($User -ne $null)
        {
            $Worksheet.Cells.Item($row, 3) = "$User"
        }
        Else{
            $Worksheet.Cells.Item($row, 3) = "Example_NoLoggedUser"
        }
        
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    else {
        # komputer nie odpowiada - zapis stanu DOWN
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Down"
        
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
    
    # reset zmiennej użytkownika przed kolejną iteracją
    $User = $null
}