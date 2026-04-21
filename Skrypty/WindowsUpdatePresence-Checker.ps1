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
$Worksheet.Cells.Item(1, 3) = "ExampleRegistryCommunication"
$Worksheet.Cells.Item(1, 4) = "ExampleUpdate_KB0000001_Path"
$Worksheet.Cells.Item(1, 5) = "ExampleUpdate_KB0000001_Status"
$Worksheet.Cells.Item(1, 6) = "ExampleUpdate_KB0000002_Path"
$Worksheet.Cells.Item(1, 7) = "ExampleUpdate_KB0000002_Status"

$row = 2
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# wczytanie listy komputerów z pliku źródłowego
$computers = gc "C:\ExamplePath\ExampleScripts\example_computers_list.txt"

# pętla po każdym komputerze
foreach ($computer in $computers)
{
    # test połączenia ping (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue

    if($ping){
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Up"
        Write-Host "Processing:$computer" -BackgroundColor "Blue"
        
        $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine

        # próba zdalnego połączenia z rejestrem
        try
        {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
            $SubKeyPath = "Example\Registry\Path\ComponentBasedServicing\Packages"
            $Status = "Pass"
        }
        catch
        {
            $Status = "Failed"
        }

        If($Status -eq "Pass")
        {
            # komunikacja z rejestrem OK
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 4  # kolor zielony
            $Worksheet.Cells.Item($row, 3) = "PASS"
            
            # pobranie listy podkluczy w ścieżce Packages
            $SubKeyNames = $reg.OpenSubKey($SubKeyPath)
            $PathSub = Foreach($sub in $SubKeyNames.GetSubKeyNames()){$sub}

            # === sprawdzenie pierwszego update'u (KB0000001) ===
            if($PathSub -eq 'Example_Package_RollupFix_Identifier_Version1')
            {
                $Worksheet.Cells.Item($row, 4).Font.ColorIndex = 4  # zielony - istnieje
                $Worksheet.Cells.Item($row, 4) = "exist"
                
                $ValuesKeyPath = 'Example\Registry\Path\ComponentBasedServicing\Packages\Example_Package_RollupFix_Identifier_Version1'
                $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
                $ValueNames = Foreach($val in $ValuesKeyNames.GetValueNames()){$val}

                if($ValueNames -match 'CurrentState')
                {
                    $ValueAL = 'CurrentState'
                    $resultAL = $ValuesKeyNames.GetValue($ValueAL)
                    
                    if($resultAL -eq '112')  # 112 oznacza zainstalowany
                    {
                        $Worksheet.Cells.Item($row, 5).Font.ColorIndex = 4  # zielony - OK
                        $Worksheet.Cells.Item($row, 5) = "OK"
                    }
                    if($resultAL -ne '112')
                    {
                        $Worksheet.Cells.Item($row, 5).Font.ColorIndex = 3  # czerwony - zła wartość
                        $Worksheet.Cells.Item($row, 5) = "Bad Value"
                    }
                }
            }
            else
            {
                $Worksheet.Cells.Item($row, 4).Font.ColorIndex = 3  # czerwony - nie istnieje
                $Worksheet.Cells.Item($row, 4) = "not exist"
            }
            
            # === sprawdzenie drugiego update'u (KB0000002) ===
            if($PathSub -eq 'Example_Package_DotNetRollup_Identifier_Version2')
            {
                $Worksheet.Cells.Item($row, 6).Font.ColorIndex = 4  # zielony - istnieje
                $Worksheet.Cells.Item($row, 6) = "exist"
                
                $ValuesKeyPath1 = 'Example\Registry\Path\ComponentBasedServicing\Packages\Example_Package_DotNetRollup_Identifier_Version2'
                $ValuesKeyNames1 = $reg.OpenSubKey($ValuesKeyPath1)
                $ValueNames1 = Foreach($val in $ValuesKeyNames1.GetValueNames()){$val}
                
                if($ValueNames1 -match 'CurrentState')
                {
                    $ValueACC = 'CurrentState'
                    $resultACC = $ValuesKeyNames1.GetValue($ValueACC)
                    
                    if($resultACC -eq '112')  # 112 oznacza zainstalowany
                    {
                        $Worksheet.Cells.Item($row, 7).Font.ColorIndex = 4  # zielony - OK
                        $Worksheet.Cells.Item($row, 7) = "OK"
                    }
                    if($resultACC -ne '112')
                    {
                        $Worksheet.Cells.Item($row, 7).Font.ColorIndex = 3  # czerwony - zła wartość
                        $Worksheet.Cells.Item($row, 7) = "Bad Value"
                    }
                }
            }
            else
            {
                $Worksheet.Cells.Item($row, 6).Font.ColorIndex = 3  # czerwony - nie istnieje
                $Worksheet.Cells.Item($row, 6) = "not exist"
            }

            [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
            $row++
        }

        If($Status -eq "Failed")
        {
            # nie udało się połączyć z rejestrem zdalnym
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 3  # kolor czerwony
            $Worksheet.Cells.Item($row, 3) = "Failed"
            [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
            $row++
        }
    }
    else {
        # komputer nie odpowiada na ping
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 3  # kolor czerwony
        $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Down"
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        $row++
    }
}