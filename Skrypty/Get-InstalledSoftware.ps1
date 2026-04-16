$path = ".\results.xls"

# Utworzenie obiektu aplikacji Excel
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

$Worksheet.Cells.Item(1, 1) = "IP Adress"      # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"         # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "Architecture"   # Kolumna C - architektura (x32/x64)
$Worksheet.Cells.Item(1, 4) = "Result"         # Kolumna D - nazwa programu
$Worksheet.Cells.Item(1, 5) = "Version"        # Kolumna E - wersja programu

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

$computers = gc "ścieżka do pliku .txt"

foreach ($computer in $computers)
{
    # Test połączenia z komputerem (jeden pakiet ping)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue

    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping){
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # Resetowanie zmiennych przed skanowaniem
        $SubKeysResult = $null
        $SubKeyPath = $null
        $result = $null
        $result1 = $null
        
        # Ścieżka do klucza rejestru z listą zainstalowanych programów (32-bit)
        $SubKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        
        # Otwarcie połączenia z rejestrem zdalnego komputera
        $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
        
        # Pobranie wszystkich podkluczy w ścieżce Uninstall
        $SubKeyNames = $reg.OpenSubKey($SubKeyPath)
        $SubKeysResult = Foreach($sub in $SubKeyNames.GetSubKeyNames()){$sub}
        
        # Iteracja przez każdy podklucz (każdy podklucz to jeden program)
        foreach($SubKeysResults in $SubKeysResult)
        {
            # Pełna ścieżka do klucza danego programu
            $KeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\$SubKeysResults"
            
            # Nazwy wartości do odczytu
            $Value = 'DisplayName'      # Nazwa programu
            $Value1 = 'DisplayVersion'  # Wersja programu
            
            # Otwarcie połączenia z rejestrem (ponowne - nieoptymalne)
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
            $key = $reg.OpenSubKey($KeyPath)
            
            # Pobranie wartości z rejestru
            $result = $key.GetValue($Value)
            $result1 = $key.GetValue($Value1)

            # Jeśli nazwa programu nie jest pusta - zapisz do Excela
            if($result -eq $null)
            {
                # Pominięcie pustych wpisów
            } 
            else
            {
                # Zapis danych w Excelu
                $Worksheet.Cells.Item($row, 3) = "x32"      # Architektura 32-bit
                $Worksheet.Cells.Item($row, 4) = "$result"  # Nazwa programu
                $Worksheet.Cells.Item($row, 5) = "$result1" # Wersja programu
                $row++  # Przejście do następnego wiersza
            }
        }
        
        $Worksheet.Cells.Item($row, 1) = "------"
        $Worksheet.Cells.Item($row, 2) = "------"
        $Worksheet.Cells.Item($row, 3) = "------"
        $Worksheet.Cells.Item($row, 4) = "-----------------------------"
        $Worksheet.Cells.Item($row, 5) = "-----------------------------"
        $row++

        
        # Resetowanie zmiennych przed skanowaniem
        $SubKeysResult = $null
        $SubKeyPath = $null
        $result = $null
        $result1 = $null
        
        # Ścieżka do klucza rejestru z listą zainstalowanych programów (64-bit)
        # WOW6432Node zawiera programy 32-bitowe działające na 64-bitowym systemie
        $SubKeyPath = "Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        
        # Otwarcie połączenia z rejestrem zdalnego komputera
        $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
        
        # Pobranie wszystkich podkluczy w ścieżce Uninstall
        $SubKeyNames = $reg.OpenSubKey($SubKeyPath)
        $SubKeysResult = Foreach($sub in $SubKeyNames.GetSubKeyNames()){$sub}
        
        # Iteracja przez każdy podklucz (każdy podklucz to jeden program)
        foreach($SubKeysResults in $SubKeysResult)
        {
            # Pełna ścieżka do klucza danego programu
            $KeyPath = "Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$SubKeysResults"
            
            # Nazwy wartości do odczytu
            $Value = 'DisplayName'      # Nazwa programu
            $Value1 = 'DisplayVersion'  # Wersja programu
            
            # Ponowne otwarcie połączenia z rejestrem
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
            $key = $reg.OpenSubKey($KeyPath)
            
            # Pobranie wartości z rejestru
            $result = $key.GetValue($Value)
            $result1 = $key.GetValue($Value1)

            # Jeśli nazwa programu nie jest pusta - zapisz do Excela
            if($result -eq $null)
            {
                # Pominięcie pustych wpisów
            } 
            else
            {
                $Worksheet.Cells.Item($row, 3) = "x64"      # Architektura 64-bit
                $Worksheet.Cells.Item($row, 4) = "$result"  # Nazwa programu
                $Worksheet.Cells.Item($row, 5) = "$result1" # Wersja programu
                $row++ 
            }
        }

        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    }
    else
    {
        # Jeśli komputer NIE odpowiada na ping (jest niedostępny)
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "DOWN" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        
        # Dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
    }
}
