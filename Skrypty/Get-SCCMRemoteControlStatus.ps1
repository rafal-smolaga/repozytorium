# Ścieżka do pliku wynikowego Excel
$path = ".\results.xls"

# Utworzenie obiektu aplikacji Excel
$Excel = new-object -comobject excel.application

# Sprawdzenie czy plik wynikowy już istnieje
if (Test-Path $path)
{ 
    # Jeśli plik istnieje - otwórz go
    $Workbook = $Excel.WorkBooks.Open($path) 
    $Worksheet = $Workbook.Worksheets.Item(1) 
}
else 
{ 
    # Jeśli plik nie istnieje - utwórz nowy skoroszyt
    $Workbook = $Excel.Workbooks.Add() 
    $Worksheet = $Workbook.Worksheets.Item(1)
}

# Ustawienie widoczności Excela (TRUE = pokaż okno Excela)
# Dzięki temu wyniki są widoczne w czasie rzeczywistym
$Excel.Visible = $True

$Worksheet.Cells.Item(1, 1) = "IP Address"      # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"          # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "Reg Status"      # Kolumna C - czy Remote Control istnieje (Yes/No)
$Worksheet.Cells.Item(1, 4) = "Value"           # Kolumna D - wartość klucza Permission Required (0/1)

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

$computers = gc "ścieżka do pliku .txt"


foreach ($computer in $computers)
{
    # Resetowanie zmiennej wynikowej
    $result = $null
    
    # Test połączenia z komputerem (jeden pakiet ping)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping){
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        write-host "Processing $computer"
        
        # Ścieżka do klucza rejestru SCCM Client Components
        $SubKeyPath = "SOFTWARE\Microsoft\SMS\Client\Client Components"
        
        # Otwarcie połączenia z rejestrem zdalnego komputera
        $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
        
        # Otwarcie podklucza i pobranie nazw wszystkich podkluczy
        $SubKeyNames = $reg.OpenSubKey($SubKeyPath)
        $RemoteControlSub = Foreach($sub in $SubKeyNames.GetSubKeyNames()){$sub}
        
        # Sprawdzenie czy istnieje podklucz 'Remote Control'
        if($RemoteControlSub -eq 'Remote Control'){
            
            # Ustawienie koloru zielonego (ColorIndex 4) dla statusu "Yes"
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 4
            $Worksheet.Cells.Item($row, 3) = "Yes"
            
            # Ścieżka do klucza z ustawieniami Remote Control
            $ValuesKeyPath = 'SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control'
            
            # Otwarcie podklucza z wartościami
            $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
            
            # Pobranie nazw wszystkich wartości
            $ValueNames = Foreach($val in $ValuesKeyNames.GetValueNames()){$val}
            
            # Odczytanie wartości 'Permission Required'
            $ValueV = 'Permission Required'
            $resultV = $ValuesKeyNames.GetValue($ValueV)
            
            # INTERPRETACJA WARTOŚCI PERMISSION REQUIRED
            # 0 = Nie wymaga pozwolenia (automatyczna zgoda)
            # 1 = Wymaga pozwolenia (użytkownik musi zaakceptować)
            
            if($resultV -eq 0)
            {
                # Wartość 0 - nie wymaga pozwolenia
                $Worksheet.Cells.Item($row, 4).Font.ColorIndex = 5
                $Worksheet.Cells.Item($row, 4) = "$resultV"
            }
            if($resultV -eq 1)
            {
                # Wartość 1 - wymaga pozwolenia
                $Worksheet.Cells.Item($row, 4).Font.ColorIndex = 3
                $Worksheet.Cells.Item($row, 4) = "$resultV"
            }
        }
        # Jeśli podklucz 'Remote Control' nie istnieje - pozostaw puste
        
        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
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