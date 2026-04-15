$adapterTypes = @{
    '-2' = 'Unknown'                          # Nieznany
    '-1' = 'Unknown'                          # Nieznany
    '0' = 'VGA'                               # VGA (analogowe)
    '1' = 'S-Video'                           # S-Video
    '2' = 'Composite'                         # Composite (RCA)
    '3' = 'Component'                         # Component (YPbPr)
    '4' = 'DVI'                               # DVI
    '5' = 'HDMI'                              # HDMI
    '6' = 'LVDS'                              # LVDS (wbudowane matryce laptopów)
    '8' = 'D-Jpn'                             # Japońskie złącze D
    '9' = 'SDI'                               # SDI (profesjonalne)
    '10' = 'DisplayPort (external)'           # DisplayPort zewnętrzny
    '11' = 'DisplayPort (internal)'           # DisplayPort wewnętrzny (eDP)
    '12' = 'Unified Display Interface'        # UDI
    '13' = 'Unified Display Interface (embedded)' # UDI wbudowane
    '14' = 'SDTV dongle'                      # SDTV dongle
    '15' = 'Miracast'                         # Miracast (bezprzewodowe)
    '16' = 'Internal'                         # Wewnętrzne
    '2147483648' = 'Internal'                 # Wewnętrzne (alternatywny kod)
}

# Inicjalizacja pustej tablicy na monitory
$arrMonitors = @()

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

# Ustawienie widoczności Excela (TRUE = pokaż okno Excela), dzięki temu wyniki są widoczne w czasie rzeczywistym
$Excel.Visible = $True

$Worksheet.Cells.Item(1, 1) = "IP Adress"    # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"       # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "Video"        # Kolumna C - typy złączy monitorów

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Wczytanie listy komputerów z pliku tekstowego
$Computers = gc "ścieżka do pliku .txt"

# Pętla przechodząca przez każdy komputer z listy
foreach ($computer in $computers)
{
    # Test połączenia z komputerem (jeden pakiet ping)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Jeśli komputer odpowiada na ping
    if($ping){
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # Wyświetlenie nazwy komputera w konsoli (informacja o postępie)
        write-host "$computer"
        
        # Klasa WmiMonitorID - podstawowe informacje o monitorze (nazwa, producent)
        $monitors = gwmi -ComputerName $computer WmiMonitorID -Namespace root/wmi -ErrorAction SilentlyContinue
        
        # Klasa WmiMonitorConnectionParams - informacje o typie złącza wideo
        $connections = gwmi -ComputerName $computer WmiMonitorConnectionParams -Namespace root/wmi -ErrorAction SilentlyContinue
        
        # Wyzerowanie tablicy monitorów dla bieżącego komputera
        $arrMonitors = ''
        $monitor = ''
        
        # Pętla przechodząca przez wszystkie wykryte monitory
        foreach ($monitor in $monitors)
        {
            # Pobranie przyjaznej nazwy monitora jeżeli istnieje
             $name = $monitor.UserFriendlyName
			
            # Pobranie typu złącza wideo dla bieżącego monitora
            # InstanceName musi być zgodne między obiema klasami WMI
            $connectionType = ($connections | ? {$_.InstanceName -eq $monitor.InstanceName}).VideoOutputTechnology
            
            # Konwersja nazwy producenta z tablicy bajtów na string
            if ($manufacturer -ne $null) {$manufacturer =[System.Text.Encoding]::ASCII.GetString($manufacturer -ne 0)}
            
            # Konwersja nazwy monitora z tablicy bajtów na string
            if ($name -ne $null) {$name =[System.Text.Encoding]::ASCII.GetString($name -ne 0)}
            
            # Mapowanie kodu liczbowego na czytelną nazwę typu złącza
            $connectionType = $adapterTypes."$connectionType"
            
            # Jeśli typ złącza nie został rozpoznany - ustaw jako 'Unknown'
            if ($connectionType -eq $null){$connectionType = 'Unknown'}
            
            # Dodanie typu złącza do tablicy (w nawiasach)
            if($name -ne $null){$arrMonitors += "($connectionType)"}
        }
        
        # FORMATOWANIE WYNIKÓW DLA EXCEL-a
        $strMonitors = ''
        $i = 0
        
        # Jeśli znaleziono jakieś monitory
        if ($arrMonitors.Count -gt 0){
            foreach ($monitor in $arrMonitors){
                # Pierwszy element - bez separatora
                if ($i -eq 0){$strMonitors += $arrMonitors[$i]}
                # Kolejne elementy - z separatorem nowej linii
                else{$strMonitors += "`n"; $strMonitors += $arrMonitors[$i]}
                $i++
            }
        }
        
        # Jeśli nie znaleziono żadnego monitora
        if ($strMonitors -eq ''){$strMonitors = 'None Found'}
        
        # Zapisanie sformatowanych typów złączy w kolumnie C
        $Worksheet.Cells.Item($row, 3) = "$strMonitors"
        
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