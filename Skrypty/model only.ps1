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

$Worksheet.Cells.Item(1, 1) = "IP Address"    # Kolumna A - nazwa komputera/adres
$Worksheet.Cells.Item(1, 2) = "Status"        # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "Model"         # Kolumna C - model komputera

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()


$computers = gc "ścieżka do pliku .txt"


foreach ($computer in $computers)
{
    # Test połączenia z komputerem (jeden pakiet ping)
    # -ea silentlycontinue - pomija błędy (np. brak odpowiedzi)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping){
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Pobranie modelu komputera za pomocą WMI
        # Win32_ComputerSystem - klasa WMI zawierająca informacje o komputerze
        # Model - właściwość przechowująca model urządzenia
        $Model = Get-WmiObject Win32_ComputerSystem -computername "$computer" | Select Model -ExpandProperty Model
        
        # Zapis modelu w kolumnie C
        $Worksheet.Cells.Item($row, 3) = "$Model"
        
        # Ponowne dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        $row++
    }
    else
    {
        # Jeśli komputer NIE odpowiada na ping (jest niedostępny)
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Dzięki temu status "DOWN" będzie wyróżniony kolorem czerwonym
        $Worksheet.Cells.Item($row, 2).Font.ColorIndex = 3
        
        # Zapis statusu "DOWN" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        
        # Dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
    }
}