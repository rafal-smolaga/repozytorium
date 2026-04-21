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
$Worksheet.Cells.Item(1, 3) = "Users"           # Kolumna C - liczba profili użytkowników

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Ścieżka do pliku z listą komputerów (DO MODYFIKACJI)
$Computers = Get-Content "ścieżka do pliku .txt"

foreach ($computer in $computers)
{
    # Test połączenia z komputerem (jeden pakiet ping)
    # -ea silentlycontinue - pomija błędy (np. brak odpowiedzi)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Resetowanie zmiennej wynikowej
    $result = $null
    
    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping){
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # ====================================================
        # ZLICZANIE PROFILI UŻYTKOWNIKÓW
        # ====================================================
        # Get-ChildItem "\\$Computer\c$\users" - pobiera listę folderów w C:\Users
        # -exclude Public,Default,Administrator - pomija foldery systemowe
        #   - Public - folder publiczny
        #   - Default - domyślny profil użytkownika
        #   - Administrator - profil administratora lokalnego
        # Measure-Object | Select Count - zlicza liczbę folderów
        $result = (Get-ChildItem "\\$Computer\c$\users" -exclude Public,Default,Administrator | Measure-Object).Count
        
        # Zapis liczby użytkowników w kolumnie C
        # ColorIndex 4 - kolor zielony (pozytywny wynik)
        $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 4
        $Worksheet.Cells.Item($row, 3) = "$result"
        
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
