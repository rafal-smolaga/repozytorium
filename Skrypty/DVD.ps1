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
$Excel.Visible = $True

# Utworzenie nagłówków kolumn w pierwszym wierszu
$Worksheet.Cells.Item(1, 1) = "IP Adress"    # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"       # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "DVD"          # Kolumna C - obecność napędu DVD

# Wczytanie listy komputerów z pliku tekstowego
$machines = gc "ścieżka do pliku .txt"

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Pętla przechodząca przez każdy komputer z listy
$machines | foreach-object {
    
    # Wyzerowanie zmiennej ping przed każdym testem
    $ping = $null
    
    # Pobranie bieżącej nazwy komputera
    $machine = $_
    
    # Test połączenia z komputerem (jeden pakiet ping)
    $ping = Test-Connection $machine -Count 1 -ea silentlycontinue
    
    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping) {
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $machine
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Pobranie informacji o napędzie DVD na zdalnym komputerze
        # Win32_CDROMDrive - klasa WMI reprezentująca napędy optyczne
        $DVD = Get-WmiObject Win32_CDROMDrive -ComputerName $machine | select -expandproperty Caption
        
        # Sprawdzenie czy napęd DVD istnieje
        if($DVD)
        {
            # Jeśli napęd istnieje - wpisz "Yes" w kolumnie C
            $Worksheet.Cells.Item($row, 3) = "Yes"
        }
        Else
        {
            # Jeśli napęd nie istnieje - wpisz "No" w kolumnie C
            $Worksheet.Cells.Item($row, 3) = "No"
        }
        
        # Dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
    }
    else 
    {
        # Jeśli komputer NIE odpowiada na ping (jest niedostępny)
        
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $machine
        
        # Zapis statusu "DOWN" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        
        # Dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
    }
}

# Uwaga: Skrypt nie zapisuje ani nie zamyka pliku Excel!
# Należy to zrobić ręcznie lub dodać brakujące linie:
# $Workbook.Save()
# $Workbook.Close()
# $Excel.Quit()