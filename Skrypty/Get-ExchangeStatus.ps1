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

# Tworzenie nagłówków kolumn
$Worksheet.Cells.Item(1, 1) = "Name"              # Kolumna A - nazwa użytkownika
$Worksheet.Cells.Item(1, 2) = "Exchange Status"   # Kolumna B - status Exchange/targetAddress

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Wczytanie listy użytkowników z pliku tekstowego
# Każda linia pliku to jedna nazwa użytkownika
$users = gc "c:\dell\users.txt"

# Pętla przechodząca przez wszystkich użytkowników z listy
foreach ($user in $users)
{
    # Zapis nazwy użytkownika w kolumnie A
    $Worksheet.Cells.Item($row, 1) = $user
    
    # Pobranie atrybutu targetAddress z Active Directory
    # targetAddress - atrybut przechowujący docelowy adres pocztowy użytkownika
    # Pusta wartość oznacza skrzynkę w bieżącym Exchange
    # Wartość w formacie "SMTP:adres@domena.com" oznacza przekierowanie
    $Exchange = Get-ADUser -Identity "$user" -Properties * | Select targetAddress -expandproperty targetAddress
    
    # Zapis wartości targetAddress w kolumnie B
    $Worksheet.Cells.Item($row, 2) = "$Exchange"
    
    # Dopasowanie szerokości kolumn po wpisaniu danych
    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    
    # Przejście do następnego wiersza
    $row++
}