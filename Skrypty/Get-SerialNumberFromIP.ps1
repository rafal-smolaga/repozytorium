# Ścieżka do pliku wynikowego Excel
$path = ".\results.xls"

# Utworzenie obiektu aplikacji Excel
$Excel = new-object -comobject excel.application

# Sprawdzenie czy plik wynikowy już istnieje
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

# Ustawienie widoczności Excela (TRUE = pokaż okno Excela)
# Dzięki temu wyniki są widoczne w czasie rzeczywistym
$Excel.Visible = $True

$Worksheet.Cells.Item(1, 1) = "IP Address"    # Kolumna A - adres IP
$Worksheet.Cells.Item(1, 2) = "Status"        # Kolumna C - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "SerialNumber"  # Kolumna B - numer seryjny


# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# WCZYTANIE LISTY ADRESÓW IP
$IPs = gc "ścieżka do pliku .txt"


foreach ($IP in $IPs)
{
    # Zapis adresu IP w kolumnie A (wiersz bieżący)
    $Worksheet.Cells.Item($row, 1) = $IP
    
    # Test połączenia z komputerem (2 pakiety ping)
    $ping = (Test-Connection -ComputerName "$IP" -Count 2).IPV4Address
    
    # Jeśli ping zakończył się sukcesem (komputer odpowiada)
    if($ping){
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # Pobranie numeru seryjnego komputera z BIOS/WMI
        # Klasa Win32_BIOS zawiera informacje o płycie głównej/urządzeniu
        # SerialNumber to numer seryjny (dla Dell, HP, Lenovo itp.)
        $SerialNumber = (Get-WmiObject win32_bios -computername $IP).serialnumber
        $Worksheet.Cells.Item($row, 3) = "$SerialNumber"
        
        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        $row++
    }
    else
    {
        # Jeśli komputer NIE odpowiada na ping (jest nieosiągalny)
        
        # Zapis statusu "Down" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "Down"
        
        # Dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        $row++
    }
}
