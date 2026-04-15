# Ścieżka do pliku wynikowego Excel
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

# Ustawienie widoczności Excela (TRUE = pokaż okno Excela)
# Dzięki temu wyniki są widoczne w czasie rzeczywistym
$Excel.Visible = $True


$Worksheet.Cells.Item(1, 1) = "IP Adress"        # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"           # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "HDD Smart Status" # Kolumna C - status SMART dysku

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# WCZYTANIE LISTY KOMPUTERÓW
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
        
        # Pobranie statusu dysku twardego za pomocą WMI
        # Win32_DiskDrive - klasa WMI reprezentująca fizyczne dyski twarde
        # Status - właściwość określająca stan dysku (Ok, Degraded, Pred Fail itp.)
        $diskStatus = Get-WmiObject win32_diskdrive -ComputerName $computer | Select-Object -ExpandProperty Status
        
        # Sprawdzenie statusu dysku
        If ($diskStatus -eq "Ok")
        {
            # Jeśli dysk jest zdrowy:
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 4
            $Worksheet.Cells.Item($row, 3) = "OK"
        }
        else
        {
            # Jeśli dysk ma problemy lub jest uszkodzony:
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 3
            $Worksheet.Cells.Item($row, 3) = "Fail"
        }
        
        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        $row++
    }
    else
    {
        
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
