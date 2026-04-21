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
$Worksheet.Cells.Item(1, 1) = "IP Adress"      # Kolumna A - nazwa komputera
$Worksheet.Cells.Item(1, 2) = "Status"         # Kolumna B - status (UP/DOWN)
$Worksheet.Cells.Item(1, 3) = "Folder2Backup"           # Kolumna C - czy folder istnieje

# Rozpoczęcie od wiersza nr 2 (ponieważ wiersz 1 to nagłówki)
$row = 2

# Automatyczne dopasowanie szerokości kolumn do zawartości
[void]$Worksheet.UsedRange.EntireColumn.AutoFit()

# Pobranie komputerów z pierwszej jednostki organizacyjnej (przykładowa ścieżka AD)
$OU1 = Get-ADComputer -Filter * -SearchBase "OU=ExampleOU1,OU=Workstations,OU=Computers,OU=ExampleDepartment,OU=ExampleRegion,DC=example,DC=company,DC=com" | select Name -ExpandProperty Name

# Pobranie komputerów z drugiej jednostki organizacyjnej (przykładowa ścieżka AD)
$OU2 = Get-ADComputer -Filter * -SearchBase "OU=ExampleOU2,OU=Workstations,OU=Computers,OU=ExampleDepartment,OU=ExampleRegion,DC=example,DC=company,DC=com" | select Name -ExpandProperty Name

# Połączenie obu list i sortowanie alfabetyczne
$computers = $OU1 + $OU2 | Sort-Object

# Główna pętla przechodząca przez wszystkie komputery
foreach ($computer in $computers)
{
    # Test połączenia z komputerem (jeden pakiet ping, ignorowanie błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    $result = $null
    
    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping){
        # Zapis nazwy komputera w kolumnie A
        $Worksheet.Cells.Item($row, 1) = $computer
        
        # Zapis statusu "UP" w kolumnie B
        $Worksheet.Cells.Item($row, 2) = "UP"
        
        # Sprawdzenie czy folder example_folder istnieje na dysku C:\
        $result = Get-ChildItem "\\$Computer\c$" -Force | Select Name -ExpandProperty Name | where name -match 'example_folder'
        
        # Jeśli folder istnieje
        if($result)
        {
            # Kolor zielony (ColorIndex 4) dla statusu "Yes"
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 4
            $Worksheet.Cells.Item($row, 3) = "Yes"
            
            # Ścieżka do serwera backup (przykładowa)
            $backupServer = "\\server.example.company.com\share\BackupFolder"
            
            # Sprawdzenie czy folder dla danego komputera już istnieje na serwerze
            $tp = Test-Path "$backupServer\$computer"
            
            # Jeśli folder na serwerze istnieje
            if($tp)
            {
                # Pobranie bieżącej daty w formacie _mm_dd_YYYY
                $gd = Get-Date -UFormat "_%m_%d_%Y"
                
                # Sprawdzenie czy plik zip z dzisiejszą datą już istnieje
                $ftp = Test-Path "$backupServer\$computer\$computer$gd.zip"
                
                # Jeśli plik zip z dzisiejszą datą istnieje - dodaj numer wersji
                if($ftp)
                {
                    $counter = 1
                    $newname = "$backupServer\$computer\$computer$gd" + "_" + $counter + ".zip"
                    
                    # Zwiększaj licznik aż do znalezienia wolnej nazwy
                    while(Test-Path $newname)
                    {
                        $counter++
                        $newname = "$backupServer\$computer\$computer$gd" + "_" + $counter + ".zip"
                    }
                    $Gd = $gd + "_" + $counter
                    $GCIS = Get-ChildItem -Path "\\$Computer\c$\ExampleFolder\*" -Force
                    $GCID = "$backupServer\$computer\$computer$gd"
                    
                    # Kompresja folderu za pomocą 7-zip
                    & "C:\Program Files\7-zip\7z.exe" a -tzip $GCID $GCIS > $null
                }
                else
                {
                    # Plik zip nie istnieje - tworzenie nowego bez numeru
                    $GCIS = Get-ChildItem -Path "\\$Computer\c$\ExampleFolder\*" -Force
                    $GCID = "$backupServer\$computer\$computer$gd"
                    
                    # Kompresja folderu za pomocą 7-zip
                    & "C:\Program Files\7-zip\7z.exe" a -tzip $GCID $GCIS > $null
                }
            }
            else
            {
                # Folder dla komputera nie istnieje na serwerze - utwórz go
                New-Item -Path "$backupServer\$computer" -ItemType Directory | Out-null
                
                # Pobranie bieżącej daty
                $gd = Get-Date -UFormat "_%m_%d_%Y"
                
                # Pobranie zawartości folderu ExampleFolder
                $GCIS = Get-ChildItem -Path "\\$Computer\c$\ExampleFolder\*" -Force
                $GCID = "$backupServer\$computer\$computer$gd"
                
                # Kompresja folderu za pomocą 7-zip
                & "C:\Program Files\7-zip\7z.exe" a -tzip $GCID $GCIS > $null
            }
        }
        else
        {
            # Folder nie istnieje - kolor czerwony (ColorIndex 3) i wpis "No"
            $Worksheet.Cells.Item($row, 3).Font.ColorIndex = 3
            $Worksheet.Cells.Item($row, 3) = "No"
        }
        
        # Dopasowanie szerokości kolumn po wpisaniu danych
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
    }
    else {
        # Komputer nie odpowiada na ping - zapis statusu "DOWN"
        $Worksheet.Cells.Item($row, 1) = $computer
        $Worksheet.Cells.Item($row, 2) = "DOWN"
        
        # Dopasowanie szerokości kolumn
        [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
        
        # Przejście do następnego wiersza
        $row++
    }
}