# Załadowanie biblioteki do obsługi grafiki
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
# Załadowanie biblioteki do obsługi okien i klawiszy (SendKeys)
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

# Zmienna sterująca pętlą główną
$start = $true

# Pobranie procesów AcroRd32 (Adobe Reader) - błędy ignorowane
$process = Get-Process AcroRd32 2>$null
if($process)
{
    # Zatrzymanie wszystkich procesów Adobe Reader
    $process| Stop-Process -Force
    # Odczekanie 2 sekund na całkowite zamknięcie
    Start-Sleep -s 2
}
$OneHourCounter = 0

Do{
    # CO GODZINĘ (120 iteracji) - ZABICIE PROCESÓW ADOBE READER
    if($OneHourCounter -eq 120)
    {
        $process = Get-Process AcroRd32 2>$null
        if($process)
        {
            $process| Stop-Process -Force
            Start-Sleep -s 2
        }
        $OneHourCounter = 0  # Reset licznika
    }
    

    $process = Get-Process AcroRd32 2>$null
    

    if($process)
    {
        # Liczba plików w folderze Player (tylko pliki, bez folderów)
        $counter = 0
        $counter = (Get-ChildItem "ścieżka do katalogu zawierającego prezentację .pdf" | Where-Object {$_.PSIsContainer -eq $false}).count
        
        # Jeśli w folderze jest więcej niż 1 plik
        if($counter -gt 1)
        {
            # Zabicie procesu Adobe Reader
            $process| Stop-Process -Force
            
            # ARCHIWIZACJA STARYCH PLIKÓW PDF
            # Pobranie wszystkich plików PDF posortowanych według daty utworzenia (malejąco)
            $Check = Get-ChildItem "ścieżka do katalogu zawierającego prezentację .pdf" *.pdf | Where-Object {$_.PSIsContainer -eq $false} | sort creationtime -Descending | select Name,CreationTime
            
            # Pobranie daty i nazwy najnowszego pliku
            $CheckDate = $Check | Select-Object -First 1 | Select CreationTime -ExpandProperty CreationTime
            $CheckName = $Check | Select-Object -First 1 | Select Name -ExpandProperty Name
            
            # Archiwizacja wszystkich plików starszych niż najnowszy
            foreach($File in $Check)
            {
                $FileCD = $File | Select CreationTime -ExpandProperty CreationTime
                if($CheckDate -gt $FileCD)
                {
                    $FileTD = ($File | Select Name -ExpandProperty Name)
                    $Date = Get-Date -Format "MM-dd-yyyy_HH-mm" 
                    # Dodanie daty do nazwy pliku
                    $FileBN = $FileTD.Name.Split('.')[0] +'_' + $Date + ".pdf"
                    # Zmiana nazwy pliku
                    Rename-Item "ścieżka do katalogu zawierającego prezentację .pdf\$FileTD" -NewName "$FileBN" 
                    # Przeniesienie pliku do folderu Archive
                    Get-Item "ścieżka do katalogu zawierającego prezentację .pdf\$FileBN" | Move-Item -Destination \\LOD1MSHOWL4\Player\Archive\
                }
            }
            
            # PRZYGOTOWANIE NAJNOWSZEGO PLIKU DO OTWORZENIA
            $File2Run = $CheckName | Select Name -ExpandProperty Name
            $TCounter = 0
            
            # Sprawdzenie czy plik nie jest zablokowany (pętla do 30 prób co 2 sekundy)
            do
            {
                # Próba otwarcia pliku do zapisu - jeśli się uda, plik jest dostępny
                try { [IO.File]::OpenWrite("ścieżka do katalogu zawierającego prezentację .pdf\$File2Run").close();$FileTest = $true}
                catch {$FileTest = $false;Timeout /NoBreak 2 | Out-Null}
                $TCounter++
            } until($FileTest -eq $true -or $TCounter -eq 30)
            
            if($FileTest -eq $true){
                # Usunięcie polskich znaków i spacji z nazwy pliku
                $File2Translate = $File2Run.Replace("ą","a").Replace("ć","c").Replace("ę","e").Replace("ł","l").Replace("ń","n").Replace("ó","o").Replace("ś","s").Replace("ź","z").Replace("ż","z").Replace(" ","")
                
                # Zmiana nazwy pliku
                Rename-Item "\\LOD1MSHOWL4\Player\$File2Run" -NewName "$File2Translate" 
                $FinalFile2Run = $File2Translate
                
                # Uruchomienie Adobe Reader z maksymalizowanym oknem
                start-process "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" "ścieżka do katalogu zawierającego prezentację .pdf\$FinalFile2Run" -WindowStyle Maximized
                
                # Odczekanie 6 sekund na uruchomienie Adobe Reader
                Start-Sleep -s 6
                
                # Wysłanie skrótu klawiszowego Ctrl+L (przejście do paska adresu)
                [System.Windows.Forms.SendKeys]::SendWait("^{l}")
            }
        }
    }
    else
    {
        # JEŚLI PROCES ADOBE READER NIE ISTNIEJE
        
        # Liczba plików w folderze Player
        $counter = (Get-ChildItem "ścieżka do katalogu zawierającego prezentację .pdf" | Where-Object {$_.PSIsContainer -eq $false}).count
        
        # Jeśli w folderze jest więcej niż 1 plik
        if($counter -gt 1)
        {
            # Archiwizacja starych plików (taki sam kod jak wyżej)
            $Check = Get-ChildItem "ścieżka do katalogu zawierającego prezentację .pdf" *.pdf | Where-Object {$_.PSIsContainer -eq $false} | sort creationtime -Descending | select Name,CreationTime
            $CheckDate = $Check | Select-Object -First 1 | Select CreationTime -ExpandProperty CreationTime
            $CheckName = $Check | Select-Object -First 1 | Select Name -ExpandProperty Name
            
            foreach($File in $Check)
            {
                $FileCD = $File | Select CreationTime -ExpandProperty CreationTime
                if($CheckDate -gt $FileCD)
                {
                    $FileTD = ($File | Select Name -ExpandProperty Name)
                    $Date = Get-Date -Format "MM-dd-yyyy_HH-mm" 
                    $FileBN = $FileTD.Name.Split('.')[0] +'_' + $Date + ".pdf"
                    Rename-Item "ścieżka do katalogu zawierającego prezentację .pdf\$FileTD" -NewName "$FileBN" 
                    Get-Item "ścieżka do katalogu zawierającego prezentację .pdf\$FileBN" | Move-Item -Destination \\LOD1MSHOWL4\Player\Archive\
                }
            }
            
            # Otwarcie najnowszego pliku
            $File2Run = $CheckName | Select Name -ExpandProperty Name
            $TCounter = 0
            do
            {
                try { [IO.File]::OpenWrite("ścieżka do katalogu zawierającego prezentację .pdf\$File2Run").close();$FileTest = $true}
                catch {$FileTest = $false;Timeout /NoBreak 2 | Out-Null}
                $TCounter++
            } until($FileTest -eq $true -or $TCounter -eq 30)
            
            if($FileTest -eq $true){
                $File2Translate = $File2Run.Replace("ą","a").Replace("ć","c").Replace("ę","e").Replace("ł","l").Replace("ń","n").Replace("ó","o").Replace("ś","s").Replace("ź","z").Replace("ż","z").Replace(" ","")
                Rename-Item "\\LOD1MSHOWL4\Player\$File2Run" -NewName "$File2Translate" 
                $FinalFile2Run = $File2Translate
                start-process "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" "ścieżka do katalogu zawierającego prezentację .pdf\$FinalFile2Run" -WindowStyle Maximized
                Start-Sleep -s 6
                [System.Windows.Forms.SendKeys]::SendWait("^{l}")
            }
        }
        
        # JEŚLI W FOLDERZE JEST 1 PLIK
        if($counter -eq 1)
        {
            # Otwarcie jedynego pliku w folderze
            $Check = Get-ChildItem "ścieżka do katalogu zawierającego prezentację .pdf" *.pdf | Where-Object {$_.PSIsContainer -eq $false} | sort creationtime -Descending | select Name,CreationTime
            $CheckName = $Check | Select-Object -First 1 | Select Name -ExpandProperty Name
            $File2Run = $CheckName | Select Name -ExpandProperty Name
            $TCounter = 0
            do
            {
                try { [IO.File]::OpenWrite("ścieżka do katalogu zawierającego prezentację .pdf\$File2Run").close();$FileTest = $true}
                catch {$FileTest = $false;Timeout /NoBreak 2 | Out-Null}
                $TCounter++
            } until($FileTest -eq $true -or $TCounter -eq 30)
            
            if($FileTest -eq $true){
                $File2Translate = $File2Run.Replace("ą","a").Replace("ć","c").Replace("ę","e").Replace("ł","l").Replace("ń","n").Replace("ó","o").Replace("ś","s").Replace("ź","z").Replace("ż","z").Replace(" ","")
                Rename-Item "\\LOD1MSHOWL4\Player\$File2Run" -NewName "$File2Translate" 
                $FinalFile2Run = $File2Translate
                start-process "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" "ścieżka do katalogu zawierającego prezentację .pdf\$FinalFile2Run" -WindowStyle Maximized
                Start-Sleep -s 6
                [System.Windows.Forms.SendKeys]::SendWait("^{l}")
            }
        }
    }
    
    # SPRAWDZENIE CZASU DZIAŁANIA SYSTEMU I RESTART o 3 AM
    $GD = $null
    $GD = Get-Date -Format 'htt'  # Pobranie przybliżonej godziny (AM/PM)
    
    # Jeśli jest około 3 nad ranem
    if($GD -eq "3AM")
    {
        # Obliczenie godzin od ostatniego restartu
        $LR = (get-date) - (gcim Win32_OperatingSystem).LastBootUpTime | Select TotalHours -ExpandProperty TotalHours
        
        # Jeśli system działa dłużej niż 20 godzin
        if($LR -gt 20)
        {
            # Restart systemu (natychmiastowy, wymuszony)
            shutdown -r -t 0 /f
        }
    }
    
    # OCZEKIWANIE PRZED NASTĘPNĄ ITERACJĄ
    Start-Sleep -s 30  # Odczekanie 30 sekund
    $OneHourCounter++   # Zwiększenie licznika godzinowego

} While($start)  # Nieskończona pętla ($start zawsze $true)