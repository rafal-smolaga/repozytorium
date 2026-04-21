# Wczytanie listy komputerów z pliku tekstowego
$computers = gc "ścieżka do pliku .txt"

# Pętla przechodząca przez każdy komputer z listy
ForEach ($Computer in $Computers) {
    
    # Test połączenia z komputerem (jeden pakiet ping)
    # -ea silentlycontinue - pomija błędy (np. brak odpowiedzi)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping)
    {
        # Uruchomienie narzędzia DelProf2.exe do usuwania starych profili użytkowników
        # -c:$computer - zdalny komputer docelowy
        # -d:30 - usuwanie profili nieużywanych od 30 dni
        # -ntuserini - usuwanie również pliku NTUSER.INI
        # -ed:adm* - wykluczenie profili administratorów (wszystkie zaczynające się od "adm")
        # -u - tryb cichy (bez interakcji z użytkownikiem)
        Start-Process -FilePath 'C:\ManufacturingAccountsManager\DelProf2.exe' -ArgumentList "-c:$computer -d:30 -ntuserini -ed:adm* -u"
        
        # Wyświetlenie informacji o przetwarzaniu komputera (zielone tło)
        Write-Host "Processing $Computer..." -BackgroundColor "Green"
    }
    else
    {
        # Komputer nie odpowiada - wyświetlenie błędu (czerwone tło)
        Write-Host "$Computer Not Responding" -BackgroundColor "Red"
    }
}