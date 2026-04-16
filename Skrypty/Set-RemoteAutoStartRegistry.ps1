# Wczytanie listy nazw komputerów z pliku tekstowego
# Plik powinien zawierać jedną nazwę komputera w każdym wierszu
$Computers = Get-Content "ścieżka do pliku .txt"

# Rozpoczęcie pętli przechodzącej przez każdy komputer z listy
ForEach ($Computer in $Computers) {
    
    # Wysłanie pojedynczego pinga do komputera (1 pakiet)
    # -ea silentlycontinue - pomija błędy (np. brak odpowiedzi)
    # Wynik zapisywany w zmiennej $ping (wartość $true lub $false)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Jeśli ping zakończył się sukcesem (komputer odpowiada)
    if($ping)
    {
        # Dodanie wpisu do rejestru na zdalnym komputerze:
        # reg add        - dodawanie wpisu w rejestrze
        # \\$computer\   - ścieżka do zdalnego komputera
        # HKLM\...\Run\  - klucz autostartu (uruchamianie przy starcie systemu)
        # /f             - wymuszenie nadpisania istniejącego wpisu
        # /v nazwa       - nazwa wartości (klucza) w rejestrze
        # /t REG_SZ      - typ wartości (ciąg tekstowy)
        # /d "ścieżka"   - dane (ścieżka do programu)
        reg add "\\$computer\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" /f /v nazwa /t REG_SZ /d "ścieżka do uruchomienia"
        
        # Wyświetlenie informacji o powodzeniu na zielonym tle
        Write-Host "Registry changed on $computer" -BackgroundColor "Green"
    }
    else
    {
        # Wyświetlenie informacji o braku dostępu do komputera na czerwonym tle
        Write-Host "Computer offline or unreachable: $computer" -BackgroundColor "Red"
    }
}

pause
