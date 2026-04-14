# Wczytanie listy nazw komputerów z pliku tekstowego
# Plik powinien zawierać jedną nazwę komputera w każdym wierszu
$Computers = Get-Content "ścieżka do pliku .txt"

# Rozpoczęcie pętli przechodzącej przez każdy komputer z listy
ForEach ($Computer in $Computers) {
    
    # Wysłanie pojedynczego pinga do komputera (1 pakiet)
    # -ea silentlycontinue - pomija błędy (np. brak odpowiedzi)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    # Jeśli ping zakończył się sukcesem (komputer odpowiada)
    if($ping)
    {
        # Usunięcie wpisu z rejestru na zdalnym komputerze:
        # reg delete     - usuwanie wpisu z rejestru
        # \\$computer\   - ścieżka do zdalnego komputera
        # HKLM\...\Run\  - klucz autostartu
        # /v "nazwa"    - nazwa wartości (klucza) do usunięcia
        # /f             - wymuszenie usunięcia bez potwierdzenia
        reg delete "\\$computer\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" /v "nazwa" /f
        
        # Wyświetlenie informacji o powodzeniu na zielonym tle
        Write-Host "Registry entry removed on $computer" -BackgroundColor "Green"
    }
    else
    {
        # Wyświetlenie informacji o braku dostępu do komputera na czerwonym tle
        Write-Host "Computer offline or unreachable: $computer" -BackgroundColor "Red"
    }
}

pause