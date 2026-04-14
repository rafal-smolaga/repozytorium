# Wczytanie listy nazw komputerów z pliku tekstowego
# Plik powinien zawierać jedną nazwę komputera w każdym wierszu
$Computers = Get-Content "ścieżka do pliku .txt"

# Rozpoczęcie pętli przechodzącej przez każdy komputer z listy
foreach($Computer in $computers)
{
    # Usunięcie komputera z Active Directory:
    # -Identity     : nazwa komputera do usunięcia
    # -Confirm:$False : pominięcie pytania o potwierdzenie (cicha operacja)
    Remove-ADComputer -Identity "$Computer" -Confirm:$False
    
    # Wyświetlenie informacji o usuniętym komputerze w konsoli
    Write-Output "$Computer removed"
}

pause
