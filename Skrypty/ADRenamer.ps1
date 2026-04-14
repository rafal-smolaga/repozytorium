# Importowanie modułu ImportExcel który umożliwia odczyt plików Excel
Import-Module ImportExcel

# Pobranie aktualnej nazwy zalogowanego użytkownika (domena\użytkownik)
$whoami = whoami

# Pobranie poświadczeń (login i hasło) użytkownika z dostępem do domeny
# Globalna zmienna będzie dostępna w całej sesji PowerShell
$global:admc = Get-Credential -credential $whoami

# Ścieżka do pliku Excel zawierającego pary starych i nowych nazw komputerów wpisane w kolumnach obok siebie z nagłówkami OLD oraz NEW
$FilePath = 'ścieżka do pliku .xlsx'

# Import danych z pliku Excel
$importcomputers = Import-Excel -Path $FilePath

# Wyodrębnienie tylko starych nazw komputerów (kolumna OLD)
$machines = $importcomputers | Select OLD, NEW -ExpandProperty OLD

# Pętla przechodząca przez każdy stary komputer
foreach($machine in $machines)
{
    # Pobranie nowej nazwy dla bieżącego komputera (kolumna New)
    $newmachine = $machine | Select OLD, NEW -ExpandProperty New
    
    # Wyświetlenie informacji o zmianie nazwy dla danego komputera
    write-host "old:$machine, new:$newmachine"
    
    # Zmiana nazwy komputera w domenie:
    # -ComputerName: stara nazwa komputera
    # -NewName: nowa nazwa komputera  
    # -DomainCredential: poświadczenia z dostępem do domeny
    # -Force: wymuszenie zmiany nawet jeśli wystąpią problemy
    # -Restart: automatyczne ponowne uruchomienie komputera po zmianie
    Rename-Computer -ComputerName "$machine" -NewName "$newmachine" -DomainCredential $global:admc -Force -Restart
}
pause