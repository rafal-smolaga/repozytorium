# Odczytanie listy nazw komputerów z pliku tekstowego
# Plik powinien zawierać jedną nazwę komputera w każdym wierszu
$Computers = Get-Content "ścieżka do pliku .txt"

# USTAWIENIE KOMUNIKATU NA KAŻDYM KOMPUTERZE
foreach ($computer in $computers)
{

    # reg add - dodanie wpisu do rejestru
    # \\$computer\ - ścieżka do rejestru zdalnego komputera
    # HKEY_LOCAL_MACHINE\...\System - klucz rejestru z ustawieniami logowania
    # /v legalnoticecaption - nazwa wartości (tytuł okna komunikatu)
    # /t reg_sz - typ danych (ciąg tekstowy)
    # /d Informacja: - dane (treść tytułu)
    # /f - wymuszenie nadpisania bez pytania o potwierdzenie
    reg add "\\$computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v legalnoticecaption /t reg_sz /d Informacja: /f
    

    # legalnoticetext - właściwa treść komunikatu wyświetlana użytkownikowi
    # Komunikat informuje o planowanej wymianie komputera i konieczności
    # zabezpieczenia danych przed wymianą
    reg add "\\$computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v legalnoticetext /t reg_sz /d "Stacja robocza zostanie wkrótce wymieniona na nową jednostkę, uprzejmie proszę o zabezpieczenie swoich danych przechowywanych na komputerze. Nacisnij Enter aby kontynuować." /f
}