# Bezpośrednia zmiana wartości EnableADAL na 0 dla bieżącego użytkownika
# Ścieżka: HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\Identity
# EnableADAL = 0 - wyłączenie nowoczesnego uwierzytelniania
Set-Itemproperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\Identity -Name "EnableADAL" -Value 0

# Pobranie bieżącej nazwy użytkownika (domena\użytkownik)
$User = whoami

# Pobranie identyfikatora SID (Security Identifier) dla bieżącego użytkownika
# SID jest unikalnym identyfikatorem użytkownika w domenie/lokalnym systemie
$TSID = ([System.Security.Principal.NTAccount]("$User")).Translate([System.Security.Principal.SecurityIdentifier]).Value

# Zmiana wartości EnableADAL w gałęzi HKEY_USERS dla konkretnego użytkownika
# reg add - dodawanie wpisu w rejestrze
# HKEY_USERS\$TSID\... - ścieżka do klucza rejestru dla danego SID
# /v EnableADAL - nazwa wartości
# /t reg_dword - typ danych (32-bitowa liczba całkowita)
# /d 00000000 - dane (wartość 0)
# /f - wymuszenie nadpisania bez pytania o potwierdzenie
reg add "HKEY_USERS\$TSID\Software\Microsoft\Office\15.0\Common\Identity" /v EnableADAL /t reg_dword /d 00000000 /f


# Sprawdzenie czy proces Outlook jest uruchomiony w bieżącej sesji użytkownika
# Get-Process -Name outlook - pobranie wszystkich procesów Outlook
# where { $_.SessionId -eq ([System.Diagnostics.Process]::GetCurrentProcess().SessionId) }
# - filtrowanie tylko tych procesów, których SessionId zgadza się z bieżącą sesją
$ProcessActive = Get-Process -Name outlook | where { $_.SessionId -eq ([System.Diagnostics.Process]::GetCurrentProcess().SessionId) }

# Pętla wykonuje się dopóki proces Outlook NIE zostanie wykryty ($ProcessActive -ne $null)
# Oznacza to, że skrypt będzie czekał, aż użytkownik uruchomi Outlook
Do {
    
    # WEWNĘTRZNA PĘTLA DO...UNTIL - powtarzanie zmiany rejestru
    $Counter = 0  # Inicjalizacja licznika
    
    Do {
        # Ponowna zmiana wartości EnableADAL (na wypadek gdyby coś ją nadpisało)
        Set-Itemproperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\Identity -Name "EnableADAL" -Value 0
        
        # Ponowna zmiana w gałęzi HKEY_USERS
        reg add "HKEY_USERS\$TSID\Software\Microsoft\Office\15.0\Common\Identity" /v EnableADAL /t reg_dword /d 00000000 /f
        
        $Counter++
        
        # Odczekanie 2 sekund przed kolejną iteracją
        Start-Sleep -s 2
    }
    Until ($Counter -eq 5)  # Wykonaj 5 razy (łącznie 10 sekund oczekiwania)
    
    # Po wykonaniu 5 prób, ponownie sprawdź czy Outlook jest uruchomiony
    $ProcessActive = Get-Process -Name outlook | where { $_.SessionId -eq ([System.Diagnostics.Process]::GetCurrentProcess().SessionId) }
}
Until ($ProcessActive -ne $null)  # Zakończ pętlę gdy Outlook zostanie wykryty

# UWAGA: Po wykryciu Outlooka skrypt kończy działanie
# Wartość EnableADAL została już ustawiona na 0 przed uruchomieniem Outlook