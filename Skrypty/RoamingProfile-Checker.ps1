# wczytanie listy użytkowników z pliku źródłowego
$users = get-content "C:\ExamplePath\example_users_list.txt"

# pętla po każdym użytkowniku
foreach($user in $users){
    # pobranie informacji z Active Directory - atrybut profilepath (ścieżka profilu roamingowego)
    $userprofile = get-aduser $user -Properties profilepath -Server exampledomain.example.com | select name, profilepath
    
    # wyświetlenie informacji w konsoli
    # jeśli ścieżka profilu istnieje - użytkownik ma profil roamingowy
    if ($userprofile.profilepath -ne $null) {
        write-host "$user is roaming" -backgroundcolor Green -ForegroundColor DarkBlue
    }
    else {
        # brak ścieżki profilu - użytkownik nie ma profilu roamingowego (lokalny)
        write-host "$user is NOT roaming" -backgroundcolor Red -ForegroundColor DarkBlue
    }
}