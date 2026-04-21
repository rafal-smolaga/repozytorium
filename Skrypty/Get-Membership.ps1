
# Nazwa użytkownika do sprawdzenia (należy wpisać właściwą nazwę)
# Przykład: "j_kowalski", "a_nowak"
$user = "login"

# Nazwa grupy AD do sprawdzenia (należy wpisać właściwą nazwę)
# Przykład: "Citrix_Users", "Domain Admins", "IT_Team"
$group = "nazwa_grupy"

# KONFIGURACJA SERWERA AD (DO MODYFIKACJI)
# Serwer domenowy Active Directory (należy wpisać właściwą nazwę)
# Przykład: "domena.com", "dc01.domena.local", "10.10.10.10"
$adServer = "domena.com"

# SPRAWDZANIE CZŁONKOSTWA W GRUPIE

# Pobranie wszystkich członków grupy (rekurencyjnie - również z zagnieżdżonych grup)
# Get-ADGroupMember - pobiera członków grupy AD
# -Server $adServer - określa który kontroler domeny ma być użyty
# -Identity $group - nazwa grupy do sprawdzenia
# -Recursive - uwzględnia również użytkowników z zagnieżdżonych grup
# Select -ExpandProperty Name - wyodrębnia tylko nazwy użytkowników
$members = Get-ADGroupMember -Server $adServer -Identity $group -Recursive | Select -ExpandProperty Name

# Sprawdzenie czy użytkownik znajduje się na liście członków grupy
If ($members -contains $user) {
    # Użytkownik należy do grupy - komunikat zielony
    Write-Host "$user exists in the group" -ForegroundColor "Green"
} Else {
    # Użytkownik NIE należy do grupy - komunikat czerwony
    Write-Host "$user not exists in the group" -ForegroundColor "Red"
}