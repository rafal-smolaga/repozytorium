# Przykład konwersji stringa z przecinkami do formatu binarnego (dla rejestru)
# Oryginalny string z przecinkami (przykładowe dane zamiast długiego ciągu)
$originalString = "9f,03,00,00,4c,00,00,00,01,14,02,00,00,00,00,00,c0,00"

# Usunięcie przecinków i znaków nowej linii z oryginalnego stringa
# Najpierw usuwamy backslashe i spacje (znaki kontynuacji linii)
$modifiedString = $originalString -replace '\\\s+', ""
# Następnie usuwamy wszystkie przecinki
$modifiedString = $modifiedString -replace ",", ""

# Wyświetlenie wyników
Write-Host "Oryginalny string z przecinkami: $originalString"
Write-Host ""
Write-Host "String po usunięciu przecinków: $modifiedString"
Write-Host ""

# Skopiowanie wyniku do schowka (można wkleić do reg add)
$modifiedString | clip
Write-Host "Skopiowano do schowka!"

# Przykład użycia w reg add (wartość binarna)
# reg add "HKLM\SOFTWARE\Example" /v BinaryValue /t REG_BINARY /d $modifiedString /f