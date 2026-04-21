# wczytanie listy komputerów z pliku źródłowego
$Computers = Get-Content "C:\ExamplePath\ExampleAccountsManager\computers.txt"

# pętla po każdym komputerze z listy
ForEach ($Computer in $Computers)
{
    # reset zmiennej ping przed każdym testem
    $ping = $null
    
    # test połączenia z komputerem (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    if($ping)
    {
        # komputer odpowiada - wyświetlenie komunikatu
        Write-Host "$Computer responding" -BackgroundColor "Green"
        
        # pobranie sesji użytkowników na zdalnym komputerze (quser)
        $quserResult = quser /server:$computer 2>&1
        
        # zamiana wielokrotnych spacji na przecinki (formatowanie CSV)
        $quserRegex = $quserResult | ForEach-Object -Process { $_ -replace '\s{2,}',',' }
        
        # konwersja na obiekt CSV
        $quserObject = $quserRegex | ConvertFrom-Csv
        
        # filtrowanie sesji dla konkretnego użytkownika i pobranie nazwy sesji
        $quserFinal = $quserObject | Where-Object { $_.username -match "example_username_pattern" } | Select username, sessionname -ExpandProperty sessionname
        
        if($quserFinal)
        {
            # wylogowanie znalezionej sesji
            logoff $quserFinal /server:$Computer
            Write-Host "Session terminated on $Computer" -BackgroundColor "Green"
        }
        else
        {
            # brak sesji dla wskazanego użytkownika
            Write-Host "Nothing to terminate on $Computer" -BackgroundColor "Blue"
        }
    }
    else
    {
        # komputer nie odpowiada na ping
        Write-Host "$Computer does not responding" -BackgroundColor "red"
    }
}