# pobranie listy komputerów z pierwszej jednostki organizacyjnej (OU_Common)
# z pominięciem maszyn o nazwach pasujących do wzorca "ExamplePrefix*"
$OU1 = Get-ADComputer -Server exampledomain.example.com -Filter {Name -notlike "ExamplePrefix*"} -SearchBase "OU=ExampleCommon,OU=ExampleWorkstations,OU=ExampleComputers,OU=ExampleEMFP,OU=ExampleManufacturing,DC=exampledomain,DC=example,DC=com" | select Name -expandproperty Name

# pobranie listy komputerów z drugiej jednostki organizacyjnej (OU_Win10)
$OU2 = Get-ADComputer -Server exampledomain.example.com -Filter * -SearchBase "OU=ExampleWin10,OU=ExampleWorkstations,OU=ExampleComputers,OU=ExampleEMFP,OU=ExampleManufacturing,DC=exampledomain,DC=example,DC=com" | select Name -expandproperty Name

# połączenie i sortowanie listy komputerów
$computers = $OU1 + $OU2 | Sort-Object Name

# pętla po każdym komputerze
ForEach ($Computer in $Computers) {
    # test połączenia ping (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    if($ping)
    {
        # wykonanie restartu zdalnego komputera
        # -r: restart, -t x: brak opóźnienia (x jako przykład), -f: wymuszenie zamknięcia aplikacji, -m: komputer zdalny
        shutdown -r -t example_time_value -f /m $computer
        Write-Host "Restart Scheduled on $Computer" -BackgroundColor "Green"
    }
    else
    { 
        Write-Host "$Computer Not Responding" -BackgroundColor "Red"
    }
}