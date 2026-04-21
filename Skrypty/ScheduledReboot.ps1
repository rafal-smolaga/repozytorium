# wczytanie listy komputerów z pliku źródłowego
$computers = gc "C:\ExamplePath\example_computers_list.txt"

# pętla po każdym komputerze
foreach($computer in $computers)
{
    # wykonanie restartu zdalnego komputera z opóźnieniem
    # /r: restart, /m: komputer zdalny, /t: czas opóźnienia w sekundach (example_value)
    shutdown /r /m \\$computer /t example_delay_seconds
    
    # komunikat potwierdzający wysłanie polecenia
    write-host "Restart command sent to $computer" -BackgroundColor "Green"
}