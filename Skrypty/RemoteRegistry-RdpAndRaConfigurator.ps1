# wczytanie listy komputerów z pliku źródłowego
$Computers = Get-Content "C:\ExampleRemoteTool\example_computers_list.txt"

# pętla po każdym komputerze
ForEach ($Computer in $Computers) 
{
    # test połączenia ping (1 pakiet, bez komunikatów błędów)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
    
    if($ping){
        Write-Host "$computer Responding" -BackgroundColor "Green"
        
        # definicja ścieżek rejestru dla zdalnej pomocy i RDP
        $ValuesKeyPath = 'Example\Registry\Path\RemoteAssistance'
        $ValuesKeyPath1 = 'Example\Registry\Path\TerminalServer'
        $ValuesKeyPath2 = 'Example\Registry\Path\TerminalServer\RDP-Tcp'
        
        $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
        
        # otwarcie kluczy rejestru na zdalnym komputerze
        $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
        $ValueNames = Foreach($val in $ValuesKeyNames.GetValueNames()){$val}

        $ValuesKeyNames1 = $reg.OpenSubKey($ValuesKeyPath1)
        $ValueNames1 = Foreach($val in $ValuesKeyNames1.GetValueNames()){$val}

        $ValuesKeyNames2 = $reg.OpenSubKey($ValuesKeyPath2)
        $ValueNames2 = Foreach($val in $ValuesKeyNames2.GetValueNames()){$val}

        # === konfiguracja fAllowFullControl (pełna kontrola zdalnej pomocy) ===
        if($ValueNames -match 'example_AllowFullControl_Value')
        {
            $ValueFC = 'example_AllowFullControl_Value'
            $resultFC = $ValuesKeyNames.GetValue($ValueFC)
            if($resultFC -eq '1')
            {
                Write-Host "AllowFullControl value is already 1 on $computer" -BackgroundColor "Blue"
            }
            if($resultFC -ne '1')
            {
                reg add "\\$computer\HKLM\Example\Registry\Path\RemoteAssistance" /v "example_AllowFullControl_Value" /t reg_dword /d 00000001 /f | Out-Null
                Write-Host "AllowFullControl value changed to 1 on $computer" -BackgroundColor "Green"
            }
        }
        else
        {
            reg add "\\$computer\HKLM\Example\Registry\Path\RemoteAssistance" /v "example_AllowFullControl_Value" /t reg_dword /d 00000001 /f | Out-Null
            Write-Host "AllowFullControl created and value has been set to 1 on $computer" -BackgroundColor "Green"
        }

        # === konfiguracja fAllowToGetHelp (możliwość otrzymania pomocy zdalnej) ===
        if($ValueNames -match 'example_AllowToGetHelp_Value')
        {
            $ValueATGH = 'example_AllowToGetHelp_Value'
            $resultATGH = $ValuesKeyNames.GetValue($ValueATGH)
            if($resultATGH -eq '1')
            {
                Write-Host "AllowToGetHelp value is already 1 on $computer" -BackgroundColor "Blue"
            }
            if($resultATGH -ne '1')
            {
                reg add "\\$computer\HKLM\Example\Registry\Path\RemoteAssistance" /v "example_AllowToGetHelp_Value" /t reg_dword /d 00000001 /f | Out-Null
                Write-Host "AllowToGetHelp value changed to 1 on $computer" -BackgroundColor "Green"
            }
        }
        else
        {
            reg add "\\$computer\HKLM\Example\Registry\Path\RemoteAssistance" /v "example_AllowToGetHelp_Value" /t reg_dword /d 00000001 /f | Out-Null
            Write-Host "AllowToGetHelp created and value has been set to 1 on $computer" -BackgroundColor "Green"
        }

        # === konfiguracja fDenyTSConnections (odmowa połączeń Terminal Services) ===
        if($ValueNames1 -match 'example_DenyTSConnections_Value')
        {
            $ValueDC = 'example_DenyTSConnections_Value'
            $resultDC = $ValuesKeyNames1.GetValue($ValueDC)
            if($resultDC -eq '0')
            {
                Write-Host "DenyTSConnections value is already 0 on $computer" -BackgroundColor "Blue"
            }
            if($resultDC -ne '0')
            {
                reg add "\\$computer\HKLM\Example\Registry\Path\TerminalServer" /v "example_DenyTSConnections_Value" /t reg_dword /d 00000000 /f | Out-Null
                Write-Host "DenyTSConnections value changed to 0 on $computer" -BackgroundColor "Green"
            }
        }
        else
        {
            reg add "\\$computer\HKLM\Example\Registry\Path\TerminalServer" /v "example_DenyTSConnections_Value" /t reg_dword /d 00000000 /f | Out-Null
            Write-Host "DenyTSConnections created and value has been set to 0 on $computer" -BackgroundColor "Green"
        }

        # === konfiguracja UserAuthentication (uwierzytelnianie użytkownika RDP) ===
        if($ValueNames2 -match 'example_UserAuthentication_Value')
        {
            $ValueUA = 'example_UserAuthentication_Value'
            $resultUA = $ValuesKeyNames2.GetValue($ValueUA)
            if($resultUA -eq '0')
            {
                Write-Host "UserAuthentication is already 0 on $computer" -BackgroundColor "Blue"
            }
            if($resultUA -ne '0')
            {
                reg add "\\$computer\HKLM\Example\Registry\Path\TerminalServer\WinStations\RDP-Tcp" /v "example_UserAuthentication_Value" /t reg_dword /d 00000000 /f | Out-Null
                Write-Host "UserAuthentication value changed to 0 on $computer" -BackgroundColor "Green"
            }
        }
        else
        {
            reg add "\\$computer\HKLM\Example\Registry\Path\TerminalServer\WinStations\RDP-Tcp" /v "example_UserAuthentication_Value" /t reg_dword /d 00000000 /f | Out-Null
            Write-Host "UserAuthentication created and value has been set to 0 on $computer" -BackgroundColor "Green"
        }
    }
    else {
        Write-Host "$computer does not Responding" -BackgroundColor "Red"
    }
}