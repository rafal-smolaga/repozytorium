# Biblioteki
Add-Type -AssemblyName System.Windows.Forms  # Interfejs graficzny
Add-Type -AssemblyName System.Drawing        # Grafika i czcionki
Add-Type -AssemblyName PresentationCore,PresentationFramework  # Okna dialogowe

# SPRAWDZENIE UPRAWNIEŃ
$currentUser = $env:UserName

# Jeśli użytkownik zaczyna się od "adm" (np. admin, administrator)
if ($currentUser.startswith("adm","CurrentCultureIgnoreCase")) {
    # Pytanie o potwierdzenie przed kontynuacją
    $Result = [System.Windows.MessageBox]::Show("Are you agree to check that everything working correctly after use?" , "Confirmation" , 4)
    
    if($result -eq "Yes") {
        # Kontynuuj
    }
    if($result -eq "No") {
        break  # Przerwij działanie skryptu
    }
}
else {
    # Użytkownik nie ma uprawnień administracyjnych - ostrzeżenie
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "UAC Issue"
    $Messageboxbody = "Are you sure that you run this script as administrator?"   
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
    [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
    break  # Przerwij działanie skryptu
}

# KONFIGURACJA OKNA GŁÓWNEGO (FORMULARZA)
$Form1 = New-Object system.Windows.Forms.Form
$Icon = New-Object system.drawing.icon ("ścieżka do pliku Outlook2013.ico")  # Ikona aplikacji
$Form1.Text = "Outlook2013Fix"
$Form1.Icon = $Icon
$Image = [system.drawing.image]::FromFile("ścieżka do pliku Outlook2013Fix.png")  # Obraz tła
$Form1.StartPosition = "CenterScreen"
$Form1.BackgroundImage = $Image
$Form1.BackgroundImageLayout = "Center"
$Form1.Width = $Image.Width
$Form1.Height = $Image.Height
$Form1.MaximizeBox = $false
$Form1.MinimizeBox = $true
$Form1.minimumSize = New-Object System.Drawing.Size(800,600) 
$Form1.maximumSize = New-Object System.Drawing.Size(800,600)

# Zdefiniowanie czcionki dla przycisków
$Font1 = New-Object System.Drawing.Font("Tahoma",14,[System.Drawing.FontStyle]::Bold)

# Etykieta IP ADDRESS
$LabelIP = New-Object System.Windows.Forms.Label
$LabelIP.Font = New-Object System.Drawing.Font("Times New Roman",19,[System.Drawing.FontStyle]::Bold)
$LabelIP.Location = new-object System.Drawing.Size(10,10)
$LabelIP.Text = "IP Address:"
$LabelIP.BackColor = "Transparent"
$LabelIP.AutoSize = $True
$Form1.Controls.Add($LabelIP)

# Pole tekstowe do wprowadzenia adresu IP
$IPABox = New-Object System.Windows.Forms.TextBox
$IPABox.Location = New-Object -TypeName System.Drawing.Point -ArgumentList(10,60)
$IPABox.Size = New-Object -TypeName System.Drawing.Point -ArgumentList (280,25)
$IPABox.Autosize = $True
$FontSA = New-Object System.Drawing.Font("Arial",16,[System.Drawing.FontStyle]::Bold)
$IPABox.Font = $FontSA

# Walidacja - tylko cyfry i kropki (format IP)
$IPABox.Add_TextChanged({
    $this.Text = $this.Text -replace '[^1234567890.]'
})
$Form1.Controls.Add($IPABox)

# Etykieta USER NAME (pole tekstowe)
$LabelUN = New-Object System.Windows.Forms.Label
$LabelUN.Font = New-Object System.Drawing.Font("Times New Roman",19,[System.Drawing.FontStyle]::Bold)
$LabelUN.Location = new-object System.Drawing.Size(10,100)
$LabelUN.Text = "UserName:"
$LabelUN.BackColor = "Transparent"
$LabelUN.AutoSize = $True
$Form1.Controls.Add($LabelUN)

$UserNameBox = New-Object System.Windows.Forms.TextBox
$UserNameBox.Location = New-Object -TypeName System.Drawing.Point -ArgumentList(10,150)
$UserNameBox.Size = New-Object -TypeName System.Drawing.Point -ArgumentList (280,25)
$UserNameBox.Autosize = $True
$UserNameBox.Font = $FontSA

# Walidacja - dozwolone małe litery, cyfry, podkreślnik i myślnik
$UserNameBox.Add_TextChanged({
    $this.Text = $this.Text -replace '[^a-z1234567890_-]'
})
$Form1.Controls.Add($UserNameBox)

# Etykieta: OUTPUT BOX (pole wyjściowe z wynikami)
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Size(358,110) 
$outputBox.Size = New-Object System.Drawing.Size(400,350) 
$outputBox.MultiLine = $True
$outputBox.ReadOnly = $True
$outputBox.Font = New-Object System.Drawing.Font("Calibri",11,[System.drawing.FontStyle]::Bold)
$outputBox.ScrollBars = "Vertical"
$Form1.Controls.Add($outputBox)

# PRZYCISK: PROCEED - główna akcja naprawy
$ButtonP = New-Object System.Windows.Forms.Button
$ButtonP.Font = $Font1
$ButtonP.Autosize = $True
$ButtonP.Location = New-Object System.Drawing.Size(10,210)
$ButtonP.Name = "Proceed"
$ButtonP.Text = "Proceed"
$ButtonP.BackColor = 'LightGray'
$ButtonP.Add_Click({PButton})
$Form1.Controls.Add($ButtonP)

# PRZYCISK: CLEAR VALUES - czyszczenie pól IP i UserName
$ButtonCV = New-Object System.Windows.Forms.Button
$ButtonCV.Font = $Font1
$ButtonCV.Autosize = $True
$ButtonCV.Location = New-Object System.Drawing.Size(130,210)
$ButtonCV.Name = "Clear Values"
$ButtonCV.Text = "Clear Values"
$ButtonCV.BackColor = 'LightGray'
$ButtonCV.Add_Click({CVButton})
$Form1.Controls.Add($ButtonCV)

# PRZYCISK: LIST LOGGED USERS - wyświetlenie zalogowanych użytkowników
$ButtonLLU = New-Object System.Windows.Forms.Button
$ButtonLLU.Font = $Font1
$ButtonLLU.Autosize = $True
$ButtonLLU.Location = New-Object System.Drawing.Size(10,420)
$ButtonLLU.Name = "List logged users"
$ButtonLLU.Text = "List logged users"
$ButtonLLU.BackColor = 'LightGray'
$ButtonLLU.Add_Click({LLUButton})
$Form1.Controls.Add($ButtonLLU)

# PRZYCISK: LIST OFFICE2013 UPDATES - wyświetlenie aktualizacji Office
$ButtonUpdates = New-Object System.Windows.Forms.Button
$ButtonUpdates.Font = $Font1
$ButtonUpdates.Autosize = $True
$ButtonUpdates.Location = New-Object System.Drawing.Size(10,360)
$ButtonUpdates.Name = "List Office2013 Updates"
$ButtonUpdates.Text = "List Office2013 Updates"
$ButtonUpdates.BackColor = 'LightGray'
$ButtonUpdates.Add_Click({UButton})
$Form1.Controls.Add($ButtonUpdates)

# PRZYCISK: CLEAR - czyszczenie okna output
$ClearButton = New-Object "System.Windows.Forms.Button"
$ClearButton.Location = New-Object System.Drawing.Size(650,500)
$ClearButton.Autosize = $True
$ClearButton.Font = $Font1
$ClearButton.BackColor = 'LightGray'
$ClearButton.Text = "Clear"
$ClearButton.Add_Click{$outputBox.Clear()}
$Form1.Controls.Add($ClearButton)

# FUNKCJA: LLUButton - lista zalogowanych użytkowników
function LLUButton()
{
    $IPA = ''
    [string]$IPA = $IPABox.Text
    
    # Sprawdzenie czy pole IP jest puste
    if($IPA -eq ''){
        $outputbox.SelectionColor='black'
        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
        $outputbox.AppendText("$gd")
        $outputBox.AppendText("`n")
        $outputbox.SelectionColor='red'
        $outputbox.AppendText("IP Address is empty!")
        $outputBox.AppendText("`n")
    }
    
    if($IPA -ne ''){
        # Sprawdzenie czy IP zaczyna się od 10 lub 172 (sieci prywatne)
        if(($IPA.Startswith("10"))-or($IPA.Startswith("172"))){
            # Sprawdzenie formatu IP (kropki)
            if($IPA -like ("*.*.*.*")) {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Checking")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $IPA")
                $outputBox.AppendText("`n")
                
                # Test ping do komputera
                $ping = Test-Connection $IPA -Count 1 -ea silentlycontinue
                
                if($ping) {
                    # Komputer odpowiada - pobranie informacji
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("$IPA")
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText(" Responding")
                    $outputBox.AppendText("`n") 

                    # Odczyt nazwy komputera z rejestru zdalnego
                    $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $IPA)
                    $ValuesKeyPath = 'SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName'
                    $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
                    $ValueNames = Foreach($val in $ValuesKeyNames.GetValueNames()){$val}
                    $ValueHN = 'Computername'
                    $ComputerHN = $ValuesKeyNames.GetValue($ValueHN)
                    
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Hostname:")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("$ComputerHN")
                    $outputBox.AppendText("`n")

                    # Pobranie listy zalogowanych użytkowników (quser)
                    $quserResult = quser /server:$IPA 2>&1
                    $quserRegex = $quserResult | ForEach-Object -Process { $_ -replace '\s{2,}',',' }
                    $quserObject = $quserRegex | ConvertFrom-Csv
                    $quserFinal = $quserObject | Select UserName -expandproperty UserName
                    
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n")
                    $outputbox.SelectionColor='DarkBlue'
                    $outputbox.AppendText("Currently Logged Users:")
                    $outputBox.AppendText("`n")
                    
                    if($quserFinal -eq $null) {
                        $outputbox.SelectionColor='blue'
                        $outputbox.AppendText("No logged users")
                        $outputBox.AppendText("`n")
                    } else {
                        Foreach($User in $quserFinal) {
                            $outputbox.SelectionColor='DarkCyan'
                            $outputbox.AppendText("$User")
                            $outputBox.AppendText("`n")
                        }
                    }
                } else {
                    # Komputer nie odpowiada na ping
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n")
                    $outputbox.SelectionColor='red'
                    $outputbox.AppendText("$IPA does not responding")
                    $outputBox.AppendText("`n")
                } 
            } else {
                # Nieprawidłowy format IP
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText("Incorrect IP Dot-decimal notation")
                $outputBox.AppendText("`n")
            }
        } else {
            # IP nie należy do dozwolonych zakresów (10.x lub 172.x)
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n")
            $outputbox.SelectionColor='red'
            $outputbox.AppendText("Incorrect IP Address Format")
            $outputBox.AppendText("`n")
        }
    }
}

# FUNKCJA: UButton - lista aktualizacji Office 2013
function UButton()
{
    $IPA = ''
    [string]$IPA = $IPABox.Text
    
    # Sprawdzenie czy pole IP jest puste
    if($IPA -eq ''){
        $outputbox.SelectionColor='black'
        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
        $outputbox.AppendText("$gd")
        $outputBox.AppendText("`n")
        $outputbox.SelectionColor='red'
        $outputbox.AppendText("IP Address is empty!")
        $outputBox.AppendText("`n")
    }
    
    if($IPA -ne ''){
        if(($IPA.Startswith("10"))-or($IPA.Startswith("172"))){
            if($IPA -like ("*.*.*.*")) {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Checking")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $IPA")
                $outputBox.AppendText("`n")
                
                $ping = Test-Connection $IPA -Count 1 -ea silentlycontinue
                
                if($ping) {
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("$IPA")
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText(" Responding")
                    $outputBox.AppendText("`n") 

                    # Odczyt nazwy komputera z rejestru
                    $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $IPA)
                    $ValuesKeyPath = 'SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName'
                    $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
                    $ValueHN = 'Computername'
                    $ComputerHN = $ValuesKeyNames.GetValue($ValueHN)
                    
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Hostname:")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("$ComputerHN")
                    $outputBox.AppendText("`n")
                    
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("Processing...")
                    $outputBox.AppendText("`n")
                    
                    # Ścieżka do klucza rejestru z zainstalowanymi programami (w tym Office)
                    $SubKeyPath = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
                    $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $IPA)
                    $SubKeyNames = $reg.OpenSubKey($SubKeyPath)
                    $PathSub = Foreach($sub in $SubKeyNames.GetSubKeyNames()){$sub}
                    
                    # Filtrowanie tylko kluczy związanych z Office 2013
                    $OfficeRegistry = $pathSub -match "0000000FF1CE}"
                    $OutputCollection = @()
                    
                    Foreach ($OfficeKB in $OfficeRegistry) {
                        $ValuesKeyPath = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$OfficeKB"
                        $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
                        $ValueNames = Foreach($val in $ValuesKeyNames.GetValueNames()){$val}
                        
                        if($ValueNames -match 'DisplayName') {
                            $output = New-Object -TypeName PSobject
                            $ValueDN = 'DisplayName'
                            $resultDN = $ValuesKeyNames.GetValue($ValueDN)
                            if($resultDN -ne $null -and $resultDN -ne '') {
                                $output | add-member NoteProperty "DisplayName" -value $resultDN
                                $OutputCollection += $output
                            }
                        }
                    }
                    
                    # Wyświetlenie wyników w osobnym oknie GridView
                    $OutputCollection | Out-GridView
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("Complete")
                    $outputBox.AppendText("`n")
                } else {
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n")
                    $outputbox.SelectionColor='red'
                    $outputbox.AppendText("$IPA does not responding")
                    $outputBox.AppendText("`n")
                } 
            } else {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText("Incorrect IP Dot-decimal notation")
                $outputBox.AppendText("`n")
            }
        } else {
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n")
            $outputbox.SelectionColor='red'
            $outputbox.AppendText("Incorrect IP Address Format")
            $outputBox.AppendText("`n")
        }
    }
}

# FUNKCJA: CVButton - czyszczenie pól IP i UserName
function CVButton()
{
    $IPABox.Text = ''      # Wyczyść pole IP
    $IPA = ''              # Wyczyść zmienną IP
    $UserNameBox.Text = '' # Wyczyść pole nazwy użytkownika
    $UN = ''               # Wyczyść zmienną nazwy użytkownika
}

# FUNKCJA: PButton - główna funkcja naprawy Outlook
# Zmiana EnableADAL, tworzenie zadania harmonogramu, kopiowanie plików
function PButton()
{
    $IPA = ''
    [string]$IPA = $IPABox.Text
    $UN = ''
    [string]$UN = $UserNameBox.Text
    
    # Walidacja pól wejściowych
    if(($IPA -eq '')-and($UN -eq '')){
        $outputbox.SelectionColor='black'
        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
        $outputbox.AppendText("$gd")
        $outputBox.AppendText("`n")
        $outputbox.SelectionColor='red'
        $outputbox.AppendText("IP Address and UserName Fields are empty!")
        $outputBox.AppendText("`n")
    }
    
    if(($IPA -eq '') -and ($UN -ne '')){
        $outputbox.SelectionColor='black'
        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
        $outputbox.AppendText("$gd")
        $outputBox.AppendText("`n")
        $outputbox.SelectionColor='red'
        $outputbox.AppendText("IP Address is empty!")
        $outputBox.AppendText("`n")
    }
    
    if(($IPA -ne '') -and ($UN -eq '')){
        $outputbox.SelectionColor='black'
        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
        $outputbox.AppendText("$gd")
        $outputBox.AppendText("`n")
        $outputbox.SelectionColor='red'
        $outputbox.AppendText("UserName is empty!")
        $outputBox.AppendText("`n")
    }
    
    if(($IPA -ne '') -and ($UN -ne '')){
        if(($IPA.Startswith("10"))-or($IPA.Startswith("172"))){
            if($IPA -like ("*.*.*.*")) {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Checking")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $IPA")
                $outputBox.AppendText("`n")
                
                $ping = Test-Connection $IPA -Count 1 -ea silentlycontinue
                
                if($ping) {
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("$IPA")
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText(" Responding")
                    $outputBox.AppendText("`n") 

                    # Odczyt nazwy komputera z rejestru
                    $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $IPA)
                    $ValuesKeyPath = 'SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName'
                    $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
                    $ValueHN = 'Computername'
                    $ComputerHN = $ValuesKeyNames.GetValue($ValueHN)
                    
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n")
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Hostname:")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText("$ComputerHN")
                    $outputBox.AppendText("`n")

                    # Sprawdzenie czy użytkownik jest zalogowany
                    $quserResult = quser /server:$IPA 2>&1
                    $quserRegex = $quserResult | ForEach-Object -Process { $_ -replace '\s{2,}',',' }
                    $quserObject = $quserRegex | ConvertFrom-Csv
                    $quserFinal = $quserObject | Select UserName -expandproperty UserName

                    if($quserFinal -eq $UN) {
                        # Pobranie ID sesji użytkownika
                        $UserIDSession = $quserObject | Select UserName, ID | where Username -eq $UN | Select ID -expandproperty ID
                        
                        if(($UserIDSession -eq "disc")-or ($UserIDSession -eq '')) {
                            $outputbox.SelectionColor='Red'
                            $outputbox.AppendText("$UN should be logged on to proceed those steps!")
                            $outputBox.AppendText("`n")
                        } else {
                            # Pobranie SID użytkownika
                            $SID = ([System.Security.Principal.NTAccount]("domena\$UN")).Translate([System.Security.Principal.SecurityIdentifier]).Value
                            
                            # Zmiana wartości EnableADAL w rejestrze zdalnego komputera
                            reg add "\\$IPA\HKEY_USERS\$SID\Software\Microsoft\Office\15.0\Common\Identity" /v EnableADAL /t reg_dword /d 00000000 /f
                            
                            # Sprawdzenie czy Outlook działa i zabicie procesu
                            $TL = Tasklist /S $IPA /fi "Session eq $UserIDSession"
                            $TLB = $TL | ForEach-Object -Process { $_ -replace '=','' }
                            $TLR = $TLB | ForEach-Object -Process { $_ -replace '\s{2,}',',' }
                            $TLR1 = $TLR | ForEach-Object -Process { $_ -replace 'console','' }
                            $TLO = $TLR1 | ConvertFrom-Csv
                            $TLF = $TLO | Select "Image Name", "PID Session Name" | where "Image Name" -eq "Outlook.exe"
                            $PIDSN = $TLF | Select "PID Session Name" -expandproperty "PID Session Name"

                            if($PIDSN -match '^[0-9]') {
                                $PTK = taskkill /S $IPA /PID $PIDSN /T
                                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                                $outputbox.AppendText("$gd")
                                $outputBox.AppendText("`n")
                                $outputbox.SelectionColor='Blue'
                                $outputbox.AppendText("EnableADAL successfully changed to 0")
                                $outputBox.AppendText("`n")
                                $outputbox.SelectionColor='Blue'
                                $outputbox.AppendText("Outlook Closed On Profile ")
                                $outputbox.SelectionColor='Green'
                                $outputbox.AppendText("$UN")
                                $outputBox.AppendText("`n")
                            } else {
                                $outputbox.SelectionColor='Blue'
                                $outputbox.AppendText("EnableADAL successfully changed to 0")
                                $outputBox.AppendText("`n")
                                $outputbox.SelectionColor='Blue'
                                $outputbox.AppendText("No Outlook Session to kill")
                                $outputBox.AppendText("`n")
                            }

                            # TWORZENIE ZADANIA HARMONOGRAMU (XML)
                            # Zadanie uruchamia OutlookFix.vbs przy logowaniu lub odblokowaniu
                            $xml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2020-11-28T10:48:05.1898377</Date>
    <Author>IT</Author>
    <URI>\$UN</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <StartBoundary>$SD</StartBoundary>
      <EndBoundary>$ED</EndBoundary>
      <Enabled>true</Enabled>
      <UserId>EUROPE\$UN</UserId>
    </LogonTrigger>
    <SessionStateChangeTrigger>
      <StartBoundary>$SD</StartBoundary>
      <EndBoundary>$ED</EndBoundary>
      <Enabled>true</Enabled>
      <StateChange>SessionUnlock</StateChange>
      <UserId>EUROPE\$UN</UserId>
    </SessionStateChangeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$SID</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Dell\OutlookFix.vbs</Command>
    </Exec>
  </Actions>
</Task>
'@

                            # Generowanie daty ważności zadania (10-20 dni od teraz)
                            $random = Get-Random -Minimum 10 -Maximum 20
                            $SD = Get-Date -UFormat "%Y-%m-%dT%R:%S"
                            $ED = (Get-date $SD).AddDays($random).ToString("yyyy-MM-ddTH:mm:ss")
                            $TD = (Get-date $SD).AddDays($random).ToString("yyyy-MM-dd H:mm:ss")
                            
                            # Podstawienie zmiennych do XML
                            $xmlPrep = $xml.Replace('$UN', $UN)
                            $xmlPrep = $xmlPrep.Replace('$SD', $SD)
                            $xmlPrep = $xmlPrep.Replace('$ED', $ED)
                            $xmlPrep = $xmlPrep.Replace('$SID', $SID)
                            $xmlFinal = [xml]$xmlPrep
                            
                            # Zapisanie XML lokalnie i na serwerze raportów
                            $xmlFinal.save("Ścieżka do zapisu pliku\$UN.xml")
                            $dts = Get-Date -Format "MM-dd-yyyy-HH-mm-ss"
                            $currentUser = $env:UserName
                            $OutlookFixRaport = "$ComputerHN" + '_' + "$UN" + '_' + "$dts" + '_' + "$currentUser"
                            $xmlFinal.save("ścieżka do raportu\$OutlookFixRaport.xml")
                            
                            # Usunięcie starego zadania jeśli istnieje
                            $TaskTest = Get-ChildItem -Path \\$IPA\c$\Windows\System32\Tasks | Select Name -expandproperty Name
                            if($TaskTest -match "$UN") {
                                SCHTASKS /Delete /S $IPA /TN $UN /F
                                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                                $outputbox.AppendText("$gd")
                                $outputBox.AppendText("`n")
                                $outputbox.SelectionColor='Blue'
                                $outputbox.AppendText("Old Task Deleted for user ")
                                $outputbox.SelectionColor='Green'
                                $outputbox.AppendText("$UN")
                                $outputBox.AppendText("`n")
                            }

                            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                            $outputbox.AppendText("$gd")
                            $outputBox.AppendText("`n")
                            $outputbox.SelectionColor='Blue'
                            $outputbox.AppendText("Task Generated for user ")
                            $outputbox.SelectionColor='Green'
                            $outputbox.AppendText("$UN")
                            $outputBox.AppendText("`n")
                            
                            # Kopiowanie plików pomocniczych na zdalny komputer
                            $SourceFTP1 = "Ścieżka do zapisu pliku\OutlookFix.vbs"
                            $SourceFTP2 = "Ścieżka do zapisu pliku\OutlookFixTask.ps1"
                            $DestinationFTP = "\\$IPA\c$\Temp"
                            Copy-Item -Path $SourceFTP1 -Destination $DestinationFTP -Force
                            Copy-Item -Path $SourceFTP2 -Destination $DestinationFTP -Force
                            
                            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                            $outputbox.AppendText("$gd")
                            $outputBox.AppendText("`n")
                            $outputbox.SelectionColor='Blue'
                            $outputbox.AppendText("File ")
                            $outputbox.SelectionColor='Green'
                            $outputbox.AppendText("OutlookFix files ")
                            $outputbox.SelectionColor='Blue'
                            $outputbox.AppendText("copied to C:\Temp")
                            $outputBox.AppendText("`n")

                            # Utworzenie nowego zadania w harmonogramie
                            schtasks.exe /Create /s $IPA /XML "Ścieżka do zapisu pliku\$UN.xml" /tn $UN
                            
                            $TaskTest = Get-ChildItem -Path \\$IPA\c$\Windows\System32\Tasks | Select Name -expandproperty Name
                            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                            $outputbox.AppendText("$gd")
                            $outputBox.AppendText("`n")
                            $outputbox.SelectionColor='Blue'
                            $outputbox.AppendText("New Task Created for user ")
                            $outputbox.SelectionColor='Green'
                            $outputbox.AppendText("$UN")
                            $outputBox.AppendText("`n")
                            $outputbox.SelectionColor='Blue'
                            $outputbox.AppendText("New Task is valid until ")
                            $outputbox.SelectionColor='Green'
                            $outputbox.AppendText("$TD")
                            $outputBox.AppendText("`n")

                            # Usunięcie lokalnego pliku XML po utworzeniu zadania
                            if($TaskTest -match "$UN") {
                                Remove-Item -Path "Ścieżka do zapisu pliku\$UN.xml" -Force
                            }
                        }
                    } else {
                        $outputbox.SelectionColor='Red'
                        $outputbox.AppendText("$UN is not logged in!")
                        $outputBox.AppendText("`n")
                    }
                } else {
                    $outputbox.SelectionColor='black'
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n")
                    $outputbox.SelectionColor='red'
                    $outputbox.AppendText("$IPA does not responding")
                    $outputBox.AppendText("`n")
                } 
            } else {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText("Incorrect IP Dot-decimal notation")
                $outputBox.AppendText("`n")
            }
        } else {
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n")
            $outputbox.SelectionColor='red'
            $outputbox.AppendText("Incorrect IP Address Format")
            $outputBox.AppendText("`n")
        }
    }
}
$Form1.ShowDialog()