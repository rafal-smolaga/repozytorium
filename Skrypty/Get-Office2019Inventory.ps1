# BIBLIOTEKI I ZALEŻNOŚCI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# POBIERANIE KOMPUTERÓW Z ACTIVE DIRECTORY

# Serwer AD (do modyfikacji - należy wpisać właściwy serwer)
# Przykład: "domena.com", "dc01.domena.local"
$adServer = "domena.com"

# Ścieżki bazowe do jednostek organizacyjnych (do modyfikacji)
$basePath = "OU=Workstations,OU=Computers,OU=DEPARTMENT,OU=REGION,DC=domain,DC=company,DC=com"
$ouWin11 = "OU=Win11,$basePath"
$ouWin10 = "OU=Win10,$basePath"

# Pobranie komputerów z pierwszej jednostki organizacyjnej (OU=Common)
$OU1 = Get-ADComputer -Server $adServer -Filter * -SearchBase $ouWin11 | select Name -ExpandProperty Name

# Pobranie komputerów z drugiej jednostki organizacyjnej (OU=Win10)
$OU2 = Get-ADComputer -Server $adServer -Filter * -SearchBase $ouWin10 | select Name -ExpandProperty Name

# Połączenie obu list i sortowanie alfabetyczne
$computers = $OU1 + $OU2 | Sort-Object

# Inicjalizacja listy wynikowej
$results = New-Object System.Collections.ArrayList

# ============================================================
# GŁÓWNA PĘTLA SKANUJĄCA KOMPUTERY
# ============================================================
ForEach ($Computer in $Computers) 
{
    # Test połączenia z komputerem (jeden pakiet ping)
    $ping = Test-Connection $computer -Count 1 -ea silentlycontinue

    # Wyświetlenie informacji o postępie w konsoli
    Write-Host "Processing $Computer..." -ForegroundColor "Yellow" -BackgroundColor "DarkCyan"

    # Jeśli komputer odpowiada na ping (jest dostępny)
    if($ping){
        $resultV = $null
        
        # SPRAWDZENIE OBECNOŚCI OFFICE 2019 W REJESTRZE
        
        # Ścieżka do klucza rejestru z listą zainstalowanych programów
        $SubKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        
        # Otwarcie połączenia z rejestrem zdalnego komputera
        $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer)
        
        # Otwarcie podklucza i pobranie nazw wszystkich podkluczy
        $SubKeyNames = $reg.OpenSubKey($SubKeyPath)
        $RemoteControlSub = Foreach($sub in $SubKeyNames.GetSubKeyNames()){$sub}

        # Sprawdzenie czy istnieje klucz odpowiadający Office 2019
        # GUID: {90160000-008C-0000-1000-0000000FF1CE} to identyfikator Office 2019
        if($RemoteControlSub -eq '{90160000-008C-0000-1000-0000000FF1CE}'){
            
            # Ścieżka do klucza rejestru Office 2019
            $ValuesKeyPath = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-008C-0000-1000-0000000FF1CE}'
            $ValuesKeyNames = $reg.OpenSubKey($ValuesKeyPath)
            $ValueNames = Foreach($val in $ValuesKeyNames.GetValueNames()){$val}
            
            # Odczytanie wartości DisplayVersion (wersja Office)
            $ValueV = 'DisplayVersion'
            $resultV = $ValuesKeyNames.GetValue($ValueV)
            
            if($resultV)
            {
                # Office 2019 zainstalowany - znaleziono wersję
                $data = [pscustomobject]@{
                    'IP Address' = $Computer
                    'Status' = "UP"
                    'Office 2019' = 'YES'
                    'Version' = $resultV
                }
            }
            else
            {
                # Office 2019 zainstalowany ale nie udało się odczytać wersji
                $data = [pscustomobject]@{
                    'IP Address' = $Computer
                    'Status' = "UP"
                    'Office 2019' = 'YES'
                    'Version' = 'Unknown'
                }
            }
        }
        else
        {
            # Office 2019 NIE jest zainstalowany
            $data = [pscustomobject]@{
                'IP Address' = $Computer
                'Status' = "UP"
                'Office 2019' = 'No'
                'Version' = "-"
            }
        } 

        # Dodanie wyniku do listy
        [void]$results.Add($data)
    }
    else
    {
        # Komputer NIE odpowiada na ping (jest niedostępny)
        $data = [pscustomobject]@{
            'IP Address' = $Computer
            'Status' = "DOWN"
            'Office 2019' = "-"
            'Version' = "-"
        }

        [void]$results.Add($data)
    }
}

# TWORZENIE OKNA Z WYNIKAMI (DataGridView)

$FormR = New-Object System.Windows.Forms.Form
$FormR.StartPosition = "CenterScreen"
$FormR.Text = 'Workstations Information'
$FormR.Size = New-Object System.Drawing.Size(1000, 400)
$FormR.FormBorderStyle = 'FixedSingle'
$FormR.MaximizeBox = $false
$FormR.MinimizeBox = $false
$FormR.minimumSize = New-Object System.Drawing.Size(1000,400) 
$FormR.maximumSize = New-Object System.Drawing.Size(1000,400)

# Panel z przewijaniem
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Fill'
$panel.AutoScroll = $true
$FormR.Controls.Add($panel)

# Kontrolka DataGridView do wyświetlania wyników
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(50, 300)
$dataGridView.Dock = 'Top'
$dataGridView.AutoSizeColumnsMode = 'Fill'
$panel.Controls.Add($dataGridView)

# Tworzenie tabeli danych
$dataTable = New-Object System.Data.DataTable
$dataTable.Columns.Add('IP Address') | Out-Null
$dataTable.Columns.Add('Status') | Out-Null
$dataTable.Columns.Add('Office 2019') | Out-Null
$dataTable.Columns.Add('Version') | Out-Null

# Wypełnienie tabeli danymi z wyników skanowania
$results | ForEach-Object {
    $row = $dataTable.NewRow()
    $row['IP Address'] = $_.'IP Address'
    $row['Status'] = $_.Status
    $row['Office 2019'] = $_.'Office 2019'
    $row['Version'] = $_.Version
    $dataTable.Rows.Add($row)
}

# Przypisanie tabeli do DataGridView
$dataGridView.DataSource = $dataTable

# PRZYCISK COPY - kopiowanie danych do schowka
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Autosize = $True
$copyButton.Location = New-Object System.Drawing.Size(300,320)
$copyButton.Text = 'Copy'
$copyButton.BackColor = 'LightGray'
$copyButton.UseVisualStyleBackColor = $True
$copyButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$copyButton.Add_Click({
    $data = New-Object System.Text.StringBuilder
    $dataGridView.Rows | ForEach-Object {
        $row = $_
        $row.Cells | ForEach-Object {
            $data.Append("$($_.Value)`t")
        }
        $data.AppendLine()
    }
    [System.Windows.Forms.Clipboard]::SetText($data.ToString())
})
$FormR.AcceptButton = $copyButton
$panel.Controls.Add($copyButton)

# Wyświetlenie okna z wynikami
$FormR.ShowDialog() | Out-Null