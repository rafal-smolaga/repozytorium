# Załadowanie biblioteki Windows Forms do tworzenia okien
Add-Type -AssemblyName System.Windows.Forms

# Włączenie nowoczesnego stylu wizualnego dla przycisków i kontrolek
[Windows.Forms.Application]::EnableVisualStyles()

# Utworzenie głównego okna aplikacji
$form = New-Object System.Windows.Forms.Form

# Ustawienie tytułu okna
$form.Text = "Computer Name and Account Selection"

# Ustawienie szerokości (400) i wysokości (950) okna
$form.Size = New-Object System.Drawing.Size(400, 950)

# Okno ma się pojawić na środku ekranu
$form.StartPosition = "CenterScreen"

# Blokada zmiany rozmiaru okna - stały rozmiar
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Ukrycie przycisku minimalizacji
$form.MinimizeBox = $false

# Ukrycie przycisku maksymalizacji
$form.MaximizeBox = $false

# Ukrycie ikony w pasku tytułowym okna
$form.ShowIcon = $false

# Definicja funkcji do odczytywania właściwości skrótów .lnk
function Get-Shortcut {
    # Parametr wejściowy - ścieżka do pliku skrótu
    param (
        [string]$ShortcutPath
    )
    # Utworzenie obiektu WScript.Shell (COM) do obsługi skrótów
    $WshShell = New-Object -ComObject WScript.Shell
    # Otwarcie skrótu i pobranie jego właściwości
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    # Zwrócenie obiektu skrótu
    return $Shortcut
}

# Utworzenie kontrolki TabControl (pozwala na tworzenie zakładek)
$tabControl = New-Object System.Windows.Forms.TabControl

# Ustawienie szerokości zakładek na 380, wysokości na 900
$tabControl.Size = New-Object System.Drawing.Size(380, 900)

# Ustawienie pozycji zakładek (10 pikseli od prawej, 10 od góry)
$tabControl.Location = New-Object System.Drawing.Point(10, 10)

# Utworzenie pierwszej zakładki
$tabMain = New-Object System.Windows.Forms.TabPage

# Nadanie nazwy zakładce - "Main"
$tabMain.Text = "Main"

# Dodanie zakładki do kontrolki TabControl
$tabControl.TabPages.Add($tabMain)

# Utworzenie etykiety (label) dla pola wprowadzania nazw komputerów
$computerLabel = New-Object System.Windows.Forms.Label

# Tekst wyświetlany na etykiecie
$computerLabel.Text = "Enter Computer Name:"

# Pozycja etykiety: X=20, Y=20
$computerLabel.Location = New-Object System.Drawing.Point(20, 20)

# Etykieta automatycznie dostosuje szerokość do tekstu
$computerLabel.AutoSize = $true

# Dodanie etykiety do zakładki Main
$tabMain.Controls.Add($computerLabel)

# Utworzenie przycisku Restart
$restartButton = New-Object System.Windows.Forms.Button

# Tekst na przycisku
$restartButton.Text = "Restart"

# Pozycja przycisku: X=20, Y=60
$restartButton.Location = New-Object System.Drawing.Point(20, 60)

# Dodanie przycisku do zakładki Main
$tabMain.Controls.Add($restartButton)

# Obsługa zdarzenia kliknięcia przycisku Restart
$restartButton.Add_Click({
    # Pobranie tekstu z pola komputerów i podzielenie na linie
    $computerNames = $computerTextBox.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

    # Sprawdzenie czy lista komputerów nie jest pusta
    if ([string]::IsNullOrEmpty($computerNames)) {
        # Wyświetlenie ostrzeżenia
        [System.Windows.Forms.MessageBox]::Show("Please enter computer names.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
    else {
        # Pętla po każdym komputerze z listy
        ForEach ($computer in $computerNames) {
            # Wysłanie pinga (1 pakiet) - ignorowanie błędów
            $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
            if($ping) {
                # Komputer odpowiada - dodanie znacznika czasu do logów
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                # Ustawienie koloru tekstu na zielony
                $outputbox.SelectionColor='green'
                $outputbox.AppendText("$computer")
                # Ustawienie koloru tekstu na niebieski
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText(" Responding")
                $outputBox.AppendText("`n")
                # Wykonanie restartu systemu (natychmiast, wymuszony, bez odliczania)
                shutdown /r /f /t 0 /m \\$computer 2>&1
                # Potwierdzenie w logach
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText("$computer")
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText(" restarted")
                $outputBox.AppendText("`n")
                # Przewinięcie okna logów na dół
                $outputbox.SelectionStart = $outputbox.Text.Length
                $outputbox.ScrollToCaret()
            }
            else {
                # Komputer nie odpowiada - błąd w logach
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText("$computer")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText(" does not responding")
                $outputBox.AppendText("`n")
                $outputbox.SelectionStart = $outputbox.Text.Length
                $outputbox.ScrollToCaret()
            }
        }
    }
})

# Utworzenie pola tekstowego do wprowadzania nazw komputerów
$computerTextBox = New-Object System.Windows.Forms.TextBox

# Pozycja pola: X=180, Y=20
$computerTextBox.Location = New-Object System.Drawing.Point(180, 20)

# Rozmiar pola: szerokość 180, wysokość 200
$computerTextBox.Size = New-Object System.Drawing.Size(180, 200)

# Pola akceptuje klawisz Enter (przechodzi do nowej linii)
$computerTextBox.AcceptsReturn = $true

# Pola nie akceptuje klawisza Tab
$computerTextBox.AcceptsTab = $false

# Pole jest wielowierszowe
$computerTextBox.Multiline = $true

# Dodanie pasków przewijania (pionowy i poziomy)
$computerTextBox.ScrollBars = 'Both'

# Ustawienie czcionki: Calibri, rozmiar 11, pogrubiona
$computerTextBox.Font = New-Object System.Drawing.Font("Calibri", 11, [System.Drawing.FontStyle]::Bold)

# Dodanie pola do zakładki Main
$tabMain.Controls.Add($computerTextBox)

# Etykieta dla wyboru konta użytkownika
$accountLabel = New-Object System.Windows.Forms.Label
$accountLabel.Text = "Select Account:"
$accountLabel.Location = New-Object System.Drawing.Point(20, 270)
$accountLabel.AutoSize = $true
$tabMain.Controls.Add($accountLabel)

# Lista rozwijana (ComboBox) z kontami użytkowników
$accountComboBox = New-Object System.Windows.Forms.ComboBox
$accountComboBox.Location = New-Object System.Drawing.Point(180, 270)
$accountComboBox.Size = New-Object System.Drawing.Size(180, 20)
# Tylko do wyboru - nie można wpisać własnej wartości
$accountComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Dodanie przykładowych kont (zanonimizowane)
$accountComboBox.Items.Add("ProcessExample01") | Out-Null
$accountComboBox.Items.Add("ProcessExample02") | Out-Null
$accountComboBox.Items.Add("ServiceExample01") | Out-Null
$accountComboBox.Items.Add("svc_example_viewer") | Out-Null
# Domyślnie wybrane konto
$accountComboBox.SelectedItem = "ServiceExample01"
$tabMain.Controls.Add($accountComboBox)

# Etykieta dla wyboru opcji konfiguracji
$optionLabel = New-Object System.Windows.Forms.Label
$optionLabel.Text = "Select Option:"
$optionLabel.Location = New-Object System.Drawing.Point(20, 310)
$optionLabel.AutoSize = $true
$tabMain.Controls.Add($optionLabel)

# Lista rozwijana z opcjami konfiguracji Taskbar
$optionComboBox = New-Object System.Windows.Forms.ComboBox
$optionComboBox.Location = New-Object System.Drawing.Point(180, 310)
$optionComboBox.Size = New-Object System.Drawing.Size(180, 20)
$optionComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Przykładowe opcje konfiguracji (zanonimizowane)
$optionComboBox.Items.Add("") | out-null
$optionComboBox.Items.Add("Config_Example01") | out-null
$optionComboBox.Items.Add("Config_Example02") | out-null
$optionComboBox.Items.Add("Config_Example03") | out-null
$optionComboBox.SelectedItem = ""
$tabMain.Controls.Add($optionComboBox)

# Przycisk Submit - wykonuje główne operacje
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(20, 450)
$tabMain.Controls.Add($submitButton)

# Checkbox - czy modyfikować preferencje Edge i rejestr
$CheckboxPreferences = New-Object System.Windows.Forms.CheckBox
$CheckboxPreferences.Text = "Preferences/reg"
$CheckboxPreferences.AutoSize = $true
$CheckboxPreferences.Location = New-Object System.Drawing.Point(200, 350)
$tabMain.Controls.Add($CheckboxPreferences)

# Checkbox - czy zabić procesy Edge
$CheckboxTaskkillEdge = New-Object System.Windows.Forms.CheckBox
$CheckboxTaskkillEdge.Text = "Taskkill EDGE"
$CheckboxTaskkillEdge.AutoSize = $true
$CheckboxTaskkillEdge.Location = New-Object System.Drawing.Point(200, 380)
$tabMain.Controls.Add($CheckboxTaskkillEdge)

# Checkbox - czy wylogować użytkownika
$CheckboxLogoff = New-Object System.Windows.Forms.CheckBox
$CheckboxLogoff.Text = "Logoff"
$CheckboxLogoff.AutoSize = $true
$CheckboxLogoff.Location = New-Object System.Drawing.Point(200, 410)
$tabMain.Controls.Add($CheckboxLogoff)

# Checkbox - czy wyczyścić cały folder Startup
$CheckboxCleanEntireStartUp = New-Object System.Windows.Forms.CheckBox
$CheckboxCleanEntireStartUp.Text = "Clean Entire Startup"
$CheckboxCleanEntireStartUp.AutoSize = $true
$CheckboxCleanEntireStartUp.Location = New-Object System.Drawing.Point(200, 440)
$tabMain.Controls.Add($CheckboxCleanEntireStartUp)

# Checkbox - czy wyczyścić tylko skróty IE/Edge ze Startup
$CheckboxCleanIEEDGEStartUp = New-Object System.Windows.Forms.CheckBox
$CheckboxCleanIEEDGEStartUp.Text = "Clean IE/EDGE Startup"
$CheckboxCleanIEEDGEStartUp.AutoSize = $true
$CheckboxCleanIEEDGEStartUp.Location = New-Object System.Drawing.Point(200, 470)
$tabMain.Controls.Add($CheckboxCleanIEEDGEStartUp)

# Logika wzajemnego wykluczania - nie można zaznaczyć obu opcji czyszczenia
$CheckboxCleanEntireStartUp.Add_CheckedChanged({
    if ($CheckboxCleanEntireStartUp.Checked) {
        $CheckboxCleanIEEDGEStartUp.Checked = $false
        $CheckboxCleanIEEDGEStartUp.Enabled = $false
    } else {
        $CheckboxCleanIEEDGEStartUp.Enabled = $true
    }
})

$CheckboxCleanIEEDGEStartUp.Add_CheckedChanged({
    if ($CheckboxCleanIEEDGEStartUp.Checked) {
        $CheckboxCleanEntireStartUp.Checked = $false
        $CheckboxCleanEntireStartUp.Enabled = $false
    } else {
        $CheckboxCleanEntireStartUp.Enabled = $true
    }
})

# Okno wyjściowe (RichTextBox) do wyświetlania logów
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Size(10, 700)
$outputBox.Size = New-Object System.Drawing.Size(350, 170)
$outputBox.MultiLine = $True
$outputBox.ReadOnly = $True
$outputBox.Font = New-Object System.Drawing.Font("Calibri", 11, [System.Drawing.FontStyle]::Bold)
$outputBox.ScrollBars = "Vertical"
$tabMain.Controls.Add($outputBox)

# Checkbox - czy dodawać argumenty uruchomieniowe dla Edge
$CheckBoxEdgeArguments = New-Object System.Windows.Forms.CheckBox
$CheckBoxEdgeArguments.Text = "Edge Arguments"
$CheckBoxEdgeArguments.AutoSize = $true
$CheckBoxEdgeArguments.Location = New-Object System.Drawing.Point(25, 500)

# Obsługa zmiany stanu checkboxa - włącza/wyłącza pola dla argumentów
$CheckBoxEdgeArguments.Add_CheckStateChanged({
    if ($CheckBoxEdgeArguments.Checked) {
        $argumentsTextBox.Enabled = $true
        $fullScreenCheckBox.Enabled = $true
        $startUpCheckBox.Enabled = $true
    } else {
        $argumentsTextBox.Enabled = $false
        $fullScreenCheckBox.Enabled = $false
        $startUpCheckBox.Enabled = $false
        $argumentsTextBox.Text = ""
        $fullScreenCheckBox.Checked = $false
        $startUpCheckBox.Checked = $false
    }
})
$tabMain.Controls.Add($CheckBoxEdgeArguments)

# Checkbox - czy uruchomić Edge w trybie pełnoekranowym
$fullScreenCheckBox = New-Object System.Windows.Forms.CheckBox
$fullScreenCheckBox.Text = "Full Screen"
$fullScreenCheckBox.Location = New-Object System.Drawing.Point(160, 500)
$fullScreenCheckBox.Enabled = $false
$fullScreenCheckBox.Checked = $false
$tabMain.Controls.Add($fullScreenCheckBox)

# Checkbox - czy dodać skrót Edge do folderu Startup
$startUpCheckBox = New-Object System.Windows.Forms.CheckBox
$startUpCheckBox.Text = "StartUp"
$startUpCheckBox.Location = New-Object System.Drawing.Point(265, 500)
$startUpCheckBox.Enabled = $false
$startUpCheckBox.Checked = $false
$startUpCheckBox.Add_CheckStateChanged({
    if ($startUpCheckBox.Checked) {
        $fileNameTextBox.Enabled = $true
    } else {
        $fileNameTextBox.Enabled = $false
        $fileNameTextBox.Text = ""
    }
})
$tabMain.Controls.Add($startUpCheckBox)

# Pole tekstowe na argumenty dla Edge (oddzielone przecinkami)
$argumentsTextBox = New-Object System.Windows.Forms.TextBox
$argumentsTextBox.Location = New-Object System.Drawing.Point(25, 550)
$argumentsTextBox.Size = New-Object System.Drawing.Size(200, 150)
$argumentsTextBox.Enabled = $false
$tabMain.Controls.Add($argumentsTextBox)

# Pole tekstowe na nazwę pliku skrótu w Startup
$fileNameTextBox = New-Object System.Windows.Forms.TextBox
$fileNameTextBox.Location = New-Object System.Drawing.Point(25, 600)
$fileNameTextBox.Size = New-Object System.Drawing.Size(200, 150)
$fileNameTextBox.Enabled = $false

# Walidacja - tylko małe litery, cyfry, spacje i podkreślnik
$fileNameTextBox.Add_TextChanged({
    if ($this.Text -match '[^a-z 0-9_]') {
        $cursorPos = $this.SelectionStart
        $this.Text = $this.Text -replace '[^a-z 0-9_]'
        $this.SelectionStart = $cursorPos - 1
        $this.SelectionLength = 0
    }
})
$tabMain.Controls.Add($fileNameTextBox)

# Przycisk Raport - zbiera informacje o komputerach
$raportButton = New-Object System.Windows.Forms.Button
$raportButton.Text = "Raport"
$raportButton.Location = New-Object System.Drawing.Point(20, 350)
$tabMain.Controls.Add($raportButton)

# Przycisk OpenPaths - otwiera foldery Taskbar i Startup w Eksploratorze
$openPathButton = New-Object System.Windows.Forms.Button
$openPathButton.Text = "OpenPaths"
$openPathButton.Location = New-Object System.Drawing.Point(20, 400)
$openPathButton.AutoSize = $true
$tabMain.Controls.Add($openPathButton)

# Obsługa kliknięcia OpenPaths
$openPathButton.Add_Click({
    $computerNames = $computerTextBox.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    $selectedAccount = $accountComboBox.SelectedItem
    $selectedOption = $optionComboBox.SelectedItem

    if ([string]::IsNullOrEmpty($computerNames)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter computer names.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } elseif ($selectedAccount -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select an account.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } elseif ($selectedOption -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select an option.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        foreach ($computer in $computerNames) {
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd`n")
            $outputbox.SelectionColor='blue'
            $outputbox.AppendText("Checking")
            $outputbox.SelectionColor='green'
            $outputbox.AppendText(" $computer`n")
            
            $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
            if($ping) {
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd`n")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText("$computer")
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText(" Responding`n")
                
                # Pobranie zalogowanego użytkownika przez WMI
                $User = (Get-WmiObject Win32_ComputerSystem -ComputerName $Computer).UserName
                # Usunięcie prefixu domeny (pierwsze 7 znaków)
                $User = $User.Substring(7)
                # Otwarcie folderu Taskbar w Eksploratorze
                explorer.exe "\\$Computer\c$\Users\$User\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
                # Otwarcie folderu Startup w Eksploratorze
                explorer.exe "\\$Computer\c$\Users\$User\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
            } else {
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd`n")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText("$computer does not responding`n")
            }
            $outputbox.SelectionStart = $outputbox.Text.Length
            $outputbox.ScrollToCaret()
        }
    }
})

# Obsługa kliknięcia Raport
$raportButton.Add_Click({
    $results = New-Object System.Collections.ArrayList
    $computerNames = $computerTextBox.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    $selectedAccount = $accountComboBox.SelectedItem
    $selectedOption = $optionComboBox.SelectedItem

    if ([string]::IsNullOrEmpty($computerNames)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter computer names.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } elseif ($selectedAccount -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select an account.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } elseif ($selectedOption -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select an option.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        ForEach ($computer in $computerNames) {
            $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
            Write-Host "Processing $Computer..." -ForegroundColor "Yellow" -BackgroundColor "DarkCyan"
            
            if($ping) {
                # Pobranie zalogowanego użytkownika
                $User = (Get-WmiObject Win32_ComputerSystem -ComputerName $Computer).UserName
                if($User -match "Process" -or $User -match "Service") {
                    $User = $User.Substring(7)
                    
                    # Zawartość folderu Taskbar
                    $directoryTB = "\\$Computer\c$\Users\$User\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
                    $allItemsTB = Get-ChildItem -Path $directoryTB -File
                    $finalItemsTB = foreach ($ItemTB in $allItemsTB) {
                        if($ItemTB.Extension -eq '.lnk' -or $ItemTB.Extension -eq '.url') {
                            $directoryPathTB = Join-Path $directoryTB $ItemTB.Name
                            $workingDirectoryTB = (Get-Shortcut -ShortcutPath $directoryPathTB).WorkingDirectory
                            if(($workingDirectoryTB -match "Internet Explorer")-or($workingDirectoryTB -match "Edge")) {
                                $argumentsTB = (Get-Shortcut -ShortcutPath $directoryPathTB).Arguments
                                $ItemTB.Name + "(" + $argumentsTB + ")"
                            } else {
                                $ItemTB.Name
                            }
                        } else {
                            $ItemTB.Name
                        }
                    }
                    $filesListTB = $finalItemsTB -join ', '
                    
                    # Zawartość folderu Startup
                    $directorySU = "\\$Computer\c$\Users\$User\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
                    $allItemsSU = Get-ChildItem -Path $directorySU -File
                    $finalItemsSU = foreach ($ItemSU in $allItemsSU) {
                        if($ItemSU.Extension -eq '.lnk' -or $ItemSU.Extension -eq '.url') {
                            $directoryPathSU = Join-Path $directorySU $ItemSU.Name
                            $workingDirectorySU = (Get-Shortcut -ShortcutPath $directoryPathSU).WorkingDirectory
                            if(($workingDirectorySU -match "Internet Explorer")-or($workingDirectorySU -match "Edge")) {
                                $argumentsSU = (Get-Shortcut -ShortcutPath $directoryPathSU).Arguments
                                $ItemSU.Name + "(" + $argumentsSU + ")"
                            } else {
                                $ItemSU.Name
                            }
                        } else {
                            $ItemSU.Name
                        }
                    }
                    $filesListSU = $finalItemsSU -join ', '
                }
                
                # Utworzenie obiektu z danymi
                $data = [pscustomobject]@{
                    'Hostname' = $Computer
                    'Status' = "UP"
                    'User' = $User
                    'Taskband' = $filesListTB
                    'Startup' = $filesListSU
                }
                [void]$results.Add($data)
            } else {
                $data = [pscustomobject]@{
                    'Hostname' = $Computer
                    'Status' = "DOWN"
                    'User' = "-"
                    'Taskband' = "-"
                    'Startup' = "-"
                }
                [void]$results.Add($data)
            }
        }
        
        # Tworzenie okna z tabelą wyników
        $FormR = New-Object System.Windows.Forms.Form
        $FormR.StartPosition = "CenterScreen"
        $FormR.Text = 'Workstation Information'
        $FormR.Size = New-Object System.Drawing.Size(1200, 400)
        $FormR.FormBorderStyle = 'FixedSingle'
        $FormR.MaximizeBox = $false
        $FormR.MinimizeBox = $false
        
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = 'Fill'
        $panel.AutoScroll = $true
        $FormR.Controls.Add($panel)
        
        $dataGridView = New-Object System.Windows.Forms.DataGridView
        $dataGridView.Size = New-Object System.Drawing.Size(50, 300)
        $dataGridView.Dock = 'Top'
        $dataGridView.AutoSizeColumnsMode = 'AllCells'
        $panel.Controls.Add($dataGridView)
        
        # Tworzenie tabeli danych
        $dataTable = New-Object System.Data.DataTable
        $dataTable.Columns.Add('Hostname') | Out-Null
        $dataTable.Columns.Add('Status') | Out-Null
        $dataTable.Columns.Add('User') | Out-Null
        $dataTable.Columns.Add('Taskband') | Out-Null
        $dataTable.Columns.Add('Startup') | Out-Null
        
        foreach ($result in $results) {
            $row = $dataTable.NewRow()
            $row['Hostname'] = $result.Hostname
            $row['Status'] = $result.Status
            $row['User'] = $result.User
            $row['Taskband'] = $result.Taskband
            $row['Startup'] = $result.Startup
            $dataTable.Rows.Add($row)
        }
        
        $dataGridView.DataSource = $dataTable
        
        # Przycisk kopiujący dane do schowka
        $copyButton = New-Object System.Windows.Forms.Button
        $copyButton.Autosize = $True
        $copyButton.Location = New-Object System.Drawing.Size(300,320)
        $copyButton.Text = 'Copy'
        $copyButton.BackColor = 'LightGray'
        $copyButton.Add_Click({
            $data = New-Object System.Text.StringBuilder
            foreach ($row in $dataGridView.Rows) {
                foreach ($cell in $row.Cells) {
                    $data.Append("$($cell.Value)`t")
                }
                $data.AppendLine()
            }
            [System.Windows.Forms.Clipboard]::SetText($data.ToString())
        })
        $FormR.AcceptButton = $copyButton
        $panel.Controls.Add($copyButton)
        
        $FormR.ShowDialog() | Out-Null
    }
})

# Obsługa kliknięcia Submit - główna akcja
$submitButton.Add_Click({
    $computerNames = $computerTextBox.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    $selectedAccount = $accountComboBox.SelectedItem
    $selectedOption = $optionComboBox.SelectedItem

    if ([string]::IsNullOrEmpty($computerNames)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter computer names.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } elseif ($selectedAccount -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select an account.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } elseif ($selectedOption -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select an option.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        foreach ($computer in $computerNames) {
            # Czyszczenie całego folderu Startup
            if($CheckboxCleanEntireStartUp.Checked) {
                $StartUpFiles = "\\$Computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
                Remove-Item "$StartUpFiles\*" -Force
            }
            
            # Czyszczenie tylko skrótów IE/Edge w Startup
            if($CheckboxCleanIEEDGEStartUp.Checked) {
                function Get-ShortcutStartIn {
                    param ([string]$ShortcutPath)
                    $WshShell = New-Object -ComObject WScript.Shell
                    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
                    return $Shortcut.WorkingDirectory
                }
                $StartUpFiles = "\\$Computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
                $files = Get-ChildItem -Path $StartUpFiles -File | Where-Object { $_.Extension -eq '.lnk' -or $_.Extension -eq '.url' }
                foreach ($file in $files) {
                    $finalPath = "\\$Computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\$file"
                    $shortcutStartIn = Get-ShortcutStartIn -ShortcutPath $finalPath
                    if(($shortcutStartIn -match "Internet Explorer")-or($shortcutStartIn -match "Edge")) {
                        Remove-Item "$finalPath" -Force
                    }
                }
            }
            
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd`n")
            $outputbox.SelectionColor='blue'
            $outputbox.AppendText("Checking")
            $outputbox.SelectionColor='green'
            $outputbox.AppendText(" $computer`n")
            
            $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
            if($ping) {
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd`n")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText("$computer")
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText(" Responding`n")
                
                # Sprawdzenie czy profil użytkownika istnieje
                $profilePath = Test-Path -Path "\\$computer\c$\Users\$selectedAccount\NTUSER.DAT"
                if($profilePath) {
                    # Zabicie procesów Edge
                    if($CheckboxTaskkillEdge.Checked) {
                        taskkill /s \\$computer /f /im msedge.exe >$null 2>&1
                        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                        $outputbox.AppendText("$gd`n")
                        $outputbox.SelectionColor='green'
                        $outputbox.AppendText("msedge.exe")
                        $outputbox.SelectionColor='blue'
                        $outputbox.AppendText(" sessions killed`n")
                    }
                    
                    # Modyfikacje preferencji i rejestru
                    if($CheckboxPreferences.Checked) {
                        $EdgePreferences = "\\$Computer\c$\Users\$selectedAccount\AppData\Local\Microsoft\Edge\User Data\Default"
                        Copy-Item -Path "$PWD\Templates\Default\Preferences" -Destination "$EdgePreferences" -Recurse -Force
                        Copy-Item -Path "$PWD\Templates\Default\Bookmarks" -Destination "$EdgePreferences" -Recurse -Force
                        $sid = ([System.Security.Principal.NTAccount]("DOMAIN\$selectedAccount")).Translate([System.Security.Principal.SecurityIdentifier]).Value
                        reg add "\\$computer\HKEY_USERS\$sid\Software\Policies\Microsoft\Windows\Explorer" /v NoPinningToTaskbar /t REG_DWORD /d 1 /f
                        reg add "\\$computer\HKEY_USERS\$sid\Software\Policies\Microsoft\Windows\Explorer" /v HidePeopleBar /t REG_DWORD /d 1 /f
                        reg add "\\$computer\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge" /v "HideRestoreDialogEnabled" /t REG_DWORD /d 1 /f
                        reg add "\\$computer\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge" /v "HideFirstRunExperience" /t REG_DWORD /d 1 /f
                        reg add "\\$computer\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge" /v "ShowRecommendationsEnabled" /t REG_DWORD /d 0 /f
                        reg add "\\$computer\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge" /v "StartupBoostEnabled" /t REG_DWORD /d 0 /f
                    }
                    
                    # Kopiowanie plików konfiguracyjnych Taskbar
                    if($selectedOption -eq "Config_Example01") {
                        $TaskbarIcons = "\\$Computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
                        $TaskbarReg = "\\$Computer\c$\Users\$selectedAccount\Documents"
                        $TaskbarX = "\\$Computer\c$\Scripts"
                        Remove-Item "$TaskbarIcons\*" -Force
                        Copy-Item -Path "$PWD\Templates\Taskband\Example01\*.vbs" -Destination "$TaskbarX" -Recurse -Force
                        Copy-Item -Path "$PWD\Templates\Taskband\Example01\TaskBar\*" -Destination "$TaskbarIcons" -Recurse -Force
                        Copy-Item -Path "$PWD\Templates\Taskband\Example01\taskband.reg" -Destination "$TaskbarReg" -Recurse -Force
                        
                        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                        $outputbox.AppendText("$gd`n")
                        $outputbox.SelectionColor='blue'
                        $outputbox.AppendText("The taskband files have been copied on the ")
                        $outputbox.SelectionColor='green'
                        $outputbox.AppendText("$computer")
                        $outputbox.SelectionColor='blue'
                        $outputbox.AppendText(" to the ")
                        $outputbox.SelectionColor='green'
                        $outputbox.AppendText("$selectedAccount")
                        $outputbox.SelectionColor='blue'
                        $outputbox.AppendText(" profile`n")
                        
                        $loggedUsers = quser /server:$Computer
                        if ($loggedUsers -match $selectedAccount) {
                            $sid = ([System.Security.Principal.NTAccount]("DOMAIN\$selectedAccount")).Translate([System.Security.Principal.SecurityIdentifier]).Value
                            # Dodanie wartości rejestru Taskband (binarnych) - szczegóły w pliku konfiguracyjnym
                            reg add "\\$computer\HKEY_USERS\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Taskband" /v FavoritesResolve /t REG_BINARY /d [BINARY_REGISTRY_VALUE_FROM_TEMPLATE] /f
                            reg add "\\$computer\HKEY_USERS\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Taskband" /v Favorites /t REG_BINARY /d [BINARY_REGISTRY_VALUE_FROM_TEMPLATE] /f
                            reg add "\\$computer\HKEY_USERS\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Taskband" /v FavoritesChanges /t REG_SZ /d 1 /f
                            reg add "\\$computer\HKEY_USERS\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Taskband" /v FavoritesVersion /t REG_SZ /d 1 /f
                        }
                    }
                    # Kolejne opcje konfiguracyjne (Config_Example02, Config_Example03 itp.) są analogiczne
                    # Różnią się tylko ścieżkami do szablonów i wartościami binarnymi rejestru
                    
                    # Modyfikacja skrótu Edge z argumentami
                    if($CheckBoxEdgeArguments.Checked) {
                        $EdgeShortcutTestPath = Test-Path -Path "\\$computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk"
                        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                        $outputbox.AppendText("$gd`n")
                        $outputbox.SelectionColor='blue'
                        $outputbox.AppendText("Checking for Edge Shortcut...`n")
                        
                        if($EdgeShortcutTestPath) {
                            $argumentsTextBoxArrays = @($argumentsTextBox.Text -split ',')
                            if($argumentsTextBox.Text -eq "" -or $argumentsTextBox.Text -match '^\s*$') {
                                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                                $outputbox.AppendText("$gd`n")
                                $outputbox.SelectionColor='red'
                                $outputbox.AppendText("Arguments textbox is empty`n")
                            } else {
                                $EdgeShortcutPath = "\\$computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk"
                                $EdgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                                $wshShell = New-Object -ComObject WScript.Shell
                                $shortcut = $wshShell.CreateShortcut($EdgeShortcutPath)
                                $shortcut.TargetPath = $EdgePath
                                $shortcut.WindowStyle = 3
                                if($fullScreenCheckBox.Checked) {
                                    $argumentsTextBoxArrays = $argumentsTextBoxArrays + " --start-fullscreen"
                                }
                                $shortcut.Arguments = $argumentsTextBoxArrays -join ' '
                                $shortcut.WorkingDirectory = (Get-Item $EdgePath).Directory.FullName
                                $shortcut.IconLocation = $EdgePath
                                $shortcut.Save()
                                
                                # Odświeżenie przypięcia do paska zadań
                                $shell = New-Object -ComObject Shell.Application
                                $folder = $shell.Namespace("\\$computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
                                $item = $folder.ParseName('Microsoft Edge.lnk')
                                $item.InvokeVerb('taskbarpin')
                                
                                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                                $outputbox.AppendText("$gd`n")
                                $outputbox.SelectionColor='green'
                                $outputbox.AppendText("Done`n")
                                
                                # Kopiowanie do Startup
                                if($startUpCheckBox.Checked) {
                                    if($fileNameTextBox.Text -eq "" -or $fileNameTextBox.Text -match '^\s*$') {
                                        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                                        $outputbox.AppendText("$gd`n")
                                        $outputbox.SelectionColor='red'
                                        $outputbox.AppendText("FileName textbox is empty`n")
                                    } else {
                                        $fileName = $fileNameTextBox.Text
                                        Copy-Item -Path "\\$computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk" -Destination "\\$computer\c$\Users\$selectedAccount\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\$fileName.lnk"
                                    }
                                }
                            }
                        } else {
                            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                            $outputbox.AppendText("$gd`n")
                            $outputbox.SelectionColor='red'
                            $outputbox.AppendText("The Edge Shortcut is not present or immutable`n")
                        }
                    }
                    
                    # Wylogowanie użytkownika
                    if($CheckboxLogoff.Checked) {
                        $quserResult = quser /server:$computer 2>&1
                        $quserRegex = $quserResult | ForEach-Object { $_ -replace '\s{2,}',',' }
                        $quserObject = $quserRegex | ConvertFrom-Csv
                        $quserFinal = $quserObject | Where-Object { $_.username -match "$selectedAccount" } | Select -ExpandProperty sessionname
                        if($quserFinal) {
                            logoff $quserFinal /server:$Computer
                            Write-Host "Session terminated on $Computer" -BackgroundColor "Green"
                        } else {
                            Write-Host "Nothing to terminate on $Computer" -BackgroundColor "Blue"
                        }
                    }
                } else {
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd`n")
                    $outputbox.SelectionColor='red'
                    $outputbox.AppendText("$selectedAccount account with valid NTUSER.DAT file not found on the $computer`n")
                }
            } else {
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd`n")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText("$computer does not responding`n")
            }
            $outputbox.SelectionStart = $outputbox.Text.Length
            $outputbox.ScrollToCaret()
        }
    }
})

$form.Controls.Add($tabControl)
$form.ShowDialog()
$form.Dispose()