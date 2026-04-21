# BIBLIOTEKI I ZALEŻNOŚCI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# FUNKCJA: CLButton - otwiera listę komputerów
function CLButton() {
    # Otwarcie pliku z listą komputerów w notatniku
    start-process C:\Scripts\CopyRemove\computers.txt
    $outputbox.SelectionColor='black'
    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
    $outputbox.AppendText("$gd")
    $outputBox.AppendText("`n")
    $outputbox.SelectionColor='blue'
    $outputbox.AppendText("Computer list")
    $outputBox.AppendText("`n")
}

# FUNKCJA: F2CButton - otwiera folder z plikami do skopiowania
function F2CButton() {
    start-process C:\Scripts\CopyRemove\ToCopy
    $outputbox.SelectionColor='black'
    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
    $outputbox.AppendText("$gd")
    $outputBox.AppendText("`n")
    $outputbox.SelectionColor='blue'
    $outputbox.AppendText("2Copy")
    $outputBox.AppendText("`n")
}

# FUNKCJA: F2RButton - otwiera folder z plikami do usunięcia
function F2RButton() {
    start-process C:\Scripts\CopyRemove\ToRemove\
    $outputbox.SelectionColor='black'
    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
    $outputbox.AppendText("$gd")
    $outputBox.AppendText("`n")
    $outputbox.SelectionColor='blue'
    $outputbox.AppendText("2Remove")
    $outputBox.AppendText("`n")
}

# FUNKCJA: CTButton - kopiowanie plików na zdalne komputery
function CTButton() {
    # Wczytanie listy komputerów z pliku
    $Computers = Get-Content "C:\Scripts\CopyRemove\computers.txt"
    
    ForEach ($Computer in $Computers) {
        # Test połączenia ping
        $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
        
        if($ping) {
            # Komputer odpowiada
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n") 
            $outputbox.SelectionColor='green'
            $outputbox.AppendText("$computer")
            $outputbox.SelectionColor='blue'
            $outputbox.AppendText(" Responding")
            $outputBox.AppendText("`n")
            
            # Źródło plików do kopiowania
            $SourceTC = 'C:\Scripts\CopyRemove\ToCopy\*'
            
            # Kopiowanie do Favorites wszystkich użytkowników
            if ($Checkbox1.Checked -eq $true) {
                $DestinationFAU = "\\" + $Computer + "\c$\Users\*\Favorites"
                Get-ChildItem $DestinationFAU | ForEach-Object {Copy-Item -Path $SourceTC -Destination $_ -Force -Recurse}
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Favorites for all users successfully copied on")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $Computer")
                $outputBox.AppendText("`n")
            }
            
            # Kopiowanie do Favorites\Links wszystkich użytkowników
            if ($Checkbox2.Checked -eq $true) {
                $DestinationFAUB = "\\" + $Computer + "\c$\Users\*\Favorites\Links"
                Get-ChildItem $DestinationFAUB | ForEach-Object {Copy-Item -Path $SourceTC -Destination $_ -Force -Recurse}
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Favorites Bar for all users successfully copied on")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $Computer")
                $outputBox.AppendText("`n")
            }
            
            # Kopiowanie na pulpit wszystkich użytkowników
            if ($Checkbox3.Checked -eq $true) {
                $DestinationDAU = "\\" + $Computer + "\c$\Users\*\Desktop"
                Get-ChildItem $DestinationDAU | ForEach-Object {Copy-Item -Path $SourceTC -Destination $_ -Force -Recurse}
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Desktop for all users successfully copied on")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $Computer")
                $outputBox.AppendText("`n")
            }
            
            # Kopiowanie do domyślnych Favorites
            if ($Checkbox4.Checked -eq $true) {
                $DestinationDF = "\\" + $Computer + "\c$\Users\Default\Favorites"
                Copy-Item -Path $SourceTC -Destination $DestinationDF -recurse -Force
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Default Favorites successfully copied on")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $Computer")
                $outputBox.AppendText("`n")
            }
            
            # Kopiowanie na domyślny pulpit
            if ($Checkbox5.Checked -eq $true) {
                $DestinationDD = "\\" + $Computer + "\c$\Users\Default\Desktop"
                Copy-Item -Path $SourceTC -Destination $DestinationDD -recurse -Force
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Default Desktop successfully copied on")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $Computer")
                $outputBox.AppendText("`n")
            }
            
            # Kopiowanie do Menu Start
            if ($Checkbox6.Checked -eq $true) {
                $DestinationMS = "\\" + $Computer + "\c$\ProgramData\Microsoft\Windows\Start Menu\Programs"
                Copy-Item -Path C:\Scripts\CopyRemove\ToCopy\* -Destination $DestinationMS -recurse -Force
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText("Menu Start succesfully copied for")
                $outputbox.SelectionColor='green'
                $outputbox.AppendText(" $Computer")
                $outputBox.AppendText("`n")
            }
        } else {
            # Komputer nie odpowiada
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n")
            $outputbox.SelectionColor='green'
            $outputbox.AppendText("$computer")
            $outputbox.SelectionColor='red'
            $outputbox.AppendText(" does not responding")
            $outputBox.AppendText("`n")
        }
    }
}

# FUNKCJA: RTButton - usuwanie plików ze zdalnych komputerów
function RTButton() {
    $Computers = Get-Content "C:\Scripts\CopyRemove\computers.txt"
    
    ForEach ($Computer in $Computers) {
        $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
        
        if($ping) {
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n") 
            $outputbox.SelectionColor='green'
            $outputbox.AppendText("$computer")
            $outputbox.SelectionColor='blue'
            $outputbox.AppendText(" Responding")
            $outputBox.AppendText("`n")
            
            # Pobranie listy plików do usunięcia
            $Files = Get-ChildItem C:\Scripts\CopyRemove\ToRemove -Name
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n") 
            $outputbox.SelectionColor='green'
            $outputbox.AppendText("files to remove:")
            $outputbox.SelectionColor='blue'
            $outputbox.AppendText(" $files")
            $outputBox.AppendText("`n")
            
            ForEach ($File in $Files) {
                $outputbox.SelectionColor='black'
                $outputbox.AppendText("Processing file:")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='darkcyan'
                $outputbox.AppendText("$File")
                $outputBox.AppendText("`n")
                
                # Usuwanie z Favorites wszystkich użytkowników
                if ($Checkbox1.Checked -eq $true) {
                    $DestinationFAUR = "\\" + $Computer + "\c$\Users\*\Favorites\$File"
                    Remove-Item -path $DestinationFAUR -Force -Recurse -ErrorAction SilentlyContinue
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Favorites shortcuts removed for all users successfully from")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText(" $Computer")
                    $outputBox.AppendText("`n")
                }
                
                # Usuwanie z Favorites\Links wszystkich użytkowników
                if ($Checkbox2.Checked -eq $true) {
                    $DestinationFAUBR = "\\" + $Computer + "\c$\Users\*\Favorites\Links\$File"
                    Remove-Item -path $DestinationFAUBR -Force -Recurse -ErrorAction SilentlyContinue
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Favorites Bar shortcuts removed for all users successfully from")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText(" $Computer")
                    $outputBox.AppendText("`n")
                }
                
                # Usuwanie z pulpitu wszystkich użytkowników
                if ($Checkbox3.Checked -eq $true) {
                    $DestinationDAUR = "\\" + $Computer + "\c$\Users\*\Desktop\$File"
                    Remove-Item -path $DestinationDAUR -Force -Recurse -ErrorAction SilentlyContinue
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Desktop shortcuts removed for all users successfully from")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText(" $Computer")
                    $outputBox.AppendText("`n")
                }
                
                # Usuwanie z domyślnych Favorites
                if ($Checkbox4.Checked -eq $true) {
                    $DestinationDFR = "\\" + $Computer + "\c$\Users\Default\Favorites\$File"
                    Remove-Item -path $DestinationDFR -Force -Recurse -ErrorAction SilentlyContinue
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Default Favorites shortcuts removed successfully from")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText(" $Computer")
                    $outputBox.AppendText("`n")
                }
                
                # Usuwanie z domyślnego pulpitu
                if ($Checkbox5.Checked -eq $true) {
                    $DestinationDDR = "\\" + $Computer + "\c$\Users\Default\Desktop\$File"
                    Remove-Item -path $DestinationDDR -Force -Recurse -ErrorAction SilentlyContinue
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Default Desktop shortcuts removed successfully from")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText(" $Computer")
                    $outputBox.AppendText("`n")
                }
                
                # Usuwanie z Menu Start
                if ($Checkbox6.Checked -eq $true) {
                    $DestinationSMR = "\\" + $Computer + "\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\$File"
                    Remove-Item -path $DestinationSMR -Force -Recurse -ErrorAction SilentlyContinue
                    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                    $outputbox.AppendText("$gd")
                    $outputBox.AppendText("`n") 
                    $outputbox.SelectionColor='blue'
                    $outputbox.AppendText("Start Menu shortcuts removed successfully from")
                    $outputbox.SelectionColor='green'
                    $outputbox.AppendText(" $Computer")
                    $outputBox.AppendText("`n")
                }
            }
        } else {
            $outputbox.SelectionColor='black'
            $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
            $outputbox.AppendText("$gd")
            $outputBox.AppendText("`n")
            $outputbox.SelectionColor='green'
            $outputbox.AppendText("$computer")
            $outputbox.SelectionColor='red'
            $outputbox.AppendText(" does not responding")
            $outputBox.AppendText("`n")
        }
    }
}

# FUNKCJA: FnDButton - wyszukiwanie i usuwanie plików po nazwie/rozszerzeniu
function FnDButton() {
    $Computers = Get-Content "C:\Scripts\CopyRemove\computers.txt"
    
    $x = ''
    $y = ''
    $z = ''
    [string]$x = $TextBox1.Text
    [string]$y = $TextBox2.Text
    [string]$z = $TextBox3.Text
    
    # Sprawdzenie czy pola nie są puste
    if($x -eq '' -and $y -eq ''-and $z -eq '') {
        $outputbox.SelectionColor='black'
        $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
        $outputbox.AppendText("$gd")
        $outputBox.AppendText("`n")
        $outputbox.SelectionColor='red'
        $outputbox.AppendText("The fields are empty!")
        $outputBox.AppendText("`n")
    } else {
        ForEach ($Computer in $Computers) {
            $aufcounter = 0
            $ping = Test-Connection $Computer -Count 1 -ea silentlycontinue
            
            if($ping) {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n") 
                $outputbox.SelectionColor='green'
                $outputbox.AppendText("$computer")
                $outputbox.SelectionColor='blue'
                $outputbox.AppendText(" Responding")
                $outputBox.AppendText("`n")
                
                # Pobranie listy użytkowników (bez profili systemowych)
                $Users = Get-ChildItem -path "\\$computer\c$\Users\" -exclude Public,Default,ADMINI~1 | where 'Name' -notmatch 'temp' | select Name -ExpandProperty Name
                
                # Wyświetlenie listy użytkowników jeśli zaznaczono odpowiednie checkboxy
                if ($Checkbox1.Checked -eq $true -or $Checkbox2.Checked -eq $true -or $Checkbox3.Checked -eq $true) {
                    $outputbox.SelectionColor='DarkMagenta'
                    $outputbox.AppendText("Users list:")
                    $outputBox.AppendText("`n")
                    ForEach ($User in $Users) {
                        $outputbox.SelectionColor='DarkCyan'
                        $outputbox.AppendText("$User")
                        $outputBox.AppendText("`n")
                    }
                }
                
                # USUWANIE Z FAVORITES (Checkbox1)
                if ($Checkbox1.Checked -eq $true) {
                    # Różne kombinacje wyszukiwania (nazwa, rozszerzenie, lub oba)
                    if($x -ne '' -and $y -eq '' -and $z -eq '') {
                        # Tylko nazwa pliku
                        ForEach ($User in $Users) {
                            $Files = Get-ChildItem -path "\\$computer\c$\Users\$User\Favorites"
                            foreach ($File in $Files) {
                                $Results = $File | where Name -match $x | where {$_.PSIsContainer -eq $false}
                                foreach ($Result in $Results) {
                                    $Result | Remove-Item
                                    $aufcounter++
                                }
                            }
                        }
                    }
                }
                
                # Podsumowanie operacji
                if($aufcounter -gt '0') {
                    Write-Host "Successfully deleted $aufcounter files $computer" -BackgroundColor "Green"
                }
                if($aufcounter -eq '0') {
                    Write-Host "Nothing to delete on $computer" -BackgroundColor "Blue"
                }
            } else {
                $outputbox.SelectionColor='black'
                $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
                $outputbox.AppendText("$gd")
                $outputBox.AppendText("`n")
                $outputbox.SelectionColor='red'
                $outputbox.AppendText("$Computer does not responding")
                $outputBox.AppendText("`n")
            }
        }
    }
}

# FUNKCJA: Check/Uncheck - zaznaczanie/odznaczanie wszystkich checkboxów
function Check/Uncheck {
    if ($checkBox1.Checked -eq $False) { $checkbox1.Checked = $true }
    else { $checkbox1.Checked = $false }
    
    if ($checkBox2.Checked -eq $False) { $checkbox2.Checked = $true }
    else { $checkbox2.Checked = $false }
    
    if ($checkBox3.Checked -eq $False) { $checkbox3.Checked = $true }
    else { $checkbox3.Checked = $false }
    
    if ($checkBox4.Checked -eq $False) { $checkbox4.Checked = $true }
    else { $checkbox4.Checked = $false }
    
    if ($checkBox5.Checked -eq $False) { $checkbox5.Checked = $true }
    else { $checkbox5.Checked = $false }
    
    if ($checkBox6.Checked -eq $False) { $checkbox6.Checked = $true }
    else { $checkbox6.Checked = $false }
    
    $outputbox.SelectionColor='black'
    $gd = Get-Date -UFormat "%d/%m/%Y %R:%S:"
    $outputbox.AppendText("$gd")
    $outputBox.AppendText("`n")
    $outputbox.SelectionColor='blue'
    $outputbox.AppendText("Check/Uncheck")
    $outputBox.AppendText("`n")
}

# KONFIGURACJA OKNA GŁÓWNEGO FORMULARZA
$Form1 = New-Object system.Windows.Forms.Form
$Icon = New-Object system.drawing.icon ("C:\Scripts\CopyRemove\ico.ico")
$Form1.Text = "Copy/Remove"
$Form1.Icon = $Icon
$Form1.StartPosition = "CenterScreen"
$Form1.BackgroundImageLayout = "Center"
$Form1.MaximizeBox = $false
$Form1.MinimizeBox = $false
$Form1.minimumSize = New-Object System.Drawing.Size(650,640) 
$Form1.maximumSize = New-Object System.Drawing.Size(650,640)

$Font1 = New-Object System.Drawing.Font("Trebuchet",11,[System.Drawing.FontStyle]::Bold)

# KONTROLKA: TAB CONTROL (zakładki)
$TabControl = New-object System.Windows.Forms.TabControl
$TabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$TabControl.Location = New-Object System.Drawing.Point(0,210)
$TabControl.Name = "TabControl"
$TabControl.ShowToolTips = $True
$TabControlF = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
$TabControl.Font = $TabControlF
$TabControl.Height = 455
$TabControl.Width = 315
$TabControl.BackColor = 'FloralWhite'
$Form1.Controls.Add($TabControl)

# KONTROLKA: OUTPUT BOX - okno wyjściowe (logi)
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Size(325,10) 
$outputBox.Size = New-Object System.Drawing.Size(300,325) 
$outputBox.MultiLine = $True 
$outputBox.ReadOnly = $True
$outputBox.Font = New-Object System.Drawing.Font("Calibri",11,[System.drawing.FontStyle]::Bold)
$outputBox.ScrollBars = "Vertical" 
$Form1.Controls.Add($outputBox)

# PRZYCISK: CLEAR - czyszczenie okna wyjściowego
$ClearButton = New-Object "System.Windows.Forms.Button";
$ClearButton.Location = New-Object System.Drawing.Size(350,350)
$ClearButton.Autosize = $True
$ClearButton.Font = $Font1
$ClearButton.BackColor = 'FloralWhite'
$ClearButton.Text = "Clear"
$ClearButton.Add_Click{$outputBox.Clear()}
$Form1.Controls.Add($ClearButton)

# ZAKŁADKA: COPY/REMOVE
$CopyRemove = New-Object System.Windows.Forms.TabPage
$CopyRemove.DataBindings.DefaultDataSourceUpdateMode = 0
$CopyRemove.UseVisualStyleBackColor = $True
$CopyRemove.Name = "Copy/Remove"
$CopyRemove.Text = "Copy/Remove"
$CopyRemove.BackColor = 'White'
$TabControl.Controls.Add($CopyRemove)

# ZAKŁADKA: MATCH REMOVER
$MatchRemover = New-Object System.Windows.Forms.TabPage
$MatchRemover.DataBindings.DefaultDataSourceUpdateMode = 0
$MatchRemover.UseVisualStyleBackColor = $True
$MatchRemover.BackColor = 'White'
$MatchRemover.Name = "Match Remover"
$MatchRemover.Text = "Match Remover"
$TabControl.Controls.Add($MatchRemover)

# KONTROLKI DLA ZAKŁADKI MATCH REMOVER
$FontL = New-Object System.Drawing.Font("Trebuchet",12,[System.Drawing.FontStyle]::Bold)

# Pole tekstowe - pierwsza wartość do wyszukania
$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Location = New-Object System.Drawing.Size(10,10)
$TextBox1.Font = $FontL
$TextBox1.Size = New-Object System.Drawing.Size(150,50) 

# Etykieta "OR"
$LabelOR = New-Object Windows.Forms.Label
$LabelOR.Font = $FontL
$LabelOR.BackColor = 'White'
$LabelOR.Location = New-Object Drawing.Point(10,40)
$LabelOR.Size = New-Object Drawing.Point(280,30)
$LabelOR.Text = "OR"

# Pole tekstowe - druga wartość do wyszukania
$TextBox2 = New-Object System.Windows.Forms.TextBox
$TextBox2.Location = New-Object System.Drawing.Size(10,70) 
$TextBox2.Font = $FontL
$TextBox2.Size = New-Object System.Drawing.Size(150,50) 

# Etykieta "File Extension"
$LabelFE = New-Object Windows.Forms.Label
$LabelFE.Font = $FontL
$LabelFE.BackColor = 'White'
$LabelFE.Location = New-Object Drawing.Point(10,100)
$LabelFE.Size = New-Object Drawing.Point(280,30)
$LabelFE.Text = "File Extension"

# Pole tekstowe - rozszerzenie pliku
$TextBox3 = New-Object System.Windows.Forms.TextBox
$TextBox3.Location = New-Object System.Drawing.Size(10,130) 
$TextBox3.Font = $FontL
$TextBox3.Size = New-Object System.Drawing.Size(150,50) 

# Przycisk OK - zatwierdza wartości do wyszukania
$ButtonOK = New-Object System.Windows.Forms.Button
$ButtonOK.Font = $Font1
$ButtonOK.TabIndex = 0
$ButtonOK.Autosize = $True
$ButtonOK.Location = New-Object System.Drawing.Size(180,10)
$ButtonOK.Name = "OK"
$ButtonOK.Text = "OK"
$ButtonOK.BackColor = 'FloralWhite'
$ButtonOK.UseVisualStyleBackColor = $True

# Przycisk Clear - czyści pola tekstowe
$ButtonClear = New-Object System.Windows.Forms.Button
$ButtonClear.Font = $Font1
$ButtonClear.TabIndex = 0
$ButtonClear.Autosize = $True
$ButtonClear.Location = New-Object System.Drawing.Size(180,70)
$ButtonClear.Name = "Clear"
$ButtonClear.Text = "Clear"
$ButtonClear.BackColor = 'FloralWhite'
$ButtonClear.UseVisualStyleBackColor = $True

$ButtonClear.Add_Click({
    $TextBox1.Text = ''
    $TextBox2.Text = ''
    $TextBox3.Text = ''
    $global:x=$TextBox1.Text
    $global:y=$TextBox2.Text
    $global:z=$TextBox3.Text
    $x=''
    $y=''
    $z=''
    $LabelV2F.ForeColor = 'Black'
    $LabelV2F.Text = "Values2Find:`n"
})

# Etykieta wyświetlająca wartości do wyszukania
$LabelV2F = New-Object Windows.Forms.Label
$LabelV2F.Font = $FontL
$LabelV2F.BackColor = 'White'
$LabelV2F.Location = New-Object Drawing.Point(10,160)
$LabelV2F.Size = New-Object Drawing.Point(300,50)
$LabelV2F.Text = "Values2Find:`n"

# Przycisk FnD - wykonuje wyszukiwanie i usuwanie
$ButtonFnD = New-Object System.Windows.Forms.Button
$ImageRT = [system.drawing.image]::FromFile("C:\Scripts\CopyRemove\bin.png")
$ButtonFnD.BackgroundImage = $ImageRT
$ButtonFnD.BackgroundImageLayout = "Center"
$ButtonFnD.Size = New-Object System.Drawing.Size(120,120)
$ButtonFnD.Font = $Font1
$ButtonFnD.BackgroundImageLayout = "Center"
$ButtonFnD.Location = New-Object System.Drawing.Size(160,230)
$ButtonFnD.Name = "Button FnD"
$ButtonFnD.BackColor = 'FloralWhite'
$ButtonFnD.Add_Click({FnDButton}) 
$ButtonFnD.UseVisualStyleBackColor = $True

# PRZYCISKI NA GŁÓWNYM FORMULARZU

# Przycisk Computers List - otwiera listę komputerów
$ButtonCL = New-Object System.Windows.Forms.Button
$ButtonCL.Font = $Font1
$ButtonCL.TabIndex = 0
$ButtonCL.Autosize = $True
$ButtonCL.Location = New-Object System.Drawing.Size(10,10)
$ButtonCL.Name = "Computers List"
$ButtonCL.Text = "Computers List"
$ButtonCL.BackColor = 'FloralWhite'
$ButtonCL.Add_Click({CLButton}) 
$ButtonCL.UseVisualStyleBackColor = $True

# Przycisk Files2Copy - otwiera folder z plikami do kopiowania
$ButtonF2C = New-Object System.Windows.Forms.Button
$ButtonF2C.Font = $Font1
$ButtonF2C.TabIndex = 1
$ButtonF2C.Autosize = $True
$ButtonF2C.Location = New-Object System.Drawing.Size(20,10)
$ButtonF2C.Name = "Files2Copy"
$ButtonF2C.Text = "Files2Copy"
$ButtonF2C.BackColor = 'FloralWhite'
$ButtonF2C.Add_Click({F2CButton}) 
$ButtonF2C.UseVisualStyleBackColor = $True

# Przycisk Files2Remove - otwiera folder z plikami do usunięcia
$ButtonF2R = New-Object System.Windows.Forms.Button
$ButtonF2R.Font = $Font1
$ButtonF2R.TabIndex = 2
$ButtonF2R.Autosize = $True
$ButtonF2R.Location = New-Object System.Drawing.Size(150,10)
$ButtonF2R.Name = "Files2Remove"
$ButtonF2R.Text = "Files2Remove"
$ButtonF2R.BackColor = 'FloralWhite'
$ButtonF2R.Add_Click({F2RButton}) 
$ButtonF2R.UseVisualStyleBackColor = $True

# Przycisk Check/Uncheck - zaznacza/odznacza wszystkie checkboxy
$ButtonCU = New-Object System.Windows.Forms.Button
$ButtonCU.Font = $Font1
$ButtonCU.TabIndex = 0
$ButtonCU.Autosize = $True
$ButtonCU.Location = New-Object System.Drawing.Size(140,10)
$ButtonCU.Name = "Check/Uncheck"
$ButtonCU.Text = "Check/Uncheck"
$ButtonCU.BackColor = 'FloralWhite'
$ButtonCU.Add_Click({Check/Uncheck}) 
$ButtonCU.UseVisualStyleBackColor = $True

# CHECKBOXY - wybór lokalizacji dla operacji
$FontC = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)

# Checkbox1 - All users Favorites
$Checkbox1 = New-Object System.Windows.Forms.CheckBox
$Checkbox1.DataBindings.DefaultDataSourceUpdateMode = 0
$Checkbox1.Location = New-Object System.Drawing.Point(40,50)
$Checkbox1.Name = "Checkbox1"
$Checkbox1.Font = $FontC
$Checkbox1.TabIndex = 4
$Checkbox1.Checked = $True
$Checkbox1.Size = New-Object System.Drawing.Size(20, 40)
$Checkbox1.BackColor = 'Transparent'
$Checkbox1.AutoSize = $True
$Checkbox1.Text = "All users Favorites"

# Checkbox2 - All users Favorites Bar
$Checkbox2 = New-Object System.Windows.Forms.CheckBox
$Checkbox2.DataBindings.DefaultDataSourceUpdateMode = 0
$Checkbox2.Location = New-Object System.Drawing.Point(40,75)
$Checkbox2.Name = "Checkbox2"
$Checkbox2.Font = $FontC
$Checkbox2.TabIndex = 4
$Checkbox2.Checked = $True
$Checkbox2.Size = New-Object System.Drawing.Size(20, 40)
$Checkbox2.BackColor = 'Transparent'
$Checkbox2.AutoSize = $True
$Checkbox2.Text = "All users Favorites Bar"

# Checkbox3 - All users Desktop
$Checkbox3 = New-Object System.Windows.Forms.CheckBox
$Checkbox3.DataBindings.DefaultDataSourceUpdateMode = 0
$Checkbox3.Location = New-Object System.Drawing.Point(40,100)
$Checkbox3.Name = "Checkbox3"
$Checkbox3.Font = $FontC
$Checkbox3.TabIndex = 5
$Checkbox3.Checked = $True
$Checkbox3.Size = New-Object System.Drawing.Size(20, 40)
$Checkbox3.BackColor = 'Transparent'
$Checkbox3.AutoSize = $True
$Checkbox3.Text = "All users Desktop"

# Checkbox4 - Default Favorites
$Checkbox4 = New-Object System.Windows.Forms.CheckBox
$Checkbox4.DataBindings.DefaultDataSourceUpdateMode = 0
$Checkbox4.Location = New-Object System.Drawing.Point(40,125)
$Checkbox4.Name = "Checkbox4"
$Checkbox4.Font = $FontC
$Checkbox4.TabIndex = 6
$Checkbox4.Checked = $True
$Checkbox4.Size = New-Object System.Drawing.Size(20, 40)
$Checkbox4.BackColor = 'Transparent'
$Checkbox4.AutoSize = $True
$Checkbox4.Text = "Default Favorites"

# Checkbox5 - Default Desktop
$Checkbox5 = New-Object System.Windows.Forms.CheckBox
$Checkbox5.DataBindings.DefaultDataSourceUpdateMode = 0
$Checkbox5.Location = New-Object System.Drawing.Point(40,150)
$Checkbox5.Name = "Checkbox5"
$Checkbox5.Font = $FontC
$Checkbox5.TabIndex = 7
$Checkbox5.Checked = $True
$Checkbox5.Size = New-Object System.Drawing.Size(20, 40)
$Checkbox5.BackColor = 'Transparent'
$Checkbox5.AutoSize = $True
$Checkbox5.Text = "Default Desktop"

# Checkbox6 - Menu Start
$Checkbox6 = New-Object System.Windows.Forms.CheckBox
$Checkbox6.DataBindings.DefaultDataSourceUpdateMode = 0
$Checkbox6.Location = New-Object System.Drawing.Point(40,175)
$Checkbox6.Name = "Checkbox6"
$Checkbox6.Font = $FontC
$Checkbox6.TabIndex = 8
$Checkbox6.Checked = $True
$Checkbox6.Size = New-Object System.Drawing.Size(20, 40)
$Checkbox6.BackColor = 'Transparent'
$Checkbox6.AutoSize = $True
$Checkbox6.Text = "Menu Start"

# PRZYCISKI KOPIOWANIA I USUWANIA

# Przycisk Copy - kopiowanie plików
$ButtonCT = New-Object System.Windows.Forms.Button
$ImageCT = [system.drawing.image]::FromFile("C:\Scripts\CopyRemove\copy.png")
$ButtonCT.BackgroundImage = $ImageCT
$ButtonCT.BackgroundImageLayout = "Center"
$ButtonCT.Size = New-Object System.Drawing.Size(120,120)
$ButtonCT.Font = $Font1
$ButtonCT.TabIndex = 9
$ButtonCT.Location = New-Object System.Drawing.Size(10,50)
$ButtonCT.Name = "CopyTime"
$ButtonCT.BackColor = 'FloralWhite'
$ButtonCT.Add_Click({CTButton}) 
$ButtonCT.UseVisualStyleBackColor = $True

# Przycisk Remove - usuwanie plików
$ButtonRT = New-Object System.Windows.Forms.Button
$ImageRT = [system.drawing.image]::FromFile("C:\Scripts\CopyRemove\bin.png")
$ButtonRT.BackgroundImage = $ImageRT
$ButtonRT.BackgroundImageLayout = "Center"
$ButtonRT.Size = New-Object System.Drawing.Size(120,120)
$ButtonRT.Font = $Font1
$ButtonRT.TabIndex = 10
$ButtonRT.Location = New-Object System.Drawing.Size(150,50)
$ButtonRT.Name = "RemoveTime"
$ButtonRT.BackColor = 'FloralWhite'
$ButtonRT.Add_Click({RTButton}) 
$ButtonRT.UseVisualStyleBackColor = $True

# DODANIE KONTROLEK DO FORMULARZA
$Form1.Controls.Add($ButtonCL)
$Form1.Controls.Add($ButtonCU)
$CopyRemove.Controls.Add($ButtonF2C)
$CopyRemove.Controls.Add($ButtonF2R)
$CopyRemove.Controls.Add($ButtonCT)
$CopyRemove.Controls.Add($ButtonRT)

$MatchRemover.Controls.Add($TextBox1)
$MatchRemover.Controls.Add($LabelOR)
$MatchRemover.Controls.Add($LabelFE)
$MatchRemover.Controls.Add($LabelV2F)
$MatchRemover.Controls.Add($TextBox2)
$MatchRemover.Controls.Add($TextBox3)
$MatchRemover.Controls.Add($ButtonOK)
$MatchRemover.Controls.Add($ButtonClear)
$MatchRemover.Controls.Add($ButtonFnD)

$Form1.Controls.Add($Checkbox1)
$Form1.Controls.Add($Checkbox2)
$Form1.Controls.Add($Checkbox3)
$Form1.Controls.Add($Checkbox4)
$Form1.Controls.Add($Checkbox5)
$Form1.Controls.Add($Checkbox6)

$Form1.ShowDialog()