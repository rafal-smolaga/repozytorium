Add-Type –AssemblyName System.Drawing
Add-Type –AssemblyName System.Windows.Forms

# włączenie stylów wizualnych Windows
[Windows.Forms.Application]::EnableVisualStyles()

# definicja czcionki dla elementów interfejsu
$Font = New-Object System.Drawing.Font("ExampleFontName", 10, [System.Drawing.FontStyle]::Bold)

# === etykieta nagłówka ===
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(45, 10) 
$Label.Size = New-Object System.Drawing.Size(280, 20)
$Label.AutoSize = $true
$Label.Text = "ExampleLabel_Computers:"
$Label.Font = $Font

# === pole tekstowe do wprowadzania listy komputerów ===
$TextBox = New-Object System.Windows.Forms.TextBox 
$TextBox.Location = New-Object System.Drawing.Size(10, 40) 
$TextBox.Size = New-Object System.Drawing.Size(150, 200)
$TextBox.AcceptsReturn = $true      # akceptuje znak nowej linii
$TextBox.AcceptsTab = $false        # nie akceptuje tabulacji
$TextBox.Multiline = $true          # tryb wielowierszowy
$TextBox.ScrollBars = 'Both'        # paski przewijania poziomy i pionowy
$TextBox.Font = $Font

# === przycisk OK z akcją pingowania i generowania raportu ===
$OkButton = New-Object System.Windows.Forms.Button
$OkButton.Location = New-Object System.Drawing.Size(50, 250)
$OkButton.Size = New-Object System.Drawing.Size(75, 25)
$OkButton.Text = "OK"

$OkButton.Add_Click({
    # ukrycie okna formularza
    $Form.Hide()
    
    # ścieżka do pliku wynikowego Excel
    $path = ".\example_results.xls"
    
    # utworzenie obiektu aplikacji Excel
    $Excel = new-object -comobject excel.application

    # sprawdzenie czy plik już istnieje
    if (Test-Path $path)
    { 
        $Workbook = $Excel.WorkBooks.Open($path) 
        $Worksheet = $Workbook.Worksheets.Item(1) 
    }
    else 
    { 
        $Workbook = $Excel.Workbooks.Add() 
        $Worksheet = $Workbook.Worksheets.Item(1)
    }
    
    $Excel.Visible = $True

    # nagłówki kolumn raportu
    $Worksheet.Cells.Item(1, 1) = "ExampleHostName"
    $Worksheet.Cells.Item(1, 2) = "ExampleStatus"

    # podzielenie tekstu z pola tekstowego na listę komputerów (po znakach nowej linii)
    $computers = $TextBox.Text -split "`r`n"
    $row = 2
    [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
    
    # pętla po każdym komputerze
    foreach ($computer in $computers) {
        # test połączenia ping (1 pakiet, bez komunikatów błędów)
        $ping = Test-Connection $computer -Count 1 -ea silentlycontinue
        
        if($ping){
            $Worksheet.Cells.Item($row, 1) = $computer
            $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Up"
            [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
            $row++
        }
        else {
            $Worksheet.Cells.Item($row, 1) = $computer
            $Worksheet.Cells.Item($row, 2) = "ExampleStatus_Down"
            [void]$Worksheet.UsedRange.EntireColumn.AutoFit()
            $row++
        }
    }
})

# === konfiguracja głównego okna formularza ===
$Form = New-Object System.Windows.Forms.Form 
$Form.Text = $ExampleWindowTitle
$Form.Size = New-Object System.Drawing.Size(180, 320)
$Form.FormBorderStyle = 'FixedSingle'      # stały rozmiar okna
$Form.StartPosition = "CenterScreen"       # wyśrodkowanie na ekranie
$Form.AutoSizeMode = 'GrowAndShrink'
$Form.Topmost = $True                      # okno zawsze na wierzchu
$Form.ShowInTaskbar = $true
$Form.MaximizeBox = $false                 # brak możliwości maksymalizacji
$Form.MinimizeBox = $false                 # brak możliwości minimalizacji
$Form.ShowIcon = $False

# dodanie kontrolek do formularza
$Form.Controls.Add($Label)
$Form.Controls.Add($TextBox)
$Form.Controls.Add($OkButton)

# wyświetlenie okna dialogowego
$Form.ShowDialog()