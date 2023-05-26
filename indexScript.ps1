# LOGO
$base64String = "DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg
ICAgICANCiAgICAgICAgICAgICAgICAgLC0tLiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIA0KI
CAgLC0tLSwgICAgICAgLC0tLid8ICAgICwtLS0sICAgICAgICAsLS0tLC4gLC0tLCAgICAgLC0tLCAgICAgLC0tLSwuLC0uLS0tLS4gICAgDQosYC0tLicgfC
AgICwtLSw6ICA6IHwgIC4nICAuJyBgXCAgICAsJyAgLicgfCB8Jy4gXCAgIC8gLmB8ICAgLCcgIC4nIHxcICAgIC8gIFwgICANCnwgICA6ICA6LGAtLS4nYHw
gICcgOiwtLS0uJyAgICAgXCAsLS0tLicgICB8IDsgXCBgXCAvJyAvIDsgLC0tLS4nICAgfDsgICA6ICAgIFwgIA0KOiAgIHwgICd8ICAgOiAgOiAgfCB8fCAg
IHwgIC5gXCAgfHwgICB8ICAgLicgYC4gXCAgLyAgLyAuJyB8ICAgfCAgIC4nfCAgIHwgLlwgOiAgDQp8ICAgOiAgfDogICB8ICAgXCB8IDo6ICAgOiB8ICAnI
CB8OiAgIDogIHwtLCAgXCAgXC8gIC8gLi8gIDogICA6ICB8LSwuICAgOiB8OiB8ICANCicgICAnICA7fCAgIDogJyAgJzsgfHwgICAnICcgIDsgIDo6ICAgfC
AgOy98ICAgXCAgXC4nICAvICAgOiAgIHwgIDsvfHwgICB8ICBcIDogIA0KfCAgIHwgIHwnICAgJyA7LiAgICA7JyAgIHwgOyAgLiAgfHwgICA6ICAgLicgICA
gXCAgOyAgOyAgICB8ICAgOiAgIC4nfCAgIDogLiAgLyAgDQonICAgOiAgO3wgICB8IHwgXCAgIHx8ICAgfCA6ICB8ICAnfCAgIHwgIHwtLCAgIC8gXCAgXCAg
XCAgIHwgICB8ICB8LSw7ICAgfCB8ICBcICANCnwgICB8ICAnJyAgIDogfCAgOyAuJycgICA6IHwgLyAgOyAnICAgOiAgOy98ICA7ICAvXCAgXCAgXCAgJyAgI
DogIDsvfHwgICB8IDtcICBcIA0KJyAgIDogIHx8ICAgfCAnYC0tJyAgfCAgIHwgJ2AgLC8gIHwgICB8ICAgIFwuL19fOyAgXCAgOyAgXCB8ICAgfCAgICBcOi
AgICcgfCBcLicgDQo7ICAgfC4nICcgICA6IHwgICAgICA7ICAgOiAgLicgICAgfCAgIDogICAuJ3wgICA6IC8gXCAgXCAgO3wgICA6ICAgLic6ICAgOiA6LSc
gICANCictLS0nICAgOyAgIHwuJyAgICAgIHwgICAsLicgICAgICB8ICAgfCAsJyAgOyAgIHwvICAgXCAgJyB8fCAgIHwgLCcgIHwgICB8LicgICAgIA0KICAg
ICAgICAnLS0tJyAgICAgICAgJy0tLScgICAgICAgIGAtLS0tJyAgICBgLS0tJyAgICAgYC0tYCBgLS0tLScgICAgYC0tLScgICAgICAgDQogICAgICAgICAgI
CAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICANCg=="

# Decodes Logo and prints to screen
function Decode-Base64AsciiArt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Base64String
    )

    try {
        $decodedContent = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($Base64String))
        Write-Output $decodedContent
    }
    catch {
        Write-Host "Error decoding the base64 string: $_"
    }
}
function New-CSVFile {
    try {
        Add-Type -AssemblyName System.Windows.Forms

        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveFileDialog.Title = "Save the new CSV file"
        $saveFileDialog.AddExtension = $true

        if ($saveFileDialog.ShowDialog() -eq 'OK') {
            $fileName = $saveFileDialog.FileName
            $filePath = [System.IO.Path]::ChangeExtension($fileName, "csv")

            $csvData = @(
                "TERM,BOOK,PAGE,DETAILS"
            )

            $csvData | Out-File -FilePath $filePath -Encoding UTF8 -Force

            Write-Host "New CSV file created and saved: $filePath"
            return $filePath
        }
        else {
            Write-Host "No file selected. CSV file creation canceled."
            return $null
        }
    }
    catch {
        Write-Host "Error creating the CSV file: $_"
        return $null
    }
}

function Select-CSVFile {
    Write-Host "Please select a CSV"
    Start-Sleep -Milliseconds 1000

    Add-Type -AssemblyName System.Windows.Forms

    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $fileDialog.Title = "Select a CSV file"

    if ($fileDialog.ShowDialog() -eq 'OK') {
        $filePath = $fileDialog.FileName
        Write-Host "Selected file: $filePath"
        return $filePath
    }
    else {
        Write-Host "No file selected."
        return $null
    }
}

function Show-OpenOrCreateCSVForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Open or Create CSV"
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.FormBorderStyle = "FixedDialog"
    $form.StartPosition = "CenterScreen"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $openButton = New-Object System.Windows.Forms.Button
    $openButton.Location = New-Object System.Drawing.Point(50, 40)
    $openButton.Size = New-Object System.Drawing.Size(200, 30)
    $openButton.Text = "Open Existing CSV"
    $openButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Location = New-Object System.Drawing.Point(50, 80)
    $createButton.Size = New-Object System.Drawing.Size(200, 30)
    $createButton.Text = "Create New CSV"
    $createButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    })

    $form.Controls.Add($openButton)
    $form.Controls.Add($createButton)

    $form.ShowDialog() | Out-Null

    if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = Select-CSVFile
    }
    elseif ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $csvPath = New-CSVFile
    }
    else {
        $csvPath = $null
    }

    $form.Dispose()

    $csvPath
}

function Add-CSVEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath
    )

    Add-Type -AssemblyName System.Windows.Forms
    # Prompt for book using a popup dialog
    $bookInputDialog = New-Object System.Windows.Forms.Form
    $bookInputDialog.Text = "Enter Book Number"
    $bookInputDialog.AutoSize = $true
    $bookInputDialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $layoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $layoutPanel.AutoSize = $true
    $layoutPanel.ColumnCount = 2
    $layoutPanel.RowCount = 2
    $layoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
    $bookInputDialog.Controls.Add($layoutPanel)

    $bookLabel = New-Object System.Windows.Forms.Label
    $bookLabel.Text = "Book Number:"
    $bookLabel.AutoSize = $true
    $bookLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $layoutPanel.Controls.Add($bookLabel, 0, 0)

    $bookTextBox = New-Object System.Windows.Forms.TextBox
    $bookTextBox.Width = 100
    $layoutPanel.Controls.Add($bookTextBox, 1, 0)

    $bookSubmitButton = New-Object System.Windows.Forms.Button
    $bookSubmitButton.Text = "Submit"
    $bookSubmitButton.Anchor = [System.Windows.Forms.AnchorStyles]::None
    $bookSubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $layoutPanel.SetColumnSpan($bookSubmitButton, 2)
    $layoutPanel.Controls.Add($bookSubmitButton, 0, 1)

    $bookInputDialogResult = $bookInputDialog.ShowDialog()

    if ($bookInputDialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $Book = $bookTextBox.Text
    }
    else {
        return
    }

    while ($true) {
        # Create a dialog box for term and page input
        $inputDialog = New-Object System.Windows.Forms.Form
        $inputDialog.Text = "Enter Term and Page"
        $inputDialog.AutoSize = $true
        $inputDialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        $layoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
        $layoutPanel.AutoSize = $true
        $layoutPanel.ColumnCount = 2
        $layoutPanel.RowCount = 3
        $layoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
        $inputDialog.Controls.Add($layoutPanel)

        $termLabel = New-Object System.Windows.Forms.Label
        $termLabel.Text = "Term:"
        $termLabel.AutoSize = $true
        $termLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $layoutPanel.Controls.Add($termLabel, 0, 0)

        $termTextBox = New-Object System.Windows.Forms.TextBox
        $termTextBox.Width = 200
        $layoutPanel.Controls.Add($termTextBox, 1, 0)

        $pageLabel = New-Object System.Windows.Forms.Label
        $pageLabel.Text = "Page:"
        $pageLabel.AutoSize = $true
        $pageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $layoutPanel.Controls.Add($pageLabel, 0, 1)

        $pageTextBox = New-Object System.Windows.Forms.TextBox
        $pageTextBox.Width = 200
        $layoutPanel.Controls.Add($pageTextBox, 1, 1)

        $submitButton = New-Object System.Windows.Forms.Button
        $submitButton.Text = "Submit"
        $submitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $submitButton.Anchor = [System.Windows.Forms.AnchorStyles]::None
        $layoutPanel.SetColumnSpan($submitButton, 2)
        $layoutPanel.Controls.Add($submitButton, 0, 2)

        $inputDialogResult = $inputDialog.ShowDialog()

        if ($inputDialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $term = $termTextBox.Text
            $page = $pageTextBox.Text
        }
        else {
            break
        }

        # Prompt for details using a popup dialog
        $detailsInputDialog = New-Object System.Windows.Forms.Form
        $detailsInputDialog.Text = "Enter Details"
        $detailsInputDialog.AutoSize = $true
        $detailsInputDialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        $detailsLabel = New-Object System.Windows.Forms.Label
        $detailsLabel.Text = "Details:"
        $detailsLabel.AutoSize = $true
        $detailsInputDialog.Controls.Add($detailsLabel)

        $detailsTextBox = New-Object System.Windows.Forms.TextBox
        $detailsTextBox.Top = $detailsLabel.Bottom + 5
        $detailsTextBox.Width = 400
        $detailsTextBox.Height = 200
        $detailsTextBox.Multiline = $true
        $detailsInputDialog.Controls.Add($detailsTextBox)

        $detailsSubmitButton = New-Object System.Windows.Forms.Button
        $detailsSubmitButton.Text = "Submit"
        $detailsSubmitButton.Top = $detailsTextBox.Bottom + 10
        $detailsSubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $detailsInputDialog.AcceptButton = $detailsSubmitButton
        $detailsInputDialog.Controls.Add($detailsSubmitButton)

        $detailsInputDialogResult = $detailsInputDialog.ShowDialog()

        if ($detailsInputDialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $details = $detailsTextBox.Text
        }
        else {
            break
        }

        # Create a new CSV row string with values enclosed in double quotes
        $newRow = "`"$term`",`"$Book`",`"$page`",`"$details`""

        # Append the new row to the existing CSV file or create a new file if it doesn't exist
        $newRow | Out-File -FilePath $CsvPath -Append -Encoding UTF8

        # Prompt to add another entry or exit using a popup dialog
        $choice = [System.Windows.Forms.MessageBox]::Show("Do you want to add another entry?", "Add Another Entry", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($choice -eq [System.Windows.Forms.DialogResult]::No) {
            break
        }
    }
}

Decode-Base64AsciiArt -Base64String $base64String
$CsvPath = Show-OpenOrCreateCSVForm
Add-CSVEntry -CsvPath $CsvPath 
