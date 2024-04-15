<# Script to create a text string for the Alma 909 field and copy the string to the clipboard

Created: 2018-12-05 by V.M. Downey for Houghton Library Technical Services
Updated: 2022-06-14 by V.M. Downey to add new $$x codes and enlarge box size
Updated: 2024-03-12 by V.M. Downey to add $$a r for reparative description and $$p for projects
Updated: 2024-03-18 by V.M. Downey to change numbered label/list box variables to descriptive variable names 

#>

$SubfieldA = "`$`$a"
$SubfieldE = "`$`$e HO"
#$SubfieldD = "`$`$d " + (Get-Date -Format FileDate) #FileDate available for PS v. 5.0 or later
$SubfieldD = "`$`$d " + (Get-Date -Format 'yyyyMMdd')
$SubfieldF = "`$`$f " + $env:USERNAME
$SubfieldP = "`$`$p "
$SubfieldX = "`$`$x "

" " | Set-Clipboard

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object System.Windows.Forms.Form
$Form.ClientSize                 = '250,525'
$Form.Text                       = "Alma 920 Field"
$Form.TopMost                    = $false

$LabelA                          = New-Object System.Windows.Forms.Label
$LabelA.Text                     = $SubfieldA + " (cataloging type)"
$LabelA.AutoSize                 = $true
$LabelA.Width                    = 25
$LabelA.Height                   = 10
$LabelA.Location                 = New-Object System.Drawing.Point(18,8)

$ListBoxA                        = New-Object System.Windows.Forms.ListBox
$ListBoxA.Width                  = 200
$ListBoxA.Height                 = 90
$ListBoxA.Location               = New-Object System.Drawing.Point(18,30) 
[void] $ListBoxA.Items.Add('')
[void] $ListBoxA.Items.Add('accession')
[void] $ListBoxA.Items.Add('original')
[void] $ListBoxA.Items.Add('copy')
[void] $ListBoxA.Items.Add('enhanced')
[void] $ListBoxA.Items.Add('reparative description')
$ListBoxA.SetSelected(0, $true) #Change number in parentheses to change the default

$LabelD                          = New-Object System.Windows.Forms.Label
$LabelD.Text                     = $SubfieldD
$LabelD.AutoSize                 = $true
$LabelD.Width                    = 25
$LabelD.Height                   = 10
$LabelD.Location                 = New-Object System.Drawing.Point(18,130)

$LabelE                          = New-Object System.Windows.Forms.Label
$LabelE.Text                     = $SubfieldE
$LabelE.AutoSize                 = $true
$LabelE.Width                    = 25
$LabelE.Height                   = 10
$LabelE.Location                 = New-Object System.Drawing.Point(18,150)

$LabelF                          = New-Object System.Windows.Forms.Label
$LabelF.Text                     = $SubfieldF
$LabelF.AutoSize                 = $true
$LabelF.Width                    = 25
$LabelF.Height                   = 10
$LabelF.Location                 = New-Object System.Drawing.Point(18,170)

$LabelP                          = New-Object System.Windows.Forms.Label
$LabelP.Text                     = $SubfieldP + " (project code)"
$LabelP.AutoSize                 = $true
$LabelP.Width                    = 25
$LabelP.Height                   = 10
$LabelP.Location                 = New-Object System.Drawing.Point(18,190)

$ListBoxP                        = New-Object System.Windows.Forms.ListBox
$ListBoxP.Width                  = 200
$ListBoxP.Height                 = 80
$ListBoxP.Location               = New-Object System.Drawing.Point(18,210) 
[void] $ListBoxP.Items.Add('')
[void] $ListBoxP.Items.Add('new acq.')
[void] $ListBoxP.Items.Add('aargh')
[void] $ListBoxP.Items.Add('shelf-read')
[void] $ListBoxP.Items.Add('project ')
$ListBoxP.SetSelected(0, $true) #Change number in parentheses to change the default

$LabelX                          = New-Object System.Windows.Forms.Label
$LabelX.Text                     = $SubfieldX + " (curatorial department)"
$LabelX.AutoSize                 = $true
$LabelX.Width                    = 25
$LabelX.Height                   = 10
$LabelX.Location                 = New-Object System.Drawing.Point(18,290)

$ListBoxX                        = New-Object System.Windows.Forms.ListBox
$ListBoxX.Width                  = 200
$ListBoxX.Height                 = 170
$ListBoxX.Location               = New-Object System.Drawing.Point(18,310) 
[void] $ListBoxX.Items.Add('')
[void] $ListBoxX.Items.Add('earbm')
[void] $ListBoxX.Items.Add('fal')
[void] $ListBoxX.Items.Add('hew')
[void] $ListBoxX.Items.Add('hfa')
[void] $ListBoxX.Items.Add('htc')
[void] $ListBoxX.Items.Add('hyde')
[void] $ListBoxX.Items.Add('modbm')
[void] $ListBoxX.Items.Add('mus')
[void] $ListBoxX.Items.Add('pga')
[void] $ListBoxX.Items.Add('trc')
[void] $ListBoxX.Items.Add('wpr')
$ListBoxX.SetSelected(0, $true) #Change number in parentheses to change the default

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(18,480)
$OKButton.Size = New-Object System.Drawing.Size(75,25)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $OKButton

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(100,480)
$CancelButton.Size = New-Object System.Drawing.Size(75,25)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $CancelButton

$Form.controls.AddRange(@($LabelA, $ListBoxA, $LabelD, $LabelE, $LabelF, $ListBoxP, $LabelP, $LabelX, $ListBoxX, $OKButton, $CancelButton))

#region gui events {

#endregion events }

#endregion GUI }

$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)

{
    if ($ListBoxA.SelectedItem -eq $null) {$Field920 = " "} 
    elseif ($ListBoxP.SelectedItem -eq $null) {$Field920 = " "} 
    else { 
    $SubfieldA = ("`$`$a " + $ListBoxA.SelectedItem[0])
    $SubfieldP = ("`$`$p " + $ListBoxP.SelectedItem)
    $SubfieldX = ("`$`$x " + $ListBoxX.SelectedItem)
    $Field920 = ($SubfieldA, $SubfieldD, $SubfieldE, $SubfieldF, $SubfieldP, $SubfieldX -join " ")}
    #Example output: $$a o $$d 20240318 $$e HO $$f abc123 $$p new acq. $$x earbm
    
    #$Field920  | Set-Clipboard # Set-Clipboard available for PS v. 5.0 or later
    [System.Windows.Forms.Clipboard]::SetText($Field920)      
}
