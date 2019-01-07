<# Script to create a text string for the Alma 909 field and copy the string to the clipboard

Created: 2018-12-05 by V.M. Downey for Houghton Library Technical Services

#>


$subfielda = "`$`$a"
$subfielde = "`$`$e HO"
$subfieldd = "`$`$d " + (Get-Date -Format FileDate)
$subfieldf = "`$`$f " + $env:USERNAME
$subfieldx = "`$`$x "

" " | Set-Clipboard


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object System.Windows.Forms.Form
$Form.ClientSize                 = '250,400'
$Form.Text                       = "Alma 920 Field"
$Form.TopMost                    = $false

$Label0                          = New-Object System.Windows.Forms.Label
$Label0.Text                     = $subfielda + " (cataloging type)"
$Label0.AutoSize                 = $true
$Label0.Width                    = 25
$Label0.Height                   = 10
$Label0.Location                 = New-Object System.Drawing.Point(18,8)


$ListBox1                        = New-Object System.Windows.Forms.ListBox
$ListBox1.Width                  = 200
$ListBox1.Height                 = 60
$ListBox1.Location               = New-Object System.Drawing.Point(18,30) 
[void] $ListBox1.Items.Add('accession')
[void] $ListBox1.Items.Add('original')
[void] $ListBox1.Items.Add('copy')
[void] $ListBox1.Items.Add('enhanced')
$ListBox1.SetSelected(0, $true) #Change number in parentheses to change the default

$Label1                          = New-Object System.Windows.Forms.Label
$Label1.Text                     = $subfieldd
$Label1.AutoSize                 = $true
$Label1.Width                    = 25
$Label1.Height                   = 10
$Label1.Location                 = New-Object System.Drawing.Point(18,120)

$Label2                          = New-Object System.Windows.Forms.Label
$Label2.Text                     = $subfielde
$Label2.AutoSize                 = $true
$Label2.Width                    = 25
$Label2.Height                   = 10
$Label2.Location                 = New-Object System.Drawing.Point(18,140)

$Label3                          = New-Object System.Windows.Forms.Label
$Label3.Text                     = $subfieldf
$Label3.AutoSize                 = $true
$Label3.Width                    = 25
$Label3.Height                   = 10
$Label3.Location                 = New-Object System.Drawing.Point(18,160)


$Label4                          = New-Object System.Windows.Forms.Label
$Label4.Text                     = $subfieldx + " (curatorial department)"
$Label4.AutoSize                 = $true
$Label4.Width                    = 25
$Label4.Height                   = 10
$Label4.Location                 = New-Object System.Drawing.Point(18,200)

$ListBox2                        = New-Object System.Windows.Forms.ListBox
$ListBox2.Width                  = 200
$ListBox2.Height                 = 120
$ListBox2.Location               = New-Object System.Drawing.Point(18,220) 
[void] $ListBox2.Items.Add('modbm')
[void] $ListBox2.Items.Add('pga')
[void] $ListBox2.Items.Add('earbm')
[void] $ListBox2.Items.Add('hew')
[void] $ListBox2.Items.Add('htc')
[void] $ListBox2.Items.Add('hyde')
[void] $ListBox2.Items.Add('wpr')
[void] $ListBox2.Items.Add('trc')
$ListBox2.SetSelected(0, $true) #Change number in parentheses to change the default

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(18,360)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(93,360)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton


$Form.controls.AddRange(@($Label0, $ListBox1,$Label1,$Label2,$Label3,$Label4, $ListBox2, $OKButton, $CancelButton))

#region gui events {

#endregion events }

#endregion GUI }


$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)

{
    if ($ListBox1.SelectedItem -eq $null) {$field920 = " "} 
    elseif ($ListBox2.SelectedItem -eq $null) {$field920 = " "} 
    else { 
    $subfielda = ("`$`$a " + $ListBox1.SelectedItem[0])
    $subfieldx = ("`$`$x " + $ListBox2.SelectedItem)
    $field920 = ($subfielda, $subfieldd, $subfielde, $subfieldf, $subfieldx -join " ")}
    
    $field920  | Set-Clipboard      
}