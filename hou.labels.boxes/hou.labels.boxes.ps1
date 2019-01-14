#Windows PowerShell script to create box labels from ArchivesSpace container TSV files

#Function to enable filename picker dialog box
Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null
 Add-Type -AssemblyName System.Windows.Forms
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.Title = "Choose an ArchivesSpace container TSV file to create labels"
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "TSV (*.tsv)|*.tsv"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

#Run dialog box function
$tsv = Get-FileName -initialDirectory "c:fso"

#Exit script if user cancels
if ($tsv -eq "") { exit 0}

#Show error message and exit script if ASpace TSV is blank or only has a header row
if ((Get-Content $tsv | Measure-Object | Select-Object).Count -le 1)
{
(New-Object -ComObject Wscript.Shell).Popup("Box labels cannot be created for this ArchivesSpace container file.`n`nCheck the file exported from ASpace:`n`n$tsv.",0,"Script Error")
exit
}

#Copy ASpace container label data to hou.labels.boxes.tsv
$content = (Get-Content $tsv) -replace "Container`tLabel", "Container Label" #Remove stray tab supplied by ASpace between "Container" and "Label" in the header
if (-not $content[0].EndsWith("Label")) #Try to repair files that are missing newline characters by adding them before the "Repository Name" field
{
$content = $content -replace "Houghton Library, Harvard Library, Harvard University", "`r`nHoughton Library, Harvard Library, Harvard University"
} 
#Save to hou.labels.boxes.tsv
$content | Set-Content -Path "$(Get-Location)\hou.labels.boxes.tsv" -force

#Perform mail merge in Word using the transformed file
$word = New-Object -ComObject "Word.Application"
$word.Visible = $true

#ALT-TAB to bring Word labels document to front
[System.Windows.Forms.SendKeys]::SendWait("%{TAB}")

$template = "$(Get-Location)\hou.labels.boxes.dotx"

$doc = $word.Documents.Add($template)
$doc.MailMerge.OpenDataSource("$(Get-Location)\hou.labels.boxes.tsv")
$doc.MailMerge.Execute()

#Close the mail merge template
$doc.Close([ref] 0)