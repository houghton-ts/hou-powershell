#Windows PowerShell script to create folder labels from ArchivesSpace EAD files


#Function to enable filename picker dialog box
Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null
 Add-Type -AssemblyName System.Windows.Forms
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.Title = "Choose an XML file to create labels"
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "XML (*.xml)|*.xml"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

#Run dialog box function
$xml = Get-FileName -initialDirectory "c:fso"

#Exit script if user cancels
if ($xml -eq "") { exit 0}

#Run folder label XSLT against selected file
$template = "$(Get-Location)\hou.labels.folders.dotx"
$xslt = New-Object System.Xml.Xsl.XslCompiledTransform
$xslt.Load("hou.labels.folders.xsl")
$xslt.Transform($xml,"hou.labels.folders.tsv")

#Show error message and exit script if XSLT produces a TSV with 0 or 1 row(s)
if ((Get-Content "hou.labels.folders.tsv" | Measure-Object | Select-Object).Count -le 1)
{
(New-Object -ComObject Wscript.Shell).Popup("Folder labels cannot be created for this EAD file",0,"Script Error")
exit
}

# Process TSV to add missing UNITIDS, remove duplicates, and sort rows
$content = Import-Csv -Delimiter "`t" -Path hou.labels.folders.tsv
# Add columns for numerically sorting box, folder, and unitid values
$content | Add-Member -MemberType NoteProperty -Name 'SORTB' -Value $null
$content | Add-Member -MemberType NoteProperty -Name 'SORTF' -Value $null 
$content | Add-Member -MemberType NoteProperty -Name 'SORTU' -Value $null
# If unitid, box, or folder value is present, keep values and add sort columns for each (padded to 9 digits)
# If no values are present, look in C05, C04, C03, C02, C01 for unitid value
# Remove any rows lacking a folder title
# Sort values and remove duplicates
$content | ForEach-Object {
$_.UNITID, $_."FOLDER TITLE", $_."FOLDER DATES", $_.SORTB, $_.SORTF, $_.SORTU =
if ($_.UNITID -or $_.BOX -or $_.FOLDER) {$_.UNITID, $_."FOLDER TITLE", $_."FOLDER DATES", ($_.BOX -replace "[^0-9]").PadLeft(9, "0"), ($_.FOLDER -replace "[^0-9]").PadLeft(9, "0"), ($_.UNITID -replace "[^0-9]").PadLeft(9, "0")} else {
if ($_."C05 ANCESTOR" -match "\(\d") {($_."C05 ANCESTOR".split(")")[0] + ")"), ($_."C05 ANCESTOR".split(")",2)[1]), "", "000000000", "000000000", ($_."C05 ANCESTOR".split(")")[0] -replace "[^0-9]").PadLeft(9, "0")} else {
if ($_."C04 ANCESTOR" -match "\(\d") {($_."C04 ANCESTOR".split(")")[0] + ")"), ($_."C04 ANCESTOR".split(")",2)[1]), "", "000000000", "000000000", ($_."C04 ANCESTOR".split(")")[0] -replace "[^0-9]").PadLeft(9, "0")} else {
if ($_."C03 ANCESTOR" -match "\(\d") {($_."C03 ANCESTOR".split(")")[0] + ")"), ($_."C03 ANCESTOR".split(")",2)[1]), "", "000000000", "000000000", ($_."C03 ANCESTOR".split(")")[0] -replace "[^0-9]").PadLeft(9, "0")} else {
if ($_."C02 ANCESTOR" -match "\(\d") {($_."C02 ANCESTOR".split(")")[0] + ")"), ($_."C02 ANCESTOR".split(")",2)[1]), "", "000000000", "000000000", ($_."C02 ANCESTOR".split(")")[0] -replace "[^0-9]").PadLeft(9, "0")} else {
if ($_."C01 ANCESTOR" -match "\(\d") {($_."C01 ANCESTOR".split(")")[0] + ")"), ($_."C01 ANCESTOR".split(")",2)[1]), "", "000000000", "000000000", ($_."C01 ANCESTOR".split(")")[0] -replace "[^0-9]").PadLeft(9, "0")} else {
"", "", "", "000000000", "000000000", "000000000"}
}
}
}
}
}
} 

#Add suffix "xxx " to folder date if folder title ends with period or square bracket and first value in date appears in title
#Suffix is for Word mail merge identification
$content | ForEach-Object {
$_."FOLDER DATES" =
if ($_."FOLDER DATES") {
if (($_."FOLDER TITLE"[-1] -eq ".") -or ($_."FOLDER TITLE"[-1] -eq "]")) {
if ($_."FOLDER TITLE" -match ($_."FOLDER DATES" -replace "\W", " ").trim().split(" ")[0]){"xxx " + $_."FOLDER DATES"} 
else {$_."FOLDER DATES"}} 
else {$_."FOLDER DATES"}
}
 else {"xxx "}
}

#Remove trailing comma from folder titles
$content | ForEach-Object {
$_."FOLDER TITLE" =
if ($_."FOLDER TITLE"[-1] -eq ",")  {$_."FOLDER TITLE".TrimEnd(",")}
else {$_."FOLDER TITLE"}
}

$content | Where-Object {$_."Folder Title"} | sort SORTB, SORTF, SORTU, BOX, FOLDER, UNITID, "FOLDER TITLE" -Unique | Export-Csv hou.labels.folders.tsv -NoTypeInformation -Delimiter "`t" -Encoding UTF8


#Perform mail merge in Word using the transformed file
$word = New-Object -ComObject "Word.Application"
$word.Visible = $true

#ALT-TAB to bring Word labels document to front
[System.Windows.Forms.SendKeys]::SendWait("%{TAB}")

$doc = $word.Documents.Add($template)
$doc.MailMerge.OpenDataSource("$(Get-Location)\hou.labels.folders.tsv")
$doc.MailMerge.Execute()

#Close the mail merge template
$doc.Close([ref] 0)