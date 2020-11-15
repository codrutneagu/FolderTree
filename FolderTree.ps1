#Made by Codrut Neagu
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$FolderTree                       = New-Object system.Windows.Forms.Form
$FolderTree.ClientSize            = '340,248'
$FolderTree.text                  = "Get Folder Tree"
$FolderTree.BackColor             = "#ffffff"
$FolderTree.TopMost               = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.BackColor               = "#d0021b"
$Button1.text                    = "Get Folder Tree to Excel"
$Button1.width                   = 205
$Button1.height                  = 50
$Button1.location                = New-Object System.Drawing.Point(64,43)
$Button1.Font                    = 'Microsoft Sans Serif,10,style=Bold'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Tool provided by digitalcitizen.life"
$Label1.BackColor                = "#ffffff"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(125,221)
$Label1.Font                     = 'Microsoft Sans Serif,10,style=Underline'
$Label1.ForeColor                = "#4a90e2"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "The folder tree Excel file is saved in`r`nthe folder where you run this tool."
$Label2.BackColor                = "#ffffff"
$Label2.AutoSize                 = $true
$Label2.width                    = 300
$Label2.height                   = 40
$Label2.location                 = New-Object System.Drawing.Point(20,140)
$Label2.Font                     = 'Microsoft Sans Serif,10,style=Underline'
$Label2.ForeColor                = "#85200c"

$FolderTree.controls.AddRange(@($Button1,$Button2,$Label1,$Label2))

$Button1.Add_Click({
GetFolderTree
})

$Label1.Add_Click({ opendigitalcitizen })
$TextBox

function GetFolderTree {

$path = (Get-Item -Path ".\").FullName
Set-Location -Path $path

Get-ChildItem -Recurse |
     ForEach-Object {$_} | 
        Select-Object -Property Directory,Name | 
        Where-Object {$_.Directory -ne $null} | 
		Sort-Object -Property Directory,Name |
        Export-Csv -Force -NoTypeInformation "$path\FolderTree.csv"

$csv = "$path\FolderTree.csv" #Location of the source file
$xlsx = "$path\FolderTree.xlsx" #Desired location of output
$delimiter = "," #Specify the delimiter used in the file

# Create a new Excel workbook with one empty sheet
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $csv)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()

# Save & close the Workbook as XLSX.
$Workbook.SaveAs($xlsx,51)
$excel.Quit()
}

function opendigitalcitizen {
    start https://www.digitalcitizen.life
}

$FolderTree.ShowDialog()