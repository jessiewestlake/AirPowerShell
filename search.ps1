#####################################################################
###############               FUNCTIONS               ###############
#####################################################################


Function SelectFolder
{
    param([string]$Description="Select Folder",[string]$RootFolder="Desktop")

        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Application]::EnableVisualStyles()   

        $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

#############################################

Function InputBox
{
    param([string]$Caption="Type something:",[string]$Title="PII Search")
    
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$script:x = $null

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = $Title
$objForm.Size = New-Object System.Drawing.Size(300,200) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$script:x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$script:x=$objTextBox.Text;$objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = $Caption
$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$script:x
}

#############################################

Function MakeSelection
{
    param([string]$Caption="Make a selection:",[string]$Title="PII Search")
    
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$script:y = $null

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = $Title
$objForm.Size = New-Object System.Drawing.Size(300,200) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$script:y="Report";$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$ReportButton = New-Object System.Windows.Forms.Button
$ReportButton.Location = New-Object System.Drawing.Size(10,75)
$ReportButton.Size = New-Object System.Drawing.Size(75,23)
$ReportButton.Text = "Report"
$ReportButton.Add_Click({$script:y="Report";$objForm.Close()})
$objForm.Controls.Add($ReportButton)

$MoveButton = New-Object System.Windows.Forms.Button
$MoveButton.Location = New-Object System.Drawing.Size(105,75)
$MoveButton.Size = New-Object System.Drawing.Size(75,23)
$MoveButton.Text = "Move"
$MoveButton.Add_Click({$script:y="Move";$objForm.Close()})
$objForm.Controls.Add($MoveButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(200,75)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$Script:y="Cancel";$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = $Caption
$objForm.Controls.Add($objLabel) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$script:y
}

#####################################################################
###############              MAIN SCRIPT              ###############
#####################################################################

$Result = MakeSelection

If ($Result -ne "Cancel"){
    $SearchPath = SelectFolder "Select the folder you want to search:"
    $ReportPath = SelectFolder "Where do you want to save the report?"
    $ReportName = InputBox "Type a name for the report:"
    
    [System.Windows.Forms.MessageBox]::Show("Searching for files, this may take some time. Once complete an excel spreadsheet will automatically open.") 
    
    ($Files = gci "$SearchPath\*" -include *UPMR*, *SURF*, "*(PA)*", *2583*.xfdl, *SF86*, *FOUO*medical*, *AF469*, *Visitor*Request*.xls*, *eQIP*, *epr*.xfdl, *af910*.xfdl, *af911*.xfdl, *opr*.xfdl, *eval*.xfdl, *epr*.pdf, *af910*.pdf, *af911*.pdf, *opr*.pdf,*eval*.xfdl, *feedback*.xfdl, *loc*.doc*, *loc*.pdf, *lor*.doc*, *lor*.pdf, *loa*.doc*, *loa*.pdf, *alpha*roster*, *recall*, *SSAN*, *AF1466*, *AF1566*, *SGLV*, *SF182*, "*allocation notice*" -exclude *bachelor*, *baylor*, *binder*, *blank*, *cert*, *checklist*, *color*, *counselor*, *dilorenzo*, *explor*, *example*, *FAQ*, *floresca*, *Flore*, *florante*, *format*, *galore*, *Guillory*, *Gloria*, *GTC-101*, *guide*, *handbook*, *instruction*, *isolation*, *isoproponal*, *intlorg*, *invlord*, *LOAC*, *load*, *load*, *local*, *locat*, *lock*, *loctite*, *lord*, *nissan*, *Process*, *Sailor*, *sample*, *ssance*, *surfa*, *surfb*, *surfe*, *surfi*, *Suzuki*, *taylor*, *template*, *training*, *traylor*, *understanding* -force -recurse -ea SilentlyContinue | Where-Object {($_.PSIsContainer -eq $False)}) | Select Name, Directory, Length, LastWriteTime, FullName, Extension |
    Export-CSV -path "$ReportPath\$ReportName.csv" -notype

    If ($Result -eq "Move"){
    
        $DestinationPath = SelectFolder "Where do you want to store the files?"
        If (!(test-path $DestinationPath)){New-Item -ItemType Directory -Force -Path $DestinationPath}
        
        ForEach ($File in $Files){
            $FileName = Join-Path -Path $DestinationPath -ChildPath $File.Name
            $FileNameNU = $FileName
            $Num = 0
            While ((Test-Path -LiteralPath $FileNameNU) -eq $true){
                
                $FileNameNU = Join-Path $DestinationPath ($File.BaseName + " - [PSv_" + ++$Num + "]" + $File.Extension)
            }
            Move-Item -Path $File.FullName -Destination $FileNameNU -ea SilentlyContinue
        }
    }
    
    If(Test-Path ("$ReportPath\$ReportName.csv")){Invoke-Item "$ReportPath\$ReportName.csv"}
}