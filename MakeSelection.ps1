If ($Result -ne "Cancel"){
    $SearchPath = SelectFolder "Select the folder you want to search:"
    $ReportPath = SelectFolder "Where do you want to save the report?"
    $ReportName = InputBox "Type a name for the report:"
    
    [System.Windows.Forms.MessageBox]::Show("Searching for files, this may take some time. Once complete an excel spreadsheet will automatically open.") 
    
    ($Files = gci "$SearchPath\*" -include *UPMR*, *SURF*, "*(PA)*", *2583*.xfdl, *SF86*, *FOUO*medical*, *AF469*, *Visitor*Request*.xls*, *eQIP*, *epr*.xfdl, *af910*.xfdl, *af911*.xfdl, *opr*.xfdl, *eval*.xfdl, *feedback*.xfdl, *loc*.doc*, *loc*.pdf, *lor*.doc*, *lor*.pdf, *loa*.doc*, *loa*.pdf, *alpha*roster*, *recall*, *SSAN*, *AF1466*, *AF1566*, *SGLV*, *SF182*, "*allocation notice*" -exclude *bachelor*, *baylor*, *binder*, *blank*, *cert*, *checklist*, *color*, *counselor*, *dilorenzo*, *explor*, *example*, *FAQ*, *floresca*, *Flore*, *florante*, *format*, *galore*, *Guillory*, *Gloria*, *GTC-101*, *guide*, *handbook*, *instruction*, *isolation*, *isoproponal*, *intlorg*, *invlord*, *LOAC*, *load*, *load*, *local*, *locat*, *lock*, *loctite*, *lord*, *nissan*, *Process*, *Sailor*, *sample*, *ssance*, *surfa*, *surfb*, *surfe*, *surfi*, *Suzuki*, *taylor*, *template*, *training*, *traylor*, *understanding* -force -recurse -ea SilentlyContinue | Where-Object {($_.PSIsContainer -eq $False)}) | Select Name, Directory, Length, LastWriteTime |
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
