<#
.SYNOPSIS
   Opens a folder selection dialog box.
.DESCRIPTION
   Uses the System.Windows.Forms class to generate a FolderBrowserDialog object to display a GUI. The GUI 
   allows selection of a folder on the system and returns the selected folder path to the pipeline.
.EXAMPLE
   Show-GUIFolderBrowserDialog

   Opens a folder selection GUI with the default caption/description and default root folder selected.
.EXAMPLE
   Show-GUIFolderBrowserDialog -Caption 'Choose a folder to save the log file' -RootFolder 'MyDocuments' -ShowNewFolderButton

   Displays the folder selection GUI with a custom caption, the user Documents folder as the root folder, 
   and the New Folder button enabled allowing a user to create a new folder from the dialog.
.INPUTS
   None, or a named String object representating the Caption parameter.
.OUTPUTS
   A System.String object representating the full path to the selected folder. Will return exceptions if the 
   dialog is terminated.
.NOTES
   Author: TSgt J. Jessie Westlake
.COMPONENT
   None
.ROLE
   GUI
.FUNCTIONALITY
   Opens a GUI to select a folder and return its full path as a String.
#>
function Show-GUIFolderBrowserDialog
{
    [CmdletBinding(HelpUri = 'https://github.com/jessiewestlake')]
    [Alias('SelectFolder')]
    [OutputType([String])]
    Param
    (
        # Determines the text displayed in the dialog box. The default is "Select Folder".
        [Parameter(ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('Prompt')]
        [String]
        $Caption = 'Select Folder',

        # Determines the root folder of the selection dialog. The parameter enumerates all of the acceptable values. The default is "Desktop".
        [Parameter(Position=1)]
        [System.Environment+SpecialFolder]
        $RootFolder = 'Desktop',

        # When present, this Switch parameter enables the 'New Folder' button on the GUI, allowing a user to create a new folder on the fly.
        [Parameter(Position=2)]
        [Switch]
        $ShowNewFolderButton
    )

    Begin
    {
        Add-Type -AssemblyName 'System.Windows.Forms'
    }
    Process
    {
        $Form = New-Object System.Windows.Forms.FolderBrowserDialog

        $Form.Description = $Caption
        $Form.RootFolder = $RootFolder
        $Form.ShowNewFolderButton = $ShowNewFolderButton

        $ButtonPressed = $Form.ShowDialog()

        If ($ButtonPressed -eq 'OK')
        {
            Write-Output $Form.SelectedPath
        }
        ElseIf ($ButtonPressed -eq 'Cancel')
        {
            Throw 'Operation cancelled by user.'
        }
        Else
        {
            Throw 'Operation was interrupted unexpectedly.'
        }
    }
    End
    {
    }
}
