<#
.SYNOPSIS
   Opens a graphical text-input dialog box.
.DESCRIPTION
   Uses the System.Windows.Forms class to generate a custom Form object to display a GUI. The GUI allows 
   input of a text string and returns that information to the pipeline.
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   None, or named String objects representating the Caption and/or Title parameter values.
.OUTPUTS
   A System.String object representating the entered text. Will return exceptions if the dialog is terminated.
.NOTES
   Author: TSgt J. Jessie Westlake
.COMPONENT
   NONE
.ROLE
   GUI
.FUNCTIONALITY
   Opens a GUI to input text and return it as a String.
#>
function Show-GUIInputBox
{
    [CmdletBinding(HelpUri = 'https://github.com/jessiewestlake')]
    [Alias('InputBox')]
    [OutputType([String])]
    Param
    (
        # Determines the text displayed in the dialog box. The default is "Type something:".
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('Description','Message')] 
        [String]
        $Caption = 'Type something:',

        # Determines the title text displayed at the top bar of the dialog box. The default is "Input Box".
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Title = 'Input Box'
    )

    Begin
    {
        Add-Type -AssemblyName 'System.Drawing'
        Add-Type -AssemblyName 'System.Windows.Forms'
    }
    Process
    {
        $Form = New-Object System.Windows.Forms.Form 
        $Form.Size = New-Object System.Drawing.Size(300,200) 
        $Form.StartPosition = "CenterScreen"
        $Form.Text = $Title
        $Form.Topmost = $True
        $Form.Add_Shown({ $Form.Activate() })


        $Label = New-Object System.Windows.Forms.Label
        $Label.Location = New-Object System.Drawing.Size(10,20) 
        $Label.Size = New-Object System.Drawing.Size(280,20) 
        $Label.Text = $Caption
        $Form.Controls.Add($Label)


        $TextBox = New-Object System.Windows.Forms.TextBox 
        $TextBox.Location = New-Object System.Drawing.Size(10,40) 
        $TextBox.Size = New-Object System.Drawing.Size(260,20)
        $TextBox.TabStop = $True
        $TextBox.TabIndex = 0
        $TextBox.Add_TextChanged({ If($TextBox.TextLength -gt 0){$OKButton.Enabled = $True} Else{$OKButton.Enabled = $False} })
        $Form.Controls.Add($TextBox) 


        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Size(75,120)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $OKButton.Text = 'OK'
        $OKButton.Enabled = $False
        $OKButton.TabStop = $True
        $OKButton.TabIndex = 1
        $Form.AcceptButton = $OKButton
        $Form.Controls.Add($OKButton)


        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Size(150,120)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $CancelButton.Text = 'Cancel'
        $CancelButton.TabStop = $True
        $CancelButton.TabIndex = 2
        $Form.CancelButton = $CancelButton
        $Form.Controls.Add($CancelButton)


        $ButtonPressed = $Form.ShowDialog()

        If ($ButtonPressed -eq 'OK')
        {
            Write-Output $TextBox.Text
        }
        ElseIf ($ButtonPressed -eq 'Cancel')
        {
            Throw 'Operation cancelled by user.'
        }
        Else
        {
            Throw 'Operation was interrupted unexpectedly.'
        }

        $Form.Dispose()
    }
    End
    {
    }
}