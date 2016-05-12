<#
.SYNOPSIS
   Opens a graphical text-input dialog box.
.DESCRIPTION
   Uses the System.Windows.Forms class to generate a custom Form object to display a GUI. The GUI allows input of a text string and returns that information to the pipeline. The user cannot continue if the text box is blank.
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
        # Sets the title text displayed at the top bar of the dialog box.
        [Parameter(ValueFromPipelineByPropertyName=$True,
                   Position=1)]
        [ValidateNotNull()]
        [String]
        $Title = 'Input Box',

        # Sets the message displayed in the dialog box.
        [Parameter(ValueFromPipelineByPropertyName=$True,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('Description','Message')] 
        [String]
        $Caption = 'Type something:',

         # Determines the shape and design of the window border and title bar.
        [Parameter()]
        [ValidateNotNull()]
        [System.Windows.Forms.FormBorderStyle]
        $BorderStyle = 'FixedDialog',

        # Adds a predefined value to the text box. The value can be edited or changed by the user after the dialog appears.
        [Parameter(Position=2)]
        [String]
        $DefaultText,

        # Enables the minimize button and menu option for the dialog. Only relevant when the borderStyle would normally allow the ability to minimize. Disabled by default.
        [Switch]
        $EnableMinimize,

        # Enables the maximize button and menu option for the dialog. Only relevant when the borderStyle would normally allow the ability to maximize. Disabled by default.
        [Switch]
        $EnableMaximize
    )

    Begin
    {
        Add-Type -AssemblyName 'System.Drawing'
        Add-Type -AssemblyName 'System.Windows.Forms'
    }
    Process
    {
        $Form = New-Object System.Windows.Forms.Form 
        $Form.ClientSize = '300,150'
        $Form.MinimumSize = '250,150'
        $Form.StartPosition = "CenterScreen"
        $Form.Text = $Title
        $Form.FormBorderStyle = $BorderStyle
        $Form.ControlBox = $True
        $Form.MaximizeBox = $EnableMaximize
        $Form.MinimizeBox = $EnableMinimize
        $Form.Topmost = $True
        $Form.Add_Shown({ $Form.Activate() })


        $Label = New-Object System.Windows.Forms.Label
        $Label.Location = '10,25'
        $Label.Size = '280,20'
        $Label.Anchor = 'Left'
        $Label.Margin = [System.Windows.Forms.Padding]::new(10,0,0,0)
        $Label.Text = $Caption
        $Form.Controls.Add($Label)


        $TextBox = New-Object System.Windows.Forms.TextBox 
        $TextBox.AutoSize = $True
        $TextBox.Location = '10,60'
        $TextBox.Size = '280,20'
        $TextBox.Margin = '10,0,10,0'
        $TextBox.Anchor = 'Left, Right'
        $TextBox.TabStop = $True
        $TextBox.TabIndex = 0
        $TextBox.Text = $DefaultText
        $TextBox.Add_TextChanged({ If($TextBox.TextLength -gt 0){$OKButton.Enabled = $True} Else{$OKButton.Enabled = $False} })
        $Form.Controls.Add($TextBox) 


        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '65,110'
        $OKButton.Size = '75,23'
        $OKButton.Margin = '10,0,10,0'
        $OKButton.Anchor = 'Bottom'
        $OKButton.DialogResult = 'OK'
        $OKButton.Text = 'OK'
        $OKButton.Enabled = $DefaultText
        $OKButton.TabStop = $True
        $OKButton.TabIndex = 1
        $Form.AcceptButton = $OKButton
        $Form.Controls.Add($OKButton)


        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '160,110'
        $CancelButton.Size = '75,23'
        $CancelButton.Margin = '10,0,10,0'
        $CancelButton.Anchor = 'Bottom'
        $CancelButton.DialogResult = 'Cancel'
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