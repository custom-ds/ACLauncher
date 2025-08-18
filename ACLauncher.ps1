# Basic configurations to run the ACLauncher. These may need to be adjusted based on your system setup.

#The path to the ACLauncher configuration file
$configPath = "$env:USERPROFILE\OneDrive\Ham Radio\ACLauncher\ACLauncher.config.json"


Add-Type -AssemblyName System.Windows.Forms



# Function to load configuration from JSON file
function Load-Configuration {
    param(
        [string]$ConfigPath
    )
    
    if (-not (Test-Path $ConfigPath)) {
        [System.Windows.Forms.MessageBox]::Show("Configuration file not found: $ConfigPath", "Configuration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
    
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        
        # Expand environment variables in the log path
        $Config.General.SettingsPath = [System.Environment]::ExpandEnvironmentVariables($Config.General.SettingsPath)
        $Config.Users.First.Settings.LogPathFilename = [System.Environment]::ExpandEnvironmentVariables($Config.Users.First.Settings.LogPathFilename)
        $Config.Users.Second.Settings.LogPathFilename = [System.Environment]::ExpandEnvironmentVariables($Config.Users.Second.Settings.LogPathFilename)

        return $Config
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error reading configuration file: $_", "Configuration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Check if ACLog is already running before showing the form
$processName = "ACLog"
$process = Get-Process -Name $processName -ErrorAction SilentlyContinue
if ($process) {
    [System.Windows.Forms.MessageBox]::Show("ACLog is already running. Please close it first before launching a new instance.", "ACLog Already Running", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit
}


if (-not (Test-Path $configPath)) {
    [System.Windows.Forms.MessageBox]::Show("Configuration file not found!`n`nThe script is looking for:`n$configPath`n`nPlease create this file using the template 'ACLauncher.config.template.json' as a reference.", "Configuration File Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit
}

try {
    $globalConfig = Get-Content $configPath -Raw | ConvertFrom-Json
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error reading configuration file: $_`n`nPath: $configPath", "Configuration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Get user names from configuration
$userNames = $globalConfig.Users | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
if ($userNames.Count -lt 2) {
    [System.Windows.Forms.MessageBox]::Show("Configuration file must contain at least 2 users.`n`nFound: $($userNames.Count) user(s)", "Configuration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

$user1Name = $userNames[0]
$user2Name = $userNames[1]
$user1Config = $globalConfig.Users.$user1Name
$user2Config = $globalConfig.Users.$user2Name

# Function to process file based on variables
function Process-ACLogFile {
    param(
        [PSCustomObject]$Config,
        [string]$Position
    )
    
    # Define the path filenames
    $settingsPathFilename = $Config.General.SettingsPath + "Settings.xml"
    $colorsPathFilename = $Config.General.SettingsPath + "ColorAndFontSettings.xml"


    # Write the LogPathFilename to HKCU\Software\Affirmatech\N3FJP Software\ACLog\LastLogFilePath
    # This avoids a warning about the log path changing since last use
    Set-ItemProperty -Path "HKCU:\Software\Affirmatech\ACLog" -Name "LastLogFilePath" -Value $Config.Users."$Position".Settings.LogPathFilename -Force

    #Write-Host "Updating settings for user: $($Config.Users."$Position".Settings.Name) ($($Config.Users."$Position".Settings.Callsign))"

    # Update the Settings.xml file
    $lines = Get-Content $settingsPathFilename

    $updatedLines = $lines | ForEach-Object {
        # Replace <USERCALL>
        if ($_ -match '^<USERCALL>') {
            "<USERCALL>$($Config.Users."$Position".Settings.Callsign)</USERCALL>"
        } 

        # Replace REGISTRATIONCALL
        elseif ($_ -match '^<REGISTRATIONCALL>') {
            "<REGISTRATIONCALL>$($Config.Users."$Position".Settings.Callsign)</REGISTRATIONCALL>"
        }
        # Replace <USERPATHFILE>
        elseif ($_ -match '^<USERPATHFILE>') {
            "<USERPATHFILE>$($Config.Users."$Position".Settings.LogPathFilename)</USERPATHFILE>"
        }
        # Replace <USERINITIALS>
        elseif ($_ -match '^<USERINITIALS>') {
            "<USERINITIALS>$($Config.Users."$Position".Settings.Initials)</USERINITIALS>"
        }
        # Replace <USEROPERATOR>
        elseif ($_ -match '^<USEROPERATOR>') {
            "<USEROPERATOR>$($Config.Users."$Position".Settings.Callsign)</USEROPERATOR>"
        }
        # Replace <LOTWUSERNAME>
        elseif ($_ -match '^<LOTWUSERNAME>') {
            "<LOTWUSERNAME>$($Config.Users."$Position".Settings.Callsign)</LOTWUSERNAME>"
        }
        # Replace <LOTWPASSWORD>
        elseif ($_ -match '^<LOTWPASSWORD>') {
            "<LOTWPASSWORD>$($Config.Users."$Position".Settings.LOTWPassword)</LOTWPASSWORD>"
        }
        # Replace <QRZINTERNETUSERNAME> and <QRZINTERNETPASSWORD>
        elseif ($_ -match '^<QRZINTERNETUSERNAME>') {
            "<QRZINTERNETUSERNAME>$($Config.Users."$Position".Settings.Callsign)</QRZINTERNETUSERNAME>"
        }
        elseif ($_ -match '^<QRZINTERNETPASSWORD>') {
            "<QRZINTERNETPASSWORD>$($Config.Users."$Position".Settings.QRZInternetPassword)</QRZINTERNETPASSWORD>"
        }
        elseif ($_ -match '^<QRZINTERNETENABLED>') {
            #If there is no QRZInternetPassword, disable the QRZ Internet feature
            if ([string]::IsNullOrEmpty($Config.Users."$Position".Settings.QRZInternetPassword)) {
                "<QRZINTERNETENABLED>False</QRZINTERNETENABLED>"
            } else {
                "<QRZINTERNETENABLED>True</QRZINTERNETENABLED>"
            }
        }
        elseif ($_ -match '^<QRZAPIACCESSKEY>') {
            "<QRZAPIACCESSKEY>$($Config.Users."$Position".Settings.QRZAPIAccessKey)</QRZAPIACCESSKEY>"
        }
        else {
            $_
        }
    }

    # Save the updated lines back to the file
    Set-Content -Path $settingsPathFilename -Value $updatedLines -Encoding UTF8



    # Update the ColorAndFontSettings.xml file
    $lines = Get-Content $colorsPathFilename
    $updatedLines = $lines | ForEach-Object {
        # Update Fonts
        if ($_ -match '^<MAINLISTFONT>') {
            '<MAINLISTFONT>Arial, 12.0pt</MAINLISTFONT>'
        } 
        elseif ($_ -match '^<DXLISTFONT>') {
            '<DXLISTFONT>Arial, 14.0pt</DXLISTFONT>'
        }
        elseif ($_ -match '^<FORMFONT>') {
            '<FORMFONT>Arial, 15.0pt</FORMFONT>'
        }
        elseif ($_ -match '^<FORMBEARINGDISTCONTFONT>') {
            '<FORMBEARINGDISTCONTFONT>Arial, 15.0pt</FORMBEARINGDISTCONTFONT>'
        }
        elseif ($_ -match '^<FORMMENUFONT>') {
            '<FORMMENUFONT>Arial, 12.0pt</FORMMENUFONT>'
        }
        elseif ($_ -match '^<LABELFONT>') {
            '<LABELFONT>Arial, 18.0pt, style=Bold</LABELFONT>'
        }
        elseif ($_ -match '^<TEXTBOXFONT>') {
            '<TEXTBOXFONT>Arial, 17.0pt, style=Bold</TEXTBOXFONT>'
        }


        #Update Background Colors
        elseif ($_ -match '^<LABELBACKCOLOR>') {
            "<LABELBACKCOLOR>$($Config.Users."$Position".Colors.FormBackgroundColor)</LABELBACKCOLOR>"
        }
        elseif ($_ -match '^<LABELOTHERBACKCOLOR>') {
            "<LABELOTHERBACKCOLOR>$($Config.Users."$Position".Colors.FormBackgroundColor)</LABELOTHERBACKCOLOR>"
        }

        else {
            $_
        }
    }

    # Save the updated lines back to the file
    Set-Content -Path $colorsPathFilename -Value $updatedLines -Encoding UTF8

    # Launch the program
    Start-Process $Config.General.ACLogPath

}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "ACLog Launcher"
$form.Size = New-Object System.Drawing.Size(365,225)
$form.StartPosition = "CenterScreen"
$form.MinimizeBox = $false
$form.MaximizeBox = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Create intro text label
$labelIntro = New-Object System.Windows.Forms.Label
$labelIntro.Location = New-Object System.Drawing.Point(20,15)
$labelIntro.Size = New-Object System.Drawing.Size(310,30)
$labelIntro.Text = "ACLog Ham Radio Logging Software Launcher"
$labelIntro.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$labelIntro.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelIntro)

# Create First button
$buttonOne = New-Object System.Windows.Forms.Button
$buttonOne.Location = New-Object System.Drawing.Point(20,60)
$buttonOne.Size = New-Object System.Drawing.Size(310,35)
$buttonOne.Text = "$($user1Config.Settings.Name) ($($user1Config.Settings.Callsign))"
$user1Colors = $user1Config.Colors.FormBackgroundColor -split ','
$buttonOne.BackColor = [System.Drawing.Color]::FromArgb([int]$user1Colors[0], [int]$user1Colors[1], [int]$user1Colors[2])
$buttonOne.ForeColor = [System.Drawing.Color]::White
$buttonOne.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonOne)

# Create Second button
$buttonTwo = New-Object System.Windows.Forms.Button
$buttonTwo.Location = New-Object System.Drawing.Point(20,105)
$buttonTwo.Size = New-Object System.Drawing.Size(310,35)
$buttonTwo.Text = "$($user2Config.Settings.Name) ($($user2Config.Settings.Callsign))"
$user2Colors = $user2Config.Colors.FormBackgroundColor -split ','
$buttonTwo.BackColor = [System.Drawing.Color]::FromArgb([int]$user2Colors[0], [int]$user2Colors[1], [int]$user2Colors[2])
$buttonTwo.ForeColor = [System.Drawing.Color]::White
$buttonTwo.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonTwo)

# Create version label
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Location = New-Object System.Drawing.Point(20,155)
$labelVersion.Size = New-Object System.Drawing.Size(310,15)
$labelVersion.Text = "Version 1.2"
$labelVersion.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$labelVersion.Font = New-Object System.Drawing.Font("Arial", 8)
$labelVersion.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($labelVersion)

# First button click event
$buttonOne.Add_Click({
    $config = Load-Configuration -ConfigPath $configPath
    
    # Call the processing function
    Process-ACLogFile -Config $config -Position "First"
    
    # Close the form and set dialog result
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# Second button click event
$buttonTwo.Add_Click({
    $config = Load-Configuration -ConfigPath $configPath
    
    # Call the processing function
    Process-ACLogFile -Config $config -Position "Second"
    
    # Close the form and set dialog result
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# Show the form
$result = $form.ShowDialog()

# Exit the script after the form is closed
exit
