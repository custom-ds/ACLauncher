# Basic configurations to run the ACLauncher. These may need to be adjusted based on your system setup.

#The path to the ACLauncher configuration file
$scriptPath = if ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    # For compiled executables, use the current directory
    $PWD.Path
}
$configPath = Join-Path $scriptPath "ACLauncher.config.json"


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

# Get computer names from configuration
$computerNames = @()
if ($globalConfig.Computers) {
    $computerNames = $globalConfig.Computers | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
}

$user1Name = $userNames[0]
$user2Name = $userNames[1]
$user1Config = $globalConfig.Users.$user1Name
$user2Config = $globalConfig.Users.$user2Name

# Function to process file based on variables
function Process-ACLogFile {
    param(
        [PSCustomObject]$Config,
        [string]$Position,
        [string]$Computer = $null
    )
    
    # Define the path filenames
    $settingsPathFilename = $Config.General.SettingsPath + "Settings.xml"
    $colorsPathFilename = $Config.General.SettingsPath + "ColorAndFontSettings.xml"
    $rigsettingsPathFilename = $Config.General.SettingsPath + "RigSettings.xml"
    $phonesettingsPathFilename = $Config.General.SettingsPath + "PHSettings.xml" 
    $cwsettingsPathFilename = $Config.General.SettingsPath + "CWSettings.xml"
    $networkdisplaysettingsPathFilename = $Config.General.SettingsPath + "NetworkDisplaySettings.xml"


    # Write the LogPathFilename to HKCU\Software\Affirmatech\N3FJP Software\ACLog\LastLogFilePath
    # This avoids a warning about the log path changing since last use
    Set-ItemProperty -Path "HKCU:\Software\Affirmatech\ACLog" -Name "LastLogFilePath" -Value $Config.Users."$Position".Settings.LogPathFilename -Force

    #Write-Host "Updating settings for user: $($Config.Users."$Position".Settings.Name) ($($Config.Users."$Position".Settings.Callsign))"

    # Update the Settings.xml file
    $lines = Get-Content $settingsPathFilename

    #Update the general settings with AC Log
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
                #Can set this to true if you want QRZ to pull up with each callsign.
                "<QRZINTERNETENABLED>False</QRZINTERNETENABLED>"
            }
        }
        elseif ($_ -match '^<QRZAPIACCESSKEY>') {
            "<QRZAPIACCESSKEY>$($Config.Users."$Position".Settings.QRZAPIAccessKey)</QRZAPIACCESSKEY>"
        }
        elseif ($_ -match '^<SHOWSETTINGSFORM>') {
            "<SHOWSETTINGSFORM>$($Config.Users."$Position".Settings.ShowSettingsForm)</SHOWSETTINGSFORM>"
        }
        else {
            $_
        }
    }

    # Save the updated lines back to the file
    Set-Content -Path $settingsPathFilename -Value $updatedLines -Encoding UTF8



    # Update the ColorAndFontSettings.xml file to customize for the user
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



    # Update the NetworkDisplaySettings.xml with the computer name
    if (Test-Path $networkdisplaysettingsPathFilename) {
        $lines = Get-Content $networkdisplaysettingsPathFilename
        $updatedLines = $lines | ForEach-Object {
            if ($_ -match '^<THISCLIENTNAME>') {
                "<THISCLIENTNAME>$env:COMPUTERNAME</THISCLIENTNAME>"
            }
            else {
                $_
            }
        }
        Set-Content -Path $networkdisplaysettingsPathFilename -Value $updatedLines -Encoding UTF8
    }



    # Apply computer-specific settings if a computer is selected. These would be for COM port settings, and rig control
    if (-not [string]::IsNullOrEmpty($Computer)) {
        $computerConfig = $Config.Computers.$Computer
        if ($computerConfig) {
            # Computer-specific settings can be applied here
            # For example: COM port settings, rig settings, etc.
            # This is where you could update registry entries or config files
            # based on the selected computer's configuration
            
            # Update the RigSettings.xml file (only if it exists)
            if (Test-Path $rigsettingsPathFilename) {
                $lines = Get-Content $rigsettingsPathFilename

                $updatedLines = $lines | ForEach-Object {
                # Replace <COMPORTNAME>
                if ($_ -match '^<COMPORTNAME>') {
                    "<COMPORTNAME>$($computerConfig.RigSettings.ComPortName)</COMPORTNAME>"
                }
                # Replace <BAUDRATE>
                elseif ($_ -match '^<BAUDRATE>') {
                    "<BAUDRATE>$($computerConfig.RigSettings.BaudRate)</BAUDRATE>"
                }
                # Replace <PARITY>
                elseif ($_ -match '^<PARITY>') {
                    "<PARITY>$($computerConfig.RigSettings.Parity)</PARITY>"
                }
                # Replace <STOPBITS>
                elseif ($_ -match '^<STOPBITS>') {
                    "<STOPBITS>$($computerConfig.RigSettings.StopBits)</STOPBITS>"
                }
                # Replace <DATABITS>
                elseif ($_ -match '^<DATABITS>') {
                    "<DATABITS>$($computerConfig.RigSettings.DataBits)</DATABITS>"
                }
                # Replace <POWEROPTIONS>
                elseif ($_ -match '^<POWEROPTIONS>') {
                    "<POWEROPTIONS>$($computerConfig.RigSettings.PowerOptions)</POWEROPTIONS>"
                }
                # Replace <RIGNAME>
                elseif ($_ -match '^<RIGNAME>') {
                    "<RIGNAME>$($computerConfig.RigSettings.RigName)</RIGNAME>"
                }
                # Replace <POLLINGRATE>
                elseif ($_ -match '^<POLLINGRATE>') {
                    "<POLLINGRATE>$($computerConfig.RigSettings.PollingRate)</POLLINGRATE>"
                }
                # Replace <READFREQUENCYCOMMAND>
                elseif ($_ -match '^<READFREQUENCYCOMMAND>') {
                    "<READFREQUENCYCOMMAND>$($computerConfig.RigSettings.ReadFrequencyCommand)</READFREQUENCYCOMMAND>"
                }
                # Replace <READMODECOMMAND>
                elseif ($_ -match '^<READMODECOMMAND>') {
                    "<READMODECOMMAND>$($computerConfig.RigSettings.ReadModeCommand)</READMODECOMMAND>"
                }
                # Replace <CONVERTCOMMANDTOHEX>
                elseif ($_ -match '^<CONVERTCOMMANDTOHEX>') {
                    "<CONVERTCOMMANDTOHEX>$($computerConfig.RigSettings.ConvertCommandToHex)</CONVERTCOMMANDTOHEX>"
                }
                else {
                    $_
                }
                }

                # Save the updated lines back to the file
                Set-Content -Path $rigsettingsPathFilename -Value $updatedLines -Encoding UTF8
            }



            # Update the PhoneSettings.xml file (only if it exists)
            if (Test-Path $phonesettingsPathFilename) {
                $lines = Get-Content $phonesettingsPathFilename

                $updatedLines = $lines | ForEach-Object {
                # Replace <REPEATDELAY>
                if ($_ -match '^<REPEATDELAY>') {
                    "<REPEATDELAY>$($computerConfig.PhoneSettings.RepeatDelay)</REPEATDELAY>"
                }
                # Replace <F1>
                elseif ($_ -match '^<F1>') {
                    "<F1>$($computerConfig.PhoneSettings.F1)</F1>"
                }
                # Replace <F2>
                elseif ($_ -match '^<F2>') {
                    "<F2>$($computerConfig.PhoneSettings.F2)</F2>"
                }
                # Replace <F3>
                elseif ($_ -match '^<F3>') {
                    "<F3>$($computerConfig.PhoneSettings.F3)</F3>"
                }
                # Replace <F4>
                elseif ($_ -match '^<F4>') {
                    "<F4>$($computerConfig.PhoneSettings.F4)</F4>"
                }
                # Replace <F5>
                elseif ($_ -match '^<F5>') {
                    "<F5>$($computerConfig.PhoneSettings.F5)</F5>"
                }
                # Replace <F6>
                elseif ($_ -match '^<F6>') {
                    "<F6>$($computerConfig.PhoneSettings.F6)</F6>"
                }
                # Replace <F7>
                elseif ($_ -match '^<F7>') {
                    "<F7>$($computerConfig.PhoneSettings.F7)</F7>"
                }
                # Replace <F8>
                elseif ($_ -match '^<F8>') {
                    "<F8>$($computerConfig.PhoneSettings.F8)</F8>"
                }
                # Replace <F9>
                elseif ($_ -match '^<F9>') {
                    "<F9>$($computerConfig.PhoneSettings.F9)</F9>"
                }
                # Replace <F10>
                elseif ($_ -match '^<F10>') {
                    "<F10>$($computerConfig.PhoneSettings.F10)</F10>"
                }
                # Replace <F11>
                elseif ($_ -match '^<F11>') {
                    "<F11>$($computerConfig.PhoneSettings.F11)</F11>"
                }
                # Replace <F12>
                elseif ($_ -match '^<F12>') {
                    "<F12>$($computerConfig.PhoneSettings.F12)</F12>"
                }

                else {
                    $_
                }

                }

                # Save the updated lines back to the file
                Set-Content -Path $phonesettingsPathFilename -Value $updatedLines -Encoding UTF8
            }

            
            # Update the CWSettings.xml file (only if it exists)
            if (Test-Path $cwsettingsPathFilename) {
                $lines = Get-Content $cwsettingsPathFilename

            $updatedLines = $lines | ForEach-Object {
                # Replace <COMPORTNAME>
                if ($_ -match '^<COMPORTNAME>') {
                    "<COMPORTNAME>$($computerConfig.CWSettings.ComPortName)</COMPORTNAME>"
                }
                # Replace <KeyingOption>
                elseif ($_ -match '^<KEYINGOPTION>') {
                    "<KEYINGOPTION>$($computerConfig.CWSettings.KeyingOption)</KEYINGOPTION>"
                }
                # Replace <TimerOption>
                elseif ($_ -match '^<TIMEROPTION>') {
                    "<TIMEROPTION>$($computerConfig.CWSettings.TimerOption)</TIMEROPTION>"
                }
                # Replace <F1>
                elseif ($_ -match '^<F1>') {
                    "<F1>$($computerConfig.CWSettings.F1)</F1>"
                }
                # Replace <F2>
                elseif ($_ -match '^<F2>') {
                    "<F2>$($computerConfig.CWSettings.F2)</F2>"
                }
                # Replace <F3>
                elseif ($_ -match '^<F3>') {
                    "<F3>$($computerConfig.CWSettings.F3)</F3>"
                }
                # Replace <F4>
                elseif ($_ -match '^<F4>') {
                    "<F4>$($computerConfig.CWSettings.F4)</F4>"
                }
                # Replace <F5>
                elseif ($_ -match '^<F5>') {
                    "<F5>$($computerConfig.CWSettings.F5)</F5>"
                }
                # Replace <F6>
                elseif ($_ -match '^<F6>') {
                    "<F6>$($computerConfig.CWSettings.F6)</F6>"
                }
                # Replace <F7>
                elseif ($_ -match '^<F7>') {
                    "<F7>$($computerConfig.CWSettings.F7)</F7>"
                }
                # Replace <F8>
                elseif ($_ -match '^<F8>') {
                    "<F8>$($computerConfig.CWSettings.F8)</F8>"
                }
                # Replace <F9>
                elseif ($_ -match '^<F9>') {
                    "<F9>$($computerConfig.CWSettings.F9)</F9>"
                }
                # Replace <F10>
                elseif ($_ -match '^<F10>') {
                    "<F10>$($computerConfig.CWSettings.F10)</F10>"
                }
                # Replace <F11>
                elseif ($_ -match '^<F11>') {
                    "<F11>$($computerConfig.CWSettings.F11)</F11>"
                }
                # Replace <F12>
                elseif ($_ -match '^<F12>') {
                    "<F12>$($computerConfig.CWSettings.F12)</F12>"
                }
                
                else {
                    $_
                }

                }

                # Save the updated lines back to the file
                Set-Content -Path $cwsettingsPathFilename -Value $updatedLines -Encoding UTF8
            }



            #Write-Host "Using computer configuration: $($computerConfig.Name)"
        }
    }

    # Launch the program
    Start-Process $Config.General.ACLogPath

}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "ACLog Launcher"
$form.Size = New-Object System.Drawing.Size(365,290)
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

# Create computer label
$labelComputer = New-Object System.Windows.Forms.Label
$labelComputer.Location = New-Object System.Drawing.Point(20,155)
$labelComputer.Size = New-Object System.Drawing.Size(100,20)
$labelComputer.Text = "Computer:"
$labelComputer.Font = New-Object System.Drawing.Font("Arial", 9)
$form.Controls.Add($labelComputer)

# Create computer dropdown
$comboComputer = New-Object System.Windows.Forms.ComboBox
$comboComputer.Location = New-Object System.Drawing.Point(120,152)
$comboComputer.Size = New-Object System.Drawing.Size(210,25)
$comboComputer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboComputer.Font = New-Object System.Drawing.Font("Arial", 9)

# Populate the dropdown with computer names and display names
if ($computerNames.Count -gt 0) {
    foreach ($computerName in $computerNames) {
        if ($globalConfig.Computers.$computerName -and $globalConfig.Computers.$computerName.Name) {
            $computerDisplayName = $globalConfig.Computers.$computerName.Name
            $null = $comboComputer.Items.Add("$computerDisplayName ($computerName)")
        }
    }
}

# Add a default "None" option if no computers are configured
if ($comboComputer.Items.Count -eq 0) {
    $null = $comboComputer.Items.Add("None configured")
}

# Auto-select computer based on current hostname
$currentHostname = $env:COMPUTERNAME
$selectedIndex = 0

if ($comboComputer.Items.Count -gt 0) {
    # Try to find a matching computer by hostname
    for ($i = 0; $i -lt $comboComputer.Items.Count; $i++) {
        $itemText = $comboComputer.Items[$i].ToString()
        
        # Skip "None configured" option
        if ($itemText -eq "None configured") {
            continue
        }
        
        # Extract computer key from display text (e.g., "ZACKTOP (First)" -> "First")
        if ($itemText -match '\(([^)]+)\)$') {
            $computerKey = $matches[1]
            $computerConfig = $globalConfig.Computers.$computerKey
            
            # Check if the computer name matches current hostname (case-insensitive)
            if ($computerConfig -and $computerConfig.Name -and 
                ($computerConfig.Name.ToUpper() -eq $currentHostname.ToUpper())) {
                $selectedIndex = $i
                break
            }
        }
    }
    
    $comboComputer.SelectedIndex = $selectedIndex
}

$form.Controls.Add($comboComputer)

# Create status label to show auto-selection info
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(20,180)
$labelStatus.Size = New-Object System.Drawing.Size(310,15)
$labelStatus.Font = New-Object System.Drawing.Font("Arial", 8)
$labelStatus.ForeColor = [System.Drawing.Color]::DarkGray

# Set status message based on selection
if ($comboComputer.SelectedItem -and $comboComputer.SelectedItem.ToString() -ne "None configured") {
    $selectedText = $comboComputer.SelectedItem.ToString()
    if ($selectedText -match '\(([^)]+)\)$') {
        $computerKey = $matches[1]
        $computerConfig = $globalConfig.Computers.$computerKey
        if ($computerConfig -and $computerConfig.Name -and 
            ($computerConfig.Name.ToUpper() -eq $currentHostname.ToUpper())) {
            $labelStatus.Text = "Auto-selected based on hostname: $currentHostname"
            $labelStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        } else {
            $labelStatus.Text = "Current hostname: $currentHostname (no match found)"
        }
    }
} else {
    $labelStatus.Text = "Current hostname: $currentHostname (no computers configured)"
}

$labelStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($labelStatus)

# Create version label
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Location = New-Object System.Drawing.Point(20,220)
$labelVersion.Size = New-Object System.Drawing.Size(310,15)
$labelVersion.Text = "Version 1.2"
$labelVersion.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$labelVersion.Font = New-Object System.Drawing.Font("Arial", 8)
$labelVersion.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($labelVersion)

# Function to get selected computer configuration key
function Get-SelectedComputer {
    if ($comboComputer.SelectedItem -eq $null) {
        return $null
    }
    
    $selectedText = $comboComputer.SelectedItem.ToString()
    
    # Return null if "None configured" is selected
    if ($selectedText -eq "None configured") {
        return $null
    }
    
    # Extract the computer key from the display text (e.g., "ZACKTOP (First)" -> "First")
    if ($selectedText -match '\(([^)]+)\)$') {
        return $matches[1]
    }
    return $null
}

# First button click event
$buttonOne.Add_Click({
    $config = Load-Configuration -ConfigPath $configPath
    $selectedComputer = Get-SelectedComputer
    
    # Call the processing function
    Process-ACLogFile -Config $config -Position "First" -Computer $selectedComputer
    
    # Close the form and set dialog result
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# Second button click event
$buttonTwo.Add_Click({
    $config = Load-Configuration -ConfigPath $configPath
    $selectedComputer = Get-SelectedComputer
    
    # Call the processing function
    Process-ACLogFile -Config $config -Position "Second" -Computer $selectedComputer
    
    # Close the form and set dialog result
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# Show the form
$result = $form.ShowDialog()

# Exit the script after the form is closed
exit
