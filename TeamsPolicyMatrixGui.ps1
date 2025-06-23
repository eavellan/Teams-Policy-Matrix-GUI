# TEAMS POLICY MATRIX GUI VERSION
# Author: Edgar Avellan

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Teams Policy Matrix GUI by Edgar Avellan"
$form.Size = New-Object System.Drawing.Size(1000, 800)
$form.StartPosition = "CenterScreen"

# Buttons
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect to Teams"
$connectButton.Location = New-Object System.Drawing.Point(20, 20)
$connectButton.Size = New-Object System.Drawing.Size(150, 30)

$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Text = "Load All Policies"
$loadButton.Location = New-Object System.Drawing.Point(190, 20)
$loadButton.Size = New-Object System.Drawing.Size(150, 30)

$viewButton = New-Object System.Windows.Forms.Button
$viewButton.Text = "View Selected"
$viewButton.Location = New-Object System.Drawing.Point(360, 20)
$viewButton.Size = New-Object System.Drawing.Size(150, 30)

$compareButton = New-Object System.Windows.Forms.Button
$compareButton.Text = "Compare Selected"
$compareButton.Location = New-Object System.Drawing.Point(530, 20)
$compareButton.Size = New-Object System.Drawing.Size(150, 30)

# Listbox
$policyListBox = New-Object System.Windows.Forms.ListBox
$policyListBox.Location = New-Object System.Drawing.Point(20, 70)
$policyListBox.Size = New-Object System.Drawing.Size(450, 640)
$policyListBox.SelectionMode = "MultiExtended"

# Log output
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(490, 70)
$logBox.Size = New-Object System.Drawing.Size(470, 640)
$logBox.ReadOnly = $true
$logBox.BackColor = 'Black'
$logBox.ForeColor = 'Lime'
$logBox.Font = 'Courier New,9pt'

# Status bar
$statusBar = New-Object System.Windows.Forms.Label
$statusBar.Text = "Status: Not connected."
$statusBar.Location = New-Object System.Drawing.Point(20, 730)
$statusBar.Size = New-Object System.Drawing.Size(950, 30)

$form.Controls.AddRange(@(
    $connectButton, $loadButton, $viewButton, $compareButton,
    $policyListBox, $logBox, $statusBar
))

# Globals
$global:TeamsConnected = $false
$global:scriptContent = ""
$global:PolicyData = @{}

function Write-Log($text) {
    $logBox.AppendText("$text`r`n")
    $logBox.SelectionStart = $logBox.Text.Length
    $logBox.ScrollToCaret()
}

# Connect to Teams
$connectButton.Add_Click({
    try {
        Write-Log "🔐 Connecting to Microsoft Teams..."
        $statusBar.Text = "Status: Connecting..."
        Import-Module MicrosoftTeams -ErrorAction Stop
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        $global:TeamsConnected = $true
        Write-Log "✅ Connected successfully to Teams."
        $statusBar.Text = "Status: Connected."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Connection failed: $_")
        Write-Log "❌ Connection failed: $_"
        $statusBar.Text = "Status: Connection failed."
    }
})

# Load policies
$loadButton.Add_Click({
    if (-not $global:TeamsConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Teams first.")
        return
    }

    $policyTypes = @(
        "TeamsAppPermissionPolicy", "TeamsAppSetupPolicy", "TeamsAudioConferencingPolicy", "TeamsCallHoldPolicy",
        "TeamsCallingPolicy", "TeamsCallParkPolicy", "TeamsCallQueue", "TeamsChannel", "TeamsChannelsPolicy",
        "TeamsClientConfiguration", "TeamsComplianceRecordingPolicy", "TeamsCortanaPolicy",
        "TeamsDialInConferencingTenantSettings", "TeamsEmergencyCallingPolicy", "TeamsEmergencyCallRoutingPolicy",
        "TeamsEnhancedEncryptionPolicy", "TeamsEventsPolicy", "TeamsFederationConfiguration", "TeamsFeedbackPolicy",
        "TeamsFilesPolicy", "TeamsGroupPolicyAssignment", "TeamsGuestCallingConfiguration",
        "TeamsGuestMeetingConfiguration", "TeamsGuestMessagingConfiguration", "TeamsIPPhonePolicy",
        "TeamsMeetingBroadcastConfiguration", "TeamsMeetingBroadcastPolicy", "TeamsMeetingConfiguration",
        "TeamsMeetingPolicy", "TeamsMessagingPolicy", "TeamsMobilityPolicy", "TeamsNetworkRoamingPolicy",
        "TeamsOnlineVoicemailPolicy", "TeamsOnlineVoicemailUserSettings", "TeamsOnlineVoiceUser",
        "TeamsOrgWideAppSettings", "TeamsPstnUsage", "TeamsShiftsPolicy", "TeamsTemplatesPolicy",
        "TeamsTenantDialPlan", "TeamsTenantNetworkRegion", "TeamsTenantNetworkSite", "TeamsTenantNetworkSubnet",
        "TeamsTenantTrustedIPAddress", "TeamsTranslationRule", "TeamsUnassignedNumberTreatment",
        "TeamsUpdateManagementPolicy", "TeamsUpgradeConfiguration", "TeamsUpgradePolicy", "TeamsUserCallingSettings",
        "TeamsUserPolicyAssignment", "TeamsVdiPolicy", "TeamsVoiceRoute", "TeamsVoiceRoutingPolicy", "TeamsWorkloadPolicy"
    )

    $allPolicies = ""
    $policyListBox.Items.Clear()
    $loaded = @()
    $failed = @()
    $global:PolicyData.Clear()

    Write-Log "🔄 Pulling policies from tenant..."

    foreach ($type in $policyTypes) {
        try {
            $cmdlet = "Get-Cs$type"
            if (Get-Command $cmdlet -ErrorAction SilentlyContinue) {
                $items = Invoke-Expression $cmdlet
                foreach ($item in $items) {
                    $identity = $item.Identity
                    $label = "$type [$identity]"
                    $policyListBox.Items.Add($label)
                    $global:PolicyData[$label] = $item
                    $block = '{0} "{1}" {{`n' -f $type, $identity
                    foreach ($prop in $item.PSObject.Properties) {
                        if ($prop.Name -notmatch "Identity|ObjectId") {
                            $value = $prop.Value -replace "`r`n", " " -replace "`n", " "
                            $block += "    $($prop.Name) = $value`n"
                        }
                    }
                    $block += "}`n`n"
                    $allPolicies += $block
                }
                $loaded += $type
                Write-Log "✅ Loaded $type"
            } else {
                $failed += $type
                Write-Log "⚠️ Cmdlet $cmdlet not available in session."
            }
        } catch {
            $failed += $type
            Write-Log "⚠️ Failed to load ${type}: $_"
        }
    }

    $global:scriptContent = $allPolicies
    $statusBar.Text = "✅ Policies loaded."

    # Summary Report
    Write-Log "`n📊 Summary Report"
    Write-Log "✅ Successfully Loaded: $($loaded.Count)"
    $loaded | ForEach-Object { Write-Log "   - $_" }

    Write-Log "⚠️ Failed to Load: $($failed.Count)"
    $failed | ForEach-Object { Write-Log "   - $_" }

    $totalExpected = $policyTypes.Count
    $totalCaptured = $loaded.Count + $failed.Count
    Write-Log "`n🔢 Total Policies Expected: $totalExpected"
    Write-Log "📈 Total Captured: $totalCaptured"
})

# View selected policies
$viewButton.Add_Click({
    $selected = $policyListBox.SelectedItems
    if ($selected.Count -eq 0 -or $selected.Count -gt 2) {
        [System.Windows.Forms.MessageBox]::Show("Select one or two policies to view side-by-side.")
        return
    }

    $viewerForm = New-Object System.Windows.Forms.Form
    $viewerForm.Text = "Policy Viewer"
    $viewerForm.Size = New-Object System.Drawing.Size(900, 600)

    $leftBox = New-Object System.Windows.Forms.RichTextBox
    $leftBox.Location = New-Object System.Drawing.Point(10, 10)
    $leftBox.Size = New-Object System.Drawing.Size(420, 540)
    $leftBox.ReadOnly = $true
    $leftBox.Font = 'Courier New,9pt'

    $rightBox = New-Object System.Windows.Forms.RichTextBox
    $rightBox.Location = New-Object System.Drawing.Point(450, 10)
    $rightBox.Size = New-Object System.Drawing.Size(420, 540)
    $rightBox.ReadOnly = $true
    $rightBox.Font = 'Courier New,9pt'

    if ($selected.Count -ge 1) {
        $policy1 = $global:PolicyData[$selected[0]]
        $text = ""
        foreach ($prop in $policy1.PSObject.Properties) {
            $text += "{0} = {1}`n" -f $prop.Name, $prop.Value
        }
        $leftBox.Text = $text
    }
    if ($selected.Count -eq 2) {
        $policy2 = $global:PolicyData[$selected[1]]
        $text = ""
        foreach ($prop in $policy2.PSObject.Properties) {
            $text += "{0} = {1}`n" -f $prop.Name, $prop.Value
        }
        $rightBox.Text = $text
    }

    $viewerForm.Controls.AddRange(@($leftBox, $rightBox))
    $viewerForm.ShowDialog()
})

# Compare selected policies
$compareButton.Add_Click({
    $selected = $policyListBox.SelectedItems
    if ($selected.Count -ne 2) {
        [System.Windows.Forms.MessageBox]::Show("Please select exactly TWO policies to compare.")
        return
    }

    $policy1 = $global:PolicyData[$selected[0]]
    $policy2 = $global:PolicyData[$selected[1]]

    if (-not $policy1 -or -not $policy2) {
        Write-Log "❌ Unable to retrieve selected policy data."
        return
    }

    Write-Log "`n🔍 Comparing:`n - $($selected[0])`n - $($selected[1])"

    foreach ($prop in $policy1.PSObject.Properties) {
        if ($prop.Name -in $policy2.PSObject.Properties.Name) {
            $value1 = $policy1.$($prop.Name)
            $value2 = $policy2.$($prop.Name)
            if ($value1 -ne $value2) {
                Write-Log "❗ $($prop.Name): $value1 ⛔ $value2"
            } else {
                Write-Log "✅ $($prop.Name): $value1"
            }
        }
    }
})

# Start the GUI
$form.ShowDialog()
