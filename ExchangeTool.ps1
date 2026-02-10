#requires -Version 5.1
#requires -Modules ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Config = [ordered]@{
    ExchangeUri  = "https://your-exchange-server/powershell/"
    DefaultOU    = "OU=Users,DC=contoso,DC=com"
    DefaultDomain = "contoso.com"
    RemoteRoutingDomain = "tenant.mail.onmicrosoft.com"
}

$script:ExchSession = $null

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logBox.AppendText("[$timestamp] $Message`r`n")
}

function Show-Error {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, "Exchange Tool", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Prompt-Text {
    param(
        [string]$Title,
        [string]$Label,
        [string]$Default = "",
        [switch]$Password
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterParent"
    $form.Size = New-Object System.Drawing.Size(420, 160)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Label
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(12, 15)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Default
    $textBox.Size = New-Object System.Drawing.Size(380, 20)
    $textBox.Location = New-Object System.Drawing.Point(12, 40)
    $textBox.UseSystemPasswordChar = $Password.IsPresent

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "OK"
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ok.Location = New-Object System.Drawing.Point(236, 75)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancel.Location = New-Object System.Drawing.Point(317, 75)

    $form.Controls.AddRange(@($label, $textBox, $ok, $cancel))
    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }

    return $null
}

function Prompt-YesNo {
    param(
        [string]$Title,
        [string]$Message
    )

    $result = [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

function Select-Group {
    param([string]$InitialQuery = "")

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Group"
    $form.StartPosition = "CenterParent"
    $form.Size = New-Object System.Drawing.Size(520, 360)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search group (name, sam):"
    $searchLabel.AutoSize = $true
    $searchLabel.Location = New-Object System.Drawing.Point(12, 12)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(12, 34)
    $searchBox.Size = New-Object System.Drawing.Size(360, 20)
    $searchBox.Text = $InitialQuery

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Location = New-Object System.Drawing.Point(380, 32)

    $groupList = New-Object System.Windows.Forms.ListBox
    $groupList.Location = New-Object System.Drawing.Point(12, 68)
    $groupList.Size = New-Object System.Drawing.Size(480, 200)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "OK"
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ok.Location = New-Object System.Drawing.Point(316, 280)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancel.Location = New-Object System.Drawing.Point(397, 280)

    $form.Controls.AddRange(@($searchLabel, $searchBox, $searchButton, $groupList, $ok, $cancel))
    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    $searchHandler = {
        $query = $searchBox.Text
        if ([string]::IsNullOrWhiteSpace($query)) {
            return
        }

        try {
            $filter = "Name -like '*$query*' -or SamAccountName -like '*$query*'"
            $groups = Get-ADGroup -Filter $filter
            $groupList.Items.Clear()
            foreach ($g in $groups) {
                $item = [pscustomobject]@{
                    Name = $g.Name
                    SamAccountName = $g.SamAccountName
                    DistinguishedName = $g.DistinguishedName
                    Display = "{0} ({1})" -f $g.Name, $g.SamAccountName
                }
                [void]$groupList.Items.Add($item)
            }
            $groupList.DisplayMember = "Display"
        } catch {
            Show-Error "Group search failed: $($_.Exception.Message)"
        }
    }

    $searchButton.Add_Click($searchHandler)
    $searchBox.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { & $searchHandler } })

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $groupList.SelectedItem
    }

    return $null
}

function Ensure-ExchangeSession {
    if ($script:ExchSession -and $script:ExchSession.State -eq "Opened") {
        return $true
    }

    try {
        $cred = Get-Credential -Message "Enter Exchange on-prem credentials"
        $script:ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Config.ExchangeUri -Authentication Kerberos -Credential $cred
        Import-PSSession $script:ExchSession -DisableNameChecking | Out-Null
        Write-Log "Connected to Exchange on-prem."
        return $true
    } catch {
        Show-Error "Failed to connect to Exchange: $($_.Exception.Message)"
        return $false
    }
}

function Get-SelectedUser {
    if (-not $resultsList.SelectedItem) {
        Show-Error "Select a user first."
        return $null
    }
    return $resultsList.SelectedItem
}

function Get-SelectedGroup {
    if (-not $groupResultsList.SelectedItem) {
        Show-Error "Select a group first."
        return $null
    }
    return $groupResultsList.SelectedItem
}

function Search-Users {
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) {
        Show-Error "Enter a name or username to search."
        return
    }

    try {
        $filter = "Name -like '*$Query*' -or SamAccountName -like '*$Query*' -or UserPrincipalName -like '*$Query*'"
        $users = Get-ADUser -Filter $filter -Properties mail,proxyAddresses,displayName,UserPrincipalName,SamAccountName
        $resultsList.Items.Clear()
        foreach ($u in $users) {
            $display = "{0} ({1})" -f $u.DisplayName, $u.SamAccountName
            $item = [pscustomobject]@{
                DisplayName = $u.DisplayName
                SamAccountName = $u.SamAccountName
                UserPrincipalName = $u.UserPrincipalName
                Mail = $u.Mail
                Display = $display
            }
            [void]$resultsList.Items.Add($item)
        }
        $resultsList.DisplayMember = "Display"
        Write-Log "Found $($users.Count) user(s) for '$Query'."
    } catch {
        Show-Error "Search failed: $($_.Exception.Message)"
    }
}

function Search-Groups {
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) {
        Show-Error "Enter a group name to search."
        return
    }

    try {
        $filter = "Name -like '*$Query*' -or SamAccountName -like '*$Query*'"
        $groups = Get-ADGroup -Filter $filter
        $groupResultsList.Items.Clear()
        foreach ($g in $groups) {
            $item = [pscustomobject]@{
                Name = $g.Name
                SamAccountName = $g.SamAccountName
                DistinguishedName = $g.DistinguishedName
                Display = "{0} ({1})" -f $g.Name, $g.SamAccountName
            }
            [void]$groupResultsList.Items.Add($item)
        }
        $groupResultsList.DisplayMember = "Display"
        Write-Log "Found $($groups.Count) group(s) for '$Query'."
    } catch {
        Show-Error "Group search failed: $($_.Exception.Message)"
    }
}

function Add-EmailAlias {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $alias = Prompt-Text -Title "Add Alias" -Label "Enter alias local part (e.g. jdoe):"
    if (-not $alias) { return }

    $domain = Prompt-Text -Title "Add Alias" -Label "Enter domain:" -Default $Config.DefaultDomain
    if (-not $domain) { return }

    $address = "smtp:$alias@$domain"

    try {
        Set-RemoteMailbox -Identity $user.UserPrincipalName -EmailAddresses @{Add = $address} | Out-Null
        Write-Log "Added alias $address to $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to add alias: $($_.Exception.Message)"
    }
}

function Remove-EmailAlias {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $alias = Prompt-Text -Title "Remove Alias" -Label "Enter alias or SMTP address to remove:" -Default ""
    if (-not $alias) { return }

    if ($alias -notmatch "^smtp:") {
        $alias = "smtp:$alias"
    }

    try {
        Set-RemoteMailbox -Identity $user.UserPrincipalName -EmailAddresses @{Remove = $alias} | Out-Null
        Write-Log "Removed alias $alias from $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to remove alias: $($_.Exception.Message)"
    }
}

function Set-PrimarySmtp {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $address = Prompt-Text -Title "Primary SMTP" -Label "Enter primary SMTP address:" -Default $user.Mail
    if (-not $address) { return }

    try {
        Set-RemoteMailbox -Identity $user.UserPrincipalName -PrimarySmtpAddress $address | Out-Null
        Write-Log "Set primary SMTP to $address for $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to set primary SMTP: $($_.Exception.Message)"
    }
}

function Hide-FromGal {
    param([bool]$Hide)

    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    try {
        Set-RemoteMailbox -Identity $user.UserPrincipalName -HiddenFromAddressListsEnabled:$Hide | Out-Null
        $action = if ($Hide) { "Hidden" } else { "Visible" }
        Write-Log "$action in GAL: $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to update GAL visibility: $($_.Exception.Message)"
    }
}

function Add-FullAccess {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $delegate = Prompt-Text -Title "Full Access" -Label "Enter delegate UPN:" -Default ""
    if (-not $delegate) { return }

    try {
        Add-MailboxPermission -Identity $user.UserPrincipalName -User $delegate -AccessRights FullAccess -InheritanceType All -AutoMapping:$true | Out-Null
        Write-Log "Granted FullAccess to $delegate on $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to grant FullAccess: $($_.Exception.Message)"
    }
}

function Remove-FullAccess {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $delegate = Prompt-Text -Title "Remove Full Access" -Label "Enter delegate UPN:" -Default ""
    if (-not $delegate) { return }

    try {
        Remove-MailboxPermission -Identity $user.UserPrincipalName -User $delegate -AccessRights FullAccess -Confirm:$false | Out-Null
        Write-Log "Removed FullAccess for $delegate on $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to remove FullAccess: $($_.Exception.Message)"
    }
}

function Add-SendAs {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $delegate = Prompt-Text -Title "Send As" -Label "Enter delegate UPN:" -Default ""
    if (-not $delegate) { return }

    try {
        Add-RecipientPermission -Identity $user.UserPrincipalName -Trustee $delegate -AccessRights SendAs -Confirm:$false | Out-Null
        Write-Log "Granted SendAs to $delegate on $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to grant SendAs: $($_.Exception.Message)"
    }
}

function Remove-SendAs {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $delegate = Prompt-Text -Title "Remove Send As" -Label "Enter delegate UPN:" -Default ""
    if (-not $delegate) { return }

    try {
        Remove-RecipientPermission -Identity $user.UserPrincipalName -Trustee $delegate -AccessRights SendAs -Confirm:$false | Out-Null
        Write-Log "Removed SendAs for $delegate on $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to remove SendAs: $($_.Exception.Message)"
    }
}

function Convert-ToSharedMailbox {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    try {
        Set-RemoteMailbox -Identity $user.UserPrincipalName -Type Shared | Out-Null
        Write-Log "Converted to shared mailbox: $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to convert mailbox: $($_.Exception.Message)"
    }
}

function Set-Forwarding {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $forwardTo = Prompt-Text -Title "Forwarding" -Label "Forward to SMTP address (leave blank to disable):" -Default ""

    try {
        if ([string]::IsNullOrWhiteSpace($forwardTo)) {
            Set-RemoteMailbox -Identity $user.UserPrincipalName -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward:$false | Out-Null
            Write-Log "Disabled forwarding for $($user.UserPrincipalName)."
            return
        }

        $keepCopy = Prompt-YesNo -Title "Forwarding" -Message "Keep a copy in the mailbox?"
        Set-RemoteMailbox -Identity $user.UserPrincipalName -ForwardingSmtpAddress $forwardTo -DeliverToMailboxAndForward:$keepCopy | Out-Null
        Write-Log "Forwarding set to $forwardTo (Keep copy: $keepCopy) for $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to set forwarding: $($_.Exception.Message)"
    }
}

function Disable-Forwarding {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    try {
        Set-RemoteMailbox -Identity $user.UserPrincipalName -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward:$false | Out-Null
        Write-Log "Disabled forwarding for $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to disable forwarding: $($_.Exception.Message)"
    }
}

function New-RemoteMailboxUser {
    if (-not (Ensure-ExchangeSession)) { return }

    $first = Prompt-Text -Title "New Mailbox" -Label "First name:" -Default ""
    if (-not $first) { return }

    $last = Prompt-Text -Title "New Mailbox" -Label "Last name:" -Default ""
    if (-not $last) { return }

    $sam = Prompt-Text -Title "New Mailbox" -Label "SamAccountName:" -Default ""
    if (-not $sam) { return }

    $upn = Prompt-Text -Title "New Mailbox" -Label "UserPrincipalName:" -Default "$sam@$($Config.DefaultDomain)"
    if (-not $upn) { return }

    $alias = Prompt-Text -Title "New Mailbox" -Label "Alias:" -Default $sam
    if (-not $alias) { return }

    $ou = Prompt-Text -Title "New Mailbox" -Label "Target OU (DN):" -Default $Config.DefaultOU
    if (-not $ou) { return }

    $password = Prompt-Text -Title "New Mailbox" -Label "Temporary password:" -Password
    if (-not $password) { return }

    try {
        $secure = ConvertTo-SecureString $password -AsPlainText -Force
        New-RemoteMailbox -Name "$first $last" -FirstName $first -LastName $last -UserPrincipalName $upn -SamAccountName $sam -Alias $alias -OnPremisesOrganizationalUnit $ou -Password $secure -ResetPasswordOnNextLogon:$true | Out-Null
        Write-Log "Created remote mailbox for $upn."
    } catch {
        Show-Error "Failed to create mailbox: $($_.Exception.Message)"
    }
}

function Enable-RemoteMailboxExisting {
    $user = Get-SelectedUser
    if (-not $user) { return }
    if (-not (Ensure-ExchangeSession)) { return }

    $alias = Prompt-Text -Title "Enable Remote Mailbox" -Label "Alias:" -Default $user.SamAccountName
    if (-not $alias) { return }

    $primary = Prompt-Text -Title "Enable Remote Mailbox" -Label "Primary SMTP address:" -Default "$alias@$($Config.DefaultDomain)"
    if (-not $primary) { return }

    $rrd = Prompt-Text -Title "Enable Remote Mailbox" -Label "Remote routing domain:" -Default $Config.RemoteRoutingDomain
    if (-not $rrd) { return }

    $remote = "$alias@$rrd"

    try {
        Enable-RemoteMailbox -Identity $user.SamAccountName -RemoteRoutingAddress $remote -PrimarySmtpAddress $primary | Out-Null
        Write-Log "Enabled remote mailbox for $($user.UserPrincipalName) with $remote."
    } catch {
        Show-Error "Failed to enable remote mailbox: $($_.Exception.Message)"
    }
}

function New-RemoteSharedMailbox {
    if (-not (Ensure-ExchangeSession)) { return }

    $name = Prompt-Text -Title "New Shared Mailbox" -Label "Display name:" -Default ""
    if (-not $name) { return }

    $alias = Prompt-Text -Title "New Shared Mailbox" -Label "Alias:" -Default ""
    if (-not $alias) { return }

    $primary = Prompt-Text -Title "New Shared Mailbox" -Label "Primary SMTP address:" -Default "$alias@$($Config.DefaultDomain)"
    if (-not $primary) { return }

    $ou = Prompt-Text -Title "New Shared Mailbox" -Label "Target OU (DN):" -Default $Config.DefaultOU
    if (-not $ou) { return }

    try {
        New-RemoteMailbox -Shared -Name $name -Alias $alias -PrimarySmtpAddress $primary -OnPremisesOrganizationalUnit $ou | Out-Null
        Write-Log "Created shared mailbox $primary."
    } catch {
        Show-Error "Failed to create shared mailbox: $($_.Exception.Message)"
    }
}

function New-RemoteRoomMailbox {
    if (-not (Ensure-ExchangeSession)) { return }

    $name = Prompt-Text -Title "New Room Mailbox" -Label "Display name:" -Default ""
    if (-not $name) { return }

    $alias = Prompt-Text -Title "New Room Mailbox" -Label "Alias:" -Default ""
    if (-not $alias) { return }

    $primary = Prompt-Text -Title "New Room Mailbox" -Label "Primary SMTP address:" -Default "$alias@$($Config.DefaultDomain)"
    if (-not $primary) { return }

    $ou = Prompt-Text -Title "New Room Mailbox" -Label "Target OU (DN):" -Default $Config.DefaultOU
    if (-not $ou) { return }

    try {
        New-RemoteMailbox -Room -Name $name -Alias $alias -PrimarySmtpAddress $primary -OnPremisesOrganizationalUnit $ou | Out-Null
        Write-Log "Created room mailbox $primary."
    } catch {
        Show-Error "Failed to create room mailbox: $($_.Exception.Message)"
    }
}

function Update-AdAttributes {
    $user = Get-SelectedUser
    if (-not $user) { return }

    try {
        $adUser = Get-ADUser -Identity $user.SamAccountName -Properties displayName,department,physicalDeliveryOfficeName,telephoneNumber

        $display = Prompt-Text -Title "Update AD" -Label "Display name:" -Default $adUser.DisplayName
        if ($display -eq $null) { return }

        $dept = Prompt-Text -Title "Update AD" -Label "Department:" -Default $adUser.Department
        if ($dept -eq $null) { return }

        $office = Prompt-Text -Title "Update AD" -Label "Office:" -Default $adUser.physicalDeliveryOfficeName
        if ($office -eq $null) { return }

        $phone = Prompt-Text -Title "Update AD" -Label "Phone:" -Default $adUser.telephoneNumber
        if ($phone -eq $null) { return }

        $replace = @{
            displayName = $display
            department = $dept
            physicalDeliveryOfficeName = $office
            telephoneNumber = $phone
        }

        Set-ADUser -Identity $user.SamAccountName -Replace $replace | Out-Null
        Write-Log "Updated AD attributes for $($user.UserPrincipalName)."
    } catch {
        Show-Error "Failed to update AD attributes: $($_.Exception.Message)"
    }
}

function Add-UserToGroup {
    $user = Get-SelectedUser
    if (-not $user) { return }

    $group = Get-SelectedGroup
    if (-not $group) { return }

    try {
        Add-ADGroupMember -Identity $group.DistinguishedName -Members $user.SamAccountName | Out-Null
        Write-Log "Added $($user.UserPrincipalName) to group $($group.Name)."
    } catch {
        Show-Error "Failed to add to group: $($_.Exception.Message)"
    }
}

function Remove-UserFromGroup {
    $user = Get-SelectedUser
    if (-not $user) { return }

    $group = Get-SelectedGroup
    if (-not $group) { return }

    try {
        Remove-ADGroupMember -Identity $group.DistinguishedName -Members $user.SamAccountName -Confirm:$false | Out-Null
        Write-Log "Removed $($user.UserPrincipalName) from group $($group.Name)."
    } catch {
        Show-Error "Failed to remove from group: $($_.Exception.Message)"
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange Hybrid Provisioning Tool"
$form.Size = New-Object System.Drawing.Size(940, 760)
$form.StartPosition = "CenterScreen"

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search user (name, sam, upn):"
$searchLabel.Location = New-Object System.Drawing.Point(12, 15)
$searchLabel.AutoSize = $true

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(12, 40)
$searchBox.Size = New-Object System.Drawing.Size(320, 20)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Location = New-Object System.Drawing.Point(340, 38)

$resultsList = New-Object System.Windows.Forms.ListBox
$resultsList.Location = New-Object System.Drawing.Point(12, 80)
$resultsList.Size = New-Object System.Drawing.Size(420, 180)

$detailsLabel = New-Object System.Windows.Forms.Label
$detailsLabel.Location = New-Object System.Drawing.Point(12, 270)
$detailsLabel.Size = New-Object System.Drawing.Size(420, 80)

$groupSearchLabel = New-Object System.Windows.Forms.Label
$groupSearchLabel.Text = "Search group (name, sam):"
$groupSearchLabel.Location = New-Object System.Drawing.Point(12, 350)
$groupSearchLabel.AutoSize = $true

$groupSearchBox = New-Object System.Windows.Forms.TextBox
$groupSearchBox.Location = New-Object System.Drawing.Point(12, 375)
$groupSearchBox.Size = New-Object System.Drawing.Size(320, 20)

$groupSearchButton = New-Object System.Windows.Forms.Button
$groupSearchButton.Text = "Search"
$groupSearchButton.Location = New-Object System.Drawing.Point(340, 373)

$groupResultsList = New-Object System.Windows.Forms.ListBox
$groupResultsList.Location = New-Object System.Drawing.Point(12, 405)
$groupResultsList.Size = New-Object System.Drawing.Size(420, 110)

$groupDetailsLabel = New-Object System.Windows.Forms.Label
$groupDetailsLabel.Location = New-Object System.Drawing.Point(12, 520)
$groupDetailsLabel.Size = New-Object System.Drawing.Size(420, 50)

$actionsLabel = New-Object System.Windows.Forms.Label
$actionsLabel.Text = "Actions"
$actionsLabel.Location = New-Object System.Drawing.Point(460, 15)
$actionsLabel.AutoSize = $true

$btnNewMailbox = New-Object System.Windows.Forms.Button
$btnNewMailbox.Text = "Create Remote Mailbox"
$btnNewMailbox.Size = New-Object System.Drawing.Size(200, 30)
$btnNewMailbox.Location = New-Object System.Drawing.Point(460, 40)

$btnShared = New-Object System.Windows.Forms.Button
$btnShared.Text = "Create Shared Mailbox"
$btnShared.Size = New-Object System.Drawing.Size(200, 30)
$btnShared.Location = New-Object System.Drawing.Point(460, 80)

$btnRoom = New-Object System.Windows.Forms.Button
$btnRoom.Text = "Create Room Mailbox"
$btnRoom.Size = New-Object System.Drawing.Size(200, 30)
$btnRoom.Location = New-Object System.Drawing.Point(460, 120)

$btnAlias = New-Object System.Windows.Forms.Button
$btnAlias.Text = "Add Email Alias"
$btnAlias.Size = New-Object System.Drawing.Size(200, 30)
$btnAlias.Location = New-Object System.Drawing.Point(460, 160)

$btnPrimary = New-Object System.Windows.Forms.Button
$btnPrimary.Text = "Set Primary SMTP"
$btnPrimary.Size = New-Object System.Drawing.Size(200, 30)
$btnPrimary.Location = New-Object System.Drawing.Point(460, 200)

$btnHide = New-Object System.Windows.Forms.Button
$btnHide.Text = "Hide From GAL"
$btnHide.Size = New-Object System.Drawing.Size(200, 30)
$btnHide.Location = New-Object System.Drawing.Point(460, 240)

$btnUnhide = New-Object System.Windows.Forms.Button
$btnUnhide.Text = "Unhide From GAL"
$btnUnhide.Size = New-Object System.Drawing.Size(200, 30)
$btnUnhide.Location = New-Object System.Drawing.Point(460, 280)

$btnFullAccess = New-Object System.Windows.Forms.Button
$btnFullAccess.Text = "Grant Full Access"
$btnFullAccess.Size = New-Object System.Drawing.Size(200, 30)
$btnFullAccess.Location = New-Object System.Drawing.Point(460, 320)

$btnSendAs = New-Object System.Windows.Forms.Button
$btnSendAs.Text = "Grant Send As"
$btnSendAs.Size = New-Object System.Drawing.Size(200, 30)
$btnSendAs.Location = New-Object System.Drawing.Point(460, 360)

$btnEnableMailbox = New-Object System.Windows.Forms.Button
$btnEnableMailbox.Text = "Enable Remote Mailbox"
$btnEnableMailbox.Size = New-Object System.Drawing.Size(200, 30)
$btnEnableMailbox.Location = New-Object System.Drawing.Point(680, 40)

$btnRemoveAlias = New-Object System.Windows.Forms.Button
$btnRemoveAlias.Text = "Remove Email Alias"
$btnRemoveAlias.Size = New-Object System.Drawing.Size(200, 30)
$btnRemoveAlias.Location = New-Object System.Drawing.Point(680, 80)

$btnConvertShared = New-Object System.Windows.Forms.Button
$btnConvertShared.Text = "Convert to Shared"
$btnConvertShared.Size = New-Object System.Drawing.Size(200, 30)
$btnConvertShared.Location = New-Object System.Drawing.Point(680, 120)

$btnRemoveFull = New-Object System.Windows.Forms.Button
$btnRemoveFull.Text = "Remove Full Access"
$btnRemoveFull.Size = New-Object System.Drawing.Size(200, 30)
$btnRemoveFull.Location = New-Object System.Drawing.Point(680, 160)

$btnRemoveSendAs = New-Object System.Windows.Forms.Button
$btnRemoveSendAs.Text = "Remove Send As"
$btnRemoveSendAs.Size = New-Object System.Drawing.Size(200, 30)
$btnRemoveSendAs.Location = New-Object System.Drawing.Point(680, 200)

$btnForwarding = New-Object System.Windows.Forms.Button
$btnForwarding.Text = "Set Forwarding"
$btnForwarding.Size = New-Object System.Drawing.Size(200, 30)
$btnForwarding.Location = New-Object System.Drawing.Point(680, 240)

$btnRemoveForwarding = New-Object System.Windows.Forms.Button
$btnRemoveForwarding.Text = "Remove Forwarding"
$btnRemoveForwarding.Size = New-Object System.Drawing.Size(200, 30)
$btnRemoveForwarding.Location = New-Object System.Drawing.Point(680, 280)

$btnUpdateAd = New-Object System.Windows.Forms.Button
$btnUpdateAd.Text = "Update AD Attributes"
$btnUpdateAd.Size = New-Object System.Drawing.Size(200, 30)
$btnUpdateAd.Location = New-Object System.Drawing.Point(680, 320)

$btnAddGroup = New-Object System.Windows.Forms.Button
$btnAddGroup.Text = "Add to Group"
$btnAddGroup.Size = New-Object System.Drawing.Size(200, 30)
$btnAddGroup.Location = New-Object System.Drawing.Point(680, 360)

$btnRemoveGroup = New-Object System.Windows.Forms.Button
$btnRemoveGroup.Text = "Remove from Group"
$btnRemoveGroup.Size = New-Object System.Drawing.Size(200, 30)
$btnRemoveGroup.Location = New-Object System.Drawing.Point(680, 400)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(12, 580)
$logBox.Size = New-Object System.Drawing.Size(900, 150)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true

$searchButton.Add_Click({ Search-Users -Query $searchBox.Text })
$searchBox.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { Search-Users -Query $searchBox.Text } })
$resultsList.Add_SelectedIndexChanged({
    $selected = $resultsList.SelectedItem
    if ($selected) {
        $detailsLabel.Text = "UPN: {0}`r`nSam: {1}`r`nMail: {2}" -f $selected.UserPrincipalName, $selected.SamAccountName, $selected.Mail
    }
})

$groupSearchButton.Add_Click({ Search-Groups -Query $groupSearchBox.Text })
$groupSearchBox.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { Search-Groups -Query $groupSearchBox.Text } })
$groupResultsList.Add_SelectedIndexChanged({
    $selected = $groupResultsList.SelectedItem
    if ($selected) {
        $groupDetailsLabel.Text = "Name: {0}`r`nSam: {1}" -f $selected.Name, $selected.SamAccountName
    }
})

$btnNewMailbox.Add_Click({ New-RemoteMailboxUser })
$btnShared.Add_Click({ New-RemoteSharedMailbox })
$btnRoom.Add_Click({ New-RemoteRoomMailbox })
$btnAlias.Add_Click({ Add-EmailAlias })
$btnPrimary.Add_Click({ Set-PrimarySmtp })
$btnHide.Add_Click({ Hide-FromGal -Hide $true })
$btnUnhide.Add_Click({ Hide-FromGal -Hide $false })
$btnFullAccess.Add_Click({ Add-FullAccess })
$btnSendAs.Add_Click({ Add-SendAs })
$btnEnableMailbox.Add_Click({ Enable-RemoteMailboxExisting })
$btnRemoveAlias.Add_Click({ Remove-EmailAlias })
$btnConvertShared.Add_Click({ Convert-ToSharedMailbox })
$btnRemoveFull.Add_Click({ Remove-FullAccess })
$btnRemoveSendAs.Add_Click({ Remove-SendAs })
$btnForwarding.Add_Click({ Set-Forwarding })
$btnRemoveForwarding.Add_Click({ Disable-Forwarding })
$btnUpdateAd.Add_Click({ Update-AdAttributes })
$btnAddGroup.Add_Click({ Add-UserToGroup })
$btnRemoveGroup.Add_Click({ Remove-UserFromGroup })

$form.Controls.AddRange(@(
    $searchLabel, $searchBox, $searchButton,
    $resultsList, $detailsLabel,
    $groupSearchLabel, $groupSearchBox, $groupSearchButton,
    $groupResultsList, $groupDetailsLabel,
    $actionsLabel,
    $btnNewMailbox, $btnShared, $btnRoom,
    $btnAlias, $btnPrimary, $btnHide, $btnUnhide,
    $btnFullAccess, $btnSendAs,
    $btnEnableMailbox, $btnRemoveAlias, $btnConvertShared,
    $btnRemoveFull, $btnRemoveSendAs, $btnForwarding,
    $btnRemoveForwarding, $btnUpdateAd, $btnAddGroup, $btnRemoveGroup,
    $logBox
))

Write-Log "Ready. Configure ExchangeUri/DefaultOU/DefaultDomain in the script if needed."
[void]$form.ShowDialog()

if ($script:ExchSession) {
    Remove-PSSession $script:ExchSession
}
