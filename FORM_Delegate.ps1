#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.7.2" } -Version 7.1
<#
.SYNOPSIS
    Displays and updates standard folder permissions for an Exchange Online mailbox delegate.
    Originally published https://github.com/michevnew/PowerShell/blob/3d7a120a471e0e28f8b19ecb73765ee806120cf7/Change_meeting_organizer_bulk.ps1

.DESCRIPTION
    Queries mailbox status and current Calendar, Tasks, Inbox, Contacts, and Notes permissions
    for a delegate, presents them in a Windows Forms dialog, applies the selected changes when
    accepted, then re-queries Exchange Online and verifies the resulting permission state.

.PARAMETER Owner
    Mailbox whose folder permissions are being managed. The Manager alias is supported.

.PARAMETER Delegate
    User whose folder permissions are being managed.

.PARAMETER ReadOnly
    Displays the current permissions without allowing changes.

.PARAMETER Diag
    Enables diagnostic messages for optional form assets.

.PARAMETER RestoreFrom
    Restores the managed folder permissions from a backup JSON file. The script first verifies
    that the backup's mailbox GUIDs match the live Owner and Delegate, then prompts to back up
    the current permissions before restoring.
    New backups are stored under Documents\ExchangeDelegateBackups.

.EXAMPLE
    .\FORM_Delegate.ps1 -Owner mailbox@contoso.com -Delegate user@contoso.com

.EXAMPLE
    .\FORM_Delegate.ps1 -Owner mailbox@contoso.com -Delegate user@contoso.com -RestoreFrom "$HOME\Documents\ExchangeDelegateBackups\backup.json"

.NOTES
    Requires ExchangeOnlineManagement 3.7.2 or later and Exchange permissions to read and
    manage mailbox folder permissions. Windows Forms requires Windows PowerShell 7.1 or later.
    Changelog:
    2026-10-28 - Restored script parameters and ensured permission changes run after dialog
    submission; unchanged folder permissions are skipped.
    2026-10-28 - Added colored console operation logs and before/after permission summaries.
    2026-10-28 - Added live mailbox status queries, strict folder inventory checks, and
    post-change verification against Exchange Online state.
    2026-10-28 - Added pre-change permission backup consent and validated backup restore mode.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][Alias("Manager")][ValidateNotNullOrEmpty()][string]$Owner,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Delegate,
    [switch]$ReadOnly,
    [switch]$Diag,
    [ValidateNotNullOrEmpty()][string]$RestoreFrom
)

function Write-ConsoleLog {
    param(
        [Parameter(Mandatory)][ValidateSet("INFO", "CHANGE", "SUCCESS", "WARNING", "ERROR", "SKIP")][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )

    $colors = @{
        INFO    = "Cyan"
        CHANGE  = "Magenta"
        SUCCESS = "Green"
        WARNING = "Yellow"
        ERROR   = "Red"
        SKIP    = "DarkGray"
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $colors[$Level]
}

function Write-PermissionSummary {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][hashtable]$Permissions
    )

    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    Write-Host "Mailbox: $($Permissions['Owner'])    Delegate: $($Permissions['Delegate'])" -ForegroundColor Gray
    Write-Host ("{0,-12} {1}" -f "Folder", "Access right") -ForegroundColor White
    Write-Host ("{0,-12} {1}" -f "------", "------------") -ForegroundColor DarkGray

    foreach ($folder in @("Calendar", "Tasks", "Inbox", "Contacts", "Notes")) {
        $accessRight = if ($Permissions.ContainsKey($folder)) { [string]$Permissions[$folder] } else { "Unavailable" }
        $color = switch ($accessRight) {
            "None" { "DarkGray" }
            "Editor" { "Green" }
            "Custom" { "Yellow" }
            "Unavailable" { "Red" }
            default { "White" }
        }
        Write-Host ("{0,-12} {1}" -f $folder, $accessRight) -ForegroundColor $color
    }

    $calendarFlags = if ($Permissions["CalendarExt"] -and $Permissions["CalendarExt"].Count -gt 0) {
        $Permissions["CalendarExt"] -join ", "
    }
    else {
        "None"
    }
    Write-Host "Calendar delegate flags: $calendarFlags" -ForegroundColor Gray
    Write-Host ""
}

function Get-MailboxStatus {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Identity,
        [Parameter(Mandatory)][ValidateSet("Mailbox", "Delegate")][string]$Role
    )

    try {
        $mailboxes = @(Get-EXOMailbox -Identity $Identity -PropertySets Archive -ErrorAction Stop -Verbose:$false)
    }
    catch {
        Write-ConsoleLog -Level ERROR -Message "Could not query Exchange Online status for $Role '$Identity': $($_.Exception.Message)"
        throw
    }

    if ($mailboxes.Count -ne 1) {
        $message = "Expected one $Role mailbox for '$Identity', but Exchange returned $($mailboxes.Count)."
        Write-ConsoleLog -Level ERROR -Message $message
        throw $message
    }

    $mailbox = $mailboxes[0]
    if ([string]::IsNullOrWhiteSpace([string]$mailbox.PrimarySmtpAddress) -or
        [string]::IsNullOrWhiteSpace([string]$mailbox.RecipientTypeDetails) -or
        [string]::IsNullOrWhiteSpace([string]$mailbox.Guid)) {
        $message = "Exchange returned incomplete mailbox status for '$Identity'. No permission changes will be attempted."
        Write-ConsoleLog -Level ERROR -Message $message
        throw $message
    }

    [pscustomobject]@{
        Role                 = $Role
        RequestedIdentity    = $Identity
        DisplayName          = $mailbox.DisplayName
        PrimarySmtpAddress   = [string]$mailbox.PrimarySmtpAddress
        RecipientTypeDetails = [string]$mailbox.RecipientTypeDetails
        Guid                 = [string]$mailbox.Guid
        ArchiveStatus        = [string]$mailbox.ArchiveStatus
        Status               = "Mailbox object returned by Exchange Online"
    }
}

function Write-MailboxStatus {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][object[]]$Status
    )

    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    foreach ($mailbox in $Status) {
        Write-Host ("{0}: {1} ({2})" -f $mailbox.Role, $mailbox.PrimarySmtpAddress, $mailbox.RecipientTypeDetails) -ForegroundColor Green
        Write-Host ("  Display name: {0}; GUID: {1}; Archive: {2}; Status: {3}" -f $mailbox.DisplayName, $mailbox.Guid, $mailbox.ArchiveStatus, $mailbox.Status) -ForegroundColor Gray
    }
    Write-Host ""
}

function Test-PermissionSnapshot {
    param(
        [Parameter(Mandatory)][hashtable]$Expected,
        [Parameter(Mandatory)][hashtable]$Actual
    )

    $mismatches = @()
    foreach ($folder in @("Calendar", "Tasks", "Inbox", "Contacts", "Notes")) {
        if (-not $Actual.ContainsKey($folder)) {
            $mismatches += "${folder}: not returned by Exchange"
        }
        elseif ([string]$Expected[$folder] -ne [string]$Actual[$folder]) {
            $mismatches += "${folder}: expected '$($Expected[$folder])', found '$($Actual[$folder])'"
        }
    }

    $expectedFlags = @($Expected["CalendarExt"] | Sort-Object)
    $actualFlags = @($Actual["CalendarExt"] | Sort-Object)
    if (($expectedFlags -join ",") -ne ($actualFlags -join ",")) {
        $mismatches += "Calendar delegate flags: expected '$($expectedFlags -join ', ')', found '$($actualFlags -join ', ')'"
    }

    [pscustomobject]@{
        Succeeded  = ($mismatches.Count -eq 0)
        Mismatches = $mismatches
    }
}

function Read-PermissionBackup {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [Parameter(Mandatory)][object[]]$CurrentMailboxStatus
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Permission backup file not found: $Path"
    }

    try {
        $backup = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    }
    catch {
        throw "Could not read permission backup '$Path': $($_.Exception.Message)"
    }

    if ($backup.SchemaVersion -ne 1 -or $backup.Schema -ne "ExchangeDelegateTool.PermissionBackup") {
        throw "Unsupported or invalid permission backup format: $Path"
    }

    foreach ($role in @("Mailbox", "Delegate")) {
        $currentMailbox = $CurrentMailboxStatus | Where-Object Role -eq $role | Select-Object -First 1
        $backupRole = $backup.Mailboxes[$role]
        if ($null -eq $currentMailbox -or $null -eq $backupRole -or
            [string]::IsNullOrWhiteSpace([string]$backupRole.Guid) -or
            $currentMailbox.Guid -ne $backupRole.Guid) {
            throw "Backup $role identity does not match the live Exchange mailbox. No changes were made."
        }
    }

    $requiredFolders = @("Calendar", "Tasks", "Inbox", "Contacts", "Notes")
    $allowedRights = @{
        Calendar = @("None", "AvailabilityOnly", "LimitedDetails", "Reviewer", "Author", "Editor")
        Tasks    = @("None", "Reviewer", "Author", "Editor")
        Inbox    = @("None", "Reviewer", "Author", "Editor")
        Contacts = @("None", "Reviewer", "Author", "Editor")
        Notes    = @("None", "Reviewer", "Author", "Editor")
    }
    foreach ($folder in $requiredFolders) {
        if (-not $backup.Permissions.ContainsKey($folder) -or
            $backup.Permissions[$folder] -notin $allowedRights[$folder]) {
            throw "Backup contains a missing or unsupported permission for $folder. No changes were made."
        }
    }

    $allowedFlags = @("Delegate", "CanViewPrivateItems", "CanManageCategories")
    $flags = @($backup.Permissions.CalendarExt)
    if (@($flags | Where-Object { $_ -notin $allowedFlags }).Count -gt 0 -or
        ($flags -contains "CanViewPrivateItems" -and $flags -notcontains "Delegate") -or
        ($flags -contains "CanManageCategories" -and $flags -notcontains "Delegate") -or
        ($flags.Count -gt 0 -and $backup.Permissions.Calendar -ne "Editor")) {
        throw "Backup contains invalid Calendar delegate flags. No changes were made."
    }

    return $backup
}

function Save-PreChangePermissionBackup {
    param(
        [Parameter(Mandatory)][hashtable]$Permissions,
        [Parameter(Mandatory)][object[]]$MailboxStatus
    )

    $restorableRights = @{
        Calendar = @("None", "AvailabilityOnly", "LimitedDetails", "Reviewer", "Author", "Editor")
        Tasks    = @("None", "Reviewer", "Author", "Editor")
        Inbox    = @("None", "Reviewer", "Author", "Editor")
        Contacts = @("None", "Reviewer", "Author", "Editor")
        Notes    = @("None", "Reviewer", "Author", "Editor")
    }
    foreach ($folder in @("Calendar", "Tasks", "Inbox", "Contacts", "Notes")) {
        if ($Permissions[$folder] -notin $restorableRights[$folder]) {
            throw "Cannot create a restorable backup because $folder has unsupported/custom access rights '$($Permissions[$folder])'. No changes were made."
        }
    }

    $statusByRole = @{}
    foreach ($mailbox in $MailboxStatus) { $statusByRole[$mailbox.Role] = $mailbox }
    foreach ($role in @("Mailbox", "Delegate")) {
        if (-not $statusByRole.ContainsKey($role) -or [string]::IsNullOrWhiteSpace([string]$statusByRole[$role].Guid)) {
            throw "Cannot create a verified permission backup because live $role mailbox status is unavailable. No changes were made."
        }
    }

    $documentsPath = [Environment]::GetFolderPath("MyDocuments")
    if ([string]::IsNullOrWhiteSpace($documentsPath)) {
        throw "Cannot determine the Documents directory for a permission backup. No changes were made."
    }
    $backupDirectory = Join-Path $documentsPath "ExchangeDelegateBackups"
    try {
        $null = New-Item -Path $backupDirectory -ItemType Directory -Force -ErrorAction Stop
    }
    catch {
        throw "Could not create permission backup directory '$backupDirectory': $($_.Exception.Message). No changes were made."
    }

    $safeOwner = $Permissions.Owner -replace '[^a-zA-Z0-9._-]', '_'
    $safeDelegate = $Permissions.Delegate -replace '[^a-zA-Z0-9._-]', '_'
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"
    $uniqueSuffix = [guid]::NewGuid().ToString("N").Substring(0, 8)
    $backupPath = Join-Path $backupDirectory "${safeOwner}_${safeDelegate}_${timestamp}_${uniqueSuffix}.json"

    Write-ConsoleLog -Level WARNING -Message "A pre-change backup is required before modifying Exchange permissions."
    Write-Host "Backup file: $backupPath" -ForegroundColor Gray
    $answer = Read-Host "Create this backup and continue? [Y/N]"
    if ($answer -notmatch '^(?i:y|yes)$') {
        Write-ConsoleLog -Level SKIP -Message "Backup declined; no mailbox permissions were changed."
        return $null
    }

    $backup = [ordered]@{
        SchemaVersion = 1
        Schema        = "ExchangeDelegateTool.PermissionBackup"
        CreatedAtUtc  = (Get-Date).ToUniversalTime().ToString("o")
        Mailboxes     = [ordered]@{
            Mailbox  = [ordered]@{
                PrimarySmtpAddress = $statusByRole["Mailbox"].PrimarySmtpAddress
                Guid               = $statusByRole["Mailbox"].Guid
            }
            Delegate = [ordered]@{
                PrimarySmtpAddress = $statusByRole["Delegate"].PrimarySmtpAddress
                Guid               = $statusByRole["Delegate"].Guid
            }
        }
        Permissions   = [ordered]@{
            Calendar    = $Permissions.Calendar
            Tasks       = $Permissions.Tasks
            Inbox       = $Permissions.Inbox
            Contacts    = $Permissions.Contacts
            Notes       = $Permissions.Notes
            CalendarExt = @($Permissions.CalendarExt)
        }
    }

    try {
        $backup | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $backupPath -Encoding utf8NoBOM -ErrorAction Stop
        $verifiedBackup = Get-Content -LiteralPath $backupPath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        if ($verifiedBackup.Schema -ne $backup.Schema -or $verifiedBackup.Mailboxes.Mailbox.Guid -ne $statusByRole["Mailbox"].Guid) {
            throw "Backup verification did not match the source mailbox."
        }
    }
    catch {
        Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        throw "Could not save and verify pre-change permission backup: $($_.Exception.Message). No changes were made."
    }

    Write-ConsoleLog -Level SUCCESS -Message "Verified pre-change backup saved: $backupPath"
    Write-Host "Restore command: & '$PSCommandPath' -Owner '$($Permissions.Owner)' -Delegate '$($Permissions.Delegate)' -RestoreFrom '$backupPath'" -ForegroundColor Cyan
    return $backupPath
}

#For details on what the script does and how to run it, check: https://michev.info/blog/post/7306/free-tool-to-man…utlook-delegates

#Reset WhatIfPreference
$script:WhatIfPreference = $false
# Clean up before startup
Remove-Variable params, ExOSession, cUser, cFolders, cPermissions, cPermissionsNew, cFlags, cFlagsValue, FormCancelled -ErrorAction SilentlyContinue
$script:params = $PSBoundParameters

# Connect to ExO with required scopes
Write-ConsoleLog -Level INFO -Message "Connecting to Exchange Online..."
$permissionCommandCount = @(Get-Command *-MailboxFolderPermission -ErrorAction SilentlyContinue).Count
$mailboxQueryCommand = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
$script:ExOSession = Get-ConnectionInformation | Where-Object State -eq "Connected" | Select-Object -First 1
if (-not $script:ExOSession -or $permissionCommandCount -lt 4 -or -not $mailboxQueryCommand) {
    # No existing session, connect. Load only the cmdlets we need.No need for Set-Mailbox, the delegate flag toggles it automatically
    # Remember that we can also leverage the REST cmdlets for GET operations, no need to load these
    Connect-ExchangeOnline -DisableWAM -CommandName *-MailboxFolderPermission, Get-MailboxFolderStatistics, Get-MailboxFolder, Get-EXOMailbox -ShowBanner:$false -Verbose:$false -ErrorAction Stop
    $script:ExOSession = Get-ConnectionInformation | Where-Object State -eq "Connected" | Select-Object -First 1
}
else {
    # Seems we have an existing session, reuse it
    Write-Verbose "Reusing existing Exchange Online session..."
}
if (-not $script:ExOSession) {
    Write-ConsoleLog -Level ERROR -Message "No connected Exchange Online session is available."
    throw "Exchange Online connection could not be confirmed."
}
Write-ConsoleLog -Level SUCCESS -Message "Exchange Online session is ready."

# Cater to self-service scenarios, we need to know the current user
if (!(Get-Command Get-MailboxFolderStatistics -ErrorAction SilentlyContinue)) {
    Write-ConsoleLog -Level INFO -Message "Running with limited permissions; using the current user's folder cmdlets."
}
$script:cUser = ${ExOSession}?.UserPrincipalName

# Load Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Make sure we use the system theme for controls
[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-FormIcon {
    param(
        [Parameter(Mandatory)][string]$IconName
    )

    $iconPath = Join-Path -Path $PSScriptRoot -ChildPath $IconName
    if (Test-Path -Path $iconPath -PathType Leaf) {
        if ($Diag) {
            Write-Host "[Diag] Icon found: $iconPath" -ForegroundColor Green
        }
        return [System.Drawing.Image]::FromFile($iconPath)
    }

    if ($Diag) {
        Write-Warning "[Diag] Missing icon file: $iconPath"
    }

    $fallbackIcon = [System.Drawing.SystemIcons]::Information
    if ($null -ne $fallbackIcon) {
        return $fallbackIcon.ToBitmap()
    }

    return [System.Drawing.Bitmap]::new(24, 24)
}

# Main form object
$form1 = New-Object System.Windows.Forms.Form
$form1.ClientSize = New-Object System.Drawing.Size(424, 309)
$form1.Name = "Form1"
$form1.Text = "Delegate Permissions: $Delegate on $Owner"
$form1.StartPosition = "CenterScreen"
$form1.FormBorderStyle = 'FixedDialog'
$form1.MaximizeBox = $false
$form1.MinimizeBox = $false
$form1.Font = New-Object System.Drawing.Font("Segoe UI SemiBold", 8, [System.Drawing.FontStyle]::Regular)
$form1.KeyPreview = $True

# Main group box
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Name = "groupBox1"
$groupBox1.Location = New-Object System.Drawing.Size(6, 8)
$groupBox1.Size = New-Object System.Drawing.Size(412, 220)
$groupBox1.TabStop = $False #skips control in the tab order (tab-pressing)
$groupBox1.Text = "This delegate has the following permissions"
$form1.Controls.Add($groupBox1)

#Region Calendar Controls
# First Icon
$pictureBox1 = New-Object System.Windows.Forms.PictureBox
$pictureBox1.Image = Get-FormIcon -IconName "Calendar3.ico"
$pictureBox1.Location = New-Object System.Drawing.Size(10, 19)
$pictureBox1.Size = New-Object System.Drawing.Size(24, 24)
$pictureBox1.Name = "pictureBox1"
$pictureBox1.SizeMode = 2 #controls icon shrink/resize/etc
$pictureBox1.TabStop = $False
$groupBox1.Controls.Add($pictureBox1)

# First label
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Size(50, 29)
$label1.Size = New-Object System.Drawing.Size(50, 18)
$label1.Name = "label1"
$label1.UseMnemonic = $true # Use & for ALT-shortcuts; Doesnt seem to work when combobox.DropDownStyle = DropDownList so use Add_keydown event
$label1.Text = "&Calendar"
$groupBox1.Controls.Add($label1)

# First combobox
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.FormattingEnabled = $True
$comboBox1.Location = New-Object System.Drawing.Size(110, 25)
$comboBox1.Size = New-Object System.Drawing.Size(245, 21)
$comboBox1.Name = "comboBox1"
$comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox1.Items.AddRange(("None", "Reviewer (can read items)", "Author (can read and create items)", "Editor (can read, create, and modify items)"))
$comboBox1.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox1)

# First checkbox
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Location = New-Object System.Drawing.Size(54, 49)
$checkbox1.Size = New-Object System.Drawing.Size(355, 24)
$checkBox1.Name = "checkBox1"
$checkBox1.Text = "&Delegate receives copies of meeting-related messages sent to me"
$checkBox1.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox1.UseVisualStyleBackColor = $True
$groupBox1.Controls.Add($checkBox1)
#EndRegion Calendar Controls

#Region Tasks Controls
# Second Icon
$pictureBox2 = New-Object System.Windows.Forms.PictureBox
$pictureBox2.Image = Get-FormIcon -IconName "Tasks2.ico"
$pictureBox2.Location = New-Object System.Drawing.Size(10, 76)
$pictureBox2.Size = New-Object System.Drawing.Size(24, 24)
$pictureBox2.Name = "pictureBox2"
$pictureBox2.SizeMode = 2
$pictureBox2.TabStop = $False
$groupBox1.Controls.Add($pictureBox2)

# Second label
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Size(50, 85)
$label2.Size = New-Object System.Drawing.Size(50, 18)
$label2.UseMnemonic = $true # Use & for ALT-shortcuts
$label2.Name = "label2"
$label2.Text = "&Tasks"
$groupBox1.Controls.Add($label2)

# Second combobox
$comboBox2 = New-Object System.Windows.Forms.ComboBox
$comboBox2.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox2.FormattingEnabled = $True
$comboBox2.Location = New-Object System.Drawing.Size(110, 80)
$comboBox2.Size = New-Object System.Drawing.Size(245, 21)
$comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox2.Name = "comboBox2"
$comboBox2.Items.AddRange(("None", "Reviewer (can read items)", "Author (can read and create items)", "Editor (can read, create, and modify items)"))
$comboBox2.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox2)
#EndRegion Tasks Controls

#Region Inbox Controls
# Third Icon
$pictureBox3 = New-Object System.Windows.Forms.PictureBox
$pictureBox3.Image = Get-FormIcon -IconName "inbox.ico"
$pictureBox3.Location = New-Object System.Drawing.Size(10, 111)
$pictureBox3.Size = New-Object System.Drawing.Size(24, 24)
$pictureBox3.Name = "pictureBox3"
$pictureBox3.SizeMode = 2
$pictureBox3.TabStop = $False
$groupBox1.Controls.Add($pictureBox3)

# Third label
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Size(50, 120)
$label3.Size = New-Object System.Drawing.Size(50, 18)
$label3.Name = "label3"
$label3.UseMnemonic = $true # Use & for ALT-shortcuts
$label3.Text = "&Inbox"
$groupBox1.Controls.Add($label3)

# Third combobox
$comboBox3 = New-Object System.Windows.Forms.ComboBox
$comboBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox3.FormattingEnabled = $True
$comboBox3.Location = New-Object System.Drawing.Size(110, 115)
$comboBox3.Size = New-Object System.Drawing.Size(245, 21)
$comboBox3.Name = "comboBox3"
$comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox3.Items.AddRange(("None", "Reviewer (can read items)", "Author (can read and create items)", "Editor (can read, create, and modify items)"))
$comboBox3.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox3)
#EndRegion Inbox Controls

#Region Contacts Controls
$pictureBox4 = New-Object System.Windows.Forms.PictureBox
$pictureBox4.Image = Get-FormIcon -IconName "contacts2.ico"
$pictureBox4.Location = New-Object System.Drawing.Size(10, 146)
$pictureBox4.Size = New-Object System.Drawing.Size(24, 24)
$pictureBox4.Name = "pictureBox4"
$pictureBox4.SizeMode = 2
$pictureBox4.TabStop = $False
$groupBox1.Controls.Add($pictureBox4)

# Fourth label
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Size(50, 155)
$label4.Size = New-Object System.Drawing.Size(50, 18)
$label4.Name = "label4"
$label4.UseMnemonic = $true # Use & for ALT-shortcuts
$label4.Text = "C&ontacts"
$groupBox1.Controls.Add($label4)

# Fourth combobox
$comboBox4 = New-Object System.Windows.Forms.ComboBox
$comboBox4.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox4.FormattingEnabled = $True
$comboBox4.Location = New-Object System.Drawing.Size(110, 150)
$comboBox4.Size = New-Object System.Drawing.Size(245, 21)
$comboBox4.Name = "comboBox4"
$comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox4.Items.AddRange(("None", "Reviewer (can read items)", "Author (can read and create items)", "Editor (can read, create, and modify items)"))
$comboBox4.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox4)
#EndRegion Contacts Controls

#Region Notes Controls
# Fifth Icon
$pictureBox5 = New-Object System.Windows.Forms.PictureBox
$pictureBox5.Image = Get-FormIcon -IconName "notes2.ico"
$pictureBox5.Location = New-Object System.Drawing.Size(10, 181)
$pictureBox5.Size = New-Object System.Drawing.Size(24, 24)
$pictureBox5.Name = "pictureBox5"
$pictureBox5.SizeMode = 2
$pictureBox5.TabStop = $False
$groupBox1.Controls.Add($pictureBox5)

# Fifth label
$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Size(50, 190)
$label5.Size = New-Object System.Drawing.Size(50, 18)
$label5.Name = "label5"
$label5.UseMnemonic = $true # Use & for ALT-shortcuts
$label5.Text = "&Notes"
$groupBox1.Controls.Add($label5)

# Fifth combobox
$comboBox5 = New-Object System.Windows.Forms.ComboBox
$comboBox5.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox5.FormattingEnabled = $True
$comboBox5.Location = New-Object System.Drawing.Size(110, 185)
$comboBox5.Size = New-Object System.Drawing.Size(245, 21)
$comboBox5.Name = "comboBox5"
$comboBox5.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox5.Items.AddRange(("None", "Reviewer (can read items)", "Author (can read and create items)", "Editor (can read, create, and modify items)"))
$comboBox5.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox5)
#EndRegion Notes Controls

#Region Remaining Non-Groupbox Controls
#Last two checkboxes
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Location = New-Object System.Drawing.Size(6, 228)
$checkBox2.Size = New-Object System.Drawing.Size(400, 24)
$checkBox2.Name = "checkBox2"
$checkBox2.UseMnemonic = $true # Use & for ALT-shortcuts
$checkBox2.Text = "Automatically &send a message to delegate summarizing these permissions"
$checkBox2.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox2.UseVisualStyleBackColor = $True
$form1.Controls.Add($checkBox2)

$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox3.Location = New-Object System.Drawing.Size(6, 248)
$checkBox3.Size = New-Object System.Drawing.Size(200, 24)
$checkBox3.Name = "checkBox3"
$checkBox3.UseMnemonic = $true # Use & for ALT-shortcuts
$checkBox3.Text = "Delegate can see my &private items"
$checkBox3.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox3.UseVisualStyleBackColor = $True
$form1.Controls.Add($checkBox3)

$checkBox4 = New-Object System.Windows.Forms.CheckBox
$checkBox4.Location = New-Object System.Drawing.Size(206, 248)
$checkBox4.Size = New-Object System.Drawing.Size(400, 24)
$checkBox4.Name = "checkBox4"
$checkBox4.UseMnemonic = $true # Use & for ALT-shortcuts
$checkBox4.Text = "Delegate can &manage categories"
$checkBox4.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox4.UseVisualStyleBackColor = $True
$form1.Controls.Add($checkBox4)
#Endregion Remaining Non-Groupbox Controls

#Region OK and Cancel buttons
# OK button
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(132, 279)
$OKButton.Size = New-Object System.Drawing.Size(75, 23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form1.AcceptButton = $OKButton
$form1.Controls.Add($OKButton)

# Cancel button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(216, 279)
$CancelButton.Size = New-Object System.Drawing.Size(75, 23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form1.CancelButton = $CancelButton
$form1.Controls.Add($CancelButton)
#Endregion OK and Cancel buttons

#region Event Handlers
# Load event handler
$form1.add_Load({
        if ($cPermissions.ContainsKey("Calendar")) {
            if ($cPermissions["Calendar"] -notin @("Author", "Editor", "Reviewer", "None")) {
                $comboBox1.Items.Add("Custom")
                $comboBox1.SelectedItem = "Custom"
            }
            else { $comboBox1.SelectedIndex = Convert-ComboTextToInt($cPermissions["Calendar"]) }
        }
        if ($cPermissions.ContainsKey("Tasks")) {
            if ($cPermissions["Tasks"] -notin @("Author", "Editor", "Reviewer", "None")) {
                $comboBox2.Items.Add("Custom")
                $comboBox2.SelectedItem = "Custom"
            }
            else { $comboBox2.SelectedIndex = Convert-ComboTextToInt($cPermissions["Tasks"]) }
        }
        if ($cPermissions.ContainsKey("Inbox")) {
            if ($cPermissions["Inbox"] -notin @("Author", "Editor", "Reviewer", "None")) {
                $comboBox3.Items.Add("Custom")
                $comboBox3.SelectedItem = "Custom"
            }
            else { $comboBox3.SelectedIndex = Convert-ComboTextToInt($cPermissions["Inbox"]) }
        }
        if ($cPermissions.ContainsKey("Contacts")) {
            if ($cPermissions["Contacts"] -notin @("Author", "Editor", "Reviewer", "None")) {
                $comboBox4.Items.Add("Custom")
                $comboBox4.SelectedItem = "Custom"
            }
            else { $comboBox4.SelectedIndex = Convert-ComboTextToInt($cPermissions["Contacts"]) }
        }
        if ($cPermissions.ContainsKey("Notes")) {
            if ($cPermissions["Notes"] -notin @("Author", "Editor", "Reviewer", "None")) {
                $comboBox5.Items.Add("Custom")
                $comboBox5.SelectedItem = "Custom"
            }
            else { $comboBox5.SelectedIndex = Convert-ComboTextToInt($cPermissions["Notes"]) }
        }
        if ($comboBox1.SelectedItem -notmatch 'Editor') {
            $checkBox1.Enabled = $false; $checkBox1.Checked = $false
            $checkBox3.Enabled = $false; $checkBox3.Checked = $false
            $checkBox4.Enabled = $false; $checkBox4.Checked = $false
        }
        else {
            if (!$checkBox1.Checked) {
                $checkBox3.Enabled = $false; $checkBox3.Checked = $false
                $checkBox4.Enabled = $false; $checkBox4.Checked = $false
            }
        }
        if ($cPermissions.ContainsKey("CalendarExt")) {
            $checkbox1.checked = $cPermissions["CalendarExt"].Contains("Delegate")
            $checkbox3.checked = $cPermissions["CalendarExt"].Contains("CanViewPrivateItems")
            $checkbox4.checked = $cPermissions["CalendarExt"].Contains("CanManageCategories")

        }
        if ($WhatIfPreference) { $OKButton.Enabled = $false }
    })

# KeyDown event handler
$form1.Add_KeyDown({
        if (($_.KeyCode -eq "Enter") -and ($form1.ActiveControl -eq $OKButton)) { display_GUI }
        if ($_.KeyCode -eq "Escape") { CancelForm }
        if ($_.KeyCode -eq "C" -and $_.modifiers -match "ALT") { $comboBox1.DroppedDown = $True; $form1.ActiveControl = $comboBox1 }
        if ($_.KeyCode -eq "T" -and $_.modifiers -match "ALT") { $comboBox2.DroppedDown = $True; $form1.ActiveControl = $comboBox2 }
        if ($_.KeyCode -eq "I" -and $_.modifiers -match "ALT") { $comboBox3.DroppedDown = $True; $form1.ActiveControl = $comboBox3 }
        if ($_.KeyCode -eq "O" -and $_.modifiers -match "ALT") { $comboBox4.DroppedDown = $True; $form1.ActiveControl = $comboBox4 }
        if ($_.KeyCode -eq "N" -and $_.modifiers -match "ALT") { $comboBox5.DroppedDown = $True; $form1.ActiveControl = $comboBox5 }
        if ($_.KeyCode -eq "D" -and $_.modifiers -match "ALT") {
            if ($checkBox1.Enabled) {
                if ($checkBox1.Checked) { $checkBox1.Checked = $false } else { $checkBox1.Checked = $true }
                $_.SuppressKeyPress = $true
            }
        }
        if ($_.KeyCode -eq "S" -and $_.modifiers -match "ALT") {
            if ($checkBox2.Enabled) {
                if ($checkBox2.Checked) { $checkBox2.Checked = $false } else { $checkBox2.Checked = $true }
                $_.SuppressKeyPress = $true
            }
        }
        if ($_.KeyCode -eq "P" -and $_.modifiers -match "ALT") {
            if ($checkBox3.Enabled) {
                if ($checkBox3.Checked) { $checkBox3.Checked = $false } else { $checkBox3.Checked = $true }
            }
            $_.SuppressKeyPress = $true
        }
        if ($_.KeyCode -eq "M" -and $_.modifiers -match "ALT") {
            if ($checkBox4.Enabled) {
                if ($checkBox4.Checked) { $checkBox4.Checked = $false } else { $checkBox4.Checked = $true }
            }
            $_.SuppressKeyPress = $true
        }
    })

# ComboBox1 SelectedIndexChanged event handler
$comboBox1.add_SelectedIndexChanged({
        if ($comboBox1.SelectedIndex -ne '3') {
            $checkBox1.Enabled = $false; $checkBox1.Checked = $false
            #CanViewPrivateItems is only allowed when the Delegate flag is set as well
            $checkBox3.Enabled = $false; $checkBox3.Checked = $false
            #Similarly, CanManageCategories is only allowed when Delegate flag is set
            $checkBox4.Enabled = $false; $checkBox4.Checked = $false
        }
        else { $checkBox1.Enabled = $true }
    })

#Checkbox1 state change handler
$checkBox1.add_CheckedChanged({
        if ($checkBox1.Checked -eq $false) {
            #CanViewPrivateItems is only allowed when the Delegate flag is set as well
            $checkBox3.Enabled = $false; $checkBox3.Checked = $false
            #Similarly, CanManageCategories is only allowed when Delegate flag is set
            $checkBox4.Enabled = $false; $checkBox4.Checked = $false
        }
        else {
            $checkBox3.Enabled = $true
            $checkBox4.Enabled = $true
        }
    })

#Cancel button Click event handler
$CancelButton.add_Click({ CancelForm })

#OK button click event handler
$OKButton.Add_Click({ display_GUI })

#Handle closing the form via the X button
$form1.add_FormClosing({
        if ($form1.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            $script:FormCancelled = $true
        }
    })
#endregion Event Handlers

#region Helper Functions
# Fetch the set of folders for the given mailbox, we cannot hardcode them due to localization
function Get-DefaultFolders {
    #Validation should be done prior to calling this function, only do mandatory flag here
    param(
        [Parameter(Mandatory = $true)][Alias("Manager")][string]$Owner
    )

    # Fetch default folders
    Write-Verbose "Fetching default folders for mailbox $Owner ..."
    try {
        if (!(Get-Command Get-MailboxFolderStatistics -ErrorAction SilentlyContinue)) {
            # Do we run in end-user scenario?
            Write-Verbose "Get-MailboxFolderStatistics cmdlet not available in the current session. Trying Get-MailboxFolder instead, should work for end-user scenarios."
            if (($Owner -ne $cUser) -and ($Owner -ne $cUser.Split("@")[0])) {
                # Not the current user
                Throw "Get-MailboxFolder cannot be used to retrieve other users' folder information. Please run the script as an administrator or with sufficient privileges."
            }
            $defaultFolders = Get-MailboxFolder -Identity $Owner -Recurse -ResultSize Unlimited -ErrorAction Stop -Verbose:$false | Where-Object { $_.DefaultFolderType -in @("Calendar", "Inbox", "Contacts", "Tasks", "Notes") } | Select-Object @{n = "Name"; e = { $_.DefaultFolderType } }, Identity
        }
        else {
            # No server-side filter, no property selection :(
            # Get-ExOMailboxFolderStatistics is freaking slow when run against non-existent values, we get no benefit from using it
            $defaultFolders = Get-MailboxFolderStatistics -Identity $Owner -ErrorAction Stop -Verbose:$false -WarningAction SilentlyContinue | Where-Object { $_.FolderType -in @("Calendar", "Inbox", "Contacts", "Tasks", "Notes") } | Select-Object @{n = "Name"; e = { $_.FolderType } }, @{n = "Identity"; e = { ($_.Identity.ToString().replace('\', ':\')) } }
        }
    }
    catch {
        Write-ConsoleLog -Level ERROR -Message "Failed to query default folders for mailbox $Owner. No permission changes will be attempted. $($_.Exception.Message)"
        throw
    }

    $defaultFoldersHash = @{}
    $defaultFolders | foreach-object { $defaultFoldersHash[$_.Name] = $_.Identity }

    $requiredFolders = @("Calendar", "Inbox", "Contacts", "Tasks", "Notes")
    $missingFolders = @($requiredFolders | Where-Object { -not $defaultFoldersHash.ContainsKey($_) -or [string]::IsNullOrWhiteSpace([string]$defaultFoldersHash[$_]) })
    if ($missingFolders.Count -gt 0) {
        $message = "Exchange returned an incomplete default-folder inventory for '$Owner'. Missing: $($missingFolders -join ', '). No permission changes will be attempted."
        Write-ConsoleLog -Level ERROR -Message $message
        throw $message
    }

    return $defaultFoldersHash
}

# Fetch the standard set of folder permissions
function Get-StandardFolderPermissions {
    #Validation should be done prior to calling this function, only do mandatory flag here
    param(
        [Parameter(Mandatory = $true)][Alias("Manager")][string]$Owner,
        [Parameter(Mandatory = $true)][string]$Delegate
    )

    # Get the set of folders
    $script:cFolders = Get-DefaultFolders -Owner $Owner
    $folderPermissions = @{}

    # Fetch permissions for each folder
    foreach ($folder in $cFolders.GetEnumerator()) {
        Write-Verbose "Fetching permissions for folder $($folder.Value) ..."
        try {
            #No point using Get-ExOMailboxFolderPermission, it's terribly slow when erroring out
            $varPermissions = Get-MailboxFolderPermission -Identity $folder.Value.ToString() -User $Delegate -ErrorAction Stop -Verbose:$false
            if ($folder.Key -eq "Calendar") {
                if ($varPermissions.SharingPermissionFlags) {
                    $folderPermissions["CalendarExt"] = @($varPermissions.SharingPermissionFlags -split ", ").Trim() ?? @()
                }
                else { $folderPermissions["CalendarExt"] = @() } # Covered below, remove?
            }
            $folderPermissions[$($folder.Name)] = ($varPermissions.AccessRights)[0] ?? "None"
        }
        catch {
            if ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { throw "Folder ""$($folder.Value)"" not found, this should not happen..." }
            $remoteException = $_.Exception.Data['RemoteException']
            $remoteExceptionType = if ($null -ne $remoteException) { [string]$remoteException.TypeName } else { "" }
            if ($_.CategoryInfo.Reason -eq "UserNotFoundInPermissionEntryException" -or $remoteExceptionType -match '(^|\.)UserNotFoundInPermissionEntryException$') {
                $folderPermissions[$($folder.Name)] = "None"
                continue
            }
            Write-ConsoleLog -Level ERROR -Message "Failed to read permissions for $($folder.Value): $($_.Exception.Message)"
            throw
        }
    }

    if (!$folderPermissions.ContainsKey("CalendarExt")) {
        $folderPermissions["CalendarExt"] = @()
    }
    $folderPermissions["Owner"] = $Owner
    $folderPermissions["Delegate"] = $Delegate
    return $folderPermissions
}

# Get the Form values and close the form
function display_GUI {
    $script:cPermissionsNew = @{}
    $cPermissionsNew["Calendar"] = $combobox1.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Tasks"] = $combobox2.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Inbox"] = $combobox3.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Contacts"] = $combobox4.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Notes"] = $combobox5.SelectedItem.Split(" ")[0]
    $cPermissionsNew["CalendarExt"] = @()
    #SharingPermissionFlags parameter only applies to calendar folders and can only be used when the AccessRights parameter value is Editor.
    if ($checkbox1.Checked) { $cPermissionsNew["CalendarExt"] += "Delegate" }
    #CanViewPrivateItems is only allowed when the Delegate flag is set as well
    if ($checkbox3.Checked -and ($cPermissionsNew["CalendarExt"].Contains("Delegate"))) { $cPermissionsNew["CalendarExt"] += "CanViewPrivateItems" }
    #Similarly, CanManageCategories is only allowed when Delegate flag is set
    if ($checkbox4.Checked -and ($cPermissionsNew["CalendarExt"].Contains("Delegate"))) { $cPermissionsNew["CalendarExt"] += "CanManageCategories" }
    #SendNotificationToUser parameter only applies to calendar folders and can only be used with the following AccessRights parameter values: AvailabilityOnly, LimitedDetails, Reviewer, Editor
    if ($checkBox2.Checked) { $cPermissionsNew["SendSummary"] = $true }
    else { $cPermissionsNew["SendSummary"] = $false }

    $form1.Close() | Out-Null
}

# Close the form and mark it as cancelled
function CancelForm {
    $form1.Close() | Out-Null
    $script:FormCancelled = $true
}

# Process any permission changes as needed
function ApplyMBXPermissionChanges {
    # Validaitons should be done prior to calling this function

    # Check for changes in SharingPermissionFlags as they need special handling
    if (Compare-Object -ReferenceObject $cPermissions["CalendarExt"] -DifferenceObject $cPermissionsNew["CalendarExt"]) {
        Write-Verbose "Detected change in permission for SharingPermissionFlags, to be addressed as part of Calendar folder processing..."
        $script:cFlags = $true
        #Empty array is not accepted, use "None" instead
        if ($cPermissionsNew["CalendarExt"].Count -eq 0) { $script:cFlagsValue = "None" } else { $script:cFlagsValue = $cPermissionsNew["CalendarExt"] }
    }
    else { $script:cFlags = $false }
    $calendarFlags = if ($cFlags) { $cFlagsValue } else { "None" }

    #Process changes
    foreach ($perm in $cPermissionsNew.GetEnumerator()) {
        # Skip SendNotificationToUser, we don't have an initial value for it, but we respect the new value below
        # Also skip CalendarExt here, we handle SharingPermissionFlags separately
        if ($perm.Key -in @("Owner", "Delegate", "SendSummary", "CalendarExt")) { continue }

        # Skip custom permissions
        if ($perm.Value -eq "Custom") {
            Write-ConsoleLog -Level SKIP -Message "Custom permissions are not supported for $($perm.Key); no change was made."
            continue
        }

        if (-not $cFolders.ContainsKey($perm.Key) -or [string]::IsNullOrWhiteSpace([string]$cFolders[$perm.Key])) {
            Write-ConsoleLog -Level ERROR -Message "Skipping $($perm.Key): no folder identity was found for mailbox $Owner."
            continue
        }

        # Skip unchanged permissions, except for Calendar sharing flags.
        if ($cPermissions.ContainsKey($perm.Key) -and $cPermissions[$perm.Key] -eq $perm.Value) {
            if (($perm.Key -eq "Calendar") -and ($perm.Value -eq "Editor") -and $cFlags) {
                Write-ConsoleLog -Level CHANGE -Message "Updating Calendar sharing flags for $Delegate."
                try {
                    Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -SharingPermissionFlags $cFlagsValue -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                    Write-ConsoleLog -Level SUCCESS -Message "Updated Calendar sharing flags."
                }
                catch {
                    Write-ConsoleLog -Level ERROR -Message "Failed to update Calendar sharing flags: $($_.Exception.Message)"
                    continue
                }
            }
            else {
                Write-ConsoleLog -Level SKIP -Message "No change to $($perm.Key) permission."
            }
            continue
        }

        # Add
        if (($null -eq $cPermissions[$perm.Key]) -or ($cPermissions[$perm.Key] -eq "None")) {
            Write-ConsoleLog -Level CHANGE -Message "Adding $($perm.Value) permission on $($perm.Key)."
            try {
                # Calendar needs special handling
                if ($perm.Key -eq "Calendar") {
                    # Changes to the two SharingPermissionFlags should be processed for Calendar folder with Editor access right and modified SharingPermissionFlags
                    if ($perm.Value -eq "Editor") {
                        Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -SharingPermissionFlags $calendarFlags -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                    }
                    # Else just process SendNotificationToUser. We don't support LimitedDetails and AvailabilityOnly currently, but include them regardless
                    elseif ($perm.Value -in @("AvailabilityOnly", "LimitedDetails", "Reviewer")) {
                        Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                    }
                    # Else just add the access right
                    else { Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference }
                }
                # For non-calendar folders
                else {
                    Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                }
                Write-ConsoleLog -Level SUCCESS -Message "Added $($perm.Value) permission on $($perm.Key)."
            }
            catch {
                Write-ConsoleLog -Level ERROR -Message "Failed to add permission on $($perm.Key): $($_.Exception.Message)"
            }
            continue
        }

        # Remove
        if ($perm.Value -eq "None") {
            Write-ConsoleLog -Level CHANGE -Message "Removing permission on $($perm.Key)."
            try {
                # Process SendNotificationToUser only for Calendar folder
                if ($perm.Key -eq "Calendar") {
                    Remove-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate.Trim() -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                }
                else {
                    Remove-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate.Trim() -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                }
                Write-ConsoleLog -Level SUCCESS -Message "Removed permission on $($perm.Key)."
            }
            catch {
                Write-ConsoleLog -Level ERROR -Message "Failed to remove permission on $($perm.Key): $($_.Exception.Message)"
            }
            continue
        }

        # Set
        Write-ConsoleLog -Level CHANGE -Message "Setting $($perm.Key) permission to $($perm.Value)."
        try {
            # Calendar needs special handling
            if ($perm.Key -eq "Calendar") {
                # Changes to the two SharingPermissionFlags should be processed for Calendar folder with Editor access right
                if ($perm.Value -eq "Editor") {
                    Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -SharingPermissionFlags $calendarFlags -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                }
                # Process remaining access rights that support SendNotificationToUser. We don't support LimitedDetails and AvailabilityOnly currently, but include them regardless
                elseif ($perm.Value -in @("AvailabilityOnly", "LimitedDetails", "Reviewer")) {
                    Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                }
                # Else just set the access right
                else { Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference }
            }
            # For non
            else {
                Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
            }
            Write-ConsoleLog -Level SUCCESS -Message "Set $($perm.Key) permission to $($perm.Value)."
        }
        catch {
            Write-ConsoleLog -Level ERROR -Message "Failed to update $($perm.Key) permission: $($_.Exception.Message)"
        }
        continue
    }
}

# Convert combobox text to integer index
function Convert-ComboTextToInt($role) {
    switch -wildcard ($role) {
        'Reviewer*' { "1" }
        'Author*' { "2" }
        'Editor*' { "3" }
        'None' { "0" }
    }
}

function GenerateForm {
    # Validate all the values
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][Alias("Manager")][string]$Owner,
        [Parameter(Mandatory)][string]$Delegate,
        [switch]$ReadOnly = $false,
        [string]$RestoreFrom
    )

    if ($ReadOnly -and $RestoreFrom) {
        throw "-ReadOnly cannot be combined with -RestoreFrom."
    }

    # Query Exchange for mailbox identities and status before reading or changing permissions.
    $script:mailboxStatus = @(
        Get-MailboxStatus -Identity $Owner -Role Mailbox
        Get-MailboxStatus -Identity $Delegate -Role Delegate
    )
    Write-MailboxStatus -Title "EXCHANGE ONLINE: current mailbox status" -Status $mailboxStatus

    # Fetch current permissions
    $script:cPermissions = Get-StandardFolderPermissions -Owner $Owner -Delegate $Delegate
    Write-PermissionSummary -Title "BEFORE: current mailbox permissions" -Permissions $cPermissions

    if ($RestoreFrom) {
        $backup = Read-PermissionBackup -Path $RestoreFrom -CurrentMailboxStatus $mailboxStatus
        $script:cPermissionsNew = @{
            Owner       = $Owner
            Delegate    = $Delegate
            Calendar    = $backup.Permissions.Calendar
            Tasks       = $backup.Permissions.Tasks
            Inbox       = $backup.Permissions.Inbox
            Contacts    = $backup.Permissions.Contacts
            Notes       = $backup.Permissions.Notes
            CalendarExt = @($backup.Permissions.CalendarExt)
            SendSummary = $false
        }
        Write-ConsoleLog -Level INFO -Message "Loaded and validated permission backup created at $($backup.CreatedAtUtc)."
        Write-PermissionSummary -Title "RESTORE TARGET: permissions from backup" -Permissions $cPermissionsNew
        return
    }

    # Set whatIfPreference based on ReadOnly switch
    $script:WhatIfPreference = $ReadOnly.IsPresent

    # Show the form
    $form1.Text = "Delegate Permissions for $($cPermissions["Delegate"]) on $($cPermissions["Owner"])"
    $form1.ShowDialog() | Out-Null
}
#endregion Helper Functions

#Call the form
GenerateForm -Owner $Owner -Delegate $Delegate -ReadOnly:$ReadOnly -RestoreFrom $RestoreFrom
if ($FormCancelled) {
    Write-ConsoleLog -Level SKIP -Message "Permission changes cancelled; mailbox was not modified."
    return
}
if (!($cPermissionsNew) -or !($cPermissions)) {
    Throw "‼️Cannot apply permission changes, missing current or new permissions data..."
}
else {
    $backupPath = Save-PreChangePermissionBackup -Permissions $cPermissions -MailboxStatus $mailboxStatus
    if (-not $backupPath) {
        return
    }

    Write-ConsoleLog -Level INFO -Message "Applying selected mailbox permission changes."
    ApplyMBXPermissionChanges
    Write-ConsoleLog -Level INFO -Message "Re-querying mailbox status and folder permissions from Exchange Online."
    $afterMailboxStatus = @(
        Get-MailboxStatus -Identity $Owner -Role Mailbox
        Get-MailboxStatus -Identity $Delegate -Role Delegate
    )
    Write-MailboxStatus -Title "EXCHANGE ONLINE: current mailbox status after processing" -Status $afterMailboxStatus
    foreach ($mailbox in $afterMailboxStatus) {
        $beforeMailbox = $mailboxStatus | Where-Object Role -eq $mailbox.Role | Select-Object -First 1
        if ($null -eq $beforeMailbox -or $beforeMailbox.Guid -ne $mailbox.Guid) {
            $message = "Mailbox identity verification failed for role '$($mailbox.Role)'. Do not retry until the mailbox identities have been reviewed."
            Write-ConsoleLog -Level ERROR -Message $message
            throw $message
        }
    }

    $afterPermissions = Get-StandardFolderPermissions -Owner $Owner -Delegate $Delegate
    Write-PermissionSummary -Title "AFTER: verified mailbox permissions" -Permissions $afterPermissions
    $verification = Test-PermissionSnapshot -Expected $cPermissionsNew -Actual $afterPermissions
    if ($verification.Succeeded) {
        Write-ConsoleLog -Level SUCCESS -Message "Exchange reports all requested folder permissions and Calendar delegate flags are in place."
    }
    else {
        foreach ($mismatch in $verification.Mismatches) {
            Write-ConsoleLog -Level ERROR -Message "Verification mismatch: $mismatch"
        }
        throw "Mailbox permission verification failed. Review the current Exchange state above before retrying."
    }
}

Write-ConsoleLog -Level SUCCESS -Message "Mailbox permission review is complete. Pre-change backup: $backupPath"

