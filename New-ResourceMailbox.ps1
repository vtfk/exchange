[CmdletBinding(SupportsShouldProcess)]
param(
    <# New-Mailbox #>
    [Parameter(Mandatory = $True)]
    [string]$DisplayName,

    [Parameter(Mandatory = $True)]
    [string]$PrimarySmtpAddress,

    [Parameter(Mandatory = $True)]
    [ValidateSet("Equipment", "Room", "Teams")]
    [string]$ResourceType,

    [Parameter()]
    [string]$RoomMailboxPassword,

    <# Set-Place #>
    [Parameter()]
    [int]$Capacity = -1,

    [Parameter()]
    [string]$City,

    [Parameter()]
    [string]$Building,

    [Parameter()]
    [int]$Floor = -1,

    [Parameter()]
    [ValidateScript({ $_ -like "*,*;*,*" })]
    [string]$GeoCoordinates,<#

    [Parameter()]
    [bool]$IsWheelChairAccessible,#>

    [Parameter()]
    [string]$PostalCode,

    [Parameter()]
    [string]$State,

    [Parameter()]
    [string]$Street,
    
    [Parameter()]
    [string[]]$Tags,

    [Parameter()]
    [string]$AudioDeviceName,

    [Parameter()]
    [string]$VideoDeviceName,

    [Parameter()]
    [string]$DisplayDeviceName,

    <# Set-CalendarProcessing | Set-MailboxFolderPermission #>
    [Parameter()]
    [string[]]$Booking,

    [Parameter()]
    [ValidateSet("None", "AutoUpdate", "AutoAccept")]
    [string]$AutomateProcessing = "AutoAccept",

    <# Add-MailboxPermission #>
    [Parameter()]
    [string[]]$FullAccess,

    <# Set-MsolUser #>
    [Parameter()]
    [string]$UsageLocation = "NO",

    <# Set-MsolUserLicense #>
    [Parameter()]
    [ValidateScript({ $_ -like "*:*" })]
    [string[]]$Licenses,

    ## Only used when called from Acos skjema-script
    # For logging purposes and mail delivery
    [Parameter()]
    [ValidateScript({ $_ -like "*@*.no" })]
    [string]$Creator,
    
    ## Only used when called from Acos skjema-script
    # Required to do a silent run
    [Parameter()]
    [string]$Target
)

# import environment variables
$envPath = Join-Path -Path $PSScriptRoot -ChildPath ".\envs.ps1"
. $envPath

# import Tools
. "$PSScriptRoot\Tools\ConvertTo-List.ps1"

Add-LogTarget -Name CMTrace

Write-Log -Message "#######################################################################################################"

if ($ResourceType -eq "Equipment" -and ($Capacity -gt -1 -or ![string]::IsNullOrEmpty($City) -or ![string]::IsNullOrEmpty($Building) -or $Floor -gt -1 -or ![string]::IsNullOrEmpty($GeoCoordinates) -or ![string]::IsNullOrEmpty($PostalCode) -or ![string]::IsNullOrEmpty($State) -or ![string]::IsNullOrEmpty($Street) -or ($null -ne $Tags -and $Tags.Count -gt 0) -or ![string]::IsNullOrEmpty($AudioDeviceName) -or ![string]::IsNullOrEmpty($VideoDeviceName) -or ![string]::IsNullOrEmpty($DisplayDeviceName)))
{
    Write-Log -Message "Equipment doesn't support 'Set-Place' arguments. Choose Room to add 'Set-Place' arguments" -Level ERROR
    Write-Error "Equipment doesn't support 'Set-Place' arguments. Choose Room to add 'Set-Place' arguments" -ErrorAction Stop
}

if ($ResourceType -eq "Teams" -and ([string]::IsNullOrEmpty($UsageLocation) -or $Licenses.Count -le 0 -or [string]::IsNullOrEmpty($RoomMailboxPassword)))
{
    Write-Log -Message "Teams meeting rooms requires parameters 'UsageLocation', 'RoomMailboxPassword' and 'AddLicenses' (licenses is retrieved from `"Get-MsolAccountSku | Where { $_.AccountSkuId -like '*MEETING_ROOM' } | Select -ExpandProperty AccountSkuId`")" -Level ERROR
    Write-Error "Teams meeting rooms requires parameters 'UsageLocation', 'RoomMailboxPassword' and 'AddLicenses' (licenses is retrieved from 'Get-MsolAccountSku | Sort AccountSkuId')" -ErrorAction Stop
}

if ($Creator) {
    Add-LogTarget -Name Email -Configuration @{
        SMTPServer = $smtpServer
        From = $smtpFrom
        To = $Creator
        Level = "SUCCESS"
        BodyAsHtml = $True
    }

    if ($Target) {
        if ($PSCmdlet.ShouldProcess((@{Exchange=$True;MSOnline=$True;Target=$Target} | ConvertTo-Json), "Connect-Office365")) {
            Connect-Office365 -Exchange -MSOnline -Target $Target
        }
    }

    # error message for mail body
    [string]$mailMsgError = "<p><b>Feilmelding</b>:<br>%errormsg%</p>`n<p>Trenger du hjelp? Ta kontakt med <a href=`"$servicedeskMail`">$servicedeskMail</a> for feilsøking."
}

<# make sure PrimarySmtpAddress does not already exist
if ($PSCmdlet.ShouldProcess($PrimarySmtpAddress, "Get-Recipient"))
{
    $mailRecipient = Get-Recipient -Identity $PrimarySmtpAddress -ErrorAction SilentlyContinue
    if ($mailRecipient)
    {
        Write-Log -Message "PrimarySmtpAddress '$PrimarySmtpAddress' already exists: ($($mailRecipient.DisplayName))" -Level ERROR
        Write-Error "PrimarySmtpAddress '$PrimarySmtpAddress' already exists: ($($mailRecipient.DisplayName))" -ErrorAction Stop
    }
}#>

# make alias from PrimarySmtpAddress
$alias = $PrimarySmtpAddress.Split('@')[0]

<# make sure alias does not already exist
if ($PSCmdlet.ShouldProcess($alias, "Get-Recipient"))
{
    $aliasRecipient = Get-Recipient -Identity $alias -ErrorAction SilentlyContinue
    if ($aliasRecipient)
    {
        Write-Log -Message "Alias '$alias' already exists: ($($aliasRecipient.DisplayName))" -Level ERROR
        Write-Error "Alias '$alias' already exists: ($($aliasRecipient.DisplayName))" -ErrorAction Stop
    }
}#>

# if any other FullAccess users is added, add these here as well
if ($FullAccess)
{
    if (!$calFullAccess)
    {
        [string[]]$calFullAccess = @()
    }

    $FullAccess | % { $calFullAccess += $_ }
}

# place splat
$placeSplat = @{
    Identity = $PrimarySmtpAddress
}
if (![string]::IsNullOrEmpty($Building)) { $placeSplat.Add("Building", $Building) }
if ($Capacity -gt -1) { $placeSplat.Add("Capacity", $Capacity) }
if (![string]::IsNullOrEmpty($City)) { $placeSplat.Add("City", $City) }
if ($Floor -gt -1) { $placeSplat.Add("Floor", $Floor) }
if (![string]::IsNullOrEmpty($GeoCoordinates)) { $placeSplat.Add("GeoCoordinates", $GeoCoordinates) }
if (![string]::IsNullOrEmpty($PostalCode)) { $placeSplat.Add("PostalCode", $PostalCode) }
if (![string]::IsNullOrEmpty($State)) { $placeSplat.Add("State", $State) }
if (![string]::IsNullOrEmpty($Street)) { $placeSplat.Add("Street", $Street) }
if (($null -ne $Tags -and $Tags.Count -gt 0)) { $placeSplat.Add("Tags", $Tags) }
if (![string]::IsNullOrEmpty($AudioDeviceName)) { $placeSplat.Add("AudioDeviceName", $AudioDeviceName) }
if (![string]::IsNullOrEmpty($VideoDeviceName)) { $placeSplat.Add("VideoDeviceName", $VideoDeviceName) }
if (![string]::IsNullOrEmpty($DisplayDeviceName)) { $placeSplat.Add("DisplayDeviceName", $DisplayDeviceName) }

# mailbox splat
$mailboxSplat = @{
    DisplayName = $DisplayName
    Alias = $alias
    PrimarySmtpAddress = $PrimarySmtpAddress
    Name = $DisplayName
}

# add resource type to splat
if ($ResourceType -eq "Equipment")
{
    $mailboxSplat.Add("Equipment", $true)
}
elseif ($ResourceType -eq "Room")
{
    $mailboxSplat.Add("Room", $true)
}
elseif ($ResourceType -eq "Teams")
{
    $mailboxSplat.Add("Room", $true)
    $mailboxSplat.Add("RoomMailboxPassword", (ConvertTo-SecureString -String $RoomMailboxPassword -AsPlainText -Force))
    $mailboxSplat.Add("EnableRoomMailboxAccount", $true)
}

# calendarProcessing splat
$calendarProcessingSplat = @{
    Identity = $PrimarySmtpAddress
    DeleteComments = $False
    DeleteSubject = $False
    AddOrganizerToSubject = $False
    AutomateProcessing = $AutomateProcessing
}

if ($Booking -and "Alle" -notin $Booking)
{
    $calendarProcessingSplat.Add("BookInPolicy", $Booking)
    $calendarProcessingSplat.Add("AllBookInPolicy", $False)
}

# mailboxFolderPermission splat
$mailboxFolderPermissionSplat = @{
    User = "Default"
    AccessRight = "Reviewer"
}

# teamsMsolUser splat
$teamsMsolUserSplat = @{
    UserPrincipalName = $PrimarySmtpAddress
    PasswordNeverExpires = $True
}

# teamsMsolUserLicense splat
$teamsMsolUserLicenseSplat = @{
    UserPrincipalName = $PrimarySmtpAddress
}

if ($ResourceType -eq "Teams")
{
    # change FolderPermission for Default user to NonEditingAuthor
    $mailboxFolderPermissionSplat["AccessRight"] = "NonEditingAuthor"

    # set Msol User
    $teamsMsolUserSplat.Add("UsageLocation", $UsageLocation)

    # set Msol User Licenses
    $teamsMsolUserLicenseSplat.Add("AddLicenses", $Licenses)
}

# create mailbox
if ($PSCmdlet.ShouldProcess(($mailboxSplat | ConvertTo-Json), "New-Mailbox"))
{
    try
    {
        $mbx = New-Mailbox @mailboxSplat -ErrorAction Stop
        Write-Verbose "New mailbox '$($mbx.Name)' created"
        Write-Log -Message "New resource mailbox created" -Body $mailboxSplat

        <#if (!$mbx.Name)
        {
            throw "New-Mailbox got no Name in return : $(($mailboxSplat | ConvertTo-Json))"
        }#>
    }
    catch
    {
        if (!$Creator) {
            Write-Log -Message "Failed to create new resource mailbox: $_" -Body $mailboxSplat -Level ERROR
        }
        else {
            Write-Log -Message "Feilet ved opprettelse av ressurs mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
        }

        Write-Error $_ -ErrorAction Stop
    }
}

# set default CalendarProcessing
if ($PSCmdlet.ShouldProcess(($calendarProcessingSplat | ConvertTo-Json), "Set-CalendarProcessing"))
{
    Write-Verbose "Waiting 30 seconds before updating calendar processing on mailbox"
    Start-Sleep -Seconds 30
    
    try
    {
        $currentAutomateProcessing = Get-CalendarProcessing -Identity $calendarProcessingSplat.Identity | Select -ExpandProperty AutomateProcessing
        if ($currentAutomateProcessing -ne "AutoAccept")
        {
            Write-Warning "'$PrimarySmtpAddress' is not set to default AutomateProcessing value 'AutoAccpet' but has the value '$currentAutomateProcessing' ...."
            if (!$Creator) {
                Write-Log -Message "'$PrimarySmtpAddress' is not set to default AutomateProcessing value 'AutoAccpet' but has the value '$currentAutomateProcessing' ...." -Level WARNING
            }
            else {
                Write-Log -Message "'$PrimarySmtpAddress' is not set to default AutomateProcessing value 'AutoAccpet' but has the value '$currentAutomateProcessing' ...." -Level INFO
            }
        }

        Set-CalendarProcessing @calendarProcessingSplat -ErrorAction Stop
        Write-Log -Message "Calendar processing set on resource mailbox" -Body $calendarProcessingSplat
    }
    catch
    {
        if (!$Creator) {
            Write-Log -Message "Failed to set calendar processing on resource mailbox: $_" -Body $calendarProcessingSplat -Level ERROR
        }
        else {
            Write-Log -Message "Feilet ved setting av ekstra innstillinger på ressurs mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; CalendarProcessing = $calendarProcessingSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
        }

        Write-Error $_ -ErrorAction Stop
    }
}

# set Default folder permissions to Reviewer
if ($PSCmdlet.ShouldProcess(($mailboxFolderPermissionSplat | ConvertTo-Json), "Set-MailboxFolderPermission"))
{
    Write-Verbose "Waiting 30 seconds before updating folder permission on mailbox"
    Start-Sleep -Seconds 30
    
    try
    {
        $identity = $mbx.UserPrincipalName
        $calendarName = (Get-MailboxFolderStatistics -Identity $identity -FolderScope Calendar | Select-Object -First 1).Name
        $folderPermissionName = "$($identity):\$calendarName"
        $mailboxFolderPermissionSplat.Add("identity", $folderPermissionName)

        Set-MailboxFolderPermission @mailboxFolderPermissionSplat -ErrorAction Stop
        Write-Log -Message "Default folder permission set on resource mailbox" -Body $mailboxFolderPermissionSplat
    }
    catch
    {
        if (!$Creator) {
            Write-Log -Message "Failed to set default folder permission on resource mailbox: $_" -Body $mailboxFolderPermissionSplat -Level ERROR
        }
        else {
            Write-Log -Message "Feilet ved setting av innstillinger på ressurs mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; CalendarProcessing = $calendarProcessingSplat; FolderPermissions = $mailboxFolderPermissionSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
        }

        Write-Error $_ -ErrorAction Stop
    }
}

# add FullAccess groups to all
if ($calFullAccess.Count -gt 0)
{
    foreach ($calBooker in $calFullAccess)
    {
        $calMailboxPermission = @{
            User = $calBooker
            AccessRight = "FullAccess"
        }

        if ($null -eq $identity) {
            $calMailboxPermission.Add("Identity", $PrimarySmtpAddress)
        }
        else {
            $calMailboxPermission.Add("Identity", $identity)
        }

        if ($PSCmdlet.ShouldProcess(($calMailboxPermission | ConvertTo-Json), "Add-MailboxPermission"))
        {
            try
            {
                Add-MailboxPermission @calMailboxPermission -ErrorAction Stop | Out-Null
                Write-Log -Message "FullAccess permission added to resource mailbox" -Body $calMailboxPermission
            }
            catch
            {
                if (!$Creator) {
                    Write-Log -Message "Failed to add FullAccess permission to resource mailbox: $_" -Body $calMailboxPermission -Level ERROR
                }
                else {
                    Write-Log -Message "Feilet ved FullAccess tilgang på ressurs mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; CalendarProcessing = $calendarProcessingSplat; FolderPermissions = $mailboxFolderPermissionSplat; FullAccessUser = $calMailboxPermission }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
                }
                
                Write-Error $_ -ErrorAction Stop
            }
        }
    }
}

# add Booking addresses to BookInPolicy and change theire accessrights to Reviewer, set AllBookInPolicy to $False
if ($Booking -and "Alle" -notin $Booking)
{
    # change accessrights to Reviewer
    foreach ($booker in $Booking)
    {
        $bookingSplat = @{
            Identity = $folderPermissionName
            User = $booker
            AccessRight = "Reviewer"
        }
        if ($PSCmdlet.ShouldProcess(($bookingSplat | ConvertTo-Json), "Set-MailboxFolderPermission"))
        {
            try
            {
                Set-MailboxFolderPermission @bookingSplat -ErrorAction Stop
                Write-Log -Message "Folder permission added to resource mailbox" -Body $bookingSplat
            }
            catch
            {
                if (!$Creator) {
                    Write-Log -Message "Failed to add folder permission to resource mailbox: $_" -Body $bookingSplat -Level ERROR
                }
                else {
                    Write-Log -Message "Feilet ved booking tilgang på ressurs mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; CalendarProcessing = $calendarProcessingSplat; FolderPermissions = $mailboxFolderPermissionSplat; BookingUser = $bookingSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
                }

                Write-Error $_ -ErrorAction Stop
            }
        }
    }
}

# set place info
if ($placeSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($placeSplat | ConvertTo-Json), "Set-Place"))
    {
        Write-Host "Waiting 5 minutes before adding Place info"
        Start-Sleep -Seconds (5*60)

        try
        {
            Set-Place @placeSplat -ErrorAction Stop
            Write-Log -Message "Place info set on resource mailbox" -Body $placeSplat
        }
        catch
        {
            if ($_.ToString() -like "*Encountered an internal server error*")
            {
                if (!$Creator) {
                    Write-Log -Message "Place info set on resource mailbox. Although an exception were thrown: $_" -Body $placeSplat -Level WARNING
                }
                else {
                    Write-Log -Message "Place info set on resource mailbox. Although an exception were thrown: $_" -Body @{ Subject = "Feil: Exchange skjemaparser"; Generelt = $mailboxSplat; CalendarProcessing = $calendarProcessingSplat; FolderPermissions = $mailboxFolderPermissionSplat; PlaceInfo = $placeSplat; "Hva skjedde"= $mailMsgError.Replace("%errormsg%", $_) } -Level WARNING
                }

                Write-Warning -Message "$_"
            }
            else
            {
                if (!$Creator) {
                    Write-Log -Message "Failed to set place info on resource mailbox: $_" -Body $placeSplat -Level ERROR
                }
                else {
                    Write-Log -Message "Feilet ved setting av place info på ressurs mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; CalendarProcessing = $calendarProcessingSplat; FolderPermissions = $mailboxFolderPermissionSplat; PlaceInfo = $placeSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
                }

                Write-Error $_ -ErrorAction Stop
            }
        }
    }
}

if ($ResourceType -eq "Teams")
{
    # set Teams Msol User
    if ($PSCmdlet.ShouldProcess(($teamsMsolUserSplat | ConvertTo-Json), "Set-MsolUser"))
    {
        try
        {
            Set-MsolUser @teamsMsolUserSplat -ErrorAction Stop
            Write-Verbose "Teams meeting room user updated"
            Write-Log -Message "Teams meeting room user update" -Body $teamsMsolUserSplat
        }
        catch
        {
            if (!$Creator) {
                Write-Log -Message "Failed to update Teams meeting room user: $_" -Body $teamsMsolUserSplat -Level ERROR
            }
            else {
                Write-Log -Message "Feilet ved endring av Teams møteromsbruker!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; TeamsMsolUser = $teamsMsolUserSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
            }

            Write-Error $_ -ErrorAction Stop
        }
    }

    # set Teams Msol User License
    if ($PSCmdlet.ShouldProcess(($teamsMsolUserLicenseSplat | ConvertTo-Json), "Set-MsolUserLicense"))
    {
        try
        {
            Set-MsolUserLicense @teamsMsolUserLicenseSplat -ErrorAction Stop
            Write-Verbose "Teams meeting room licenses updated"
            Write-Log -Message "Teams meeting room licenses update" -Body $teamsMsolUserLicenseSplat
        }
        catch
        {
            if (!$Creator) {
                Write-Log -Message "Failed to update Teams meeting room licenses: $_" -Body $teamsMsolUserLicenseSplat -Level ERROR
            }
            else {
                Write-Log -Message "Feilet ved endring av Teams møteromslisenser!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; TeamsMsolUser = $teamsMsolUserSplat; TeamsMsolUserLicense = $teamsMsolUserLicenseSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
            }

            Write-Error $_ -ErrorAction Stop
        }
    }
}

# send success mail if Creator is filled out and everything has worked fine!
if ($Creator) {
    $mailSplat = [ordered]@{
        Generelt = $mailboxSplat
        CalendarProcessing = $calendarProcessingSplat
        FolderPermissions = $mailboxFolderPermissionSplat
    }
    if ($Booking) { $mailSplat.Add("Booking", $Booking) }
    if ($FullAccess) { $mailSplat.Add("FullAccess", $FullAccess) }

    if ($PSCmdlet.ShouldProcess(($mailSplat | ConvertTo-Json), "Write-Log")) {
        Write-Log -Message "Ny ressurs mailbox opprettet<br><br>$(ConvertTo-List -Hash $mailSplat)" -Body @{ Subject = "Exchange skjemaparser" } -Level SUCCESS
    }
}