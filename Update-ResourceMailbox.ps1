### Set-Mailbox | Set-CalendarProcessing
# Identity can be Name, Alias, DN, CanonicalName, EmailAddress, GUID, LegacyExchangeDN, SamAccountName, UPN
#
### Set-User
# Identity can be Name, DN, CanonicalName, GUID, UPN
#
[CmdletBinding(SupportsShouldProcess)]
param(
    <# Set-Mailbox | Set-User | Set-CalendarProcessing #>
    [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
    [string]$Identity,

    <# Set-Mailbox #>
    [Parameter()]
    [string]$Alias,
    
    [Parameter()]
    [string]$DisplayName,

    [Parameter()]
    [bool]$HiddenFromAddressLists,

    [Parameter()]
    [string]$MailTip,

    [Parameter()]
    [string]$Office,

    [Parameter()]
    [string[]]$EmailAddressesAdd,
    
    [Parameter()]
    [string[]]$EmailAddressesRemove,

    <# Set-Mailbox | Set-Place #>
    [Parameter()]
    [int]$ResourceCapacity = -1,

    <# Set-Mailbox | Set-MsolUserPrincipalName #>
    [Parameter()]
    [string]$UserPrincipalName,

    <# Set-User #>
    [Parameter()]
    [string]$Company,

    [Parameter()]
    [string]$Phone,

    <# Set-User | Set-Place #>
    [Parameter()]
    [string]$StreetAddress,

    [Parameter()]
    [string]$PostalCode,

    [Parameter()]
    [string]$City,

    <# Set-Place #>
    [Parameter()]
    [string]$Building,

    [Parameter()]
    [string]$CountryOrRegion,

    [Parameter()]
    [int]$Floor = -1,

    [Parameter()]
    [ValidateScript({ $_ -like "*,*;*,*" })]
    [string]$GeoCoordinates,

    [Parameter()]
    [bool]$IsWheelChairAccessible,

    [Parameter()]
    [string]$State,
    
    [Parameter()]
    [string[]]$Tags,

    [Parameter()]
    [string]$AudioDeviceName,

    [Parameter()]
    [string]$VideoDeviceName,

    [Parameter()]
    [string]$DisplayDeviceName,
    
    <# Set-CalendarProcessing #>
    [Parameter()]
    [string]$AdditionalResponse,

    [Parameter()]
    [ValidateScript({ $_ -notlike '* *' })]
    [string[]]$BookingAdd,

    [Parameter()]
    [ValidateScript({ $_ -notlike '* *' })]
    [string[]]$BookingRemove,

    <# Add-MailboxFolderPermission | Remove-MailboxFolderPermission #>
    [Parameter()]
    [ValidateScript({ $_ -notlike '* *' })]
    [string[]]$CalendarPermissionAdd,

    [Parameter()]
    [ValidateSet("Author", "AvailabilityOnly", "Contributor", "Editor", "LimitedDetails", "None", "NonEditingAuthor", "Owner", "PublishingEditor", "PublishingAuthor", "Reviewer")]
    [string[]]$CalendarPermissionAccessRights,

    [Parameter()]
    [ValidateScript({ $_ -notlike '* *' })]
    [string[]]$CalendarPermissionRemove,

    <# Add-MailboxPermission #>
    [Parameter()]
    [ValidateScript({ $_ -notlike '* *' })]
    [string[]]$FullAccessAdd,

    [Parameter()]
    [ValidateScript({ $_ -notlike '* *' })]
    [string[]]$FullAccessRemove
)

if (($EmailAddressesAdd -or $EmailAddressesRemove) -and $UserPrincipalName)
{
    Write-Error "You can't use the EmailAddresses parameters and the UserPrincipalName parameter in the same command" -ErrorAction Stop
}
if ($BookingRemove -and 'Alle' -in $BookingRemove)
{
    Write-Error "You can't remove access for 'Alle'. Specify each UPN to remove" -ErrorAction Stop
}
if ($BookingAdd -and $BookingRemove)
{
    Write-Error "You can't add and remove Booking in the same command!" -ErrorAction Stop
}
if ($CalendarPermissionAdd -and $CalendarPermissionRemove)
{
    Write-Error "You can't add and remove calendar permissions in the same command!" -ErrorAction Stop
}

if (![string]::IsNullOrEmpty($Building) -or ![string]::IsNullOrEmpty($City) -or ![string]::IsNullOrEmpty($CountryOrRegion) -or $Floor -gt -1 -or ![string]::IsNullOrEmpty($GeoCoordinates) -or $PSBoundParameters.ContainsKey("IsWheelChairAccessible") -or ![string]::IsNullOrEmpty($PostalCode) -or $ResourceCapacity -gt -1 -or ![string]::IsNullOrEmpty($State) -or ![string]::IsNullOrEmpty($StreetAddress) -or ($null -ne $Tags -and $Tags.Count -gt 0) -or ![string]::IsNullOrEmpty($AudioDeviceName) -or ![string]::IsNullOrEmpty($VideoDeviceName) -or ![string]::IsNullOrEmpty($DisplayDeviceName))
{
    $mbx = Get-Mailbox $Identity
    if ($mbx.RecipientTypeDetails -ne "RoomMailbox")
    {
        Write-Error "'Set-Place' arguments supports only RoomMailbox and not '$($mbx.RecipientTypeDetails)'." -ErrorAction Stop
    }
}

Function Set-FullAccessPermissions
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [string]$Identity,

        [Parameter(Mandatory = $True)]
        [string[]]$FullAccess,

        [Parameter(Mandatory = $True)]
        [ValidateSet("Add", "Remove")]
        [string]$Action
    )

    # setting full access rights
    foreach ($recipient in $FullAccess)
    {
        # fullAccess setting
        $fullAccessSplat = @{
            #Identity = $Name
            Identity = $Identity
            User = $recipient
            AccessRights = "FullAccess"
            InheritanceType = "All"
            Confirm = $False
        }

        if ($Action -eq "Add")
        {
            if ($PSCmdlet.ShouldProcess(($fullAccessSplat | ConvertTo-Json), "Add-MailboxPermission"))
            {
                Add-MailboxPermission @fullAccessSplat -WarningAction SilentlyContinue -ErrorAction Continue | Out-Null
            }
        }
        elseif ($Action -eq "Remove")
        {
            if ($PSCmdlet.ShouldProcess(($fullAccessSplat | ConvertTo-Json), "Remove-MailboxPermission"))
            {
                Remove-MailboxPermission @fullAccessSplat -WarningAction SilentlyContinue -ErrorAction Continue | Out-Null
            }
        }
    }
}

# mailbox settings
$mailboxSplat = @{
    Identity = $Identity
}

# user settings
$userSplat = @{
    Identity = $Identity
}

# calendarprocessing settings
$calendarProcessingSplat = @{
    Identity = $Identity
}

# calendarPermission settings
$calendarPermissionSplat = @{
    Identity = "$($Identity):\$((Get-MailboxFolderStatistics -Identity $Identity -FolderScope "Calendar" | Select-Object -First 1).Name)"
    Confirm = $False
}

# mailbox splat
if ($Alias) { $mailboxSplat.Add("Alias", $Alias.Trim()) }
if ($DisplayName) { $mailboxSplat.Add("DisplayName", $DisplayName.Trim()) }
if ($PSBoundParameters.ContainsKey("HiddenFromAddressLists")) { $mailboxSplat.Add("HiddenFromAddressListsEnabled", $HiddenFromAddressLists) }
if ($MailTip) { $mailboxSplat.Add("MailTip", $MailTip) }
if ($Office) { $mailboxSplat.Add("Office", $Office) }
if ($ResourceCapacity -gt -1) { $mailboxSplat.Add("ResourceCapacity", $ResourceCapacity) }
if ($EmailAddressesAdd -or $EmailAddressesRemove)
{
    if ($EmailAddressesAdd -and $EmailAddressesRemove) { $mailboxSplat.Add("EmailAddresses", @{Add=$EmailAddressesAdd; Remove=$EmailAddressesRemove}) }
    elseif ($EmailAddressesAdd) { $mailboxSplat.Add("EmailAddresses", @{Add=$EmailAddressesAdd}) }
    elseif ($EmailAddressesRemove) { $mailboxSplat.Add("EmailAddresses", @{Remove=$EmailAddressesRemove}) }
}
if ($UserPrincipalName)
{
    if (!$mbx)
    {
        $mbx = Get-Mailbox $Identity
    }

    $addresses = $mbx | Select -ExpandProperty EmailAddresses | % { $_.Replace("SMTP:", "smtp:") }
    $addresses += "SMTP:$UserPrincipalName"
    $mailboxSplat.Add("EmailAddresses", $addresses)
}

# user splat
if ($Company) { $userSplat.Add("Company", $Company.Trim()) }
if ($Phone) { $userSplat.Add("Phone", $Phone.Trim()) }
if ($StreetAddress) { $userSplat.Add("StreetAddress", $StreetAddress.Trim()) }
if ($PostalCode) { $userSplat.Add("PostalCode", $PostalCode.Trim()) }
if ($City) { $userSplat.Add("City", $City.Trim()) }

# place splat
$placeSplat = @{
    Identity = $Identity
}
if (![string]::IsNullOrEmpty($Building)) { $placeSplat.Add("Building", $Building) }
if (![string]::IsNullOrEmpty($City)) { $placeSplat.Add("City", $City) }
if (![string]::IsNullOrEmpty($CountryOrRegion)) { $placeSplat.Add("CountryOrRegion", $CountryOrRegion) }
if ($Floor -gt -1) { $placeSplat.Add("Floor", $Floor) }
if (![string]::IsNullOrEmpty($GeoCoordinates)) { $placeSplat.Add("GeoCoordinates", $GeoCoordinates) }
if ($PSBoundParameters.ContainsKey("IsWheelChairAccessible")) { $placeSplat.Add("IsWheelChairAccessible", $IsWheelChairAccessible) }
if (![string]::IsNullOrEmpty($PostalCode)) { $placeSplat.Add("PostalCode", $PostalCode) }
if ($ResourceCapacity -gt -1) { $placeSplat.Add("Capacity", $ResourceCapacity) }
if (![string]::IsNullOrEmpty($State)) { $placeSplat.Add("State", $State) }
if (![string]::IsNullOrEmpty($StreetAddress)) { $placeSplat.Add("Street", $StreetAddress) }
if (($null -ne $Tags -and $Tags.Count -gt 0)) { $placeSplat.Add("Tags", $Tags) }
if (![string]::IsNullOrEmpty($AudioDeviceName)) { $placeSplat.Add("AudioDeviceName", $AudioDeviceName) }
if (![string]::IsNullOrEmpty($VideoDeviceName)) { $placeSplat.Add("VideoDeviceName", $VideoDeviceName) }
if (![string]::IsNullOrEmpty($DisplayDeviceName)) { $placeSplat.Add("DisplayDeviceName", $DisplayDeviceName) }

# calendarprocessing splat
if ($PSBoundParameters.ContainsKey("AdditionalResponse"))
{
    if ($AdditionalResponse)
    {
        $calendarProcessingSplat.Add("AdditionalResponse", $AdditionalResponse.Trim())
        $calendarProcessingSplat.Add("AddAdditionalResponse", $True)
    }
    else
    {
        $calendarProcessingSplat.Add("AddAdditionalResponse", $False)
    }
}
if ($BookingAdd -or $BookingRemove)
{
    # get current bookInPolicy users
    $currentBookInPolicy = Get-CalendarProcessing $Identity | Select -ExpandProperty BookInPolicy | Get-Mailbox | Select -ExpandProperty UserPrincipalName
}
if ($BookingAdd)
{
    if ('Alle' -notin $BookingAdd)
    {
        # add bookInPolicy given in
        $bookInPolicy = @($currentBookInPolicy) + @($BookingAdd)

        $calendarProcessingSplat.Add("BookInPolicy", $bookInPolicy)
        $calendarProcessingSplat.Add("AllBookInPolicy", $False)
    }
    else
    {
        $calendarProcessingSplat.Add("BookInPolicy", $null)
        $calendarProcessingSplat.Add("AllBookInPolicy", $True)
    }
}
if ($BookingRemove)
{
    # remove bookInPolicy given in
    $bookInPolicy = $currentBookInPolicy | Where { $_ -notin $BookingRemove }

    if ($bookInPolicy)
    {
        $calendarProcessingSplat.Add("BookInPolicy", $bookInPolicy)
        $calendarProcessingSplat.Add("AllBookInPolicy", $False)
    }
    else
    {
        $calendarProcessingSplat.Add("BookInPolicy", $null)
        $calendarProcessingSplat.Add("AllBookInPolicy", $True)
    }
}

# msolUserPrincipalName splat
if ($UserPrincipalName)
{
    # get UPN as of now
    $msolUserPrincipalNameSplat = @{
        UserPrincipalName = ($mbx | Select -ExpandProperty UserPrincipalName)
        NewUserPrincipalName = $UserPrincipalName
    }
}

# execute mailbox settings
if ($mailboxSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($mailboxSplat | ConvertTo-Json), "Set-Mailbox"))
    {
        Set-Mailbox @mailboxSplat -ErrorAction Stop
    }
}

# execute user settings
if ($userSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($userSplat | ConvertTo-Json), "Set-User"))
    {
        Set-User @userSplat -ErrorAction Stop
    }
}

# set place info
if ($placeSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($placeSplat | ConvertTo-Json), "Set-Place"))
    {
        Set-Place @placeSplat -ErrorAction Stop
    }
}

# execute calendarprocessing settings
if ($calendarProcessingSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($calendarProcessingSplat | ConvertTo-Json), "Set-CalendarProcessing"))
    {
        Set-CalendarProcessing @calendarProcessingSplat -ErrorAction Stop
    }

    if ($BookingAdd -or $BookingRemove)
    {
        if ($PSCmdlet.ShouldProcess($Identity, "Get-Mailbox"))
        {
            $identity = (Get-Mailbox $Identity).UserPrincipalName
            $calendarName = (Get-MailboxFolderStatistics -Identity $identity -FolderScope Calendar | Select-Object -First 1).Name
            $folderPermissionName = "$($identity):\$calendarName"
        }
    }

    if ($BookingAdd)
    {
        if ("Alle" -notin $BookingAdd)
        {
            # change accessrights to Reviewer
            foreach ($booker in $BookingAdd)
            {
                if ($PSCmdlet.ShouldProcess($booker, "Set-MailboxFolderPermission"))
                {
                    Set-MailboxFolderPermission -Identity $folderPermissionName -User $booker -AccessRight Reviewer -ErrorAction Stop
                }
            }
        }
        else
        {
            # remove access
            foreach ($booker in $currentBookInPolicy)
            {
                if ($PSCmdlet.ShouldProcess($booker, "Remove-MailboxFolderPermission"))
                {
                    Remove-MailboxFolderPermission -Identity $folderPermissionName -User $booker -Confirm:$False
                }
            }
        }
    }
    if ($BookingRemove)
    {
        # remove access
        foreach ($booker in $BookingRemove)
        {
            if ($PSCmdlet.ShouldProcess($booker, "Remove-MailboxFolderPermission"))
            {
                Remove-MailboxFolderPermission -Identity $folderPermissionName -User $booker -Confirm:$False
            }
        }
    }

    # make sure Default user is set to Reviewer (again)
    Set-MailboxFolderPermission -Identity $folderPermissionName -User Default -AccessRight Reviewer -ErrorAction Stop
}

# add Calendar permissions
if ($CalendarPermissionAdd)
{
    if ($CalendarPermissionAccessRights) { $calendarPermissionSplat.Add("AccessRights", $CalendarPermissionAccessRights) }
    else { $calendarPermissionSplat.Add("AccessRights", "Reviewer") }
    $CalendarPermissionAdd | % {
        $calendarPermissionSplat.User = $_

        if ($PSCmdlet.ShouldProcess(($calendarPermissionSplat | ConvertTo-Json), "Add-MailboxFolderPermission"))
        {
            Add-MailboxFolderPermission @calendarPermissionSplat -ErrorAction Stop
        }
    }
}

# remove Calendar permissions
if ($CalendarPermissionRemove)
{
    $CalendarPermissionRemove | % {
        $calendarPermissionSplat.User = $_

        if ($PSCmdlet.ShouldProcess(($calendarPermissionSplat | ConvertTo-Json), "Remove-MailboxFolderPermission"))
        {
            Remove-MailboxFolderPermission @calendarPermissionSplat -ErrorAction Stop
        }
    }
}

# execute msolUserPrincipalName settings
if ($msolUserPrincipalNameSplat)
{
    if ($PSCmdlet.ShouldProcess(($msolUserPrincipalNameSplat | ConvertTo-Json), "Set-MsolUserPrincipalName"))
    {
        Set-MsolUserPrincipalName @msolUserPrincipalNameSplat -ErrorAction Stop
    }
}

if ($FullAccessAdd -or $FullAccessRemove)
{
    if (!$mbx)
    {
        $mbx = Get-Mailbox $Identity
    }
}

# grant full access rights
if ($FullAccessAdd)
{
    $mbx | Set-FullAccessPermissions -FullAccess $FullAccessAdd -Action Add
}

# revoke full access rights
if ($FullAccessRemove)
{
    $mbx | Set-FullAccessPermissions -FullAccess $FullAccessRemove -Action Remove
}