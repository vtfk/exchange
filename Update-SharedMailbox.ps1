### Set-Mailbox
# Identity can be Name, Alias, DN, CanonicalName, EmailAddress, GUID, LegacyExchangeDN, SamAccountName, UPN
#
### Set-User
# Identity can be Name, DN, CanonicalName, GUID, UPN
#
[CmdletBinding(SupportsShouldProcess)]
param(
    <# Set-Mailbox | Set-User #>
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

    <# Set-Mailbox | Set-MsolUserPrincipalName #>
    [Parameter()]
    [string]$UserPrincipalName,

    <# Set-User #>
    [Parameter()]
    # firma
    [string]$Company,

    [Parameter()]
    # telefonarbeid
    [string]$Phone,

    [Parameter()]
    # gatevei
    [string]$StreetAddress,

    [Parameter()]
    # postnummer
    [string]$PostalCode,

    [Parameter()]
    # poststed
    [string]$City,

    [Parameter()]
    # områderegion
    [string]$StateOrProvince,

    [Parameter()]
    # land
    $CountryOrRegion,

    [Parameter()]
    # mobiltelefon
    [string]$MobilePhone,

    [Parameter()]
    # faks
    [string]$Fax,

    [Parameter()]
    # telefonprivat
    [string]$HomePhone,

    [Parameter()]
    # nettside
    [string]$WebPage,

    [Parameter()]
    # notater
    [string]$Notes,

    [Parameter()]
    # stilling
    [string]$Title,

    [Parameter()]
    # avdeling
    [string]$Department,

    [Parameter()]
    # overordnet
    [string]$Manager,

    <# Add-RecipientPermission #>
    [Parameter()]
    [string[]]$SendAsAdd,

    [Parameter()]
    [string[]]$SendAsRemove,

    <# Add-MailboxPermission #>
    [Parameter()]
    [string[]]$FullAccessAdd,

    [Parameter()]
    [string[]]$FullAccessRemove
)

if (($EmailAddressesAdd -or $EmailAddressesRemove) -and $PrimarySmtpAddress)
{
    Write-Error "You can't use the EmailAddresses parameter and the UserPrincipalName parameter in the same command" -ErrorAction Stop
}

Function Set-GrantSendAsPermissions
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [string]$Identity,

        [Parameter(Mandatory = $True)]
        [string[]]$SendAs,

        [Parameter(Mandatory = $True)]
        [ValidateSet("Add", "Remove")]
        [string]$Action
    )

    # setting send as rights
    foreach ($recipient in $SendAs)
    {
        # sendAs setting
        $sendAsSplat = @{
            #Identity = $Name
            Identity = $Identity
            AccessRights = "SendAs"
            Trustee = $recipient
            Confirm = $False
        }

        if ($Action -eq "Add")
        {
            if ($PSCmdlet.ShouldProcess(($sendAsSplat | ConvertTo-Json), "Add-RecipientPermission"))
            {
                Add-RecipientPermission @sendAsSplat -WarningAction SilentlyContinue -ErrorAction Continue | Out-Null
            }
        }
        elseif ($Action -eq "Remove")
        {
            if ($PSCmdlet.ShouldProcess(($sendAsSplat | ConvertTo-Json), "Remove-RecipientPermission"))
            {
                Remove-RecipientPermission @sendAsSplat -WarningAction SilentlyContinue -ErrorAction Continue | Out-Null
            }
        }
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

# mailbox splat
if ($Alias) { $mailboxSplat.Add("Alias", $Alias.Trim()) }
if ($DisplayName) { $mailboxSplat.Add("DisplayName", $DisplayName.Trim()) }
if ($PSBoundParameters.ContainsKey("HiddenFromAddressLists")) { $mailboxSplat.Add("HiddenFromAddressListsEnabled", $HiddenFromAddressLists) }
if ($MailTip) { $mailboxSplat.Add("MailTip", $MailTip) }
if ($Office) { $mailboxSplat.Add("Office", $Office) }
if ($EmailAddressesAdd -or $EmailAddressesRemove)
{
    if ($EmailAddressesAdd -and $EmailAddressesRemove) { $mailboxSplat.Add("EmailAddresses", @{Add=$EmailAddressesAdd; Remove=$EmailAddressesRemove}) }
    elseif ($EmailAddressesAdd) { $mailboxSplat.Add("EmailAddresses", @{Add=$EmailAddressesAdd}) }
    elseif ($EmailAddressesRemove) { $mailboxSplat.Add("EmailAddresses", @{Remove=$EmailAddressesRemove}) }
}
if ($UserPrincipalName)
{
    $mbx = Get-Mailbox $Identity
    $addresses = $mbx | Select -ExpandProperty EmailAddresses | % { $_.Replace("SMTP:", "smtp:") }
    $addresses += "SMTP:$UserPrincipalName"
    $mailboxSplat.Add("EmailAddresses", $addresses)

    # also set Alias
    $mailboxSplat.Add("Alias", $UserPrincipalName.Split("@")[0])
}

# user splat
if ($Company) { $userSplat.Add("Company", $Company.Trim()) }
if ($Phone) { $userSplat.Add("Phone", $Phone.Trim()) }
if ($StreetAddress) { $userSplat.Add("StreetAddress", $StreetAddress.Trim()) }
if ($PostalCode) { $userSplat.Add("PostalCode", $PostalCode.Trim()) }
if ($City) { $userSplat.Add("City", $City.Trim()) }

if ($StateOrProvince) { $userSplat.Add("StateOrProvince", $StateOrProvince.Trim()) }
if ($CountryOrRegion) { $userSplat.Add("CountryOrRegion", $CountryOrRegion) }
if ($MobilePhone) { $userSplat.Add("MobilePhone", $MobilePhone.Trim()) }
if ($Fax) { $userSplat.Add("Fax", $Fax.Trim()) }
if ($HomePhone) { $userSplat.Add("HomePhone", $HomePhone.Trim()) }
if ($WebPage) { $userSplat.Add("WebPage", $WebPage.Trim()) }
if ($Notes) { $userSplat.Add("Notes", $Notes.Trim()) }
if ($Title) { $userSplat.Add("Title", $Title.Trim()) }
if ($Department) { $userSplat.Add("Department", $Department.Trim()) }
if ($Manager) { $userSplat.Add("Manager", $Manager) }

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

# execute msolUserPrincipalName settings
if ($msolUserPrincipalNameSplat)
{
    if ($PSCmdlet.ShouldProcess(($msolUserPrincipalNameSplat | ConvertTo-Json), "Set-MsolUserPrincipalName"))
    {
        Set-MsolUserPrincipalName @msolUserPrincipalNameSplat -ErrorAction Stop
    }
}

if ($SendAsAdd -or $SendAsRemove -or $FullAccessAdd -or $FullAccessRemove)
{
    if (!$mbx)
    {
        $mbx = Get-Mailbox $Identity
    }
}

# grant send as rights
if ($SendAsAdd)
{
    $mbx | Set-GrantSendAsPermissions -SendAs $SendAsAdd -Action Add
}

# revoke send as rights
if ($SendAsRemove)
{
    $mbx | Set-GrantSendAsPermissions -SendAs $SendAsRemove -Action Remove
}

# grant full access rights
if ($FullAccessAdd)
{
    $mbx | Set-FullAccessPermissions -FullAccess $FullAccessAdd -Action Add

    # also grant FullAccess people SendAs
    $mbx | Set-GrantSendAsPermissions -SendAs $FullAccessAdd -Action Add
}

# revoke full access rights
if ($FullAccessRemove)
{
    $mbx | Set-FullAccessPermissions -FullAccess $FullAccessRemove -Action Remove
}
