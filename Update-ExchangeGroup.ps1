### Set-DistributionGroup
# Identity can be Name, Alias, DN, CanonicalName, EmailAddress, GUID
#
[CmdletBinding(SupportsShouldProcess)]
param(
    <# Set-DistributionGroup #>
    [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
    [string]$Identity,

    [Parameter()]
    [string]$Alias,

    [Parameter()]
    [string]$DisplayName,

    [Parameter()]
    [string[]]$EmailAddressesAdd,
    
    [Parameter()]
    [string[]]$EmailAddressesRemove,

    [Parameter()]
    [bool]$HiddenFromAddressLists,

    [Parameter()]
    [string]$MailTip,

    [Parameter()]
    [string[]]$ManagedByAdd,

    [Parameter()]
    [string[]]$ManagedByRemove,

    [Parameter()]
    [string[]]$ModeratedByAdd,
    
    [Parameter()]
    [string[]]$ModeratedByRemove,

    [Parameter()]
    [bool]$ModerationEnabled,

    [Parameter()]
    [string]$Name,

    [Parameter()]
    [string]$PrimarySmtpAddress,

    [Parameter()]
    [string]$AcceptMessagesFromInternalAndExternal,

    [Parameter()]
    [string[]]$SendOnBehalfToAdd,

    [Parameter()]
    [string[]]$SendOnBehalfToRemove,

    <# Set-RecipientPermission #>
    [Parameter()]
    [string[]]$SendAsAdd,

    [Parameter()]
    [string[]]$SendAsRemove
)

if (($EmailAddressesAdd -or $EmailAddressesRemove) -and $PrimarySmtpAddress)
{
    Write-Error "You can't use the EmailAddresses parameter and the PrimarySmtpAddress parameter in the same command" -ErrorAction Stop
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

# distributionGroup settings
$distributionGroupSplat = @{
    Identity = $Identity
}

if ($Alias) { $distributionGroupSplat.Add("Alias", $Alias) }
if ($DisplayName) { $distributionGroupSplat.Add("DisplayName", $DisplayName) }
if ($EmailAddressesAdd -or $EmailAddressesRemove)
{
    if ($EmailAddressesAdd -and $EmailAddressesRemove) { $distributionGroupSplat.Add("EmailAddresses", @{Add=$EmailAddressesAdd; Remove=$EmailAddressesRemove}) }
    elseif ($EmailAddressesAdd) { $distributionGroupSplat.Add("EmailAddresses", @{Add=$EmailAddressesAdd}) }
    elseif ($EmailAddressesRemove) { $distributionGroupSplat.Add("EmailAddresses", @{Remove=$EmailAddressesRemove}) }
}
if ($PSBoundParameters.ContainsKey("HiddenFromAddressLists")) { $distributionGroupSplat.Add("HiddenFromAddressListsEnabled", $HiddenFromAddressLists) }
if ($MailTip) { $distributionGroupSplat.Add("MailTip", $MailTip) }
if ($ManagedByAdd -or $ManagedByRemove)
{
    if ($ManagedByAdd -and $ManagedByRemove) { $distributionGroupSplat.Add("ManagedBy", @{Add=$ManagedByAdd; Remove=$ManagedByRemove}) }
    elseif ($ManagedByAdd) { $distributionGroupSplat.Add("ManagedBy", @{Add=$ManagedByAdd}) }
    elseif ($ManagedByRemove) { $distributionGroupSplat.Add("ManagedBy", @{Remove=$ManagedByRemove}) }
}
if ($ModeratedByAdd -or $ModeratedByRemove)
{
    if ($ModeratedByAdd -and $ModeratedByRemove) { $distributionGroupSplat.Add("ModeratedBy", @{Add=$ModeratedByAdd; Remove=$ModeratedByRemove}) }
    elseif ($ModeratedByAdd) { $distributionGroupSplat.Add("ModeratedBy", @{Add=$ModeratedByAdd}) }
    elseif ($ModeratedByRemove) { $distributionGroupSplat.Add("ModeratedBy", @{Remove=$ModeratedByRemove}) }
}
if ($PSBoundParameters.ContainsKey("ModerationEnabled")) { $distributionGroupSplat.Add("ModerationEnabled", $ModerationEnabled) }
if ($Name) { $distributionGroupSplat.Add("Name", $Name) }
if ($PrimarySmtpAddress)
{
    $distributionGroupSplat.Add("PrimarySmtpAddress", $PrimarySmtpAddress)
    $distributionGroupSplat.Add("Alias", $PrimarySmtpAddress.Split("@")[0])
}
if ($PSBoundParameters.ContainsKey("AcceptMessagesFromInternalAndExternal"))
{
    # needs to be inverted
    $distributionGroupSplat.Add("RequireSenderAuthenticationEnabled", !$AcceptMessagesFromInternalAndExternal)
}
if ($SendOnBehalfToAdd -or $SendOnBehalfToRemove)
{
    if ($SendOnBehalfToAdd -and $SendOnBehalfToRemove) { $distributionGroupSplat.Add("GrantSendOnBehalfTo", @{Add=$SendOnBehalfToAdd; Remove=$SendOnBehalfToRemove}) }
    elseif ($SendOnBehalfToAdd) { $distributionGroupSplat.Add("GrantSendOnBehalfTo", @{Add=$SendOnBehalfToAdd}) }
    elseif ($SendOnBehalfToRemove) { $distributionGroupSplat.Add("GrantSendOnBehalfTo", @{Remove=$SendOnBehalfToRemove}) }
}

# execute distributiongroup settings
if ($distributionGroupSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($distributionGroupSplat | ConvertTo-Json), "Set-DistributionGroup"))
    {
        Set-DistributionGroup @distributionGroupSplat -BypassSecurityGroupManagerCheck -ErrorAction Stop
    }
}

if ($SendAsAdd -or $SendAsRemove)
{
    $group = Get-DistributionGroup $Identity

    if ($SendAsAdd)
    {
        $group | Set-GrantSendAsPermissions -SendAs $SendAsAdd -Action Add
    }
    if ($SendAsRemove)
    {
        $group | Set-GrantSendAsPermissions -SendAs $SendAsRemove -Action Remove
    }
}