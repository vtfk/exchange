# https://itsallinthecode.com/exchange-online-enable-group-moderation-and-sending-restrictions/
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
    [string]$Identity,

    [Parameter()]
    [bool]$ModerationEnabled,

    ## ONLY USERS (distinguishedNames)
    [Parameter()]
    [string[]]$Moderators,

    ## ONLY USERS (distinguishedNames)
    [Parameter()]
    [string[]]$BypassModeratorUsers,

    ## ONLY GROUPS (distinguishedNames)
    [Parameter()]
    [string[]]$BypassModeratorGroups,

    <#Notify all senders when their messages aren’t approved. = 6
    Notify senders in your organization when their messages aren’t approved. = 2
    Don’t notify anyone when a message isn’t approved. = 0#>
    [Parameter()]
    [ValidateSet(0, 2, 6)]
    [int]$ModerationFlags = 6
)

if ($Moderators) {
    $Moderators | % {
        if ($_ -notmatch 'CN=.+DC=.+') {
            Write-Error 'Moderators must be an array of distinguishedNames' -ErrorAction Stop
        }
    }
}

if ($BypassModeratorUsers) {
    $BypassModeratorUsers | % {
        if ($_ -notmatch "CN=.+DC=.+") {
            Write-Error 'BypassModeratorUsers must be an array of distinguishedNames' -ErrorAction Stop
        }
    }
}

if ($BypassModeratorGroups) {
    $BypassModeratorGroups | % {
        if ($_ -notmatch "CN=.+DC=.+") {
            Write-Error 'BypassModeratorGroups must be an array of distinguishedNames' -ErrorAction Stop
        }
    }
}

$groupSplat = @{
    Identity = $Identity
}

$addSplat = @{}
$replaceSplat = @{}

# get group to modify
$group = Get-ADGroup -Identity $Identity -Properties * -ErrorAction Stop

if ($PSBoundParameters.ContainsKey('ModerationEnabled')) {
    if ($group.Contains('msExchEnableModeration')) { $replaceSplat.Add('msExchEnableModeration', $ModerationEnabled) }
    else { $addSplat.Add('msExchEnableModeration', $ModerationEnabled) }
}

if ($PSBoundParameters.ContainsKey('Moderators')) {
    if ($group.Contains('msExchModeratedByLink')) { $replaceSplat.Add('msExchModeratedByLink', ($group.msExchModeratedByLink + $Moderators)) }
    else { $addSplat.Add('msExchModeratedByLink', $Moderators) }
}

if ($PSBoundParameters.ContainsKey('BypassModeratorUsers')) {
    if ($group.Contains('msExchBypassModerationLink')) { $replaceSplat.Add('msExchBypassModerationLink', ($group.msExchBypassModerationLink + $BypassModeratorUsers)) }
    else { $addSplat.Add('msExchBypassModerationLink', $BypassModeratorUsers) }
}

if ($PSBoundParameters.ContainsKey('BypassModeratorGroups')) {
    if ($group.Contains('dlMemSubmitPerms')) { $replaceSplat.Add('dlMemSubmitPerms', ($group.dlMemSubmitPerms + $BypassModeratorGroups)) }
    else { $addSplat.Add('dlMemSubmitPerms', $BypassModeratorGroups) }
}

if ($PSBoundParameters.ContainsKey('ModerationFlags')) {
    if ($group.Contains('msExchModerationFlags')) { $replaceSplat.Add('msExchModerationFlags', $ModerationFlags) }
    else { $addSplat.Add('msExchModerationFlags', $ModerationFlags) }
}

if ($addSplat.Count -gt 0) { $groupSplat.Add("Add", $addSplat) }
if ($replaceSplat.Count -gt 0) { $groupSplat.Add("Replace", $replaceSplat) }

if ($groupSplat.Count -le 1) {
    Write-Error 'Nothing to do. Aborting!' -ErrorAction Stop
}

if ($PSCmdlet.ShouldProcess(($groupSplat | ConvertTo-Json), "Set-ADGroup")) {
    Set-ADGroup @groupSplat -ErrorAction Stop
}