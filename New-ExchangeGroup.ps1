[CmdletBinding(SupportsShouldProcess)]
param(
    <# New-DistributionGroup #>
    [Parameter(Mandatory = $True)]
    [string]$DisplayName,

    [Parameter(Mandatory = $True)]
    [string]$PrimarySmtpAddress,

    [Parameter(Mandatory = $True)]
    [string[]]$ManagedBy,

    [Parameter()]
    [string[]]$Members,

    [Parameter()]
    [string[]]$ModeratedBy,

    [Parameter()]
    [string]$Notes,

    [Parameter()]
    [bool]$AcceptMessagesFromInternalAndExternal,

    [Parameter()]
    [switch]$RoomList,

    <# Set-DistributionGroup #>
    [Parameter()]
    [bool]$HiddenFromAddressLists,

    [Parameter()]
    [string[]]$EmailAddresses,

    [Parameter()]
    [string]$MailTip,

    [Parameter()]
    [string[]]$SendOnBehalfTo,

    <# Add-RecipientPermission #>
    [Parameter()]
    [string[]]$SendAs,

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

if ($Creator) {
    Add-LogTarget -Name Email -Configuration @{
        SMTPServer = $smtpServer
        From = $smtpFrom
        To = $Creator
        Level = "SUCCESS"
        BodyAsHtml = $True
    }

    if ($Target) {
        if ($PSCmdlet.ShouldProcess((@{Exchange=$True;Target=$Target} | ConvertTo-Json), "Connect-Office365")) {
            Connect-Office365 -Exchange -Target $Target
        }
    }

    # error message for mail body
    [string]$mailMsgError = "<p><b>Feilmelding</b>:<br>%errormsg%</p>`n<p>Trenger du hjelp? Ta kontakt med <a href=`"mailto:$servicedeskMail`">$servicedeskMail</a> for feilsøking."
}

# distributionGroup settings
$distributionGroupSplat = @{
    Name = $DisplayName
    Type = "Security"
    DisplayName = $DisplayName;
    Alias = ($PrimarySmtpAddress.Split('@')[0])
    PrimarySmtpAddress = $PrimarySmtpAddress
    ManagedBy = $ManagedBy
    Members = $Members
}

if ($ModeratedBy)
{
    $distributionGroupSplat.Add("ModerationEnabled", $True)
    $distributionGroupSplat.Add("ModeratedBy", $ModeratedBy)
}
if ($Notes) { $distributionGroupSplat.Add("Notes", $Notes) }
if ($PSBoundParameters.ContainsKey("AcceptMessagesFromInternalAndExternal"))
{
    # needs to be inverted
    $distributionGroupSplat.Add("RequireSenderAuthenticationEnabled", !$AcceptMessagesFromInternalAndExternal)
}
if ($RoomList)
{
    $distributionGroupSplat.Add("RoomList", $True)
    $distributionGroupSplat.Remove("Type")
    $distributionGroupSplat.Add("Type", "Distribution")
}

# configureDistributionGroup settings
$configureDistributionGroupSplat = @{
    Identity = $PrimarySmtpAddress
}

if ($PSBoundParameters.ContainsKey("HiddenFromAddressLists")) { $configureDistributionGroupSplat.Add("HiddenFromAddressListsEnabled", $HiddenFromAddressLists) }
if ($MailTip) { $configureDistributionGroupSplat.Add("MailTip", $MailTip) }
if ($SendOnBehalfTo) { $configureDistributionGroupSplat.Add("GrantSendOnBehalfTo", $SendOnBehalfTo) }

# add aliases
if ($EmailAddresses)
{
    # array with emailaddresses to set
    [string[]]$addresses = @("SMTP:$($PrimarySmtpAddress)")

    # add aliases
    foreach ($email in $EmailAddresses)
    {
        $addresses += "smtp:$($email)"
    }

    $configureDistributionGroupSplat.Add("EmailAddresses", $addresses)
}

# create security mail-enabled group
if ($PSCmdlet.ShouldProcess(($distributionGroupSplat | ConvertTo-Json), "New-DistributionGroup"))
{
    try
    {
        New-DistributionGroup @distributionGroupSplat -ErrorAction Stop | Out-Null
        Write-Log -Message "Created new distribution group" -Body $distributionGroupSplat
    }
    catch
    {
        if (!$Creator) {
            Write-Log -Message "Failed to create new distrigution group: $_" -Body $distributionGroupSplat -Level ERROR
        }
        else {
            Write-Log -Message "Feilet ved opprettelse av distribution group!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $distributionGroupSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
        }

        Write-Error $_ -ErrorAction Stop
    }
}

# configure security mail-enabled group
if ($configureDistributionGroupSplat.Count -gt 1)
{
    if ($PSCmdlet.ShouldProcess(($configureDistributionGroupSplat | ConvertTo-Json), "Set-DistributionGroup"))
    {
        # wait for group to be created in Exchange before continuing
        Start-Sleep -Seconds 5

        try
        {
            Set-DistributionGroup @configureDistributionGroupSplat -ErrorAction Stop -BypassSecurityGroupManagerCheck
            Write-Log -Message "Additional setting(s) set on distribution group" -Body $configureDistributionGroupSplat
        }
        catch
        {
            if (!$Creator) {
                Write-Log -Message "Failed to set additional setting(s) on distrigution group: $_" -Body $configureDistributionGroupSplat -Level ERROR
            }
            else {
                Write-Log -Message "Feilet ved setting av ekstra innstillinger på distribution group!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $distributionGroupSplat; AdditionalSettings = $configureDistributionGroupSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
            }

            Write-Error $_ -ErrorAction Stop
        }
    }
}

if ($SendAs)
{
    # wait for group to be created in Exchange before continuing
    Start-Sleep -Seconds 5

    # setting send as rights
    foreach ($recipient in $SendAs)
    {
        # sendAs setting
        $sendAsSplat = @{
            Identity = $PrimarySmtpAddress
            AccessRights = "SendAs"
            Trustee = $recipient
            Confirm = $False
        }

        if ($PSCmdlet.ShouldProcess(($sendAsSplat | ConvertTo-Json), "Add-RecipientPermission"))
        {
            try
            {
                Add-RecipientPermission @sendAsSplat -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                Write-Log -Message "SendAs permission added to distribution group" -Body $sendAsSplat
            }
            catch
            {
                if (!$Creator) {
                    Write-Log -Message "Failed to add SendAs permission to distrigution group: $_" -Body $sendAsSplat -Level WARNING
                }
                else {
                    Write-Log -Message "Feilet ved SendAs tilgang på distribution group!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $distributionGroupSplat; AdditionalSettings = $configureDistributionGroupSplat; SendAsUser = $sendAsSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
                }

                Write-Error $_ -ErrorAction Continue
            }
        }
    }
}

# send success mail if Creator is filled out and everything has worked fine!
if ($Creator) {
    $mailSplat = [ordered]@{
        Generelt = $distributionGroupSplat
        AdditionalSettings = $configureDistributionGroupSplat
    }
    if ($ManagedBy) { $mailSplat.Add("ManagedBy", $ManagedBy) }
    if ($Members) { $mailSplat.Add("Members", $Members) }
    if ($ModeratedBy) { $mailSplat.Add("ModeratedBy", $ModeratedBy) }
    if ($SendOnBehalfTo) { $mailSplat.Add("SendOnBehalfTo", $SendOnBehalfTo) }
    if ($SendAs) { $mailSplat.Add("SendAs", $SendAs) }

    if ($PSCmdlet.ShouldProcess(($mailSplat | ConvertTo-Json), "Write-Log")) {
        Write-Log -Message "Ny distribusjonsgruppe opprettet<br><br>$(ConvertTo-List -Hash $mailSplat)" -Body @{ Subject = "Exchange skjemaparser" } -Level SUCCESS
    }
}