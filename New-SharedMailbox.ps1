## FullAccess will also grant SendAs and GrantSendOnBehalfTo
[CmdletBinding(SupportsShouldProcess)]
param(
    <# New-Mailbox #>
    [Parameter(Mandatory = $True)]
    [string]$DisplayName,

    [Parameter(Mandatory = $True)]
    [string]$PrimarySmtpAddress,

    [Parameter()]
    [string]$FirstName,
    
    [Parameter()]
    [string]$LastName,

    <# Set-Mailbox #>
    [Parameter()]
    [bool]$HiddenFromAddressLists,

    [Parameter()]
    [string]$Office,

    [Parameter()]
    [string[]]$EmailAddresses,

    [Parameter()]
    [string[]]$SendOnBehalfTo,

    <# Add-RecipientPermission #>
    [Parameter()]
    [string[]]$SendAs,

    <# Add-MailboxPermission #>
    [Parameter()]
    [string[]]$FullAccess,

    ## Only used when called from Acos skjema-script
    # For logging purposes and mail delivery
    [Parameter()]
    [ValidateScript({ $_ -like "*@*.no" })]
    [string]$Creator,
    
    ## Only used when called from Acos skjema-script
    # Required to do a silent run
    [Parameter()]
    [string]$Target<#,

    [Parameter()]
    [switch]$SkipDisableUser#>
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

Function Add-GrantSendAsPermissions
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $True)]
        [string[]]$SendAs
    )
    
    # wait for mailbox to be available (hopefully)
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Start-Sleep -Seconds 3
    }

    # setting send as rights
    foreach ($recipient in $SendAs)
    {
        # sendAs setting
        $sendAsSplat = @{
            #Identity = $DisplayName
            Identity = $PrimarySmtpAddress
            AccessRights = "SendAs"
            Trustee = $recipient
            Confirm = $False
        }

        if ($PSCmdlet.ShouldProcess(($sendAsSplat | ConvertTo-Json), "Add-RecipientPermission"))
        {
            try
            {
                Add-RecipientPermission @sendAsSplat -WarningAction SilentlyContinue -ErrorAction Continue | Out-Null
                Write-Log -Message "SendAs permission added to shared mailbox" -Body $sendAsSplat
            }
            catch
            {
                if (!$Creator) {
                    Write-Log -Message "Failed to add SendAs permission to shared mailbox: $_" -Body $sendAsSplat -Level WARNING
                }
                else {
                    Write-Log -Message "Feilet ved SendAs tilgang til shared mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; AdditionalSettings = $mailboxConfigureSplat; SendAs = $sendAsSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level WARNING
                }
                Write-Error $_ -ErrorAction Continue
            }
        }
    }
}

Function Add-GrantSendOnBehalfToPermissions
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $True)]
        [string[]]$SendOnBehalfTo
    )
    
    # wait for mailbox to be available (hopefully)
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Start-Sleep -Seconds 3
    }

    # sendOnBehalfTo setting
    $sendOnBehalfToSplat = @{
        #Identity = $DisplayName
        Identity = $PrimarySmtpAddress
        GrantSendOnBehalfTo = $SendOnBehalfTo
    }

    # setting send on behalf to rights
    if ($PSCmdlet.ShouldProcess(($sendOnBehalfToSplat | ConvertTo-Json), "Set-Mailbox"))
    {
        try
        {
            Set-Mailbox @sendOnBehalfToSplat -WarningAction SilentlyContinue -ErrorAction Continue
            Write-Log -Message "SendOnBehalf permission added to shared mailbox" -Body $sendOnBehalfToSplat
        }
        catch
        {
            if (!$Creator) {
                Write-Log -Message "Failed to add SendOnBehalf permission to shared mailbox: $_" -Body $sendOnBehalfToSplat -Level WARNING
            }
            else {
                Write-Log -Message "Feilet ved SendOnBehalf tilgang til shared mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; AdditionalSettings = $mailboxConfigureSplat; SendOnBehalf = $sendOnBehalfToSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level WARNING
            }
            Write-Error $_ -ErrorAction Continue
        }
    }
}

# mailbox settings
$mailboxSplat = @{
    Name = $DisplayName
    Shared = $True
    DisplayName = $DisplayName
    Alias = ($PrimarySmtpAddress.Split('@')[0])
    PrimarySmtpAddress = $PrimarySmtpAddress
}

if ($LastName) { $mailboxSplat.Add("LastName", $LastName) }
if ($FirstName) { $mailboxSplat.Add("FirstName", $FirstName) }

# mailbox configure settings
$mailboxConfigureSplat = @{
    MessageCopyForSentAsEnabled = $True
    MessageCopyForSendOnBehalfEnabled = $True
}

if ($Office) { $mailboxConfigureSplat.Add("Office", $Office) }
if ($PSBoundParameters.ContainsKey("HiddenFromAddressLists")) { $mailboxConfigureSplat.Add("HiddenFromAddressListsEnabled", $HiddenFromAddressLists) }

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

    $mailboxConfigureSplat.Add("EmailAddresses", $addresses)
}

<# user block settings
$userBlockSplat = @{
    AccountEnabled = $False
}#>

# create shared mailbox
if ($PSCmdlet.ShouldProcess(($mailboxSplat | ConvertTo-Json), "New-Mailbox"))
{
    try
    {
        $sharedBox = New-Mailbox @mailboxSplat -ErrorAction Stop
        Write-Log -Message "New shared mailbox created" -Body $mailboxSplat

        #$userBlockSplat.Add("ObjectId", $sharedBox.UserPrincipalName)
    }
    catch
    {
        if (!$Creator) {
            Write-Log -Message "Failed to create new shared mailbox: $_" -Body $mailboxSplat -Level ERROR
        }
        else {
            Write-Log -Message "Feilet ved opprettelse av shared mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
        }

        Write-Error $_ -ErrorAction Stop
    }
}
else
{
    #$userBlockSplat.Add("ObjectId", $PrimarySmtpAddress)
}

# configure mailbox
if ($PSCmdlet.ShouldProcess(($mailboxConfigureSplat | ConvertTo-Json), "Set-Mailbox"))
{
    # wait for mailbox to be available (hopefully)
    Start-Sleep -Seconds 5

    try
    {
        $sharedBox | Set-Mailbox @mailboxConfigureSplat -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Log -Message "Additional setting(s) set on shared mailbox" -Body $mailboxConfigureSplat
    }
    catch
    {
        if (!$Creator) {
            Write-Log -Message "Failed to set additional setting(s) on shared mailbox: $_" -Body $mailboxConfigureSplat -Level ERROR
        }
        else {
            Write-Log -Message "Feilet ved setting av ekstra innstillinger på shared mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; AdditionalSettings = $mailboxConfigureSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
        }

        Write-Error $_ -ErrorAction Stop
    }
}

<# block user from signin
if (!$SkipDisableUser)
{
    if ($PSCmdlet.ShouldProcess(($userBlockSplat | ConvertTo-Json), "Set-AzureADUser"))
    {
        # wait for mailbox to be available (hopefully)
        Start-Sleep -Seconds 15

        Set-AzureADUser @userBlockSplat -ErrorAction Continue -WarningAction SilentlyContinue
    }
}#>

# grant send on behalf to
if ($SendOnBehalfTo)
{
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Add-GrantSendOnBehalfToPermissions -SendOnBehalfTo $SendOnBehalfTo
    }
    else
    {
        Add-GrantSendOnBehalfToPermissions -SendOnBehalfTo $SendOnBehalfTo -WhatIf
    }
}

# grant send as rights
if ($SendAs)
{
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Add-GrantSendAsPermissions -SendAs $SendAs
    }
    else
    {
        Add-GrantSendAsPermissions -SendAs $SendAs -WhatIf
    }
}

# grant full access rights
if ($FullAccess)
{
    # wait for mailbox to be available (hopefully)
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Start-Sleep -Seconds 3
    }

    # setting full access rights
    foreach ($recipient in $FullAccess)
    {
        # fullAccess setting
        $fullAccessSplat = @{
            #Identity = $DisplayName
            Identity = $PrimarySmtpAddress
            User = $recipient
            AccessRights = "FullAccess"
            InheritanceType = "All"
        }

        if ($PSCmdlet.ShouldProcess(($fullAccessSplat | ConvertTo-Json), "Add-MailboxPermission"))
        {
            try
            {
                Add-MailboxPermission @fullAccessSplat -WarningAction SilentlyContinue -ErrorAction Continue | Out-Null
                Write-Log -Message "FullAccess permission added to shared mailbox" -Body $fullAccessSplat
            }
            catch
            {
                if (!$Creator) {
                    Write-Log -Message "Failed to add FullAccess permission to shared mailbox: $_" -Body $fullAccessSplat -Level ERROR
                }
                else {
                    Write-Log -Message "Feilet ved FullAccess tilgang på shared mailbox!<br><br>$(ConvertTo-List -Hash ([ordered]@{ Generelt = $mailboxSplat; AdditionalSettings = $mailboxConfigureSplat; FullAccess = $fullAccessSplat }))<br>$($mailMsgError.Replace("%errormsg%", $_))" -Body @{ Subject = "Feil: Exchange skjemaparser"; } -Level ERROR
                }

                Write-Error $_ -ErrorAction Stop
            }
        }
    }

    # also grant FullAccess people SendOnBehalfTo
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Add-GrantSendOnBehalfToPermissions -SendOnBehalfTo $FullAccess
    }
    else
    {
        Add-GrantSendOnBehalfToPermissions -SendOnBehalfTo $FullAccess -WhatIf
    }
    
    # also grant FullAccess people SendAs
    if (!$PSBoundParameters.ContainsKey("WhatIf"))
    {
        Add-GrantSendAsPermissions -SendAs $FullAccess
    }
    else
    {
        Add-GrantSendAsPermissions -SendAs $FullAccess -WhatIf
    }
}

# send success mail if Creator is filled out and everything has worked fine!
if ($Creator) {
    $mailSplat = [ordered]@{
        Generelt = $mailboxSplat
        AdditionalSettings = $mailboxConfigureSplat
    }
    if ($SendOnBehalfTo) { $mailSplat.Add("SendOnBehalfTo", $SendOnBehalfTo) }
    if ($SendAs) { $mailSplat.Add("SendAs", $SendAs) }
    if ($FullAccess) { $mailSplat.Add("FullAccess", $FullAccess) }

    if ($PSCmdlet.ShouldProcess(($mailSplat | ConvertTo-Json), "Write-Log")) {
        Write-Log -Message "Ny shared mailbox opprettet<br><br>$(ConvertTo-List -Hash $mailSplat)" -Body @{ Subject = "Exchange skjemaparser" } -Level SUCCESS
    }
}