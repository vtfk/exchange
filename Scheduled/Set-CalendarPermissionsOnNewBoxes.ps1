param(
    [Parameter()]
    [DateTime]$FilterDate,

    [Parameter()]
    [string[]]$IdentityList
)

# import environment variables
$envPath = Join-Path -Path $PSScriptRoot -ChildPath "..\envs.ps1"
. $envPath

# set up logger
Import-Module Logger
Add-LogTarget -Name Console
if ($FilterDate) {
    Add-LogTarget -Name CMTrace -Configuration @{ Path = "%1_-$(((Get-Date) - $FilterDate).Days)_days.log" }
}
elseif ($IdentityList) {
    Add-LogTarget -Name CMTrace -Configuration @{ Path = "%1_$($IdentityList.Count)-identities_days.log" }
}
else {
    Add-LogTarget -Name CMTrace -Configuration @{ Path = "%1.log" }
}

# connect to Exchange Online
Connect-Office365 -Exchange -Target $o365AutomationUser

# current index beeing iterated over
[int]$userCount = 1

# make sure ExchangeOnlineManagement module is correctly loaded
Get-Command -Name "Get-EXOMailbox"

# filter mailboxes by createdwhen
if ($FilterDate)
{
    $lastAdded = $FilterDate
}
else
{
    $lastAdded = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays($lastAddedInDays)
}

if (!$IdentityList) {
    Write-Log -Message "Checking for user mailboxes created after '$($lastAdded.ToString('dd.MM.yyyy HH:mm:ss'))'"
    $mailboxes = Get-EXOMailbox -Filter "WhenMailboxCreated -gt '$lastAdded'" -RecipientTypeDetails UserMailbox -ResultSize Unlimited
}
else {
    Write-Log -Message "Getting mailboxes for '$($IdentityList -join "','")'"
    $mailboxes = $IdentityList | Get-EXOMailbox
}

if ($mailboxes)
{
    if (!$mailboxes.Count)
    {
        [int]$mbxCount = 1
    }
    else
    {
        [int]$mbxCount = $mailboxes.Count
    }
}
else
{
    [int]$mbxCount = 0
}

foreach ($mbx in $mailboxes)
{
    [bool]$isStudent = $mbx.PrimarySmtpAddress -like "*@skole.vtfk.no*"
    Write-Log -Message "[$userCount / $mbxCount] :: $(if ($isStudent) { 'STUDENT' } else { 'EMPLOYEE' }) :: '$($mbx.DisplayName)' ($($mbx.Name))"
    $defaultHandled = $false
    $groupHandled = $false
    do {
        try
        {
            $calendarName = Get-EXOMailboxFolderStatistics -Identity $mbx.UserPrincipalName -Folderscope Calendar | Where { $_.FolderType -eq "Calendar" } | Select -ExpandProperty Name
            $folderPermissionName = "$($mbx.UserPrincipalName):\$calendarName"

            Write-Log -Message "Changing 'Default' access to '$defaultAccessRights'"
            Set-MailboxFolderPermission -Identity $folderPermissionName -User Default -AccessRights $defaultAccessRights -WarningAction Stop -ErrorAction Stop
            Write-Log -Message "OK"
            $defaultHandled = $true
        }
        catch
        {
            $outcome = Get-EXOException -Exception $_
            $defaultHandled = $outcome.Handled

            if ($outcome.Severity -eq "Error")
            {
                Write-Log -Message $outcome.Message -Exception $_ -Level ERROR
            }
            elseif ($outcome.Severity -eq "Warning")
            {
                Write-Log -Message $outcome.Message -Exception $_ -Level WARNING
            }
            elseif ($outcome.Severity -eq "Information")
            {
                Write-Log -Message $outcome.Message -Exception $_ -Level INFO
            }

            if ($outcome.Reauthenticate)
            {
                Connect-Office365 -Exchange
            }
        }
    } while (!$defaultHandled);

    do {
        try
        {
            Write-Log -Message "Adding '$groupName' with '$groupAccessRights'"
            $null = Add-MailboxFolderPermission -Identity $folderPermissionName -User $groupName -AccessRights $groupAccessRights -WarningAction Stop -ErrorAction Stop
            Write-Log -Message "OK"
            $groupHandled = $true
        }
        catch
        {
            $outcome = Get-EXOException -Exception $_
            $groupHandled = $outcome.Handled
            
            if ($outcome.Severity -eq "Error")
            {
                Write-Log -Message $outcome.Message -Exception $_ -Level ERROR
            }
            elseif ($outcome.Severity -eq "Warning")
            {
                Write-Log -Message $outcome.Message -Exception $_ -Level WARNING
            }
            elseif ($outcome.Severity -eq "Information")
            {
                Write-Log -Message $outcome.Message -Exception $_ -Level INFO
            }

            if ($outcome.Reauthenticate)
            {
                Connect-Office365 -Exchange
            }
        }
    } while (!$groupHandled);

    $userCount++
}