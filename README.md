# Exchange scripts / modules

This is Vestfold og Telemark fylkeskommunes place for the Exchange scripts / modules we use

## New-ExchangeGroup.ps1

```PowerShell
.\New-ExchangeGroup.ps1 -DisplayName "My New Group" -PrimarySmtpAddress "mynewgroup@vtfk.no" -ManagedBy @("john.doe@vtfk.no", "kari.nordmann@vtfk.no")
```

### Optional parameters

| Parameter | Type | Comment |
| --------- | ---- | ------- |
| Members | String[] | List of email addresses
| ModeratedBy | String[] | List of email addresses
| Notes | String | Description for group
| AcceptMessagesFromInternalAndExternal | Boolean | `True` for accepting from Internal AND External. `False` for accepting only from Internal (Default)
| RoomList | Switch | Set this to indicate that this will be created as a Room List Distribution group
| HiddenFromAddressLists | Boolean | `True` for hiding it. `False` for not hiding it (default)
| EmailAddresses | String[] | List of email addresses to set as aliases for this group
| MailTip | String | Message to display as a tip when sending mail to group
| SendOnBehalfTo | String[] | List of email addresses
| SendAs | String[] | List of email addresses

## New-ResourceMailbox.ps1

```PowerShell
.\New-ResourceMailbox.ps1 -DisplayName "My New Resource" -PrimarySmtpAddress "mynewresource@vtfk.no" -ResourceType Equipment|Room|Teams
```

### Optional parameters

| Parameter | Type | Comment |
| --------- | ---- | ------- |
| RoomMailboxPassword | String | Password for mailbox. Used with ResourceType Teams
| Capacity | Int | Resource capacity
| City | String | Where is this resource located
| Building | String | Which building is this resource located at
| Floor | Int | At what floor in the building is this resource located at
| GeoCoordinates | String | Latitude/longitude in format `*,*;*,*`
| PostalCode | String | Where is this resource located
| State | String | Where is this resource located
| Street | String[] | Where is this resource located
| Tags | String[] | Additional features in the resource
| AudioDeviceName | String | Specified the audio device at the resource
| VideoDeviceName | String | Specified the video device at the resource
| DisplayDeviceName | String | Specified the display device at the resource
| Booking | String[] | List of email addresses to limit who can book the resource
| AutomateProcessing | String | None|AutoUpdate|AutoAccept - Configure the auto attendant for the resource (`Default is AutoAccept`)
| FullAccess | String[] | List of email addresses which will have full access to the resource
| UsageLocation | String | Set usage location for resource (`Default is 'NO'`)
| Licenses | String[] | Add license to Teams resource. Used with ResourceType Teams

## New-SharedMailbox.ps1

```PowerShell
.\New-SharedMailbox.ps1 -DisplayName "My New Shared" -PrimarySmtpAddress "mynewshared@vtfk.no"
```

### Optional parameters

| Parameter | Type | Comment |
| --------- | ---- | ------- |
| FirstName | String | Not used ever for a Shared mailbox...
| LastName | String | Not used ever for a Shared mailbox...
| HiddenFromAddressLists | Boolean | `True` for hiding it. `False` for not hiding it (default)
| Office | String | Where is this shared mailbox used
| EmailAddresses | String[] | List of email addresses to set as aliases for this group
| SendOnBehalfTo | String[] | List of email addresses
| SendAs | String[] | List of email addresses
| FullAccess | String[] | List of email addresses which will have full access to the shared mailbox

[Scheduled scripts here](Scheduled/)
