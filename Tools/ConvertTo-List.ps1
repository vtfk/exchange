Function ConvertTo-List
{
    param(
        [Parameter(Mandatory = $True)]
        [Hashtable]$Hash,

        [Parameter()]
        [int]$Indent = 0
    )

    $enum = $Hash.GetEnumerator()
    [string]$out = ""

    foreach ($row in $enum) {
        if ($row.Value.GetType().Name -eq "Hashtable") {
            $out += "<span$(if ($Indent -gt 0) { " style='margin-left: $($Indent * 20)px;'" })><b>$($row.Key)</b>:</span><br>`n"
            $out += ConvertTo-List -Hash $row.Value -Indent ($Indent + 1)
        }
        else {
            $out += "<span$(if ($Indent -gt 0) { " style='margin-left: $($Indent * 20)px;'" })><b>$($row.Name)</b>: $($row.Value)</span><br>`n"
        }
    }

    return $out
}