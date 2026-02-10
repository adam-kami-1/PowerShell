<#
.SYNOPSIS
Display info about pressed key.
.DESCRIPTION
Shows information about keys presses on keyboard.
.EXAMPLE
Test-Key.ps1
.INPUTS
None
.OUTPUTS
None
.NOTES
To end the script, break it with 'Break' key:
Press 'Break' (Ctrl+[Pause/Break]) and any usual key to enter debug mode.
Then use 'q' to quit script.

Press any other key to display information about this key.'
#>

Write-Output "Press 'Break' (Ctrl+[Pause/Break]) and any usual key to enter debug mode."
Write-Output "Then use 'q' to quit script."
Write-Output 'Press any other key to display information about this key.'
function KeyValue
{
    Param (
        [char] $key
    )
    if ($key -lt ' ')
    {
        return '^' + [char]([byte]$key + [byte]([char]'@'))
    }
    else {
        return $key
    }
}
while ($true) {
    $key = [System.Console]::ReadKey($true)
    # $key | Get-Member
    # $key.Key | Get-Member
    # $key.Modifiers | Get-Member
    # $key | Format-List
    "KeyChar   : {0}({1})" -F (KeyValue $Key.KeyChar), [int]$Key.KeyChar
    "Key       : {0}({1})" -F $Key.Key, $Key.Key.value__
    "Modifiers : {0}({1})" -F $Key.Modifiers, $Key.Modifiers.value__
    '=================================================='
}
<#
$key is result of `[System.Console]::ReadKey($true)`

$key                                [System.ConsoleKeyInfo]
    ToString()                      [string]
    Key                             [System.ConsoleKey]
        ToString()                  [string]
        ToString(string format)     [string]
        ToString(...)               [string]
        value__                     [int]
    KeyChar                         [char]
    Modifiers                       [System.ConsoleModifiers]
        ToString()                  [string]
        ToString(string format)     [string]
        ToString(...)               [string]
        value__                     [int] 1 - Alt; 2 - Shift; 4 - Control; 5 - Alt, Control = Right Alt
#>