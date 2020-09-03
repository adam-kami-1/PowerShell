#!/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe
<#
.Synopsis
   Clearing of a menu variable.
.DESCRIPTION
   Clearing of a menu variable.
.EXAMPLE
   Menu-Clear -MenuName menu
#>
param (
    [CmdletBinding()]

    # Name of the variable used to store menu (without starting $)
    [Parameter(Mandatory=$true,
               Position=0)]
    [String] $MenuName
)

$Menu = (Get-Variable $MenuName).Value
if (-not $Menu.ContainsKey('Name') -or ($Menu.Name -ne $MenuName))
{
    Write-Error "Variable `$${MenuName} does not contain menu initialized by Menu-Init.ps1" -Category InvalidArgument
    return
}
Remove-Variable -Name ${MenuName} -scope Global
