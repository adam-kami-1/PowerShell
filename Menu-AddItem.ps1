<#
.Synopsis
   Clearing of a menu variable.
.DESCRIPTION
   Clearing of a menu variable.
.EXAMPLE
   Menu-Init -MenuName menu
   Menu-AddItem -MenuName menu -ItemValue 1 -ItemName 'One'
   Menu-AddItem menu 2 'Two'
#>
param (
    [CmdletBinding()]

    # Name of the variable used to store menu (without starting $)
    [Parameter(Mandatory=$true,
               Position=0)]
    [String] $MenuName,

    # Item value to be returned when item is selected in menu
    [Parameter(Mandatory=$true,
               Position=1)]
    $ItemValue,

    # Item description to be displayed in menu. If not present then item value converted with ToString().
    [Parameter(Mandatory=$false,
               Position=2)]
    [String] $ItemDescr = $ItemValue.ToString()
)

$Menu = (Get-Variable $MenuName).Value
if (-not $Menu.ContainsKey('Name') -or ($Menu.Name -ne $MenuName))
{
    Write-Error "Variable `$${MenuName} does not contain menu initialized by Menu-Init.ps1" -Category InvalidArgument
    return
}
(Get-Variable ${MenuName}).Value.Item += @{ Value = $ItemValue; Descr = $ItemDescr }
