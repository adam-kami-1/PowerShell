<#
.Synopsis
   Creation of a menu variable.
.DESCRIPTION
   Creation of a menu variable.
.EXAMPLE
   Menu-Init -MenuName menu
#>
param (
    [CmdletBinding()]

    # Name of the variable used to store menu (without starting $)
    [Parameter(Mandatory=$true,
               Position=0)]
    [String] $MenuName,

    # Format of the box sorrounding menu.
    [Parameter(Mandatory=$false)]
    [ValidateSet("", "ascii", "box", "dbox")]
    [String] $Border = "box",

    # Top position of the menu box.
    [Parameter(Mandatory=$false)]
    [Int] $Top = 0,

    # Left position of the menu box.
    [Parameter(Mandatory=$false)]
    [Int] $Left = 0
)

$Menu = @{
    Name = ${MenuName};
    Top = $Top;
    Left = $Left;
    Item = @();
    Border = @{
        Width = 1;
        UL = ' ';
        UR = ' ';
        LL = ' ';
        LR = ' ';
        V = ' ';
        H = ' ';
        U = '^';
        D = 'v'}
}

switch ($Border)
{
    "ascii"
    {
        #"++++|-^v"
        $Menu.Border.Width = 1
        $Menu.Border.UL = '+'
        $Menu.Border.UR = '+'
        $Menu.Border.LL = '+'
        $Menu.Border.LR = '+'
        $Menu.Border.V = '|'
        $Menu.Border.H = '-'
        $Menu.Border.U = '^'
        $Menu.Border.D = 'v'
    }
    "box"
    {
        #"┌┐└┘│─▲▼"
        $Menu.Border.Width = 1
        $Menu.Border.UL = '┌'
        $Menu.Border.UR = '┐'
        $Menu.Border.LL = '└'
        $Menu.Border.LR = '┘'
        $Menu.Border.V = '│'
        $Menu.Border.H = '─'
        $Menu.Border.U = '▲'
        $Menu.Border.D = '▼'
    }
    "dbox"
    {
        #"╔╗╚╝║═▲▼"
        $Menu.Border.Width = 1
        $Menu.Border.UL = '╔'
        $Menu.Border.UR = '╗'
        $Menu.Border.LL = '╚'
        $Menu.Border.LR = '╝'
        $Menu.Border.V = '║'
        $Menu.Border.H = '═'
        $Menu.Border.U = '▲'
        $Menu.Border.D = '▼'
    }
}

Set-Variable -Name ${MenuName} -scope Global -Value $Menu
