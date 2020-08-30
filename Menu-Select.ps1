<#
.Synopsis
   Display a menu and allow user to select one item.
.DESCRIPTION
   Display a menu and allow user to select one item.
.EXAMPLE
   Menu-Init -MenuName menu
   Menu-AddItem -MenuName menu -ItemValue 1 -ItemName 'One'
   Menu-AddItem menu 2 'Two'
#>
param (
    # Name of the variable used to store menu (without starting $)
    [Parameter(Mandatory=$true,
               Position=0)]
    [String] $MenuName
)


function Min ( $Val1, $Val2 )
{
    if ($Val1 -le $Val2)
    {
        return $Val1
    }
    else
    {
        return $Val2
    }
}


function Max ( $Val1, $Val2 )
{
    if ($Val1 -ge $Val2)
    {
        return $Val1
    }
    else
    {
        return $Val2
    }
}


function DisplayBox ( [int] $Left, [int] $Top, [int] $Width, [int] $Height )
{
    [System.Console]::SetCursorPosition($Left, $Top)
    Write-Host -NoNewline ($Menu.Border.UL+$($Menu.Border.H * ($Width-2))+$Menu.Border.UR)
    for ($r = 1; $r -le ($Height-2); $r++)
    {
        [System.Console]::SetCursorPosition($Left, $Top+$r)
        Write-Host -NoNewline $Menu.Border.V
        [System.Console]::SetCursorPosition($Left+$Width-1, $Top+$r)
        Write-Host -NoNewline $Menu.Border.V
    }
    [System.Console]::SetCursorPosition($Left, $Top+$Height-1)
    Write-Host -NoNewline ($Menu.Border.LL+$($Menu.Border.H * ($Width-2))+$Menu.Border.LR)
}


function getItemRow ([int] $itemNo)
{
    return ($Menu.Top + 1 + ($itemNo - $Menu.View.Start))
}


function DisplayItem ([int] $itemNo, [bool] $sel)
{
    [System.Console]::SetCursorPosition($Menu.Left + 1, $(getItemRow $itemNo))
    if ( $sel )
    {
        # Store old colors and invert
        $b = [System.Console]::BackgroundColor
        $f = [System.Console]::ForegroundColor
        [System.Console]::BackgroundColor = $f
        [System.Console]::ForegroundColor = $b
    }
    [System.Console]::Write($Menu.Item[$itemNo].Descr)
    if ( $sel )
    {
        [System.Console]::BackgroundColor = $b
        [System.Console]::ForegroundColor = $f
    }
}

function ShowOverflow ()
{
    if ($Menu.View.Start -ne $Menu.Start)
    {
        [System.Console]::SetCursorPosition($Menu.Left + 1, $(getItemRow ($Menu.View.Start-1)))
        Write-Host -NoNewline ($Menu.Border.H+$($Menu.Border.U * ($Menu.Width-2))+$Menu.Border.H)
    }
    else
    {
        [System.Console]::SetCursorPosition($Menu.Left + 1, $(getItemRow ($Menu.View.Start-1)))
        Write-Host -NoNewline $($Menu.Border.H * ($Menu.Width))
    }
    #Write-Host -NoNewline (" ("+($Menu.Top).GetType()+")")
    if ($Menu.View.End -ne $Menu.End)
    {
        [System.Console]::SetCursorPosition($Menu.Left + 1, $(getItemRow ($Menu.View.End+1)))
        Write-Host -NoNewline ($Menu.Border.H+$($Menu.Border.D * ($Menu.Width-2))+$Menu.Border.H)
    }
    else
    {
        [System.Console]::SetCursorPosition($Menu.Left + 1, $(getItemRow ($Menu.View.End+1)))
        Write-Host -NoNewline $($Menu.Border.H * ($Menu.Width))
    }
}


function DisplayMenu ()
{
    for ($itemNo = $Menu.View.Start; $itemNo -le $Menu.View.End; $itemNo++)
    {
        DisplayItem $itemNo ($itemNo -eq $Menu.Current)
    }
    ShowOverflow
}

function SaveStatus ()
{
    $Menu.Cursor = [System.Console]::CursorVisible
    [System.Console]::CursorVisible = $false
}


function RestoreStatus ()
{
    [System.Console]::CursorVisible = $Menu.Cursor
}


function JumpUp ( [int] $Step )
{
    if ($Menu.Current -eq $Menu.Start)
    {
        break
    }
    DisplayItem $Menu.Current $false
    $Menu.Current -= $Step
    $Menu.Current = Max $Menu.Current $Menu.Start
    if ($Menu.Current -lt $Menu.View.Start)
    {
        if (($Menu.View.Start - $Step) -lt $Menu.Start)
        {
            $Menu.View.Start = $Menu.Start
            $Menu.View.End = $Menu.Start+$Menu.View.Height-1
        }
        else
        {
            $Menu.View.Start -= $Step
            $Menu.View.End -= $Step
        }
        DisplayMenu
    }
    else
    {
        DisplayItem $Menu.Current $true
        ShowOverflow
    }
}


function JumpDown ( [int] $Step )
{
    if ($Menu.Current -eq $Menu.End)
    {
        break
    }
    DisplayItem $Menu.Current $false
    $Menu.Current += $Step
    $Menu.Current = Min $Menu.Current $Menu.End
    if ($Menu.Current -gt $Menu.View.End)
    {
        if (($Menu.View.End + $Step) -gt $Menu.End)
        {
            $Menu.View.Start = $Menu.End-$Menu.View.Height+1
            $Menu.View.End = $Menu.End
        }
        else
        {
            $Menu.View.Start += $Step
            $Menu.View.End += $Step
        }
        DisplayMenu
    }
    else
    {
        DisplayItem $Menu.Current $true
        ShowOverflow
    }
}


$Menu = (Get-Variable $MenuName).Value
if (-not $Menu.ContainsKey('Name') -or ($Menu.Name -ne $MenuName))
{
    Write-Error "Variable `$${MenuName} does not contain menu initialized by Menu-Init.ps1" -Category InvalidArgument
    return
}

# Width of menu items
$Menu.Width = 3
# Find the longest Descr
foreach ( $item in $Menu.Item )
{
    if ( $Menu.Width -lt $item.Descr.Length )
    {
        $Menu.Width = $item.Descr.Length
    }
}
# Make all Descr the sam length
for ( $i = 0; $i -lt $Menu.Item.Count; $i++ )
{
    if ( $Menu.Item[$i].Descr.Length -lt $Menu.Width)
    {
        $Menu.Item[$i].Descr += (' ' * ($Menu.Width - $Menu.Item[$i].Descr.Length))
    }
}

# Number of items in menu
$Menu.Height = $Menu.Item.Count
# First item in menu
$Menu.Start = 0
# Last item in menu
$Menu.End = $Menu.Start+$Menu.Height-1
# Currently selected item
$Menu.Current = $Menu.Start

# Number of shown items
$Menu.View = @{Height = $Menu.Height}
if (($Menu.Top+$Menu.View.Height+2) -gt [System.Console]::WindowHeight)
{
    $Menu.View.Height = [System.Console]::WindowHeight-$Menu.Top-2
}
# First shown item
$Menu.View.Start = $Menu.Start
# Last shown item
$Menu.View.End = $Menu.View.Start+$Menu.View.Height-1


SaveStatus

DisplayBox $Menu.Left $Menu.Top ($Menu.Width+2) ($Menu.View.Height+2)

DisplayMenu

while ($true)
{
    $k = [System.Console]::ReadKey($true)
    switch ($k.Key)
    {
        ([System.ConsoleKey]::Home)
            {
                JumpUp $Menu.Current-$Menu.Start
            }
        ([System.ConsoleKey]::End)
            {
                JumpDown $Menu.End-$Menu.Current
            }
        ([System.ConsoleKey]::UpArrow)
            {
                JumpUp 1
            }
        ([System.ConsoleKey]::DownArrow)
            {
                JumpDown 1
            }
        ([System.ConsoleKey]::PageUp)
            {
                JumpUp $Menu.View.Height
            }
        ([System.ConsoleKey]::PageDown)
            {
                JumpDown $Menu.View.Height
            }
        ([System.ConsoleKey]::Escape)
            {
                RestoreStatus
                return $null
            }
        ([System.ConsoleKey]::Enter)
            {
                RestoreStatus
                return $Menu.Item[$Menu.Current].Value
            }
        default
            {
                [System.Console]::Beep()
            }
    }
}

