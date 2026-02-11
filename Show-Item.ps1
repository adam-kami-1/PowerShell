<#
.SYNOPSIS
Display the contents of input file/object.
.DESCRIPTION
The contents of the file/object is diplayed using 'less.exe' pager
if the length of the output is longer then terminal height.
.EXAMPLE
Displying regular text file:
Show-Item.ps1 -Path .\file.txt
.EXAMPLE
Displying results of a command:
Show-Item.ps1 -InputObject (Get-Content .\file.txt)
.EXAMPLE
Displying results of a command using a pipe:
Get-Content .\file.txt | Show-Item.ps1
.INPUTS
System.String
.INPUTS
System.String[]
.OUTPUTS
None
.NOTES
InputObject parameter should be a string or array of strings.
When piping large InputObject than there are first displayed
few initial lines, the script waits until whole object is
piped before starting pager.
#>
# [OutputType([$null])]
param (
    [Parameter(Mandatory=$true,
               ParameterSetName="Path",
               Position=0)]
    [System.String] $Path = '',

    [Parameter(Mandatory=$true,
               ParameterSetName="InputObject",
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [System.Management.Automation.PSObject] $InputObject = $null
)

Begin
{
    function Show-Object
    {
        Param (
            [System.Management.Automation.PSObject] $InputObject
        )
        if (($InputObject -is [string]) -or ($InputObject.Count -le $Height))
        {
            Write-Output $InputObject
        }
        else
        {
            Write-Output $InputObject | less.exe
        }
    } # function Show-Object

    # Check the terminal type. In smart terminal (not dumb terminal)
    # we have acces to System.Console and can use less.exe.
    $SmartTerminal = $true
    if ($env:TERM -eq 'emacs')
    {
        $SmartTerminal = $false
    }
    $Height = -1
    $Width = 80
    if ($SmartTerminal)
    {
        $Height = [System.Console]::WindowHeight - 1
        $Width = [System.Console]::WindowWidth
    }

    # Check which Parameter Set is used
    $Count = -1
    if ('' -ne $Path)
    {
        # Parameter Set Name is "Path"
        $InputObject = Get-Content -Path $Path
        Show-Object $InputObject
    }
    elseif ($null -ne $InputObject)
    {
        # Parameter Set Name is "InputObject",
        # $InputObject received as parameter
        Show-Object $InputObject
    }
    else
    {
        # Parameter Set Name is "InputObject",
        # $InputObject will be piped in Process block
        $Count = 0
        if ($SmartTerminal)
        {
            $TmpFile = 'D:\Users\Adam\TMP\less.buf'
            Out-File -InputObject '' -FilePath $TmpFile
        }
    }
}

Process
{
    if ($Count -ge 0)
    {
        if ($SmartTerminal)
        {
            Out-File -InputObject $InputObject -FilePath $TmpFile -Append
        }
        if (($Count -lt $Height) -or -not $SmartTerminal)
        {
            # # Display at least only first $Width characters of every line
            # Write-Host ($InputObject -replace ('(.{'+$Width+'})(.*)'),'$1')
            # $Count += 1

            # Display whole line
            Write-Host $InputObject
            $Count += [int]($InputObject.Length / $Width + 0.5)
        }
        else
        {
            Write-Host ("`r" + '-\|/'[[int]($count / 100) % 4]) -NoNewline
            $Count += 1
        }
    }
}
End
{
    if ($Count -ge 0)
    {
        if (($Count -ge $Height) -and $SmartTerminal)
        {
            less.exe $TmpFile
        }
        if ($SmartTerminal)
        {
            Remove-Item $TmpFile
        }
    }
}
