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
    $Height = [System.Console]::WindowHeight - 1
    $Width = [System.Console]::WindowWidth

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
            Write-Output $InputObject | less.exe -S
        }
    } # function Show-Object

    $Count = -1
    if ('' -ne $Path)
    {
        $InputObject = Get-Content -Path $Path
        Show-Object $InputObject
    }
    elseif ($null -eq $InputObject)
    {
        $Count = 0
        $TmpFile = 'D:\Users\Adam\TMP\less.buf'
        Out-File -InputObject '' -FilePath $TmpFile
    }
    else
    {
        Show-Object $InputObject
    }
}
Process
{
    if ($Count -ge 0)
    {
        Out-File -InputObject $InputObject -FilePath $TmpFile -Append
        if ($Count -lt $Height)
        {
            # Display at least only first $Width characters of every line
            Write-Host ($InputObject -replace ('(.{'+$Width+'})(.*)'),'$1')
        }
        else
        {
            Write-Host ("`r" + '-\|/'[[int]($count / 10) % 4]) -NoNewline
        }
        $Count += 1
    }
}
End
{
    if ($Count -ge 0)
    {
        if ($Count -ge $Height)
        {
            less.exe -SF $TmpFile
        }
        Remove-Item $TmpFile
    }
}
