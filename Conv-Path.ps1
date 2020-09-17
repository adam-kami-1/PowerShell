[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [System.String] $Path
)

function Conv-Path
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.String] $Path
    )

    # Use Convert-Path to convert $Path to
    # absolute path.
    $Path = Convert-Path $Path
    $Drive = Split-Path -Qualifier $Path -ErrorAction SilentlyContinue
    if ($Drive -eq $null)
    {
        # For most providers Convert-Path removes drive part of the path
        $Drive = Split-Path -Qualifier $PWD
        $Path = $Drive+'\'+$Path
    }
    # Get the name of provider
    $Provider = (Get-PSDrive $Drive.Substring(0,$Drive.Length-1)).Provider.Name

    Write-Verbose "Analysing absolute path: $Path"
    # Split abolute path into path slices
    $Slices = $Path.split('\')
    # Drive need to be uperased.
    # This is required for FileSystem.
    # Other providers ignore this.
    $Path = $Slices[0].ToUpper()
    for ( $i = 1; $i -lt $Slices.Count ; $i++ )
    {
        Write-Verbose ("Checking element no ${i}: " + $Slices[$i])
        if ( $Slices[$i] -ne '' )
        {
            Write-Verbose ("Filtering " + $Path + "\ with " + $Slices[$i])
            if ($Provider -eq 'FileSystem')
            {
                $Path += '\' + (Get-ChildItem ($Path+'\') -Filter $Slices[$i] `
                                                          -ErrorAction SilentlyContinue)[0].Name
            }
            else
            {
                $Items = Get-ChildItem ($Path+'\') -ErrorAction SilentlyContinue
                if ($Provider -eq 'Registry')
                {
                    $Items = $Items.PSChildName
                }
                else
                {
                    $Items = $Items.Name
                }
                foreach ($Item in $Items)
                {
                    if ($Item -eq $Slices[$i])
                    {
                        $Path += '\' + $Item
                    }
                }
            }
        }
        else
        {
            Write-Verbose "Last element is empty"
            $Path += '\'
        }
    }
    Write-Verbose "Result of function is: $Path"
    if ($Provider -eq 'FileSystem')
    {
        return $Path
    }
    else
    {
        return Convert-Path $Path
    }
} # Conv-Path #

Conv-Path @PSBoundParameters
