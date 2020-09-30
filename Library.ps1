<#
.FORWARDHELPTARGETNAME Get-Help
.FORWARDHELPCATEGORY Cmdlet 
#>
function Help ()
{
    #[CmdletBinding(DefaultParameterSetName='AllUsersView', HelpUri='https://go.microsoft.com/fwlink/?LinkID=113316')]
    param(
        [Parameter(Position=0, ValueFromPipelineByPropertyName=$true)]
        [string]
        ${Name},

        [string]
        ${Path},

        [ValidateSet('Alias','Cmdlet','Provider','General','FAQ','Glossary','HelpFile','ScriptCommand','Function','Filter',
    'ExternalScript','All','DefaultHelp','Workflow','DscResource','Class','Configuration')]
        [string[]]
        ${Category},

        [string[]]
        ${Component},

        [string[]]
        ${Functionality},

        [string[]]
        ${Role},

        [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
        [switch]
        ${Detailed},

        [Parameter(ParameterSetName='AllUsersView')]
        [switch]
        ${Full},

        [Parameter(ParameterSetName='Examples', Mandatory=$true)]
        [switch]
        ${Examples},

        [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
        [string]
        ${Parameter},

        [Parameter(ParameterSetName='Online', Mandatory=$true)]
        [switch]
        ${Online},

        [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
        [switch]
        ${ShowWindow})

    #Set the outputencoding to Console::OutputEncoding.
    # More.com doesn't work well with Unicode.
    #$outputEncoding=[System.Console]::OutputEncoding
    # The above problem is not a problem when using less.

    if (($PSBoundParameters['Online'] -eq $true) -or
        ($PSBoundParameters['ShowWindow'] -eq $true))
    {
        # When parameters Online or ShowWindow are used,
        # then use Get-Help without any change. Nothing
        # will be displayed here.
        Get-Help @PSBoundParameters
    }
    elseif (($PSBoundParameters['Parameter'] -ne $null) -and
            ($PSBoundParameters['Name'] -match '[]*?[]'))
    {
        # When parameter Parameter is used and
        # Name contains wildcards, then use Get-Help
        # without any change, but use less to view
        # better the list of found items.
        Get-Help @PSBoundParameters | less
    }
    else
    {
        # When we use less it is better to always
        # display full help and browse it with less,
        # then displaying parts of help info.

        # Ignore Detailed parameter
        if ($PSBoundParameters['Detailed'] -ne $null)
        {
            $PSBoundParameters.Remove('Detailed')
        }
        # Ignore Examples parameter
        if ($PSBoundParameters['Examples'] -ne $null)
        {
            $PSBoundParameters.Remove('Examples')
        }
        # Ignore Parameter parameter
        if ($PSBoundParameters['Parameter'] -ne $null)
        {
            $PSBoundParameters.Remove('Parameter')
        }
        <#
        if (($Name.IndexOf('-') -eq -1) -or
            ($Name -like 'about_*'))
        {
            # Force Full parameter
            if ($PSBoundParameters['Full'] -eq $null)
            {
                $PSBoundParameters['Full'] = $True
            }
            Get-Help @PSBoundParameters | less
        }
        else
        {
        #>
            pomoc.ps1 @PSBoundParameters | less
        #}
    }
} # Help #
