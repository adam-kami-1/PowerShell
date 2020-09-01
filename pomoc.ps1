<#
.SYNOPSIS
pomoc displays full help on the selected item.
.DESCRIPTION
If -Test switch is not used then pomoc uses Get-Help -Full to retrieve help information.
.EXAMPLE
pomoc -Name Get_Command
#>
[CmdletBinding(DefaultParameterSetName='AllUsersView', HelpUri='https://go.microsoft.com/fwlink/?LinkID=113316')]
param (

    # Gets help about the specified command or concept. Enter the name of a cmdlet, function, provider, script, or workflow, such as `Get-Member`, a conceptual article name, such as `about_Objects`, or an alias, such as `ls`. Wildcard characters are permitted in cmdlet and provider names, but you can't use wildcard characters to find the names of function help and script help articles.
    #
    # To get help for a script that isn't located in a path that's listed in the `$env:Path` environment variable, type the script's path and file name.
    #
    # If you enter the exact name of a help article, `Get-Help` displays the article contents.
    #
    # If you enter a word or word pattern that appears in several help article titles, `Get-Help` displays a list of the matching titles.
    #
    # If you enter a word that doesn't match any help article titles, `Get-Help` displays a list of articles that include that word in their contents.
    #
    # The names of conceptual articles, such as `about_Objects`, must be entered in English, even in non-English versions of PowerShell.
    [Parameter(Position=0, ValueFromPipelineByPropertyName=$true)]
    [String] $Name = '',

    # Gets help that explains how the cmdlet works in the specified provider path. Enter a PowerShell provider path.
    #
    # This parameter gets a customized version of a cmdlet help article that explains how the cmdlet works in the specified PowerShell provider path. This parameter is effective only for help about a provider cmdlet and only when the provider includes a custom version of the provider cmdlet help article in its help file. To use this parameter, install the help file for the module that includes the provider.
    #
    # To see the custom cmdlet help for a provider path, go to the provider path location and enter a `Get-Help` command or, from any path location, use the Path parameter of `Get-Help` to specify the provider path. You can also find custom cmdlet help online in the provider help section of the help articles.
    #
    # For more information about PowerShell providers, see about_Providers (./About/about_Providers.md).
    [string] ${Path},

    # Displays help only for items in the specified category and their aliases. Conceptual articles are in the HelpFile category.
    [ValidateSet('Alias','Cmdlet','Provider','General','FAQ','Glossary','HelpFile','ScriptCommand','Function','Filter','ExternalScript','All','DefaultHelp','Workflow','DscResource','Class','Configuration')]
    [string[]] ${Category},

    # Displays commands with the specified component value, such as Exchange . Enter a component name. Wildcard characters are per mitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [string[]] ${Component},

    # Displays help for items with the specified functionality. Enter the functionality. Wildcard characters are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [string[]] ${Functionality},

    # Displays help customized for the specified user role. Enter a role. Wildcard characters are permitted.
    #
    # Enter the role that the user plays in an organization. Some cmdlets display different text in their help files based on the value of this parameter. This parameter has no effect on help for the core cmdlets.
    [string[]] ${Role},

    # Adds parameter descriptions and examples to the basic help display. This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
    [switch] ${Detailed},

    # Displays the entire help article for a cmdlet. Full includes parameter descriptions and attributes, examples, input and output object types, and additional notes.
    #
    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='AllUsersView')]
    [switch] ${Full},

    # Displays only the name, synopsis, and examples. To display only the examples, type `(Get-Help <cmdlet-name>).Examples`.
    #
    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='Examples', Mandatory=$true)]
    [switch] ${Examples},

    # Displays only the detailed descriptions of the specified parameters. Wildcards are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
    [string] ${Parameter},

    # Displays the online version of a help article in the default browser. This parameter is valid only for cmdlet, function, workflow, and script help articles. You can't use the Online parameter with `Get-Help` in a remote session.
    #
    # For information about supporting this feature in help articles that you write, see about_Comment_Based_Help (./About/about_Comment_Based_Help.md), and Supporting Online Help (/powershell/scripting/developer/module/supporting-online-help), and Writing Help for PowerShell Cmdlets (/powershell/scripting/developer/help/writing-help-for-windows-powershell-cmdlets).
    [Parameter(ParameterSetName='Online', Mandatory=$true)]
    [switch] ${Online},

    # Displays the help topic in a window for easier reading. The window includes a Find search feature and a Settings box that lets you set options for the display, including options to display only selected sections of a help topic.
    #
    # The ShowWindow parameter supports help topics for commands (cmdlets, functions, CIM commands, workflows, scripts) and conceptual About articles. It does not support provider help.
    #
    # This parameter was introduced in PowerShell 3.0.
    [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
    [switch] ${ShowWindow},
    
    # Run my version of help
    [Parameter(Mandatory=$false)]
    [Switch] $Test,

    # If -Test then rescan help files
    [Parameter(Mandatory=$false)]
    [Switch] $Rescan
)

function ToCamelCase ( [String] $Str )
{
    $ChArr = $Str.ToCharArray()
    $Str = ''
    for ($i = 0; $i -lt $ChArr.Length; $i++)
    {
        if (($i -eq 0) -or ('.-_ '.Contains($ChArr[$i-1])))
        {
            $Str += (($ChArr[$i]).ToString()).ToUpper()
        }
        else
        {
            $Str += $ChArr[$i]
        }
    }
    return $Str
} # ToCamelCase #

#####################################################################
# Displaying stuff
#####################################################################

function DisplayTxtHelpFile ($Item)
{
    $Help = Get-Content $Item.File
    Write-Host "Help File"$Item.File"contains"$Help.Count"lines"
    $Paragraph = @()
    $para = ''
    $LineNo = 0
    while ($LineNo -lt $Help.Count)
    {
        $Line = $Help[$LineNo]
        if ($Line -ne '')
        {
            if ($para -eq '')
            {
                $para = $Line
            }
            else
            {
                $para += "`n"+$Line
            }
        }
        elseif ($para -ne '')
        {
            $Paragraph += $para
            $para = ''
        }
        $LineNo++
    }
    Write-Host "It contains"$Paragraph.Count"paragraphs"
    $DisplayName = $Item.Name.Replace('_', ' ')
    $DisplayName = ToCamelCase $DisplayName
    Write-Host "Its name is $DisplayName"
    for ($LineNo = 0; $LineNo -lt $Paragraph.Count; $LineNo++)
    {
        $Line = $Paragraph[$LineNo]
        switch -Regex ($Line.Trim())
        {
            '^SHORT DESCRIPTION$'
                {
                    Write-Host -ForegroundColor Magenta "SYNOPSIS"
                }
            '^LONG DESCRIPTION$'
                {
                    Write-Host -ForegroundColor Magenta "DESCRIPTION"
                }
            "^(TOPIC|$DisplayName)`$"
                {
                    Write-Host -ForegroundColor Magenta "NAME"
                    Write-Host "    $DisplayName"
                }
            '^SEE ALSO$'
                {
                    Write-Host -ForegroundColor Magenta "RELATED LINKS"
                }
            default
                {
                    #Write-Host $Line
                }
        }
    }
} # DisplayTxtHelpFile #

function DisplayXmlHelpFile ($Item)
{
    Write-Host ("Displaying file: "+$Item.File+" Item no: "+$Item.Index)
    Write-Host ("Item name: "+$Item.Name+" Synopsis: "+$Item.Synopsis)
} # DisplayXmlHelpFile #

function DisplayHelpItem ($Item)
{
    switch ($Item.Format)
    {
        "txt"
            {
                DisplayTxtHelpFile $Item
            }
        "xml"
            {
                DisplayXmlHelpFile $Item
            }
    }
} # DisplayHelpItem #


#####################################################################
# File searching stuff
#####################################################################


function CheckTxtHelpFiles ( $ModuleName, $Path )
{
    $Items = @()
    $Files = (Get-ChildItem $Path\*.help.txt).Name
    if ($Files.Count -gt 0)
    {
        foreach ($File in $Files)
        {
            $Items += [pscustomobject]@{Name = $File.Replace('.help.txt','');
                                        ModuleName = $ModuleName;
                                        File = "$Path\$File";
                                        Format = 'txt';
                                        Index = -1;
                                        Category = 'HelpFile';
                                        Synopsis = ''; # SHORT DESCRIPTION
                                        }
        }
    }
    return $Items
} # CheckTxtHelpFiles #

function CheckXMLFile ( $ModuleName, $Path, $File )
{
    $Items = @()
    $XML = [xml](Get-Content "$Path\$File")
    if ($XML.helpItems -ne $null)
    {
        if ($XML.helpItems.command -ne $null)
        {
            if (($XML.helpItems.command).GetType().Name -eq "Object[]")
            {
                for ($Index = 0; $Index -lt ($XML.helpItems.command).Count; $Index++)
                {
                    $command = ($XML.helpItems.command)[$Index]
                    $Items += [pscustomobject]@{Name = $command.details.name;
                                                ModuleName = $ModuleName;
                                                File = "$Path\$File";
                                                Format = 'xml';
                                                Index = $Index;
                                                Category = 'Cmdlet';
                                                Synopsis = $command.details.description.para
                                                }
                }
            }
        }
    }
    return $Items
} # CheckXMLFile #

function CheckXmlHelpFiles ( $ModuleName, $Path )
{
    $Items = @()
    $Files = (Get-ChildItem $Path\*.dll-help.xml).Name
    if ($Files.Count -gt 0)
    {
        foreach ($File in $Files)
        {
            if ($ModuleName -eq '')
            {
                $MN = $File.Remove($File.Length - '.dll-help.xml'.Length)
                switch ($MN)
                {
                    'System.Management.Automation'
                        {
                            $MN = 'Microsoft.PowerShell.Core'
                        }
                    'Microsoft.PowerShell.Consolehost'
                        {
                            $MN = 'Microsoft.PowerShell.Host'
                        }
                    'Microsoft.PowerShell.Commands.Diagnostics'
                        {
                            $MN = 'Microsoft.PowerShell.Diagnostics'
                        }
                    'Microsoft.PowerShell.Commands.Management'
                        {
                            $MN = 'Microsoft.PowerShell.Management'
                        }
                    'Microsoft.PowerShell.Commands.Utility'
                        {
                            $MN = 'Microsoft.PowerShell.Utility'
                        }
                }
                $Items += CheckXMLFile $MN $Path $File
            }
            else
            {
                $Items += CheckXMLFile $ModuleName $Path $File
            }
        }
    }
    return $Items
} # CheckXmlHelpFiles #

function FindHelpFiles ()
{
    $Help_Items = @()

    $Help_Items += CheckTxtHelpFiles '' $PSHOME\$PSUICulture
    $Help_Items += CheckXmlHelpFiles '' $PSHOME\$PSUICulture
    foreach ($Module in ((Get-Childitem $PSHOME\Modules -Directory).Name))
    {
        #Write-Host "Checking module: $Module"
        if (Test-Path $PSHOME\Modules\$Module\$PSUICulture)
        {
            $Help_Items += CheckTxtHelpFiles $Module $PSHOME\Modules\$Module\$PSUICulture
            $Help_Items += CheckXmlHelpFiles $Module $PSHOME\Modules\$Module\$PSUICulture
        }
    }
    return $Help_Items
} # FindHelpFiles #



#####################################################################
# Main body
#####################################################################

# Below is used by function help ()
#Set the outputencoding to Console::OutputEncoding. More.com doesn't work well with Unicode.
#$outputEncoding=[System.Console]::OutputEncoding

#$Test = $true
#$Name = 'about_functions'
#$Name = 'Get-Help'
#$Name = '*'


###########################################################
# Call standard Get-Help if no -Test
if (-not $Test)
{
    Get-Help @PSBoundParameters | C:\Git\usr\bin\less.exe
    return
}


###########################################################
# Find all *.help.txt and  *.dll-help.xml files HelpFiles
if ((Test-Path $env:HOME\.PS-pomoc.xml) -and -not $Rescan)
{
    $Help_Items = Import-Clixml -Path $env:HOME\.PS-pomoc.xml
}
else
{
    $Help_Items = FindHelpFiles
    Export-Clixml -Path $env:HOME\.PS-pomoc.xml -InputObject ($Help_Items | Sort-Object -Property Name)
}


###########################################################
# Display the first matching Help Item (it's a temporary solution)
$found = @()
for ($i = 0; $i -lt $Help_Items.Count; $i++)
{
    if ($Help_Items[$i].Name -like $Name)
    {
        $found += $i
        #DisplayHelpItem $Help_Items[$i]
        #return
    }
}
switch ($found.Count)
{
    0
        {
            Write-Error "Unable to find $Name"
        }
    1
        {
            DisplayHelpItem $Help_Items[$found[0]]
        }
    default
        {
            for ($i = 0; $i -lt $found.Count; $i++)
            {
                #Write-Host $Help_Items[$found[$i]].Name
                [pscustomobject]@{Name = $Help_Items[$found[$i]].Name;
                                  Category = $Help_Items[$found[$i]].Category;
                                  Module = $Help_Items[$found[$i]].ModuleName;
                                  Synopsis = $Help_Items[$found[$i]].Synopsis;
                                  File = $Help_Items[$found[$i]].File;
                                  Index = $Help_Items[$found[$i]].Index
                                 }
            }
        }
}
