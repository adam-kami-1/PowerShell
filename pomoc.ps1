﻿#!/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe
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
    [System.String] $Name = '',

    # Gets help that explains how the cmdlet works in the specified provider path. Enter a PowerShell provider path.
    #
    # This parameter gets a customized version of a cmdlet help article that explains how the cmdlet works in the specified PowerShell provider path. This parameter is effective only for help about a provider cmdlet and only when the provider includes a custom version of the provider cmdlet help article in its help file. To use this parameter, install the help file for the module that includes the provider.
    #
    # To see the custom cmdlet help for a provider path, go to the provider path location and enter a `Get-Help` command or, from any path location, use the Path parameter of `Get-Help` to specify the provider path. You can also find custom cmdlet help online in the provider help section of the help articles.
    #
    # For more information about PowerShell providers, see about_Providers (./About/about_Providers.md).
    [System.String] ${Path},

    # Displays help only for items in the specified category and their aliases. Conceptual articles are in the HelpFile category.
    [ValidateSet('Alias','Cmdlet','Provider','General','FAQ','Glossary','HelpFile','ScriptCommand','Function','Filter','ExternalScript','All','DefaultHelp','Workflow','DscResource','Class','Configuration')]
    [System.String[]] ${Category},

    # Displays commands with the specified component value, such as Exchange . Enter a component name. Wildcard characters are per mitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [System.String[]] ${Component},

    # Displays help for items with the specified functionality. Enter the functionality. Wildcard characters are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [System.String[]] ${Functionality},

    # Displays help customized for the specified user role. Enter a role. Wildcard characters are permitted.
    #
    # Enter the role that the user plays in an organization. Some cmdlets display different text in their help files based on the value of this parameter. This parameter has no effect on help for the core cmdlets.
    [System.String[]] ${Role},

    # Adds parameter descriptions and examples to the basic help display. This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
    [Switch] ${Detailed},

    # Displays the entire help article for a cmdlet. Full includes parameter descriptions and attributes, examples, input and output object types, and additional notes.
    #
    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='AllUsersView')]
    [Switch] ${Full},

    # Displays only the name, synopsis, and examples. To display only the examples, type `(Get-Help <cmdlet-name>).Examples`.
    #
    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='Examples', Mandatory=$true)]
    [Switch] ${Examples},

    # Displays only the detailed descriptions of the specified parameters. Wildcards are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
    [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
    [System.String] ${Parameter},

    # Displays the online version of a help article in the default browser. This parameter is valid only for cmdlet, function, workflow, and script help articles. You can't use the Online parameter with `Get-Help` in a remote session.
    #
    # For information about supporting this feature in help articles that you write, see about_Comment_Based_Help (./About/about_Comment_Based_Help.md), and Supporting Online Help (/powershell/scripting/developer/module/supporting-online-help), and Writing Help for PowerShell Cmdlets (/powershell/scripting/developer/help/writing-help-for-windows-powershell-cmdlets).
    [Parameter(ParameterSetName='Online', Mandatory=$true)]
    [Switch] ${Online},

    # Displays the help topic in a window for easier reading. The window includes a Find search feature and a Settings box that lets you set options for the display, including options to display only selected sections of a help topic.
    #
    # The ShowWindow parameter supports help topics for commands (cmdlets, functions, CIM commands, workflows, scripts) and conceptual About articles. It does not support provider help.
    #
    # This parameter was introduced in PowerShell 3.0.
    [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
    [Switch] ${ShowWindow},
    
    # Run my version of help
    [Parameter(Mandatory=$false)]
    [Switch] $Test,

    # If -Test then rescan help files
    [Parameter(Mandatory=$false)]
    [Switch] $Rescan
)

function ToCamelCase ( [System.String] $Str )
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
} # Max #


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
} # Min #


Function VerCmp ( [System.String] $Version1, [System.String] $Version2 )
{
    $Ver1 = $Version1.Split('.')
    $Ver2 = $Version2.Split('.')
    for ($i = 0; $i -lt (Min $Ver1.Count $Ver2.Count); $i++)
    {
        if ($Ver1[$i] -ne $Ver2[$i])
        {
            return $Ver1[$i] - $Ver2[$i]
        }
    }
    if ($Ver1.Count -eq $Ver2.Count)
    {
        return 0
    }
    return $Ver1.Count - $Ver2.Count
} # VerCmp #


#####################################################################
# Displaying stuff
#####################################################################

function DisplayTxtHelpFile ( [System.Collections.Hashtable] $Item )
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

function DisplayXmlHelpFile ( [System.Collections.Hashtable] $Item )
{
    Write-Host ("Displaying file: "+$Item.File+" Item no: "+$Item.Index)
    Write-Host ("Item name: "+$Item.Name+" Synopsis: "+$Item.Synopsis)
} # DisplayXmlHelpFile #

function DisplayHelpItem ( [System.Collections.Hashtable] $Item )
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


$HelpInfo = @{ Items = @(); ItemIndex = @{}; Functions = @{} }

function AddItem ( [System.Boolean] $MarkFunc, [System.Collections.Hashtable] $Item )
{
    # $Item = [pscustomobject]$I
    # Write-Host ("Adding Item "+$Item.Name)
    if ($HelpInfo.ItemIndex[$Item.Name] -eq $null)
    {
        if ($HelpInfo.Functions[$Item.Name] -ne $null)
        {
            $Item.Category = 'Function'
            if ($MarkFunc)
            {
                $HelpInfo.Functions[$Item.Name] = $null
            }
        }
        $HelpInfo.ItemIndex[$Item.Name] = $HelpInfo.Items.Count
        $HelpInfo.Items += $Item
    }
} # AddItem #


function CheckTxtHelpFiles ( [System.String] $ModuleName, [System.String] $Path )
{
    $Files = (Get-ChildItem $Path\*.help.txt).Name
    if ($Files.Count -gt 0)
    {
        foreach ($File in $Files)
        {
            AddItem $true @{Name = $File.Replace('.help.txt','');
                            ModuleName = $ModuleName;
                            File = "$Path\$File";
                            Format = 'txt';
                            Index = -1;
                            Category = 'HelpFile';
                            Synopsis = ''}
        }
    }
    return
} # CheckTxtHelpFiles #


function CheckXMLFile ( [System.Collections.Hashtable] $LocalFuncs, [System.String] $ModuleName, [System.String] $Path, [System.String] $File )
{
    $XML = [System.Xml.XmlDocument](Get-Content "$Path\$File")
    if ($XML.helpItems -ne $null)
    {
        if ($XML.helpItems.command -ne $null)
        {
            if (($XML.helpItems.command).GetType().Name -eq "Object[]")
            {
                for ($Index = 0; $Index -lt ($XML.helpItems.command).Count; $Index++)
                {
                    $command = ($XML.helpItems.command)[$Index]
                    $Category = 'Cmdlet'
                    if ($LocalFuncs[$command.details.name] -ne $null)
                    {
                        $Category = 'Function'
                        $LocalFuncs[$command.details.name] = $null
                    }
                    AddItem $true @{Name = $command.details.name;
                                    ModuleName = $ModuleName;
                                    File = "$Path\$File";
                                    Format = 'xml';
                                    Index = $Index;
                                    Category = $Category;
                                    Synopsis = $command.details.description.para}
                }
            }
        }
    }
    foreach ($Function in $LocalFuncs.keys)
    {
        if ($LocalFuncs[$Function] -ne $null)
        {
            AddItem $false @{Name = $Function;
                             ModuleName = $ModuleName;
                             File = '';
                             Format = '';
                             Index = -1;
                             Category = 'Function';
                             Synopsis = '...'}
        }
    }
} # CheckXMLFile #


function CheckXmlHelpFiles ( [System.Object[]] $LocalFunctions, [System.String] $ModuleName, [System.String] $Path, [System.String] $Pattern )
{
    $LocalFuncs = @{}
    if ($LocalFunctions -ne $null)
    {
        for ($i = 0; $i -lt $LocalFunctions.Count; $i++)
        {
            $LocalFuncs[$LocalFunctions[$i]] = 'Function'
        }
    }
    $Files = (Get-ChildItem $Path\*$Pattern).Name
    if ($Files.Count -gt 0)
    {
        foreach ($File in $Files)
        {
            if ($ModuleName -eq '')
            {
                $MN = $File.Remove($File.Length - $Pattern.Length)
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
                CheckXMLFile $LocalFuncs $MN $Path $File
            }
            else
            {
                CheckXMLFile $LocalFuncs $ModuleName $Path $File
            }
        }
    }
} # CheckXmlHelpFiles #


function CheckModule ( [System.String] $Path, [System.String] $ModuleName, [System.String] $Version)
{
    if ($Version -eq '')
    {
        $Path = "$Path\$ModuleName"
    }
    else
    {
        $Path = "$Path\$ModuleName\$Version"
    }
    $LocalFuncs = @()
    if (Test-Path "$Path\$ModuleName.psd1")
    {
        $LocalFuncs = (Import-PowerShellDataFile -Path "$Path\$ModuleName.psd1" -ErrorAction SilentlyContinue).FunctionsToExport
    }
    CheckTxtHelpFiles $ModuleName "$Path\$PSUICulture"
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.dll-help.xml'
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.cdxml-help.xml'
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.psd1-help.xml'
    #CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '.psm1-help.xml'
    CheckXmlHelpFiles $LocalFuncs $ModuleName "$Path\$PSUICulture" '-help.xml'
} # CheckModule #


function FindHelpFiles ()
{
    Get-ChildItem function: | ForEach-Object { $HelpInfo.Functions[$_.Name] = 'Function' }
    CheckTxtHelpFiles '' $PSHOME\$PSUICulture
    $LocalFuncs = @()
    CheckXmlHelpFiles $LocalFuncs '' $PSHOME\$PSUICulture '.dll-help.xml'
    CheckXmlHelpFiles $LocalFuncs '' $PSHOME\$PSUICulture '-help.xml'
    foreach ($ModulePath in (($env:PSModulePath).Split(';')))
    {
        foreach ($Module in ((Get-Childitem $ModulePath -Directory).Name))
        {
            #if ($Module -eq 'BranchCache') {
            #    Write-Host "Checking module: $Module"
            #}
            $Version = ''
            foreach ($SubDir in ((Get-Childitem $ModulePath\$Module -Directory).Name))
            {
                if ($SubDir -eq $PSUICulture)
                {
                    CheckModule $ModulePath $Module ''
                }
                elseif ($SubDir -match '^([0-9]\.)+[0-9]+$')
                {
                    if ((VerCmp $Version $SubDir) -le 0)
                    {
                        $Version = $SubDir
                    }
                }
            }
            if ($Version -ne '')
            {
                #Write-Host "Checking module: $Module\$SubDir"
                CheckModule $ModulePath $Module $Version
            }
        }
    }
    Get-ChildItem alias: | 
        ForEach-Object `
        {
            # Alias  $_.Name  ->  $_.Definition
            #
            if ($HelpInfo.ItemIndex[$_.Definition] -ne $null)
            {
                $Item = $HelpInfo.Items[$HelpInfo.ItemIndex[$_.Definition]]
                AddItem $true @{Name = $_.Name;
                                ModuleName = $Item.ModuleName;
                                File = $Item.File;
                                Format = $Item.Format;
                                Index = $Item.Index;
                                Category = 'Alias';
                                Synopsis = $_.Definition}
            }
        }
    foreach ($Function in $HelpInfo.Functions.keys)
    {
        if ($HelpInfo.Functions[$Function] -ne $null)
        {
            AddItem $false @{Name = $Function;
                             ModuleName = '';
                             File = '';
                             Format = '';
                             Index = -1;
                             Category = 'Function';
                             Synopsis = '...'}
        }
    }
    $HelpInfo.Functions = @{}
} # FindHelpFiles #



#####################################################################
# Main body
#####################################################################

# Below is used by function help ()
#Set the outputencoding to Console::OutputEncoding. More.com doesn't work well with Unicode.
#$outputEncoding=[System.Console]::OutputEncoding

#$Test = $true
#$Rescan = $true
#$Name = 'Add-Computer'
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
    $HelpInfo = Import-Clixml -Path $env:HOME\.PS-pomoc.xml
}
else
{
    FindHelpFiles
    Export-Clixml -Path $env:HOME\.PS-pomoc.xml -Encoding UTF8 -InputObject $HelpInfo
}


###########################################################
# Display the first matching Help Item (it's a temporary solution)
if ($HelpInfo.ItemIndex[$Name] -ne $null)
{
    DisplayHelpItem $HelpInfo.Items[$HelpInfo.ItemIndex[$Name]]
    return
}
$found = @()
for ($i = 0; $i -lt $HelpInfo.Items.Count; $i++)
{
    if ($HelpInfo.Items[$i].Name -like $Name)
    {
        $found += $i
    }
}
switch ($found.Count)
{
    0
        {
            Write-Error "Unable to find $Name"
        }
    #1
    #    {
    #        DisplayHelpItem $HelpInfo.Items[$found[0]]
    #    }
    default
        {
            for ($i = 0; $i -lt $found.Count; $i++)
            {
                #Write-Host $HelpInfo.Items[$found[$i]].Name
                [pscustomobject]@{Name = $HelpInfo.Items[$found[$i]].Name;
                                  Category = $HelpInfo.Items[$found[$i]].Category;
                                  Module = $HelpInfo.Items[$found[$i]].ModuleName;
                                  Synopsis = $HelpInfo.Items[$found[$i]].Synopsis;
                                  File = $HelpInfo.Items[$found[$i]].File;
                                  Index = $HelpInfo.Items[$found[$i]].Index
                                 }
            }
        }
}
