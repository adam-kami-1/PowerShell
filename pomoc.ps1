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

    # Displays the online version of a help article in the default browser. This parameter is valid only for cmdlet, function, workflow, and script help articles. You can't use the Online parameter with `Get-Help` in a remote session.
    #
    # For information about supporting this feature in help articles that you write, see about_Comment_Based_Help (./About/about_Comment_Based_Help.md), and Supporting Online Help (/powershell/scripting/developer/module/supporting-online-help), and Writing Help for PowerShell Cmdlets (/powershell/scripting/developer/help/writing-help-for-windows-powershell-cmdlets).
    [Parameter(ParameterSetName='Online', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] ${Online},

    # Displays the help topic in a window for easier reading. The window includes a Find search feature and a Settings box that lets you set options for the display, including options to display only selected sections of a help topic.
    #
    # The ShowWindow parameter supports help topics for commands (cmdlets, functions, CIM commands, workflows, scripts) and conceptual About articles. It does not support provider help.
    #
    # This parameter was introduced in PowerShell 3.0.
    [Parameter(ParameterSetName='ShowWindow', Mandatory=$true)]
    [System.Management.Automation.SwitchParameter] ${ShowWindow},

    # Specifies the number of characters in each line of output. If this parameter is not used, the width is determined by the characteristics of the host. The default for the PowerShell console is 80 characters.
    #
    [Parameter(Mandatory=$false)]
    [System.Int32] $Width,

    # Run my version of help
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter] $Test,

    # If -Test then rescan help files
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.SwitchParameter] $Rescan
)


#!/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe

#    # Adds parameter descriptions and examples to the basic help display. This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
#    [System.Management.Automation.SwitchParameter] ${Detailed},

#    # Displays the entire help article for a cmdlet. Full includes parameter descriptions and attributes, examples, input and output object types, and additional notes.
#    #
#    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='AllUsersView')]
#    [System.Management.Automation.SwitchParameter] ${Full},

#    # Displays only the name, synopsis, and examples. To display only the examples, type `(Get-Help <cmdlet-name>).Examples`.
#    #
#    # This parameter is effective only when the help files are installed on the computer. It has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='Examples', Mandatory=$true)]
#    [System.Management.Automation.SwitchParameter] ${Examples},

#    # Displays only the detailed descriptions of the specified parameters. Wildcards are permitted. This parameter has no effect on displays of conceptual ( About_ ) help.
#    [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
#    [System.String] ${Parameter},



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

$HelpInfo = @{
    Items = @()
    }

$Work = @{
    Colors = @{};
    ItemIndex = @{};
    Functions = @{};
    Output = [System.String[]] @()
    }
#Write-Host ($Work.Colors.F_Magenta+'PROBA'+$Work.Colors.F_Default)


function DisplayParagraph ( [System.Int32] $IndentLevel, [System.String] $Format, [System.String] $Text = '')
{
    $Indent = 4 * $IndentLevel
    $TextWidth = $Width-$Indent
    switch ($Format)
    {
        'empty'
            {
                $Work.Output += @('')
            }
        'para'
            {
                $Para = $Text
                $Lines = @()
                while ($Para.Length -gt $TextWidth)
                {
                    $Line = $Para.Substring(0, $TextWidth)
                    $NextLine = ''
                    while ($Line.Substring($Line.Length-1, 1) -ne ' ')
                    {
                        $NextLine = $NextLine.Insert(0, $Line.Substring($Line.Length-1, 1))
                        $Line = $Line.Substring(0, $Line.Length-1)
                    }
                    $Lines += @($Line.TrimEnd())
                    $Para = $NextLine+$Para.Substring($TextWidth)

                }
                $Lines += @($Para)
                foreach ($Line in $Lines)
                {
                    $Work.Output += @((' ' * $Indent)+$Line)
                }
                $Work.Output += @('')
            }
        'code'
            {
                $Work.Output += @((' ' * $Indent)+$Text)
            }
        'sect'
            {
                #Write-Host ($Work.Colors.F_Magenta+'TEST'+$Work.Colors.F_Default)
                $Work.Output += @((' ' * $Indent)+$Work.Colors.F_Magenta+$Text+$Work.Colors.F_Default)
            }
        'subsect'
            {
                $Work.Output += @((' ' * $Indent)+$Text)
            }
    }
} # DisplayParagraph #


function DisplayXmlHelpFile ( [System.Xml.XmlElement] $command )
# [System.Collections.Hashtable] $Item )
{
    # Write-Host ("Item name: "+$Item.Name)
    # Write-Host ("Synopsis: "+$Item.Synopsis)
    # Write-Host ("Category: "+$Item.Category)
    DisplayParagraph 0 empty

    DisplayParagraph 0 sect "NAME"
    DisplayParagraph 1 para $command.details.name

    DisplayParagraph 0 sect "SYNOPSIS"
    DisplayParagraph 1 para $command.details.description.para

    DisplayParagraph 0 sect "SYNTAX"
    DisplayParagraph 1 para 'Syntax will be described later !!!'

    DisplayParagraph 0 sect "DESCRIPTION"
    for ($i = 0; $i -lt $command.description.para.Count; $i++)
    {
        DisplayParagraph 1 para $command.description.para[$i]
    }

    DisplayParagraph 0 sect "PARAMETERS"
    DisplayParagraph 1 para 'Parameters will be described later !!!'

    DisplayParagraph 0 sect "INPUTS"
    DisplayParagraph 1 para 'Inputs will be described later !!!'

    DisplayParagraph 0 sect "OUTPUTS"
    DisplayParagraph 1 para 'Outputs will be described later !!!'

    DisplayParagraph 0 sect "NOTES"
    DisplayParagraph 1 para 'Notes will be described later !!!'

    DisplayParagraph 0 sect "EXAMPLES"
    DisplayParagraph 1 para 'Parameters will be described later !!!'

    DisplayParagraph 0 sect "RELATED LINKS"
    DisplayParagraph 1 para 'Related links will be described later !!!'

    DisplayParagraph 0 sect "REMARKS"
    DisplayParagraph 1 para 'Remarks will be described later !!!'
} # DisplayXmlHelpFile #

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
                # Write-Host ("Displaying file: "+$Item.File+" Item no: "+$Item.Index)
                $XML = [System.Xml.XmlDocument](Get-Content $Item.File)
                DisplayXmlHelpFile ($XML.helpItems.command)[$Item.Index]
            }
    }
} # DisplayHelpItem #


#####################################################################
# File searching stuff
#####################################################################


function AddItem ( [System.Boolean] $MarkFunc, [System.Collections.Hashtable] $Item )
{
    # $Item = [pscustomobject]$I
    # Write-Host ("Adding Item "+$Item.Name)
    if ($Work.ItemIndex[$Item.Name] -eq $null)
    {
        if ($Work.Functions[$Item.Name] -ne $null)
        {
            $Item.Category = 'Function'
            if ($MarkFunc)
            {
                $Work.Functions[$Item.Name] = $null
            }
        }
        $Work.ItemIndex[$Item.Name] = $HelpInfo.Items.Count
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
                                    Component = '';
                                    Functionality = '';
                                    Role = '';
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
                             Component = '';
                             Functionality = '';
                             Role = '';
                             Synopsis = '...'}
        }
    }
} # CheckXMLFile #


function CheckXmlHelpFiles ( [System.String[]] $LocalFunctions, [System.String] $ModuleName, [System.String] $Path, [System.String] $Pattern )
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
    Get-ChildItem function: | ForEach-Object { $Work.Functions[$_.Name] = 'Function' }
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
            if ($Work.ItemIndex[$_.Definition] -ne $null)
            {
                $Item = $HelpInfo.Items[$Work.ItemIndex[$_.Definition]]
                AddItem $true @{Name = $_.Name;
                                ModuleName = $Item.ModuleName;
                                File = $Item.File;
                                Format = $Item.Format;
                                Index = $Item.Index;
                                Category = 'Alias';
                                Component = '';
                                Functionality = '';
                                Role = '';
                                Synopsis = $_.Definition}
            }
        }
    foreach ($Function in $Work.Functions.keys)
    {
        if ($Work.Functions[$Function] -ne $null)
        {
            AddItem $false @{Name = $Function;
                             ModuleName = '';
                             File = '';
                             Format = '';
                             Index = -1;
                             Category = 'Function';
                             Component = '';
                             Functionality = '';
                             Role = '';
                             Synopsis = '...'}
        }
    }
    $Work.Functions = @{}
} # FindHelpFiles #


#####################################################################
# Main body
#####################################################################

# Below two lines are in function help ()
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

if ($Width -eq 0)
{
    if ( $psISE -ne $null )
    {
        # This is PowerShell ISE
        $Width = 80
    }
    else
    {
        $Width = [System.Console]::BufferWidth
    }
}
if ($psISE -eq $null)
{
    $Esc=[char]0x1B;
    $Work.Colors = @{
        F_Default = "${Esc}[39m";
        F_Black = "${Esc}[30m";
        F_DarkRed = "${Esc}[31m";
        F_DarkGreen = "${Esc}[32m";
        F_DarkYellow = "${Esc}[33m";
        F_DarkBlue = "${Esc}[34m";
        F_DarkMagenta = "${Esc}[35m";
        F_DarkCyan = "${Esc}[36m";
        F_Gray = "${Esc}[37m";
        F_DarkGray = "${Esc}[90m";
        F_Red = "${Esc}[91m";
        F_Green = "${Esc}[92m";
        F_Yellow = "${Esc}[93m";
        F_Blue = "${Esc}[94m";
        F_Magenta = "${Esc}[95m";
        F_Cyan = "${Esc}[96m";
        F_White = "${Esc}[97m"}
}
else
{
    $Work.Colors = @{
        F_Default = ''
        F_Black = ''
        F_DarkRed = ''
        F_DarkGreen = ''
        F_DarkYellow = ''
        F_DarkBlue = ''
        F_DarkMagenta = ''
        F_DarkCyan = ''
        F_Gray = ''
        F_DarkGray = ''
        F_Red = ''
        F_Green = ''
        F_Yellow = ''
        F_Blue = ''
        F_Magenta = ''
        F_Cyan = ''
        F_White = ''}
}
#Write-Host ($Work.Colors.F_Magenta+'PROBA'+$Work.Colors.F_Default)
# Empty parameter $Name means: list all help items
if ($Name -eq '')
{
    $Name = '*'
}
$found = @()
for ($i = 0; $i -lt $HelpInfo.Items.Count; $i++)
{
    if (($HelpInfo.Items[$i].Name -like $Name) -and
        (($Category.Count -eq 0) -or ($Category -contains $HelpInfo.Items[$i].Category)) -and
        (($Component.Count -eq 0) -or ($Component -contains $HelpInfo.Items[$i].Component)) -and
        (($Functionality.Count -eq 0) -or ($Functionality -contains $HelpInfo.Items[$i].Functionality)) -and
        (($Role.Count -eq 0) -or ($Role -contains $HelpInfo.Items[$i].Role)))
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
    1
        {
            DisplayHelpItem $HelpInfo.Items[$found[0]]
            $Work.Output
        }
    default
        {
            for ($i = 0; $i -lt $found.Count; $i++)
            {
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
